"""
RD Station CRM -> SharePoint (aba Vendas / Tabela4)
Clinica Rejuvenesce

Grava apenas negocios FECHADOS com data de fechamento >= CUTOFF.
Uma linha por produto do negocio. Nao toca em nada anterior ao CUTOFF.
Colunas L:O sao formulas da planilha e sao replicadas em R1C1 (nunca sobrescritas por valor).
"""

import os, sys, time, logging, datetime as dt
import requests, msal

# ─── configuracao (tudo via Secrets do GitHub) ────────────────────────────────
_REQ = ["RD_TOKEN", "AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET", "SHAREPOINT_DRIVE_ID"]
_falta = [k for k in _REQ if not os.environ.get(k, "").strip()]
if _falta:
    print("ERRO: secrets ausentes ou vazios -> " + ", ".join(_falta))
    print("Cadastre em Settings > Secrets and variables > Actions (nome exato, maiusculas).")
    raise SystemExit(1)

RD_TOKEN      = os.environ["RD_TOKEN"].strip()
TENANT_ID     = os.environ["AZURE_TENANT_ID"].strip()
CLIENT_ID     = os.environ["AZURE_CLIENT_ID"].strip()
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"].strip()
DRIVE_ID      = os.environ["SHAREPOINT_DRIVE_ID"].strip()
FILE_ID       = os.environ.get("SHAREPOINT_FILE_ID", "").strip()
FILE_PATH     = os.environ.get("SHAREPOINT_FILE_PATH", "/VENDAS_E_CAMPANHA.xlsx").strip()

SHEET   = os.environ.get("SHEET_NAME", "Vendas")
TABLE   = os.environ.get("TABLE_NAME", "Tabela4")
CUTOFF  = dt.date.fromisoformat(os.environ.get("CUTOFF_DATE", "2026-08-09"))
STAGE   = os.environ.get("STAGE_MATCH", "FECHAMENTO").upper()
PIPES   = [p.strip() for p in os.environ.get("PIPELINE_IDS", "").split(",") if p.strip()]
MAXPAG  = int(os.environ.get("MAX_PAGES", "250"))
THREADS = int(os.environ.get("THREADS", "12"))
DRYRUN  = "--dry-run" in sys.argv
LISTAR  = "--listar" in sys.argv
DIAG    = "--diagnostico" in sys.argv

RD_URL = "https://crm.rdstation.com/api/v1/deals"
GRAPH  = "https://graph.microsoft.com/v1.0"
WB     = None  # definido em resolver_arquivo()

CF = {  # ids dos campos personalizados do RD
    "avaliador":      os.environ.get("CF_AVALIADOR",  "691b0d0ab5e2d0001db1085d"),
    "meio_avaliacao": os.environ.get("CF_MEIO",       "6740c07d840a380026d05b3e"),
    "data_avaliacao": os.environ.get("CF_DATA_AVAL",  "691e0f68034fef0015ca1a3f"),
}

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)-7s %(message)s",
                    datefmt="%H:%M:%S")
log = logging.getLogger("sync")

EPOCH = dt.date(1899, 12, 30)


# ─── helpers ──────────────────────────────────────────────────────────────────
def serial(d):
    """date -> numero de serie do Excel"""
    return (d - EPOCH).days if d else None


def parse_dt(s):
    if not s:
        return None
    try:
        return dt.datetime.fromisoformat(str(s).replace("Z", "+00:00")).date()
    except Exception:
        for f in ("%Y-%m-%d", "%d/%m/%Y"):
            try:
                return dt.datetime.strptime(str(s)[:10], f).date()
            except Exception:
                pass
    return None


def cf_value(deal, field_id):
    for c in deal.get("deal_custom_fields", []):
        cid = c.get("custom_field_id") or (c.get("custom_field") or {}).get("_id")
        if cid == field_id:
            return c.get("value")
    return None


def norm(v):
    return str(v).strip().upper() if v is not None else ""


# ─── RD Station ───────────────────────────────────────────────────────────────
S = requests.Session()
SAFETY_DAYS = int(os.environ.get("SAFETY_DAYS", "540"))
LOOKBACK = int(os.environ.get("LOOKBACK_ROWS", "800"))   # 0 = ler a planilha inteira


ESPERAS = [3, 8, 20, 45, 90]


def rd_get(params, tolerante=False):
    """Chamada ao endpoint de negocios. tolerante=True devolve None em vez de abortar."""
    p = dict(params); p["token"] = RD_TOKEN
    ultimo = ""
    for tent, espera in enumerate(ESPERAS):
        try:
            r = S.get(RD_URL, params=p, timeout=90)
        except Exception as e:
            ultimo = f"conexao: {e}"
            log.warning("  tentativa %s falhou (%s) — aguardando %ss", tent + 1, ultimo, espera)
            time.sleep(espera); continue
        if r.ok:
            time.sleep(0.4)                       # gentileza com o rate limit
            try:
                return r.json()
            except Exception:
                ultimo = f"HTTP 200 nao-JSON: {r.text[:200]}"
                break
        ultimo = f"HTTP {r.status_code}: {r.text[:200]}"
        if r.status_code in (401, 403):
            raise SystemExit(f"ERRO RD: token invalido ou sem permissao.\n{ultimo}")
        if r.status_code == 400:
            log.info("  RD recusou os parametros (400): %s", r.text[:200])
            return None
        ra = r.headers.get("Retry-After")
        pausa = int(ra) if (ra or "").isdigit() else espera
        log.warning("  tentativa %s -> %s | aguardando %ss", tent + 1, ultimo, pausa)
        time.sleep(pausa)
    if tolerante:
        log.warning("  desistindo desta estrategia: %s", ultimo)
        return None
    raise SystemExit(f"ERRO RD: sem resposta valida apos {len(ESPERAS)} tentativas.\n"
                     f"Ultima: {ultimo}")


def diagnostico():
    """Testa combinacoes de parametros e mostra o closed_at real de cada negocio."""
    hoje = dt.date.today()
    amanha = (hoje + dt.timedelta(days=1)).isoformat()
    ini = CUTOFF.isoformat()
    testes = [
        ("A base",              {"limit": 3}),
        ("B win",               {"limit": 3, "win": "true"}),
        ("C closed_period",     {"limit": 3, "closed_at_period": "true",
                                 "start_date": ini, "end_date": amanha}),
        ("D closed+win",        {"limit": 3, "closed_at_period": "true", "win": "true",
                                 "start_date": ini, "end_date": amanha}),
        ("E order closed_at",   {"limit": 3, "order_by": "closed_at", "direction": "desc"}),
        ("F updated_period",    {"limit": 3, "updated_at_period": "true",
                                 "start_date": ini, "end_date": amanha}),
    ]
    for nome, par in testes:
        q = dict(par); q["token"] = RD_TOKEN
        try:
            r = S.get(RD_URL, params=q, timeout=60)
        except Exception as e:
            log.info("%-18s EXCECAO %s", nome, e); continue
        if not r.ok:
            log.info("%-18s HTTP %s | %s", nome, r.status_code, r.text[:160]); time.sleep(1.5); continue
        try:
            j = r.json()
        except Exception:
            log.info("%-18s HTTP 200 nao-JSON", nome); time.sleep(1.5); continue
        ds = j.get("deals", [])
        log.info("%-18s HTTP 200 | total=%s devolvidos=%s", nome, j.get("total"), len(ds))
        for d in ds:
            st = (d.get("deal_stage") or {}).get("name")
            log.info("%-18s   %-28s closed_at=%-28s status=%-9s win=%-5s etapa=%s",
                     "", str(d.get("name"))[:28], str(d.get("closed_at")),
                     d.get("status"), d.get("win"), st)
        time.sleep(1.5)


def _fechou_apos_corte(d):
    f = parse_dt(d.get("closed_at"))
    return bool(f and f >= CUTOFF)


def estrategia_periodo():
    """win=true + closed_at_period. O RD filtra em UTC, entao alargo 1 dia de cada lado
    e reaplico o corte local em horario de Brasilia dentro de elegivel()."""
    base = {"limit": 200, "win": "true", "closed_at_period": "true",
            "start_date": (CUTOFF - dt.timedelta(days=1)).isoformat(),
            "end_date": (dt.date.today() + dt.timedelta(days=2)).isoformat()}
    j = rd_get(dict(base, page=1), tolerante=True)
    if j is None:
        return None
    ds = j.get("deals", [])
    if not ds:
        log.warning("Filtro devolveu 0 ganhos no periodo — conferindo pelo cursor por seguranca")
        return None
    if not any(parse_dt(d.get("closed_at")) for d in ds):
        log.info("closed_at veio vazio — filtro nao confiavel, usando cursor")
        return None
    log.info("Filtro na origem aceito (win + closed_at_period) — total=%s", j.get("total"))
    todos, pag = list(ds), 2
    while len(ds) == 200 and pag <= 50:
        j = rd_get(dict(base, page=pag), tolerante=True)
        if not j:
            break
        ds = j.get("deals", [])
        todos.extend(ds)
        if len(ds) < 200:
            break
        pag += 1
    return todos


def estrategia_cursor():
    """Percorre os negocios ganhos com next_page. Sem teto de 10 mil."""
    log.info("Cursor: percorrendo negocios ganhos (win=true)")
    limite_criacao = CUTOFF - dt.timedelta(days=SAFETY_DAYS)
    todos, params, n, secas = [], {"limit": 200, "win": "true"}, 0, 0
    while n < 400:
        j = rd_get(params)
        if not j:
            log.warning("RD recusou na pagina %s — interrompendo", n + 1)
            break
        lote = j.get("deals", [])
        if not lote:
            break
        todos.extend(lote)
        n += 1
        criacoes = [parse_dt(d.get("created_at")) for d in lote]
        if criacoes and all(c and c < limite_criacao for c in criacoes if c):
            secas += 1
            if secas >= 2:
                log.info("Alcancado %s (corte - %s dias) — parando", limite_criacao, SAFETY_DAYS)
                break
        else:
            secas = 0
        nxt = j.get("next_page")
        if not nxt or not j.get("has_more"):
            break
        params = {"limit": 200, "win": "true", "next_page": nxt}
        if n % 10 == 0:
            log.info("  %s paginas, %s negocios", n, len(todos))
    return todos


def buscar_deals():
    deals = estrategia_periodo()
    if deals is None:
        deals = estrategia_cursor()
    log.info("RD: %s negocios lidos", len(deals))
    return deals


def elegivel(d):
    st = d.get("deal_stage") or {}
    if STAGE not in norm(st.get("name")) and norm(st.get("nickname")) != "F":
        return False
    if PIPES:
        pid = (d.get("deal_pipeline") or {}).get("id") or (d.get("deal_pipeline") or {}).get("_id")
        if pid not in PIPES:
            return False
    fech = parse_dt(d.get("closed_at"))
    return bool(fech and fech >= CUTOFF)


def linhas_do_deal(d):
    """Uma linha por produto. Colunas A..K; L..O ficam None (sao formulas)."""
    fech = parse_dt(d.get("closed_at"))
    cria = parse_dt(d.get("created_at"))
    aval = parse_dt(cf_value(d, CF["data_avaliacao"]))
    nome = (d.get("name") or "").strip()
    if not nome:
        cts = d.get("contacts") or []
        nome = (cts[0].get("name") if cts else "") or ""
    fonte = (d.get("deal_source") or {}).get("name") or ""
    resp = (d.get("user") or {}).get("name") or ""
    meio = cf_value(d, CF["meio_avaliacao"]) or ""
    avaliador = cf_value(d, CF["avaliador"]) or ""
    etapa = (d.get("deal_stage") or {}).get("name") or "FECHAMENTO"

    prods = d.get("deal_products") or []
    if not prods:
        prods = [{"name": "", "total": d.get("amount_unique") or d.get("amount_total") or 0}]

    out = []
    for p in prods:
        try:
            valor = float(p.get("total") or (float(p.get("price") or 0) * float(p.get("amount") or 1)))
        except Exception:
            valor = 0.0
        out.append([
            nome,                       # A Nome
            etapa,                      # B Etapa
            valor,                      # C Valor Único
            serial(cria),               # D Data de criação
            serial(fech),               # E Data de fechamento
            fonte,                      # F Fonte
            resp,                       # G Responsável
            (p.get("name") or "").strip(),  # H Produtos
            meio,                       # I Meio que Avaliação foi realizada:
            avaliador,                  # J Avaliador
            serial(aval),               # K MM/AAAA da Avaliação
            None, None, None, None,     # L..O formulas
        ])
    return out


# ─── Microsoft Graph ──────────────────────────────────────────────────────────
def token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET)
    r = app.acquire_token_for_client(["https://graph.microsoft.com/.default"])
    if "access_token" not in r:
        raise SystemExit("Falha no token Graph: %s" % r.get("error_description"))
    return r["access_token"]


SESSAO = {"id": None}
ESPERAS_G = [5, 12, 25, 45, 70, 100]


def g(method, url, tk, **kw):
    h = {"Authorization": f"Bearer {tk}", "Content-Type": "application/json"}
    if SESSAO["id"] and "/workbook" in url:
        h["workbook-session-id"] = SESSAO["id"]
    ultimo = ""
    for tent, espera in enumerate(ESPERAS_G):
        try:
            r = requests.request(method, url, headers=h, timeout=120, **kw)
        except Exception as e:
            ultimo = f"conexao: {e}"
            log.warning("  Graph tentativa %s: %s — aguardando %ss", tent + 1, ultimo, espera)
            time.sleep(espera); continue
        if r.ok:
            return r.json() if r.text else {}
        ultimo = f"HTTP {r.status_code}: {r.text[:250]}"
        if r.status_code in (429, 503, 504, 509):
            ra = r.headers.get("Retry-After")
            pausa = int(ra) if (ra or "").isdigit() else espera
            log.warning("  Graph tentativa %s: %s | aguardando %ss", tent + 1, ultimo, pausa)
            time.sleep(min(pausa, 120)); continue
        if r.status_code in (409, 423):     # arquivo travado / conflito de escrita
            log.warning("  Graph tentativa %s: %s | arquivo em uso, aguardando %ss",
                        tent + 1, ultimo, espera)
            time.sleep(espera); continue
        raise RuntimeError(f"Graph {ultimo}")
    raise RuntimeError(f"Graph: sem sucesso apos {len(ESPERAS_G)} tentativas. Ultima -> {ultimo}")


def abrir_sessao(tk):
    """Sessao persistente reduz throttling e mantem o recalculo consistente."""
    try:
        j = g("POST", f"{WB}/createSession", tk, json={"persistChanges": True})
        SESSAO["id"] = j.get("id")
        log.info("Sessao do workbook aberta")
    except Exception as e:
        log.warning("Sem sessao persistente (%s) — seguindo sem ela", str(e)[:120])


def fechar_sessao(tk):
    if not SESSAO["id"]:
        return
    try:
        g("POST", f"{WB}/closeSession", tk)
    except Exception:
        pass
    SESSAO["id"] = None


def listar_drive(tk):
    """Imprime os arquivos da biblioteca para descobrir o caminho correto."""
    def anda(url, prefixo, nivel):
        for it in g("GET", url, tk).get("value", []):
            tipo = "DIR " if "folder" in it else "FILE"
            log.info("  %s %s%s   id=%s", tipo, prefixo, it["name"], it["id"])
            if "folder" in it and nivel < 2:
                anda(f"{GRAPH}/drives/{DRIVE_ID}/items/{it['id']}/children",
                     prefixo + it["name"] + "/", nivel + 1)
    log.info("Conteudo da biblioteca:")
    anda(f"{GRAPH}/drives/{DRIVE_ID}/root/children", "/", 0)


def resolver_arquivo(tk):
    """Define WB a partir do FILE_ID ou, na falta dele, do caminho."""
    global WB, FILE_ID
    if not FILE_ID:
        cam = FILE_PATH if FILE_PATH.startswith("/") else "/" + FILE_PATH
        url = f"{GRAPH}/drives/{DRIVE_ID}/root:{requests.utils.quote(cam)}"
        try:
            it = g("GET", url, tk)
        except RuntimeError as e:
            raise SystemExit(
                f"Arquivo nao encontrado em '{cam}'. Rode com --listar para ver os caminhos.\n{e}")
        FILE_ID = it["id"]
        log.info("Arquivo: %s  (id %s)", it.get("name"), FILE_ID)
    WB = f"{GRAPH}/drives/{DRIVE_ID}/items/{FILE_ID}/workbook"


def col_letra(n):
    s = ""
    while n > 0:
        n, m = divmod(n - 1, 26)
        s = chr(65 + m) + s
    return s


def ler_existentes(tk):
    """Le A:K da aba e devolve o conjunto de chaves ja gravadas a partir do CUTOFF."""
    ws = f"{WB}/worksheets('{SHEET}')"
    ur = g("GET", f"{ws}/usedRange(valuesOnly=true)?$select=address,rowCount", tk)
    total = int(ur.get("rowCount") or 0)
    chaves, CH = set(), 2000
    inicio = 2
    if LOOKBACK and total > LOOKBACK:
        inicio = max(2, total - LOOKBACK + 1)
        log.info("Lendo apenas as ultimas %s linhas (a partir da %s)", LOOKBACK, inicio)
    for ini in range(inicio, total + 1, CH):
        fim = min(total, ini + CH - 1)
        rg = g("GET", f"{ws}/range(address='A{ini}:K{fim}')?$select=values", tk)
        for v in rg.get("values", []):
            chaves.add(chave(v))
    chaves.discard(None)
    log.info("Planilha: %s linhas lidas", total - 1)
    return chaves, total


def chave(v):
    """Nome | data fechamento | valor | produto — so para linhas a partir do CUTOFF."""
    try:
        e = v[4]
        if e in (None, ""):
            return None
        d = EPOCH + dt.timedelta(days=int(float(e)))
        if d < CUTOFF:
            return None
        val = round(float(v[2] or 0), 2)
        return f"{norm(v[0])}|{d.isoformat()}|{val}|{norm(v[7])}"
    except Exception:
        return None


def formulas_modelo(tk):
    """Le L2:O2 em R1C1 — independente do numero da linha."""
    rg = g("GET", f"{WB}/worksheets('{SHEET}')/range(address='L2:O2')?$select=formulasR1C1", tk)
    return rg["formulasR1C1"][0]


def inserir(tk, linhas, modelo):
    """Adiciona na Tabela4 e replica as formulas L:O na linha nova."""
    add = g("POST", f"{WB}/tables('{TABLE}')/rows/add", tk,
            json={"index": None, "values": linhas})
    idx = add.get("index")
    if idx is None:
        log.warning("rows/add nao devolveu index — formulas L:O nao replicadas")
        return
    primeira = idx + 2  # linha 1 = cabecalho
    for i in range(len(linhas)):
        r = primeira + i
        g("PATCH", f"{WB}/worksheets('{SHEET}')/range(address='L{r}:O{r}')", tk,
          json={"formulasR1C1": [modelo]})


# ─── main ─────────────────────────────────────────────────────────────────────
def main():
    log.info("Config: drive=%s... file_id=%s path=%s", DRIVE_ID[:12], FILE_ID or "(vazio)", FILE_PATH)
    if DIAG:
        diagnostico()
        return
    if LISTAR:
        listar_drive(token())
        return

    log.info("CUTOFF %s | aba %s | tabela %s | dry-run=%s", CUTOFF, SHEET, TABLE, DRYRUN)

    deals = buscar_deals()
    elegiveis = [d for d in deals if elegivel(d)]
    log.info("Elegiveis (fechados a partir de %s): %s", CUTOFF, len(elegiveis))

    novas = []
    for d in elegiveis:
        novas.extend(linhas_do_deal(d))
    log.info("Linhas candidatas (1 por produto): %s", len(novas))
    if not novas:
        log.info("Nada a fazer.")
        return

    tk = token()
    resolver_arquivo(tk)
    abrir_sessao(tk)
    existentes, _ = ler_existentes(tk)

    inserir_agora, vistas = [], set()
    for l in novas:
        k = chave(l)
        if k is None or k in existentes or k in vistas:
            continue
        vistas.add(k)
        inserir_agora.append(l)

    log.info("Novas apos deduplicacao: %s", len(inserir_agora))
    if not inserir_agora:
        return
    if DRYRUN:
        for l in inserir_agora[:20]:
            log.info("  DRY %s", l[:8])
        fechar_sessao(tk)
        return

    modelo = formulas_modelo(tk)
    for i in range(0, len(inserir_agora), 20):
        bloco = inserir_agora[i:i + 20]
        inserir(tk, bloco, modelo)
        log.info("  gravadas %s/%s", min(i + 20, len(inserir_agora)), len(inserir_agora))

    fechar_sessao(tk)
    log.info("OK — %s linhas inseridas", len(inserir_agora))


if __name__ == "__main__":
    main()
