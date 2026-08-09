"""
RD Station CRM -> SharePoint (aba Vendas / Tabela4)
Clinica Rejuvenesce

Grava apenas negocios FECHADOS com data de fechamento >= CUTOFF.
Uma linha por produto do negocio. Nao toca em nada anterior ao CUTOFF.
Colunas L:O sao formulas da planilha e sao replicadas em R1C1 (nunca sobrescritas por valor).
"""

import os, sys, time, logging, datetime as dt
from concurrent.futures import ThreadPoolExecutor
import requests, msal

# ─── configuracao (tudo via Secrets do GitHub) ────────────────────────────────
RD_TOKEN      = os.environ["RD_TOKEN"]
TENANT_ID     = os.environ["AZURE_TENANT_ID"]
CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
DRIVE_ID      = os.environ["SHAREPOINT_DRIVE_ID"]
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


def page(n):
    for tent in range(4):
        try:
            r = S.get(RD_URL, params={"token": RD_TOKEN, "page": n, "limit": 200}, timeout=60)
            if r.status_code == 429:
                time.sleep(2 ** tent)
                continue
            r.raise_for_status()
            return r.json().get("deals", [])
        except Exception as e:
            if tent == 3:
                raise
            time.sleep(2 ** tent)
    return []


def buscar_deals():
    """Pagina ate o fim. A API do RD ignora filtros de pipeline/data em query."""
    todos, n, vazias = [], 1, 0
    while n <= MAXPAG:
        lote = list(range(n, min(n + THREADS, MAXPAG + 1)))
        with ThreadPoolExecutor(max_workers=THREADS) as ex:
            res = list(ex.map(page, lote))
        for r in res:
            todos.extend(r)
        if any(len(r) == 0 for r in res):
            vazias += 1
            if vazias >= 1:
                break
        n += THREADS
        if n % 60 == 1:
            log.info("  %s paginas, %s negocios", n - 1, len(todos))
    log.info("RD: %s negocios lidos", len(todos))
    return todos


def elegivel(d):
    st = d.get("deal_stage") or {}
    if STAGE not in norm(st.get("name")) and norm(st.get("nickname")) != "F":
        return False
    if PIPES:
        pid = (d.get("deal_pipeline") or {}).get("id") or (d.get("deal_pipeline") or {}).get("_id")
        if pid not in PIPES:
            return False
    fech = parse_dt(d.get("closed_at") or d.get("updated_at"))
    return bool(fech and fech >= CUTOFF)


def linhas_do_deal(d):
    """Uma linha por produto. Colunas A..K; L..O ficam None (sao formulas)."""
    fech = parse_dt(d.get("closed_at") or d.get("updated_at"))
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


def g(method, url, tk, **kw):
    for tent in range(4):
        r = requests.request(method, url, headers={"Authorization": f"Bearer {tk}",
                                                   "Content-Type": "application/json"},
                             timeout=120, **kw)
        if r.status_code in (429, 503, 504):
            time.sleep(int(r.headers.get("Retry-After", 2 ** tent)))
            continue
        if not r.ok:
            raise RuntimeError(f"Graph {r.status_code}: {r.text[:400]}")
        return r.json() if r.text else {}
    raise RuntimeError("Graph: excesso de tentativas")


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
    for ini in range(2, total + 1, CH):
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
        return

    modelo = formulas_modelo(tk)
    for i in range(0, len(inserir_agora), 20):
        bloco = inserir_agora[i:i + 20]
        inserir(tk, bloco, modelo)
        log.info("  gravadas %s/%s", min(i + 20, len(inserir_agora)), len(inserir_agora))

    log.info("OK — %s linhas inseridas", len(inserir_agora))


if __name__ == "__main__":
    main()
