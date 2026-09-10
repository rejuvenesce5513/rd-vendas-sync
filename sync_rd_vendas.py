"""
RD Station CRM -> SharePoint (aba Vendas / Tabela4)
Clinica Rejuvenesce

Grava apenas negocios FECHADOS com data de fechamento >= CUTOFF.
Uma linha por produto do negocio. Nao toca em nada anterior ao CUTOFF.
Colunas L:O sao formulas da planilha e sao replicadas em R1C1 (nunca sobrescritas por valor).
"""

import os, sys, time, logging, datetime as dt
import re
import unicodedata
import base64
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
STAGE_MATCH = os.environ.get("STAGE_MATCH", "FECHAMENTO,REVISAO DE PENDENCIAS,FINALIZADO")
PIPES   = [p.strip() for p in os.environ.get("PIPELINE_IDS", "").split(",") if p.strip()]
MAXPAG  = int(os.environ.get("MAX_PAGES", "250"))
THREADS = int(os.environ.get("THREADS", "12"))
DRYRUN  = "--dry-run" in sys.argv
LISTAR  = "--listar" in sys.argv
DIAG    = "--diagnostico" in sys.argv
PIPES_D = "--pipelines" in sys.argv
CAMPOS_PV = "--campos-pv" in sys.argv
BRUTO_PV  = "--bruto-pv" in sys.argv
FEEGOW_D  = "--feegow" in sys.argv
FEEGOW_BF = "--feegow-backfill" in sys.argv
FG_PROC   = "--feegow-procedimentos" in sys.argv
FG_AGENDA = "--feegow-agenda" in sys.argv
FG_CIR    = "--cirurgias" in sys.argv
FG_CIRD   = "--cirurgias-diag" in sys.argv
SO_PV   = "--prevendas" in sys.argv
SEM_PV  = "--sem-prevendas" in sys.argv
CAMPOS  = "--campos" in sys.argv

RD_URL = "https://crm.rdstation.com/api/v1/deals"
GRAPH  = "https://graph.microsoft.com/v1.0"
WB     = None  # definido em resolver_arquivo()

CF = {  # ids dos campos personalizados do RD
    "avaliador":      os.environ.get("CF_AVALIADOR",  "691b0d0ab5e2d0001db1085d"),
    "meio_avaliacao": os.environ.get("CF_MEIO",       "6740c07d840a380026d05b3e"),
    "data_avaliacao": os.environ.get("CF_DATA_AVAL",  "691e0f68034fef0015ca1a3f"),
    "mes_avaliacao":  os.environ.get("CF_MES_AVAL",   "69a5a8dc76cd5b00139e5063"),
    "data_cirurgia":  os.environ.get("CF_DATA_CIR",   "691621537df794001c2666ae"),
    "cirurgia_marcada": os.environ.get("CF_CIR_MARC", "6932c2a93aa1e5001318f8c7"),
}

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)-7s %(message)s",
                    datefmt="%H:%M:%S")
log = logging.getLogger("sync")

EPOCH = dt.date(1899, 12, 30)
def _idx(letra):
    n = 0
    for ch in letra.strip().upper():
        n = n * 26 + ord(ch) - 64
    return n - 1


COL_ID_L   = os.environ.get("COL_ID", "N").strip().upper()
COL_CIR_L  = os.environ.get("COL_CIRURGIA", "O").strip().upper()
COL_MARC_L = os.environ.get("COL_CIR_MARCADA", "P").strip().upper()
LER_ATE    = os.environ.get("LER_ATE", "P").strip().upper()
FORMULAS   = [tuple(x.split(":")) for x in
              os.environ.get("FORMULA_RANGES", "L:M,Q:X").split(",") if ":" in x]

I_ID, I_CIR, I_MARC = _idx(COL_ID_L), _idx(COL_CIR_L), _idx(COL_MARC_L)
COL_FEE_L = os.environ.get("COL_FEEGOW", "").strip().upper()      # vazio = nao grava
I_FEE     = (ord(COL_FEE_L) - 65) if COL_FEE_L else -1
CF_FEEGOW = os.environ.get("CF_ID_FEEGOW", "6a6b5afb84ec2f001de5df5a")

LARGURA = max(I_ID, I_CIR, I_MARC, I_FEE) + 1

# ─── Feegow: aba Marcacoes ────────────────────────────────────────────────────
FG_TOKEN  = os.environ.get("FEEGOW_TOKEN", "")
FG_URL    = "https://api.feegow.com/v1/api"
SHEET_MC  = os.environ.get("SHEET_MARCACOES", "Marcações")
FG_PROCS  = [x.strip() for x in os.environ.get("FEEGOW_PROCEDIMENTOS", "1,3,16").split(",") if x.strip()]
FG_DIAS_TRAS = int(os.environ.get("FEEGOW_DIAS_TRAS", "30"))
FG_DIAS_FRENTE = int(os.environ.get("FEEGOW_DIAS_FRENTE", "90"))
FG_ST_ATENDIDO = os.environ.get("FEEGOW_STATUS_ATENDIDO", "Atendido")
FG_MAX_NOMES = int(os.environ.get("FEEGOW_MAX_NOMES", "60"))
FG_PAC_PATHS = [x.strip() for x in os.environ.get(
    "FEEGOW_PACIENTE_PATHS",
    "/patient/search,/patient/informations,/patient/information,/patient/list").split(",") if x.strip()]
_FG_PAC_OK = [None]


def fg_nome_paciente(pid):
    """A doc nao expoe o caminho: tenta os candidatos e memoriza o que responder."""
    caminhos = [_FG_PAC_OK[0]] if _FG_PAC_OK[0] else FG_PAC_PATHS
    for c in caminhos:
        j = fg_get(c, {"paciente_id": pid}, tolerante=True)
        itens = fg_conteudo(j)
        for it in itens:
            if not isinstance(it, dict):
                continue
            nm = it.get("nome") or it.get("nome_completo") or it.get("name") or ""
            if nm:
                if _FG_PAC_OK[0] != c:
                    log.info("  endpoint de paciente: %s", c)
                    _FG_PAC_OK[0] = c
                return str(nm).strip()
    return ""
# colunas da aba Marcacoes (A..O)
MC = {"data": 0, "hora": 1, "paciente": 2, "idPac": 3, "tipo": 4, "prof": 5,
      "status": 6, "agendou": 7, "marcado": 8, "realizado": 9, "evento": 10,
      "idAgd": 11, "vendPre": 12, "vendCom": 13, "abertoCom": 14}
MC_LARG = 15


def fg_get(caminho, params=None, tolerante=False):
    if not FG_TOKEN:
        raise RuntimeError("FEEGOW_TOKEN nao configurado")
    url = FG_URL + caminho
    for tent, espera in enumerate(ESPERAS):
        try:
            r = S.get(url, params=params or {}, headers={"x-access-token": FG_TOKEN}, timeout=90)
        except Exception as e:
            log.warning("  feegow %s: conexao %s — aguardando %ss", caminho, str(e)[:80], espera)
            time.sleep(espera); continue
        if r.ok:
            time.sleep(0.3)
            try:
                j = r.json()
            except Exception:
                log.error("  feegow %s: resposta nao-JSON", caminho)
                return None
            if isinstance(j, dict) and j.get("success") is False:
                log.warning("  feegow %s: success=false — %s", caminho, str(j.get("content"))[:120])
                return None if tolerante else j
            return j
        if r.status_code in (429, 500, 502, 503, 504):
            log.warning("  feegow %s: HTTP %s — aguardando %ss", caminho, r.status_code, espera)
            time.sleep(espera); continue
        if r.status_code == 404 and tolerante:
            return None
        log.error("  feegow %s: HTTP %s — %s", caminho, r.status_code, r.text[:160])
        return None
    return None


def fg_conteudo(j):
    if isinstance(j, dict):
        c = j.get("content")
        return c if isinstance(c, list) else ([c] if isinstance(c, dict) else [])
    return j if isinstance(j, list) else []


def fg_procedimentos():
    """Lista os procedimentos e destaca os que serao importados."""
    j = fg_get("/procedures/list")
    itens = fg_conteudo(j)
    if not itens:
        log.error("Nao consegui listar procedimentos. Confira o FEEGOW_TOKEN.")
        return
    alvo = set(FG_PROCS)
    log.info("%s procedimento(s) cadastrados. Marcados com >>> serao importados:", len(itens))
    for it in itens:
        pid = str(it.get("procedimento_id") or it.get("id") or "")
        nome = it.get("nome") or it.get("procedimento") or ""
        if pid in alvo:
            log.info("  >>> %-6s %s", pid, nome)
    log.info("--- demais (amostra) ---")
    n = 0
    for it in itens:
        pid = str(it.get("procedimento_id") or it.get("id") or "")
        if pid in alvo:
            continue
        log.info("      %-6s %s", pid, (it.get("nome") or it.get("procedimento") or "")[:52])
        n += 1
        if n >= 15:
            log.info("      ... e mais %s", len(itens) - len(alvo) - n)
            break


# ─── planilha de Cirurgias (arquivo separado) ─────────────────────────────────
ARQ_CIR   = os.environ.get("ARQ_CIRURGIAS", "")
SHEET_CIR = os.environ.get("SHEET_CIRURGIAS", "Cirurgias")
CIR_PROCS = [x.strip() for x in os.environ.get("CIRURGIA_PROCEDIMENTOS", "8").split(",") if x.strip()]
CIR_NOME  = os.environ.get("CIRURGIA_NOME_CONTEM", "CIRURGIA")
CIR_STATUS = os.environ.get("CIRURGIA_STATUS",
    "Aguardando,Chamando,Marcado - não confirmado,Aguardando pagamento,"
    "Atendido,Em atendimento,Marcado - confirmado")
# janela movel: o passado nao muda, entao olhamos pouco para tras e 180 dias a frente
CIR_DE   = os.environ.get("CIRURGIA_DE", "")            # vazio = hoje - CIRURGIA_DIAS_TRAS
CIR_ATE  = os.environ.get("CIRURGIA_ATE", "")           # vazio = hoje + CIRURGIA_DIAS_FRENTE
CIR_TRAS   = int(os.environ.get("CIRURGIA_DIAS_TRAS", "30"))
CIR_FRENTE = int(os.environ.get("CIRURGIA_DIAS_FRENTE", "180"))
# colunas: DATA | ID AGENDAMENTO | PROCEDIMENTO | ID FEEGOW | SITUACAO | VALOR
CC = {"data": 0, "idAgd": 1, "proc": 2, "idPac": 3, "situacao": 4, "valor": 5}
CC_LARG = 6


FG_JANELA_DIAS = int(os.environ.get("FEEGOW_JANELA_DIAS", "150"))   # a API recusa >= 6 meses


def fg_janelas(de, ate):
    """Divide o intervalo em pedacos aceitos pela API (< 6 meses)."""
    out, ini = [], de
    while ini <= ate:
        fim = min(ate, ini + dt.timedelta(days=FG_JANELA_DIAS - 1))
        out.append((ini, fim))
        ini = fim + dt.timedelta(days=1)
    return out


def _pagina(params):
    saida, start = [], 0
    while start < 30000:
        j = fg_get("/appoints/search", dict(params, start=start, offset=200))
        lote = fg_conteudo(j)
        if not lote:
            break
        saida.extend(lote)
        if len(lote) < 200:
            break
        start += 200
    return saida


def fg_buscar_cirurgias(alvo, de, ate):
    """Tenta o filtro de procedimento na origem; se vier vazio, varre tudo e filtra aqui."""
    todos, vistos = [], set()
    filtroOK = None
    for a, b in fg_janelas(de, ate):
        base = {"data_start": a.strftime("%d-%m-%Y"), "data_end": b.strftime("%d-%m-%Y"),
                "list_procedures": 1}
        lote = []
        if filtroOK is not False:
            for pid in sorted(alvo):
                lote += _pagina(dict(base, procedimento_id=pid))
            if lote and filtroOK is None:
                filtroOK = True
                log.info("  filtro de procedimento aceito na origem")
        if not lote:
            if filtroOK is None:
                log.info("  filtro de procedimento ignorado pela API — varrendo e filtrando aqui")
            filtroOK = False
            lote = [x for x in _pagina(base)
                    if str(x.get("procedimento_id") or "") in alvo]
        for x in lote:
            k = str(x.get("agendamento_id") or "")
            if k and k not in vistos:
                vistos.add(k); todos.append(x)
        log.info("  %s a %s: %s cirurgia(s) | acumulado %s",
                 a.strftime("%d/%m"), b.strftime("%d/%m"), len(lote), len(todos))
    return todos


def cir_procedimentos_alvo():
    """IDs de procedimento considerados cirurgia: lista fixa ou nome contendo o termo."""
    if CIR_PROCS:
        return set(CIR_PROCS), {}
    mapa = fg_mapa("/procedures/list", "procedimento_id", "nome")
    alvo = {pid for pid, nome in mapa.items() if CIR_NOME and CIR_NOME.upper() in norm(nome)}
    return alvo, mapa


def cir_intervalo():
    """Janela movel por padrao; datas fixas so se informadas em CIRURGIA_DE / _ATE."""
    hoje = dt.date.today()
    def pd(x, padrao):
        if not str(x).strip():
            return padrao
        try:
            d, m, a = str(x).split("-"); return dt.date(int(a), int(m), int(d))
        except Exception:
            return padrao
    de = pd(CIR_DE, hoje - dt.timedelta(days=CIR_TRAS))
    ate = pd(CIR_ATE, hoje + dt.timedelta(days=CIR_FRENTE))
    return de, ate


def cirurgias_diag():
    """Mostra quais procedimentos entram como cirurgia e um agendamento cru."""
    import json as _j
    alvo, mapa = cir_procedimentos_alvo()
    if not mapa:
        mapa = fg_mapa("/procedures/list", "procedimento_id", "nome")
    log.info("Termo de busca no nome: %r | lista fixa: %s", CIR_NOME, CIR_PROCS or "(vazia)")
    log.info("Procedimentos que SERAO importados como cirurgia:")
    for pid in sorted(alvo, key=lambda x: (len(x), x)):
        log.info("   >>> %-6s %s", pid, mapa.get(pid, ""))
    if not alvo:
        log.warning("Nenhum procedimento casou. Lista completa para escolher os IDs:")
        for pid, nome in sorted(mapa.items(), key=lambda x: x[1]):
            log.info("       %-6s %s", pid, nome[:60])
        return
    de, ate = cir_intervalo()
    log.info("Janelas de busca (limite de 6 meses da API):")
    for a, b in fg_janelas(de, ate):
        log.info("   %s a %s", a.strftime("%d-%m-%Y"), b.strftime("%d-%m-%Y"))
    # o que existe de fato na agenda, por procedimento
    a0, b0 = de, min(ate, de + dt.timedelta(days=40))
    todos = _pagina({"data_start": a0.strftime("%d-%m-%Y"), "data_end": b0.strftime("%d-%m-%Y"),
                     "list_procedures": 1})
    log.info("--- agendamentos de %s a %s: %s no total ---",
             a0.strftime("%d/%m"), b0.strftime("%d/%m"), len(todos))
    cont = {}
    for x in todos:
        k = str(x.get("procedimento_id") or "(sem procedimento_id)")
        cont[k] = cont.get(k, 0) + 1
    log.info("distribuicao por procedimento_id:")
    for k, n in sorted(cont.items(), key=lambda x: -x[1])[:20]:
        log.info("   %-8s %-46s %s", k, mapa.get(k, "(fora do cadastro)")[:46], n)
    if todos:
        chaves = sorted({k for x in todos for k in x.keys()})
        log.info("campos do agendamento: %s", ", ".join(chaves))
        val = [k for k in chaves if "valor" in k.lower() or "preco" in k.lower()]
        log.info("campos de valor: %s", val or "NENHUM")
        ex = next((x for x in todos if x.get("procedimentos")), todos[0])
        log.info("exemplo cru: %s", _j.dumps(ex, ensure_ascii=False)[:700])
    itens = fg_buscar_cirurgias(alvo, a0, b0)
    log.info("--- casando com os IDs configurados %s: %s encontrada(s) ---", sorted(alvo), len(itens))
    if not itens:
        log.warning("Nenhuma casou. Escolha o ID certo na distribuicao acima e"
                    " coloque em CIRURGIA_PROCEDIMENTOS.")


# ─── aba Prevendas ────────────────────────────────────────────────────────────
SHEET_PV   = os.environ.get("SHEET_PV", "Prevendas")
TABLE_PV   = os.environ.get("TABLE_PV", "")           # vazio = grava por intervalo
PIPE_PV    = os.environ.get("PIPELINE_PV", "6706cd6fb3284c0025da0e80")  # PRE-VENDAS
PIPE_PV_NM = os.environ.get("PIPELINE_PV_NOME", "")   # ou trecho do nome
# a API nao devolve o funil na listagem: identificamos pelo nome da etapa
PV_DETALHE = int(os.environ.get("PV_DETALHE_MAX", "0"))   # 0 = nao busca detalhe
PV_CONTATO = int(os.environ.get("PV_CONTATO_MAX", "0"))   # 0 = nao busca o ID do contato
ETAPAS_PV  = os.environ.get("ETAPAS_PV",
    "NOVO CONTATO,QUALIFICACAO E INTERESSE,AVALIACAO,REALIZADAS,NO SHOW,DIA 1,DIA 2,DIA 3,DIA 4,DIA 5,DIA 6,DIA 7")
CUTOFF_PV  = dt.date.fromisoformat(os.environ.get("CUTOFF_PV", os.environ.get("CUTOFF_DATE", "2026-08-01")))
CF_PV = {
    "primeiro":  os.environ.get("CF_PV_PRIMEIRO", ""),
    "ultimo":    os.environ.get("CF_PV_ULTIMO", ""),
    "agendou":   os.environ.get("CF_PV_AGENDOU", "6740c0264636ba001da07a13"),
    "meio":      os.environ.get("CF_PV_MEIO", "6740c07d840a380026d05b3e"),
    "realizada": os.environ.get("CF_PV_REALIZADA", "68a4b62611150a0014b02f4b"),
    "avaliador": os.environ.get("CF_PV_AVALIADOR", "691b0d0ab5e2d0001db1085d"),
    "dataAval":  os.environ.get("CF_PV_DATA_AVAL", "691e0f68034fef0015ca1a3f"),
    "feegow":    os.environ.get("CF_PV_FEEGOW", "6a6b5afb84ec2f001de5df5a"),
}
IDX_CMP = list(range(11)) + [I_CIR, I_MARC] + ([I_FEE] if I_FEE >= 0 else [])


# ─── helpers ──────────────────────────────────────────────────────────────────
def serial(d):
    """date -> numero de serie do Excel"""
    return (d - EPOCH).days if d else None


def parse_dt(v):
    if v in (None, ""):
        return None
    txt = str(v).strip()
    try:
        return dt.datetime.fromisoformat(txt.replace("Z", "+00:00")).date()
    except Exception:
        pass
    for f in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(txt[:10], f).date()
        except Exception:
            pass
    m = re.fullmatch(r"(\d{1,2})[/-](\d{4})", txt)          # MM/AAAA -> dia 1
    if m:
        return dt.date(int(m.group(2)), int(m.group(1)), 1)
    m = re.fullmatch(r"(\d{4})[/-](\d{1,2})", txt)          # AAAA-MM -> dia 1
    if m:
        return dt.date(int(m.group(1)), int(m.group(2)), 1)
    return None


def cf_value(deal, field_id):
    for c in deal.get("deal_custom_fields", []):
        cid = c.get("custom_field_id") or (c.get("custom_field") or {}).get("_id")
        if cid == field_id:
            return c.get("value")
    return None


def norm(v):
    """Maiusculas, sem acento, espacos colapsados — comparacao tolerante."""
    if v is None:
        return ""
    t = unicodedata.normalize("NFD", str(v))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", t).strip().upper()


# ─── RD Station ───────────────────────────────────────────────────────────────
S = requests.Session()
SAFETY_DAYS = int(os.environ.get("SAFETY_DAYS", "540"))
LOOKBACK = int(os.environ.get("LOOKBACK_ROWS", "800"))   # 0 = ler a planilha inteira
MAX_UPD  = int(os.environ.get("MAX_UPDATES", "60"))
MARCAR_REV = os.environ.get("MARCAR_REVERTIDAS", "1") not in ("0", "false", "")      # teto por execucao; o resto vai no proximo ciclo


ESPERAS = [3, 8, 20, 45, 90]


def rd_get(params, tolerante=False, caminho=None):
    """Chamada a API do RD. caminho=None usa o endpoint de negocios.
    tolerante=True devolve None em vez de abortar."""
    p = dict(params); p["token"] = RD_TOKEN
    url = (RD_URL.rsplit("/deals", 1)[0] + caminho) if caminho else RD_URL
    ultimo = ""
    for tent, espera in enumerate(ESPERAS):
        try:
            r = S.get(url, params=p, timeout=90)
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


def listar_pipelines():
    """Descobre os funis e as etapas de cada um."""
    vistos = {}
    params = {"limit": 200}
    for _ in range(6):
        j = rd_get(params)
        if not j:
            break
        lote = j.get("deals", [])
        if not lote:
            break
        for d in lote:
            pl = d.get("deal_pipeline") or {}
            pid = pl.get("id") or pl.get("_id") or "?"
            nome = pl.get("name") or "(sem nome)"
            et = (d.get("deal_stage") or {}).get("name") or "?"
            vistos.setdefault((pid, nome), {})
            vistos[(pid, nome)][et] = vistos[(pid, nome)].get(et, 0) + 1
        nxt = j.get("next_page")
        if not nxt or not j.get("has_more"):
            break
        params = {"limit": 200, "next_page": nxt}
    conta = {"pre-vendas": 0, "comercial/outro": 0}
    for (pid, nome), etapas in vistos.items():
        for e, n in etapas.items():
            conta["pre-vendas" if norm(e) in etapas_pv() else "comercial/outro"] += n
    log.info("Classificacao pela etapa: pre-vendas=%s | comercial/outro=%s",
             conta["pre-vendas"], conta["comercial/outro"])
    log.info("Etapas tratadas como pre-vendas: %s", ", ".join(sorted(etapas_pv())))
    log.info("Funis encontrados (a API nao devolve o funil na listagem):")
    for (pid, nome), etapas in sorted(vistos.items(), key=lambda x: -sum(x[1].values())):
        log.info("  %-26s id=%s", nome[:26], pid)
        for e, n in sorted(etapas.items(), key=lambda x: -x[1]):
            log.info("        %-34s %s", e[:34], n)


def campos_prevendas():
    """Mostra os campos personalizados de negocios do funil de pre-vendas."""
    params, achados = {"limit": 200}, 0
    for _ in range(8):
        j = rd_get(params)
        if not j:
            break
        for d in j.get("deals", []):
            if not do_funil_pv(d):
                continue
            log.info("=== %s | etapa %s | estado win=%s", d.get("name"),
                     nome_etapa(d), d.get("win"))
            for c in d.get("deal_custom_fields", []):
                cid = c.get("custom_field_id") or (c.get("custom_field") or {}).get("_id")
                lab = (c.get("custom_field") or {}).get("label") or c.get("label") or ""
                log.info("    %-26s %-34s = %r", cid, lab[:34], c.get("value"))
            log.info("    contatos: %r", [(x.get("id"), x.get("name")) for x in (d.get("contacts") or [])][:2])
            achados += 1
            if achados >= 3:
                return
        nxt = j.get("next_page")
        if not nxt:
            break
        params = {"limit": 200, "next_page": nxt}
    if not achados:
        log.warning("Nenhum negocio nas etapas de pre-vendas. Etapas procuradas: %s", sorted(etapas_pv()))


def bruto_prevendas():
    """Estrutura crua de um negocio de pre-vendas: revela campos nativos."""
    import json as _j
    params = {"limit": 200}
    for _ in range(8):
        j = rd_get(params)
        if not j:
            break
        for d in j.get("deals", []):
            if not do_funil_pv(d):
                continue
            log.info("=== %s | etapa %s", d.get("name"), nome_etapa(d))
            log.info("--- chaves de primeiro nivel ---")
            for k in sorted(d.keys()):
                v = d[k]
                if isinstance(v, (dict, list)):
                    log.info("  %-28s %s", k, _j.dumps(v, ensure_ascii=False)[:220])
                else:
                    log.info("  %-28s %r", k, v)
            det = buscar_deal(str(d.get("id") or d.get("_id") or ""))
            if isinstance(det, dict):
                extras = [k for k in det.keys() if k not in d]
                log.info("--- chaves extras no detalhe do negocio ---")
                for k in sorted(extras):
                    v = det[k]
                    log.info("  %-28s %s", k,
                             (_j.dumps(v, ensure_ascii=False)[:220] if isinstance(v, (dict, list)) else repr(v)))
                ct = (det.get("contacts") or [{}])[0]
                log.info("--- primeiro contato do detalhe ---")
                log.info("  %s", _j.dumps(ct, ensure_ascii=False)[:400])
            return
        nxt = j.get("next_page")
        if not nxt:
            break
        params = {"limit": 200, "next_page": nxt}
    log.warning("Nenhum negocio de pre-vendas encontrado.")


def _chaves_da_linha(v, cut):
    """Mesmas chaves do sync, mas com corte proprio (o backfill olha mais para tras)."""
    try:
        e = v[4]
        if e in (None, ""):
            return None, None
        d = e if isinstance(e, dt.date) else EPOCH + dt.timedelta(days=int(float(e)))
        if d < cut:
            return None, None
        sk = f"{norm(v[0])}|{d.isoformat()}|{norm(v[7])}"
        cs = f"{d.isoformat()}|{round(float(v[2] or 0), 2)}|{norm(v[7])}|{norm(v[6])}"
        return sk, cs
    except Exception:
        return None, None


def backfill_feegow():
    """Preenche SO a coluna do ID Feegow em linhas antigas. Nunca insere,
    nunca sobrescreve valor existente, nunca toca em outra coluna."""
    if I_FEE < 0:
        log.error("Configure COL_FEEGOW antes de rodar o backfill.")
        return
    de = dt.date.fromisoformat(os.environ.get("FEEGOW_DE", "2026-06-01"))
    ate = dt.date.fromisoformat(os.environ.get("FEEGOW_ATE", CUTOFF.isoformat()))
    cf = CF_FEEGOW
    colL = chr(65 + I_FEE)
    log.info("Backfill do ID Feegow: fechamentos de %s a %s (exclusivo) | coluna %s", de, ate, colL)

    # ---- 1. negocios do RD no intervalo ----
    base = {"limit": 200, "win": "true", "closed_at_period": "true",
            "start_date": de.isoformat(), "end_date": ate.isoformat()}
    deals, pag = [], 1
    while pag <= 80:
        j = rd_get(dict(base, page=pag), tolerante=True)
        if not j:
            break
        lote = j.get("deals", [])
        deals.extend(lote)
        if len(lote) < 200:
            break
        pag += 1
    eleg = [d for d in deals if elegivel_periodo(d, de, ate)]
    log.info("RD: %s ganho(s) no intervalo | %s no funil comercial", len(deals), len(eleg))

    porId, porSk, porCs = {}, {}, {}
    semCampo = 0
    for d in eleg:
        v = cf_value(d, cf)
        if v in (None, "", 0):
            semCampo += 1
            continue
        l = linhas_do_deal(d)[0]
        porId[str(d.get("id") or d.get("_id") or "")] = v
        sk, cs = _chaves_da_linha(l, de)
        if sk:
            porSk.setdefault(sk, v)
        if cs:
            porCs.setdefault(cs, v)
    log.info("Com ID Feegow: %s | sem o campo: %s", len(porId), semCampo)

    # ---- 2. planilha ----
    tk = token(); resolver_arquivo(tk); abrir_sessao(tk)
    try:
        ws = f"{WB}/worksheets('{SHEET}')"
        ur = g("GET", f"{ws}/usedRange(valuesOnly=true)?$select=rowCount", tk)
        n = int(ur.get("rowCount") or 0)
        alvo, jaTem, naoAchou = [], 0, 0
        CH = 2000
        for ini in range(2, max(n, 1) + 1, CH):
            fim = min(n, ini + CH - 1)
            rg = g("GET", f"{ws}/range(address='A{ini}:{colL}{fim}')?$select=values", tk)
            for k, v in enumerate(rg.get("values", [])):
                lin = ini + k
                if len(v) <= I_FEE:
                    continue
                if str(v[I_FEE] or "").strip():
                    jaTem += 1
                    continue
                sk, cs = _chaves_da_linha(v, de)
                if sk is None:
                    continue                      # fora do intervalo
                rid = str(v[I_ID] or "").strip() if len(v) > I_ID else ""
                val = porId.get(rid) or porSk.get(sk) or porCs.get(cs)
                if val:
                    alvo.append((lin, val, v[0]))
                else:
                    naoAchou += 1
        log.info("Linhas no intervalo: %s a preencher | %s ja tinham | %s sem correspondencia",
                 len(alvo), jaTem, naoAchou)
        for lin, val, nome in alvo[:8]:
            log.info("   linha %s (%s) -> %s", lin, str(nome)[:34], val)
        if DRYRUN:
            log.info("dry-run: nada gravado.")
            return
        for lin, val, nome in alvo:
            g("PATCH", f"{ws}/range(address='{colL}{lin}')", tk, json={"values": [[val]]})
        log.info("Backfill concluido: %s linha(s) preenchida(s)", len(alvo))
    finally:
        fechar_sessao(tk)


def elegivel_periodo(d, de, ate):
    if not etapa_de_venda(d):
        return False
    f = parse_dt(d.get("closed_at"))
    return bool(f and de <= f < ate)


def cobertura_feegow():
    """Quantas linhas da aba Vendas teriam o ID Feegow preenchido hoje."""
    cf = os.environ.get("CF_ID_FEEGOW", "6a6b5afb84ec2f001de5df5a")
    tk = token(); resolver_arquivo(tk); abrir_sessao(tk)
    try:
        mapa, por_id, por_dvp, total = ler_existentes(tk)
    finally:
        fechar_sessao(tk)
    existentes = por_id
    log.info("Aba %s: %s linha(s) no total | %s com ID RD", SHEET, max(0, total - 1), len(existentes))

    # ---- lado da planilha: quantas linhas ja tem a coluna do ID preenchida ----
    if I_FEE >= 0:
        tk2 = token(); abrir_sessao(tk2)
        try:
            ws = f"{WB}/worksheets('{SHEET}')"
            ur = g("GET", f"{ws}/usedRange(valuesOnly=true)?$select=rowCount", tk2)
            n = int(ur.get("rowCount") or 0)
            colL = chr(65 + I_FEE)
            preench = comIdRD = 0
            CH = 2000
            for ini in range(2, max(n, 1) + 1, CH):
                fim = min(n, ini + CH - 1)
                rg = g("GET", f"{ws}/range(address='{COL_ID_L}{ini}:{colL}{fim}')?$select=values", tk2)
                for v in rg.get("values", []):
                    if v and str(v[0]).strip():
                        comIdRD += 1
                        if len(v) > (I_FEE - I_ID) and str(v[I_FEE - I_ID]).strip():
                            preench += 1
            log.info("Coluna %s da aba %s: %s de %s linha(s) com ID RD ja tem ID Feegow (%.0f%%)",
                     colL, SHEET, preench, comIdRD, (preench / comIdRD * 100) if comIdRD else 0)
        finally:
            fechar_sessao(tk2)
    else:
        log.info("COL_FEEGOW nao configurado — medindo apenas o lado do RD.")

    deals = buscar_deals()
    eleg = [d for d in deals if elegivel(d)]
    log.info("RD: %s ganho(s) no periodo | %s no funil comercial elegivel", len(deals), len(eleg))

    porMes = {}
    naPlanilha = comId = semId = 0
    exemplos = []
    for d in eleg:
        rid = str(d.get("id") or d.get("_id") or "")
        v = cf_value(d, cf)
        fech = parse_dt(d.get("closed_at"))
        ym = fech.strftime("%Y-%m") if fech else "sem data"
        a = porMes.setdefault(ym, [0, 0])
        a[0] += 1
        if v not in (None, "", 0):
            a[1] += 1
            comId += 1
            if len(exemplos) < 3:
                exemplos.append((d.get("name"), v))
        else:
            semId += 1
        if rid in existentes:
            naPlanilha += 1

    log.info("")
    log.info("%-12s %8s %8s %8s", "mes", "negocios", "com ID", "%")
    for ym in sorted(porMes):
        n, c = porMes[ym]
        log.info("%-12s %8s %8s %7.0f%%", ym, n, c, (c / n * 100) if n else 0)
    tot = comId + semId
    log.info("")
    log.info("FUNIL COMERCIAL: %s negocio(s) | %s com ID Feegow (%.0f%%) | %s sem",
             tot, comId, (comId / tot * 100) if tot else 0, semId)
    log.info("Desses, %s ja estao na aba %s pelo ID RD", naPlanilha, SHEET)
    for nome, v in exemplos:
        log.info("   exemplo: %-40s ID Feegow = %r", str(nome)[:40], v)
    if not comId:
        log.warning("Nenhum negocio com o campo preenchido. Confira o id do campo em CF_ID_FEEGOW "
                    "rodando o modo 'campos'.")


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
    for rot, par in [("etapas com win", {"limit": 200, "win": "true"}),
                     ("etapas sem win", {"limit": 200})]:
        q = dict(par); q["token"] = RD_TOKEN
        try:
            j = S.get(RD_URL, params=q, timeout=60).json()
        except Exception as e:
            log.info("%-18s excecao %s", rot, e); continue
        c = {}
        for d in j.get("deals", []):
            k = (d.get("deal_stage") or {}).get("name") or "(sem etapa)"
            c[k] = c.get(k, 0) + 1
        log.info("%-18s %s", rot, ", ".join(f"{k}={v}" for k, v in sorted(c.items(), key=lambda x: -x[1])))
        time.sleep(1.5)
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


def buscar_deal(did):
    """Consulta um negocio pelo id. Devolve dict, "404" ou None se nao conseguiu."""
    for tent, espera in enumerate(ESPERAS):
        try:
            r = S.get(f"{RD_URL}/{did}", params={"token": RD_TOKEN}, timeout=60)
        except Exception:
            time.sleep(espera); continue
        if r.status_code == 404:
            return "404"
        if r.ok:
            try:
                return r.json()
            except Exception:
                return None
        if r.status_code in (429, 500, 502, 503, 504):
            time.sleep(espera); continue
        return None
    return None


def buscar_deals():
    deals = estrategia_periodo()
    if deals is None:
        deals = estrategia_cursor()
    log.info("RD: %s negocios lidos", len(deals))
    return deals


def nome_etapa(d):
    return (d.get("deal_stage") or {}).get("name") or ""


_STAGES = None


def etapas_venda():
    global _STAGES
    if _STAGES is None:
        _STAGES = [norm(x) for x in STAGE_MATCH.split(",") if x.strip()]
    return _STAGES


def etapa_de_venda(d):
    n = norm(nome_etapa(d))
    return any(t in n for t in etapas_venda())


def elegivel(d):
    if not etapa_de_venda(d):
        return False
    if PIPES:
        pid = (d.get("deal_pipeline") or {}).get("id") or (d.get("deal_pipeline") or {}).get("_id")
        if pid not in PIPES:
            return False
    fech = parse_dt(d.get("closed_at"))
    return bool(fech and fech >= CUTOFF)


def linhas_do_deal(d):
    """UMA linha por negocio. Produtos concatenados por ', ' — mesmo padrao das
    5.806 linhas existentes, de onde a formula da coluna L deriva TC/TEC/BH/etc."""
    fech = parse_dt(d.get("closed_at"))
    cria = parse_dt(d.get("created_at"))
    aval = parse_dt(cf_value(d, CF["mes_avaliacao"])) or parse_dt(cf_value(d, CF["data_avaliacao"]))
    if aval:
        aval = aval.replace(day=1)
    nome = (d.get("name") or "").strip()
    if not nome:
        cts = d.get("contacts") or []
        nome = (cts[0].get("name") if cts else "") or ""

    prods = d.get("deal_products") or []
    nomes, soma = [], 0.0
    for p in prods:
        n = (p.get("name") or "").strip()
        if n:
            nomes.append(n)
        try:
            soma += float(p.get("total") or
                          (float(p.get("price") or 0) * float(p.get("amount") or 1)))
        except Exception:
            pass

    valor = soma
    if not valor:
        for campo in ("amount_total", "amount_unique"):
            try:
                v = float(d.get(campo) or 0)
            except Exception:
                v = 0.0
            if v:
                valor = v
                break

    base = [
        nome,                                                   # A Nome
        (d.get("deal_stage") or {}).get("name") or "FECHAMENTO",  # B Etapa
        valor,                                                  # C Valor Único
        serial(cria),                                           # D Data de criação
        serial(fech),                                           # E Data de fechamento
        (d.get("deal_source") or {}).get("name") or "",         # F Fonte
        (d.get("user") or {}).get("name") or "",                # G Responsável
        ", ".join(nomes),                                       # H Produtos
        cf_value(d, CF["meio_avaliacao"]) or "",                # I Meio da avaliação
        cf_value(d, CF["avaliador"]) or "",                     # J Avaliador
        serial(aval),                                           # K Data da avaliação
    ]
    linha = base + [None] * (LARGURA - len(base))
    linha[I_ID]   = str(d.get("id") or d.get("_id") or "")
    linha[I_CIR]  = serial(parse_dt(cf_value(d, CF["data_cirurgia"])))
    linha[I_MARC] = cf_value(d, CF["cirurgia_marcada"]) or ""
    if I_FEE >= 0:
        linha[I_FEE] = cf_value(d, CF_FEEGOW) or ""
    return [linha]


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
        if r.status_code == 400 and "InvalidSession" in r.text:
            log.warning("  Graph tentativa %s: sessao expirou — reabrindo", tent + 1)
            SESSAO["id"] = None
            abrir_sessao(tk)
            if SESSAO["id"]:
                h["workbook-session-id"] = SESSAO["id"]
            else:
                h.pop("workbook-session-id", None)
            time.sleep(2); continue
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
    mapa, por_id, por_dvp, CH = {}, {}, {}, 2000
    inicio = 2
    if LOOKBACK and total > LOOKBACK:
        inicio = max(2, total - LOOKBACK + 1)
        log.info("Lendo apenas as ultimas %s linhas (a partir da %s)", LOOKBACK, inicio)
    for ini in range(inicio, total + 1, CH):
        fim = min(total, ini + CH - 1)
        rg = g("GET", f"{ws}/range(address='A{ini}:{LER_ATE}{fim}')?$select=values", tk)
        for j, v in enumerate(rg.get("values", [])):
            linha = ini + j
            rid = str(v[I_ID]).strip() if len(v) > I_ID and v[I_ID] not in (None, "") else ""
            sk = softkey(v)
            if rid:
                por_id[rid] = (linha, v)
            if sk:
                mapa[sk] = (linha, v)
            k3 = chave_sem_nome(v)
            if k3:
                por_dvp.setdefault(k3, []).append((linha, v))
    log.info("Planilha: %s linhas lidas | %s com ID RD", total - 1, len(por_id))
    return mapa, por_id, por_dvp, total


def softkey(v):
    """Nome | data de fechamento | produtos — sem o valor, para permitir correcao.
    So considera linhas a partir do CUTOFF."""
    try:
        e = v[4]
        if e in (None, ""):
            return None
        d = e if isinstance(e, dt.date) else EPOCH + dt.timedelta(days=int(float(e)))
        if d < CUTOFF:
            return None
        return f"{norm(v[0])}|{d.isoformat()}|{norm(v[7])}"
    except Exception:
        return None


def chave_sem_nome(v):
    """Fechamento | valor | produtos | responsavel — casa negocios renomeados."""
    try:
        e = v[4]
        if e in (None, ""):
            return None
        d = e if isinstance(e, dt.date) else EPOCH + dt.timedelta(days=int(float(e)))
        if d < CUTOFF:
            return None
        return f"{d.isoformat()}|{round(float(v[2] or 0), 2)}|{norm(v[7])}|{norm(v[6])}"
    except Exception:
        return None


def valor_de(v):
    try:
        return round(float(v[2] or 0), 2)
    except Exception:
        return 0.0


def formulas_modelo(tk):
    """Le cada bloco de colunas calculadas em R1C1 — independente da linha."""
    mod = []
    for a, b in FORMULAS:
        rg = g("GET", f"{WB}/worksheets('{SHEET}')/range(address='{a}2:{b}2')?$select=formulasR1C1", tk)
        mod.append((a, b, rg["formulasR1C1"][0]))
    return mod


def fim_da_tabela(tk):
    """Ultima linha da Tabela4 na planilha."""
    rg = g("GET", f"{WB}/tables('{TABLE}')/range?$select=address,rowCount", tk)
    addr = str(rg["address"]).split("!")[-1]
    m = re.search(r"[A-Z]+(\d+):[A-Z]+(\d+)", addr)
    return int(m.group(2))


def inserir(tk, linhas, modelo):
    """Escreve direto no intervalo logo abaixo da tabela.

    Nao usa tables/rows/add: com 5.800+ linhas, formula volatil (TODAY) e tabelas
    dinamicas, aquele endpoint estoura o gateway do Graph em 504. Escrita por
    endereco e idempotente — repetir a chamada grava o mesmo valor, sem duplicar.
    """
    ultima = fim_da_tabela(tk)
    ini, fim = ultima + 1, ultima + len(linhas)
    ws = f"{WB}/worksheets('{SHEET}')"

    valores = [l[:11] for l in linhas]                      # A..K
    g("PATCH", f"{ws}/range(address='A{ini}:K{fim}')", tk, json={"values": valores})

    ids = [[l[I_ID]] for l in linhas]
    if any(x[0] for x in ids):
        g("PATCH", f"{ws}/range(address='{COL_ID_L}{ini}:{COL_ID_L}{fim}')", tk, json={"values": ids})

    cir = [[l[I_CIR], l[I_MARC]] for l in linhas]
    if any(x[0] not in (None, "") or x[1] not in (None, "") for x in cir):
        g("PATCH", f"{ws}/range(address='{COL_CIR_L}{ini}:{COL_MARC_L}{fim}')", tk,
          json={"values": cir})

    if I_FEE >= 0:
        fees = [[l[I_FEE]] for l in linhas]
        if any(str(x[0] or "").strip() for x in fees):
            g("PATCH", f"{ws}/range(address='{COL_FEE_L}{ini}:{COL_FEE_L}{fim}')", tk,
              json={"values": fees})

    for a, b, mod in (modelo or []):
        g("PATCH", f"{ws}/range(address='{a}{ini}:{b}{fim}')", tk,
          json={"formulasR1C1": [mod for _ in linhas]})

    nova_ultima = fim_da_tabela(tk)
    if nova_ultima >= fim:
        log.info("  linhas %s-%s gravadas e absorvidas pela %s", ini, fim, TABLE)
    else:
        log.warning("  linhas %s-%s gravadas, mas a %s ainda termina na %s — "
                    "expanda a tabela manualmente (as tabelas dinamicas nao verao os dados)",
                    ini, fim, TABLE, nova_ultima)


# ─── sincronizacao da aba Prevendas ───────────────────────────────────────────
def _dh(iso):
    """devolve (serial da data, fracao do dia) para gravar em colunas separadas"""
    if not iso:
        return "", ""
    t = str(iso).replace("Z", "+00:00")
    try:
        x = dt.datetime.fromisoformat(t)
    except Exception:
        d = parse_dt(iso)
        return (serial(d), "") if d else ("", "")
    # a planilha guarda hora:minuto, sem segundos — incluir segundos gera diferenca em toda linha
    return serial(x.date()), (x.hour * 60 + x.minute) / 1440.0


def estado_pt(d):
    if d.get("win") is True:
        return "Vendida"
    if d.get("win") is False:
        return "Perdida"
    return "Em Andamento"


def primeiro_contato(d):
    """A API nao tem esse campo. Usa o campo personalizado, se houver, ou a saida
    da primeira etapa registrada em deal_stage_histories (vem so no detalhe)."""
    if CF_PV["primeiro"]:
        v = cf_value(d, CF_PV["primeiro"])
        if v:
            return v
    hs = d.get("deal_stage_histories") or []
    fim = [h.get("end_date") for h in hs if h.get("end_date")]
    return min(fim) if fim else None


PV_PREFIXO = os.environ.get("PV_PREFIXO", r"^\s*Pr[eé]-?\s*vendas\s*-\s*[^-]*-\s*")


def nome_prevenda(d):
    """A negociacao as vezes tem nome curto ("Rafael"); o contato as vezes e outra
    pessoa (conjuge, indicante). Compara por palavras: mesma pessoa -> nome mais
    completo; pessoas diferentes -> vale o da negociacao."""
    bruto = (d.get("name") or "").strip()
    try:
        limpo = re.sub(PV_PREFIXO, "", bruto, flags=re.I).strip() or bruto
    except Exception:
        limpo = bruto
    cts = d.get("contacts") or []
    ct = ((cts[0].get("name") if cts else "") or "").strip()
    if not ct:
        return limpo
    if not limpo:
        return ct
    ta, tb = set(norm(limpo).split()), set(norm(ct).split())
    menor, maior = (ta, tb) if len(ta) <= len(tb) else (tb, ta)
    if menor and menor <= maior:                      # mesma pessoa
        return limpo if len(limpo) >= len(ct) else ct
    if len(limpo) <= 3:
        # so aceita o contato se o apelido for inicio de alguma palavra dele
        # ("Ba" -> Basilios) ou as iniciais ("DG" -> Douglas Guadalupe).
        pal = norm(ct).split()
        curto = norm(limpo).replace(".", "")
        iniciais = "".join(w[0] for w in pal if w)
        if curto and (any(w.startswith(curto) for w in pal) or curto == iniciais
                      or (len(curto) == 2 and iniciais.startswith(curto[0]) and curto[1] in iniciais[1:])):
            return ct
    return limpo


def linha_prevenda(d):
    cri = _dh(d.get("created_at"))
    pri = _dh(primeiro_contato(d))
    ult = _dh((cf_value(d, CF_PV["ultimo"]) if CF_PV["ultimo"] else None)
              or d.get("last_activity_at"))
    fec = _dh(d.get("closed_at"))
    cts = d.get("contacts") or []
    perda = (d.get("deal_lost_reason") or {}).get("name") or ""
    return [
        nome_prevenda(d),                                     # A Nome
        nome_etapa(d),                                        # B Etapa
        estado_pt(d),                                         # C Estado
        perda,                                                # D Motivo de Perda
        cri[0], cri[1],                                       # E,F criacao
        pri[0], pri[1],                                       # G,H primeiro contato
        ult[0], ult[1],                                       # I,J ultimo contato
        fec[0], fec[1],                                       # K,L fechamento
        (d.get("deal_source") or {}).get("name") or "",       # M Fonte
        (d.get("user") or {}).get("name") or "",              # N Responsavel
        cf_value(d, CF_PV["agendou"]) or "",                  # O Agendou avaliacao?
        cf_value(d, CF_PV["meio"]) or "",                     # P Meio
        cf_value(d, CF_PV["realizada"]) or "",                # Q Avaliacao Realizada?
        cf_value(d, CF_PV["avaliador"]) or "",                # R Avaliador
        serial(parse_dt(cf_value(d, CF_PV["dataAval"]))) or "",  # S Data da avaliacao
        cf_value(d, CF_PV["feegow"]) or "",                   # T ID Feegow
        str(d.get("id") or d.get("_id") or ""),               # U ID
        (d.get("_contato_id")
         or (str(cts[0].get("id") or cts[0].get("_id") or "") if cts else "")),  # V ID do Contato
    ]


_ETPV = None


def etapas_pv():
    global _ETPV
    if _ETPV is None:
        _ETPV = {norm(x) for x in ETAPAS_PV.split(",") if x.strip()}
    return _ETPV


def do_funil_pv(d):
    pl = d.get("deal_pipeline") or {}
    pid = pl.get("id") or pl.get("_id") or ""
    if PIPE_PV and pid:
        return pid == PIPE_PV
    if PIPE_PV_NM and pl.get("name"):
        return norm(PIPE_PV_NM) in norm(pl.get("name"))
    return norm(nome_etapa(d)) in etapas_pv()   # identificacao pela etapa


def buscar_prevendas():
    """Tenta filtro por periodo de criacao na origem; cai para cursor."""
    base = {"limit": 200, "created_at_period": "true",
            "start_date": CUTOFF_PV.isoformat(),
            "end_date": (dt.date.today() + dt.timedelta(days=2)).isoformat()}
    j = rd_get(dict(base, page=1), tolerante=True)
    usa_filtro = False
    if j:
        ds = j.get("deals", [])
        dentro = [x for x in ds if (c := parse_dt(x.get("created_at"))) and c >= CUTOFF_PV - dt.timedelta(days=1)]
        usa_filtro = bool(ds) and len(dentro) >= len(ds) * 0.8
    todos = []
    if usa_filtro:
        log.info("Pre-vendas: filtro de criacao aceito na origem (total=%s)", j.get("total"))
        todos = list(j.get("deals", []))
        pag = 2
        while len(j.get("deals", [])) == 200 and pag <= 60:
            j = rd_get(dict(base, page=pag), tolerante=True)
            if not j:
                break
            todos.extend(j.get("deals", []))
            if len(j.get("deals", [])) < 200:
                break
            pag += 1
    else:
        log.info("Pre-vendas: usando cursor")
        params, n, secas = {"limit": 200}, 0, 0
        while n < 400:
            j = rd_get(params)
            if not j:
                break
            lote = j.get("deals", [])
            if not lote:
                break
            todos.extend(lote)
            n += 1
            cs = [parse_dt(x.get("created_at")) for x in lote]
            if cs and all(c and c < CUTOFF_PV for c in cs if c):
                secas += 1
                if secas >= 2:
                    break
            else:
                secas = 0
            nxt = j.get("next_page")
            if not nxt or not j.get("has_more"):
                break
            params = {"limit": 200, "next_page": nxt}
            if n % 20 == 0:
                log.info("  %s paginas, %s negocios", n, len(todos))
    eleg = [d for d in todos if do_funil_pv(d)
            and (c := parse_dt(d.get("created_at"))) and c >= CUTOFF_PV]
    log.info("Pre-vendas: %s lidos, %s no funil desde %s", len(todos), len(eleg), CUTOFF_PV)
    return eleg


_CT_DEBUG = [0]


def contato_do_deal(rid):
    """GET /deals/{id}/contacts — a listagem de negocios nao traz o id do contato."""
    import json as _j
    try:
        j = rd_get({}, caminho=f"/deals/{rid}/contacts", tolerante=True)
    except Exception as e:
        if _CT_DEBUG[0] < 2:
            _CT_DEBUG[0] += 1
            log.warning("    contato %s: erro %s", rid, str(e)[:120])
        return ""
    lista = j if isinstance(j, list) else (
        (j.get("contacts") or j.get("data") or j.get("items") or []) if isinstance(j, dict) else [])
    for c in lista:
        if not isinstance(c, dict):
            continue
        cid = c.get("id") or c.get("_id") or c.get("contact_id") or ""
        if cid:
            return str(cid)
    if _CT_DEBUG[0] < 2:                       # mostra so as 2 primeiras falhas
        _CT_DEBUG[0] += 1
        log.warning("    contato %s sem id — resposta: %s", rid,
                    _j.dumps(j, ensure_ascii=False)[:300] if j is not None else "(vazia)")
    return ""


def enriquecer_contatos(deals, faltantes):
    if not PV_CONTATO:
        return
    alvo = [d for d in deals if id_do(d) in faltantes][:PV_CONTATO]
    if not alvo:
        return
    log.info("Buscando o ID do contato de %s negocio(s)...", len(alvo))
    ok = 0
    for d in alvo:
        cid = contato_do_deal(id_do(d))
        if cid:
            d["_contato_id"] = cid
            ok += 1
    log.info("  id do contato obtido para %s", ok)


def enriquecer(deals, faltantes):
    """Busca o detalhe so dos negocios sem primeiro contato, com teto por execucao."""
    if not PV_DETALHE:
        return
    alvo = [d for d in deals if id_do(d) in faltantes][:PV_DETALHE]
    if not alvo:
        return
    log.info("Buscando detalhe de %s negocio(s) para achar o primeiro contato...", len(alvo))
    ok = 0
    for d in alvo:
        det = buscar_deal(id_do(d))
        if isinstance(det, dict) and det.get("deal_stage_histories"):
            d["deal_stage_histories"] = det["deal_stage_histories"]
            if det.get("last_activity_at") and not d.get("last_activity_at"):
                d["last_activity_at"] = det["last_activity_at"]
            ok += 1
    log.info("  detalhe obtido para %s", ok)


def id_do(d):
    return str(d.get("id") or d.get("_id") or "")


def sincronizar_prevendas(tk):
    ws = f"{WB}/worksheets('{SHEET_PV}')"
    ur = g("GET", f"{ws}/usedRange(valuesOnly=true)?$select=address,rowCount", tk)
    total = int(ur.get("rowCount") or 0)
    existentes, CH = {}, 2000
    for ini in range(2, max(total, 1) + 1, CH):
        fim = min(total, ini + CH - 1)
        rg = g("GET", f"{ws}/range(address='A{ini}:V{fim}')?$select=values", tk)
        for j2, v in enumerate(rg.get("values", [])):
            rid = str(v[20]).strip() if len(v) > 20 and v[20] not in (None, "") else ""
            if rid:
                existentes[rid] = (ini + j2, v)
    log.info("Aba %s: %s linhas | %s com ID", SHEET_PV, max(0, total - 1), len(existentes))

    deals = buscar_prevendas()
    if PV_DETALHE:
        semPri = {rid for rid, (lin, v) in existentes.items()
                  if not (len(v) > 6 and v[6] not in (None, ""))}
        semPri |= {id_do(d) for d in deals if id_do(d) not in existentes}
        enriquecer(deals, semPri)
    if PV_CONTATO:
        semCt = {rid for rid, (lin, v) in existentes.items()
                 if not (len(v) > 21 and v[21] not in (None, ""))}
        semCt |= {id_do(d) for d in deals if id_do(d) not in existentes}
        enriquecer_contatos(deals, semCt)
    novas, atualiza = [], []
    for d in deals:
        l = linha_prevenda(d)
        rid = l[20]
        if not rid:
            continue
        alvo = existentes.get(rid)
        if alvo is None:
            novas.append(l)
            continue
        lin, atual = alvo
        def igual(a, b):
            if isinstance(a, (int, float)) or isinstance(b, (int, float)):
                try:
                    return abs(float(a or 0) - float(b or 0)) <= 1e-6
                except Exception:
                    pass
            return norm(a) == norm(b)
        dif = [i for i in range(22)
               if not igual(l[i], atual[i] if i < len(atual) else None)
               and not (l[i] in (None, "") and (atual[i] if i < len(atual) else None) not in (None, ""))]
        if dif:
            atualiza.append((lin, l, dif, list(atual)))

    if MAX_UPD and len(atualiza) > MAX_UPD:
        log.warning("Pre-vendas: %s atualizacoes pendentes, processando %s", len(atualiza), MAX_UPD)
        atualiza = atualiza[:MAX_UPD]
    log.info("Pre-vendas: %s nova(s) | %s atualizacao(oes)", len(novas), len(atualiza))

    COLS_PV = ["Nome", "Etapa", "Estado", "MotivoPerda", "DataCri", "HoraCri",
               "DataPrimContato", "HoraPrimContato", "DataUltContato", "HoraUltContato",
               "DataFech", "HoraFech", "Fonte", "Responsavel", "Agendou", "Meio",
               "Realizada", "Avaliador", "DataAval", "IDFeegow", "ID", "IDContato"]

    def mostraPV(x):
        if x in (None, ""):
            return "(vazio)"
        if isinstance(x, (int, float)):
            f = float(x)
            if 20000 < f < 80000:
                return (EPOCH + dt.timedelta(days=int(f))).strftime("%d/%m/%Y")
            if 0 < f < 1:
                m = round(f * 1440)
                return f"{m//60:02d}:{m%60:02d}"
        return str(x)[:34]

    for lin, l, dif, antes in atualiza:
        det = ", ".join(f"{COLS_PV[i]}: {mostraPV(antes[i] if i < len(antes) else None)} -> {mostraPV(l[i])}"
                        for i in dif)
        if DRYRUN:
            log.info("  DRY pv linha %s: %s", lin, det)
            continue
        g("PATCH", f"{ws}/range(address='A{lin}:V{lin}')", tk,
          json={"values": [mesclar(antes, l, dif, 22)]})
        log.info("  pv linha %s: %s", lin, det[:160])
    if novas and not DRYRUN:
        ini = total + 1
        for k in range(0, len(novas), 50):
            bloco = novas[k:k + 50]
            a, b = ini + k, ini + k + len(bloco) - 1
            g("PATCH", f"{ws}/range(address='A{a}:V{b}')", tk, json={"values": bloco})
            log.info("  pv gravadas %s/%s", min(k + 50, len(novas)), len(novas))
    elif novas:
        for l in novas[:15]:
            log.info("  DRY pv nova: %s", l[:6])


# ─── sincronizacao da aba Marcacoes a partir do Feegow ────────────────────────
def mesclar(atual, novo, dif, largura):
    """Escreve so o que mudou: parte da linha existente e troca as posicoes de dif.
    Sem isso, gravar um intervalo apaga as colunas que este sync nao preenche."""
    saida = list(atual)[:largura]
    while len(saida) < largura:
        saida.append("")
    for i in dif:
        if i < largura:
            saida[i] = novo[i]
    return saida


def fg_mapa(caminho, chaveId, chaveNome):
    j = fg_get(caminho)
    m = {}
    for it in fg_conteudo(j):
        pid = str(it.get(chaveId) or it.get("id") or "")
        nome = it.get(chaveNome) or it.get("nome") or ""
        if pid:
            m[pid] = str(nome).strip()
    return m


def fg_data_serial(txt):
    """'07-08-2024' -> serial do Excel"""
    try:
        d, m, a = str(txt).split("-")
        return serial(dt.date(int(a), int(m), int(d)))
    except Exception:
        return ""


def fg_hora_frac(txt):
    try:
        p = str(txt).split(":")
        return (int(p[0]) * 60 + int(p[1])) / 1440.0
    except Exception:
        return ""


def fg_dh(txt):
    """'2024-05-09 18:36:10' -> serial + fracao (sem segundos)"""
    if not txt:
        return "", ""
    try:
        x = dt.datetime.fromisoformat(str(txt).replace("Z", "").strip())
        return serial(x.date()), (x.hour * 60 + x.minute) / 1440.0
    except Exception:
        return "", ""


def fg_buscar_agenda():
    de = dt.date.today() - dt.timedelta(days=FG_DIAS_TRAS)
    ate = dt.date.today() + dt.timedelta(days=FG_DIAS_FRENTE)
    base = {"data_start": de.strftime("%d-%m-%Y"), "data_end": ate.strftime("%d-%m-%Y"),
            "list_procedures": 1}
    todos, start, OFF = [], 0, 200
    while start < 20000:
        j = fg_get("/appoints/search", dict(base, start=start, offset=OFF))
        lote = fg_conteudo(j)
        if not lote:
            break
        todos.extend(lote)
        if len(lote) < OFF:
            break
        start += OFF
        if start % 1000 == 0:
            log.info("  feegow: %s agendamentos lidos", len(todos))
    log.info("Feegow: %s agendamento(s) de %s a %s", len(todos), de, ate)
    return todos


def sincronizar_marcacoes(tk):
    if not FG_TOKEN:
        log.warning("FEEGOW_TOKEN ausente — pulando a aba %s", SHEET_MC)
        return
    ws = f"{WB}/worksheets('{SHEET_MC}')"
    ur = g("GET", f"{ws}/usedRange(valuesOnly=true)?$select=rowCount", tk)
    total = int(ur.get("rowCount") or 0)

    porAgd, nomePorId, CH = {}, {}, 2000
    for ini in range(2, max(total, 1) + 1, CH):
        fim = min(total, ini + CH - 1)
        rg = g("GET", f"{ws}/range(address='A{ini}:O{fim}')?$select=values", tk)
        for k, v in enumerate(rg.get("values", [])):
            lin = ini + k
            aid = str(v[MC["idAgd"]]).strip().replace(".0", "") if len(v) > MC["idAgd"] else ""
            if aid:
                porAgd[aid] = (lin, v)
            pid = str(v[MC["idPac"]]).strip().replace(".0", "") if len(v) > MC["idPac"] else ""
            nm = str(v[MC["paciente"]]).strip() if len(v) > MC["paciente"] else ""
            if pid and nm:
                nomePorId.setdefault(pid, nm)
    log.info("Aba %s: %s linha(s) | %s com ID de agendamento | %s nome(s) em cache",
             SHEET_MC, max(0, total - 1), len(porAgd), len(nomePorId))

    status = fg_mapa("/appoints/status", "id", "status")
    procs = fg_mapa("/procedures/list", "procedimento_id", "nome")
    profs = fg_mapa("/professional/list", "profissional_id", "nome")
    log.info("Tabelas: %s status | %s procedimentos | %s profissionais",
             len(status), len(procs), len(profs))

    bruto = fg_buscar_agenda()
    vistos, agenda = set(), []
    for a in bruto:
        aid = str(a.get("agendamento_id") or "")
        if aid and aid in vistos:
            continue
        vistos.add(aid); agenda.append(a)
    if len(agenda) != len(bruto):
        log.info("  %s duplicata(s) da paginacao descartada(s)", len(bruto) - len(agenda))
    alvo = set(FG_PROCS)
    agora = dt.datetime.now()
    novas, atualiza, semNome, ignorados = [], [], set(), 0

    for a in agenda:
        aid = str(a.get("agendamento_id") or "")
        if not aid:
            continue
        pid = str(a.get("paciente_id") or "")
        proc = str(a.get("procedimento_id") or "")
        existente = porAgd.get(aid)
        if existente is None and proc not in alvo:
            ignorados += 1
            continue                       # fora da lista e nao esta na planilha
        st = status.get(str(a.get("status_id") or ""), "")
        dS, hS = fg_data_serial(a.get("data")), fg_hora_frac(a.get("horario"))
        mkD, mkH = fg_dh(a.get("agendado_em"))
        marcado = (mkD + mkH) if mkD != "" else ""
        nome = nomePorId.get(pid, "")
        if not nome and existente:
            nome = str(existente[1][MC["paciente"]] or "").strip()
        if not nome:
            semNome.add(pid)
        linha = [""] * MC_LARG
        linha[MC["data"]] = dS
        linha[MC["hora"]] = hS
        linha[MC["paciente"]] = nome
        linha[MC["idPac"]] = int(pid) if str(pid).isdigit() else pid
        linha[MC["tipo"]] = procs.get(proc, "")
        linha[MC["prof"]] = profs.get(str(a.get("profissional_id") or ""), "")
        linha[MC["status"]] = st
        linha[MC["agendou"]] = (a.get("agendado_por") or "").strip()
        linha[MC["marcado"]] = marcado
        linha[MC["idAgd"]] = int(aid) if str(aid).isdigit() else aid
        # "Realizado em" nao existe na API: registra quando o sync viu virar Atendido
        if norm(st) == norm(FG_ST_ATENDIDO):
            ja = existente and len(existente[1]) > MC["realizado"] and \
                 str(existente[1][MC["realizado"]] or "").strip()
            linha[MC["realizado"]] = existente[1][MC["realizado"]] if ja else \
                serial(agora.date()) + (agora.hour * 60 + agora.minute) / 1440.0
        if existente is None:
            novas.append(linha)
        else:
            lin, atual = existente
            dif = []
            campos = [MC["data"], MC["hora"], MC["tipo"], MC["prof"], MC["status"],
                      MC["agendou"], MC["marcado"], MC["realizado"], MC["idPac"]]
            # o nome do Feegow as vezes e mais curto que o da planilha: so preenche vazio
            nomeAtual = str(atual[MC["paciente"]] or "").strip() if len(atual) > MC["paciente"] else ""
            if not nomeAtual:
                campos.append(MC["paciente"])
            DATAS = (MC["data"], MC["hora"], MC["marcado"], MC["realizado"])
            for i in campos:
                novo = linha[i]
                velho = atual[i] if i < len(atual) else None
                if novo in (None, "") and velho not in (None, ""):
                    continue               # nunca apaga campo preenchido
                if isinstance(novo, (int, float)) or isinstance(velho, (int, float)):
                    try:
                        a1, b1 = float(novo or 0), float(velho or 0)
                        # datas e horas sao gravadas com precisao de minuto:
                        # comparar em fracao de dia acusa diferenca de segundos que nao existe
                        if i in DATAS:
                            if round(a1 * 1440) == round(b1 * 1440):
                                continue
                        elif abs(a1 - b1) <= 1e-9:
                            continue
                    except Exception:
                        pass
                elif norm(novo) == norm(velho):
                    continue
                dif.append(i)
            if dif:
                atualiza.append((lin, linha, dif, list(atual)))

    log.info("Feegow: %s nova(s) | %s atualizacao(oes) | %s ignorado(s) (fora da lista)",
             len(novas), len(atualiza), ignorados)
    if semNome:
        log.info("Pacientes sem nome em cache: %s (buscando ate %s)", len(semNome), FG_MAX_NOMES)
        achou = 0
        for pid in list(semNome)[:FG_MAX_NOMES]:
            nm = fg_nome_paciente(pid)
            if nm:
                nomePorId[pid] = nm; achou += 1
            elif _FG_PAC_OK[0] is None and achou == 0:
                log.warning("  nenhum caminho de paciente funcionou — nomes ficarao em branco")
                break
        log.info("  %s nome(s) obtido(s)", achou)
        for l in novas:
            if not l[MC["paciente"]]:
                l[MC["paciente"]] = nomePorId.get(l[MC["idPac"]], "")
        for _, l, dif, _a in atualiza:
            if not l[MC["paciente"]]:
                l[MC["paciente"]] = nomePorId.get(l[MC["idPac"]], "")

    COLS_MC = {v: k for k, v in MC.items()}

    COL_DATA = {MC["data"], MC["marcado"], MC["realizado"], MC["evento"]}

    def mostraMC(x, col=None):
        if x in (None, ""):
            return "(vazio)"
        if isinstance(x, (int, float)):
            f = float(x)
            if col in COL_DATA and 20000 < f < 80000:
                d0 = EPOCH + dt.timedelta(days=f)
                return d0.strftime("%d/%m/%Y %H:%M" if abs(f - int(f)) > 1e-9 else "%d/%m/%Y")
            if col == MC["hora"] and 0 < f < 1:
                m = round(f * 1440)
                return f"{m//60:02d}:{m%60:02d}"
            return str(int(f)) if f == int(f) else str(f)
        return str(x)[:30] + " [texto]"

    if MAX_UPD and len(atualiza) > MAX_UPD:
        log.warning("Feegow: %s atualizacoes pendentes, processando %s", len(atualiza), MAX_UPD)
        atualiza = atualiza[:MAX_UPD]
    for lin, l, dif, antes in atualiza:
        det = ", ".join(f"{COLS_MC[i]}: {mostraMC(antes[i] if i < len(antes) else None, i)}"
                        f" -> {mostraMC(l[i], i)}" for i in dif)
        if DRYRUN:
            log.info("  DRY mc linha %s (%s): %s", lin, str(l[MC['paciente']])[:24], det)
            continue
        g("PATCH", f"{ws}/range(address='A{lin}:O{lin}')", tk,
          json={"values": [mesclar(antes, l, dif, MC_LARG)]})
        log.info("  mc linha %s: %s", lin, det[:150])
    if novas:
        if DRYRUN:
            for l in novas[:10]:
                log.info("  DRY mc nova: %s | %s | %s | %s", mostraMC(l[0], MC["data"]),
                         mostraMC(l[1], MC["hora"]), str(l[2])[:26], l[6])
        else:
            ini = total + 1
            for k in range(0, len(novas), 50):
                bloco = novas[k:k + 50]
                a2, b2 = ini + k, ini + k + len(bloco) - 1
                g("PATCH", f"{ws}/range(address='A{a2}:O{b2}')", tk, json={"values": bloco})
                log.info("  mc gravadas %s/%s", min(k + 50, len(novas)), len(novas))


# ─── sincronizacao da planilha de Cirurgias ───────────────────────────────────
def sincronizar_cirurgias(tk):
    if not FG_TOKEN:
        log.warning("FEEGOW_TOKEN ausente — pulando Cirurgias"); return
    if not ARQ_CIR:
        log.warning("ARQ_CIRURGIAS nao configurado — pulando"); return
    ref = resolver_arquivo_url(tk, ARQ_CIR)
    wb2 = f"{GRAPH}/drives/{ref['drive']}/items/{ref['item']}/workbook"
    ws = f"{wb2}/worksheets('{SHEET_CIR}')"
    log.info("Cirurgias: arquivo %s", ref.get("nome"))
    guard = SESSAO["id"]; SESSAO["id"] = None      # a sessao global e do outro arquivo
    try:
        ur = g("GET", f"{ws}/usedRange(valuesOnly=true)?$select=rowCount", tk)
        total = int(ur.get("rowCount") or 0)
        porAgd, CH = {}, 2000
        for ini in range(2, max(total, 1) + 1, CH):
            fim = min(total, ini + CH - 1)
            rg = g("GET", f"{ws}/range(address='A{ini}:F{fim}')?$select=values", tk)
            for k, v in enumerate(rg.get("values", [])):
                aid = str(v[CC["idAgd"]]).strip().replace(".0", "") if len(v) > CC["idAgd"] else ""
                if aid:
                    porAgd[aid] = (ini + k, v)
        log.info("Cirurgias: %s linha(s) | %s com ID de agendamento", max(0, total - 1), len(porAgd))

        alvo, mapa = cir_procedimentos_alvo()
        if not alvo:
            log.error("Nenhum procedimento de cirurgia identificado. Rode o modo cirurgias-diag.")
            return
        status = fg_mapa("/appoints/status", "id", "status")
        okSt = {norm(x) for x in CIR_STATUS.split(",") if x.strip()}
        de, ate = cir_intervalo()
        log.info("Procedimentos de cirurgia: %s | janela %s a %s", sorted(alvo), de, ate)
        agenda = fg_buscar_cirurgias(alvo, de, ate)
        log.info("Feegow: %s agendamento(s) de cirurgia", len(agenda))

        novas, atualiza, fora = [], [], 0
        for a in agenda:
            aid = str(a.get("agendamento_id") or "")
            st = status.get(str(a.get("status_id") or ""), "")
            existente = porAgd.get(aid)
            if norm(st) not in okSt and existente is None:
                fora += 1; continue
            linha = [""] * CC_LARG
            linha[CC["data"]] = fg_data_serial(a.get("data"))
            linha[CC["idAgd"]] = int(aid) if aid.isdigit() else aid
            linha[CC["proc"]] = mapa.get(str(a.get("procedimento_id") or ""), "")
            pid2 = str(a.get("paciente_id") or "")
            linha[CC["idPac"]] = int(pid2) if pid2.isdigit() else pid2
            linha[CC["situacao"]] = st
            v1 = toNumF(a.get("valor"))
            v2 = toNumF(a.get("valor_total_agendamento"))
            linha[CC["valor"]] = v1 if isinstance(v1, float) and v1 else (v2 if isinstance(v2, float) and v2 else "")
            if existente is None:
                novas.append(linha); continue
            lin, atual = existente
            dif = []
            for i in (CC["data"], CC["proc"], CC["idPac"], CC["situacao"], CC["valor"]):
                novo, velho = linha[i], (atual[i] if i < len(atual) else None)
                if novo in (None, "") and velho not in (None, ""):
                    continue
                try:
                    if isinstance(novo, (int, float)) or isinstance(velho, (int, float)):
                        if abs(float(novo or 0) - float(velho or 0)) <= 1e-6:
                            continue
                    elif norm(novo) == norm(velho):
                        continue
                except Exception:
                    if norm(novo) == norm(velho):
                        continue
                dif.append(i)
            if dif:
                atualiza.append((lin, linha, dif, list(atual)))

        log.info("Cirurgias: %s nova(s) | %s atualizacao(oes) | %s fora dos status", 
                 len(novas), len(atualiza), fora)
        COLS_CC = {v: k for k, v in CC.items()}

        def mostraCC(x, col=None):
            if x in (None, ""): return "(vazio)"
            if isinstance(x, (int, float)):
                f = float(x)
                if col == CC["data"] and 20000 < f < 80000:
                    return (EPOCH + dt.timedelta(days=int(f))).strftime("%d/%m/%Y")
                return str(int(f)) if f == int(f) else str(f)
            return str(x)[:30]

        if MAX_UPD and len(atualiza) > MAX_UPD:
            log.warning("Cirurgias: %s pendentes, processando %s", len(atualiza), MAX_UPD)
            atualiza = atualiza[:MAX_UPD]
        for lin, l, dif, antes in atualiza:
            det = ", ".join(f"{COLS_CC[i]}: {mostraCC(antes[i] if i < len(antes) else None, i)}"
                            f" -> {mostraCC(l[i], i)}" for i in dif)
            if DRYRUN:
                log.info("  DRY cir linha %s: %s", lin, det); continue
            g("PATCH", f"{ws}/range(address='A{lin}:F{lin}')", tk,
              json={"values": [mesclar(antes, l, dif, CC_LARG)]})
            log.info("  cir linha %s: %s", lin, det[:150])
        if novas:
            if DRYRUN:
                for l in novas[:10]:
                    log.info("  DRY cir nova: %s | agd %s | %s | %s",
                             mostraCC(l[0], 0), l[1], str(l[2])[:22], l[4])
            else:
                ini = total + 1
                for k in range(0, len(novas), 50):
                    bloco = novas[k:k + 50]
                    a2, b2 = ini + k, ini + k + len(bloco) - 1
                    g("PATCH", f"{ws}/range(address='A{a2}:F{b2}')", tk,
                      json={"values": bloco})
                    log.info("  cir gravadas %s/%s", min(k + 50, len(novas)), len(novas))
    finally:
        SESSAO["id"] = guard


def toNumF(x):
    """Aceita 'R$ 1.500,50', '1500.5', 1500.5 — devolve float ou string vazia."""
    if x in (None, ""):
        return ""
    if isinstance(x, (int, float)):
        return float(x)
    t = re.sub(r"[^\d,.-]", "", str(x))
    if not t or t in ("-", ".", ","):
        return ""
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return float(t)
    except Exception:
        return ""


def _share_id(url):
    b = base64.b64encode(url.encode("utf-8")).decode().rstrip("=")
    return "u!" + b.replace("/", "_").replace("+", "-")


def resolver_arquivo_url(tk, url):
    it = g("GET", f"{GRAPH}/shares/{_share_id(url)}/driveItem?$select=id,name,parentReference", tk)
    return {"drive": it["parentReference"]["driveId"], "item": it["id"], "nome": it.get("name")}


# ─── main ─────────────────────────────────────────────────────────────────────
def main():
    log.info("Config: drive=%s... file_id=%s path=%s", DRIVE_ID[:12], FILE_ID or "(vazio)", FILE_PATH)
    if PIPES_D:
        listar_pipelines()
        return
    if FG_PROC:
        fg_procedimentos()
        return
    if FG_CIRD:
        cirurgias_diag(); return
    if FG_CIR:
        tk0 = token(); resolver_arquivo(tk0)
        try: sincronizar_cirurgias(tk0)
        finally: pass
        return
    if FG_AGENDA:
        tk0 = token(); resolver_arquivo(tk0); abrir_sessao(tk0)
        try: sincronizar_marcacoes(tk0)
        finally: fechar_sessao(tk0)
        return
    if FEEGOW_BF:
        backfill_feegow()
        return
    if FEEGOW_D:
        cobertura_feegow()
        return
    if BRUTO_PV:
        bruto_prevendas()
        return
    if CAMPOS_PV:
        campos_prevendas()
        return
    if CAMPOS:
        for d in [x for x in buscar_deals() if elegivel(x)][:3]:
            log.info("=== %s | fechado %s", d.get("name"), d.get("closed_at"))
            for c in d.get("deal_custom_fields", []):
                cid = c.get("custom_field_id") or (c.get("custom_field") or {}).get("_id")
                lab = (c.get("custom_field") or {}).get("label") or c.get("label") or ""
                log.info("    %-26s %-34s = %r", cid, lab[:34], c.get("value"))
            log.info("    deal_products: %r", d.get("deal_products"))
        return
    if DIAG:
        diagnostico()
        return
    if LISTAR:
        listar_drive(token())
        return

    log.info("CUTOFF %s | aba %s | tabela %s | dry-run=%s", CUTOFF, SHEET, TABLE, DRYRUN)

    if SO_PV:
        tk0 = token(); resolver_arquivo(tk0); abrir_sessao(tk0)
        sincronizar_prevendas(tk0); fechar_sessao(tk0); return

    deals = buscar_deals()
    elegiveis = [d for d in deals if elegivel(d)]
    log.info("Elegiveis (fechados a partir de %s): %s", CUTOFF, len(elegiveis))

    novas = []
    for d in elegiveis:
        novas.extend(linhas_do_deal(d))
    log.info("Linhas candidatas (1 por negocio): %s", len(novas))
    if not novas:
        log.info("Nada a fazer.")
        return

    tk = token()
    resolver_arquivo(tk)
    abrir_sessao(tk)
    existentes, por_id, por_dvp, _ = ler_existentes(tk)

    def igual(a, b):
        if isinstance(a, (int, float)) or isinstance(b, (int, float)):
            try:
                return abs(float(a or 0) - float(b or 0)) <= 0.005
            except Exception:
                pass
        return norm(a) == norm(b)

    inserir_agora, atualizar, vistas = [], [], set()
    for l in novas:
        rid = l[I_ID]
        if rid and rid in vistas:
            continue
        vistas.add(rid)

        alvo = por_id.get(rid)
        origem = "id"
        if alvo is None:                       # linha antiga, ainda sem ID gravado
            sk = softkey(l)
            if sk and sk in existentes:
                alvo = existentes[sk]
                origem = "chave"

        if alvo is None:                       # negocio renomeado no RD
            k3 = chave_sem_nome(l)
            cands = [c for c in por_dvp.get(k3, [])
                     if not (len(c[1]) > I_ID and str(c[1][I_ID]).strip())]
            if len(cands) == 1:
                alvo = cands[0]
                origem = "renomeado"
            elif len(cands) > 1:
                log.warning("  %s: %s linhas candidatas com mesma data/valor/produto — inserindo nova",
                            str(l[0])[:40], len(cands))

        if alvo is None:
            inserir_agora.append(l)
            continue

        linha, atual = alvo

        def vazio(x):
            return x is None or (isinstance(x, str) and not x.strip())

        dif = []
        for i in IDX_CMP:
            cur = atual[i] if i < len(atual) else None
            if igual(l[i], cur):
                continue
            if vazio(l[i]) and not vazio(cur):
                continue            # RD sem dado nao apaga o que ja existe
            dif.append(i)
        for i in IDX_CMP:
            if i not in dif:
                l[i] = atual[i] if i < len(atual) else l[i]
        falta_id = origem in ("chave", "renomeado") or not (len(atual) > I_ID and str(atual[I_ID]).strip())
        if dif or falta_id:
            atualizar.append((linha, l, dif, falta_id, origem, list(atual)))

    if MAX_UPD and len(atualizar) > MAX_UPD:
        log.warning("Atualizacoes pendentes: %s — processando %s agora, resto no proximo ciclo",
                    len(atualizar), MAX_UPD)
        atualizar = atualizar[:MAX_UPD]
    ids_rd = {l[I_ID] for l in novas if l[I_ID]}
    candidatas = [(rid, lin, v) for rid, (lin, v) in por_id.items()
                  if rid and rid not in ids_rd
                  and any(t in norm(v[1] if len(v) > 1 else "") for t in etapas_venda())]
    if candidatas:
        log.info("Conferindo no RD %s linha(s) que sumiram da consulta...", len(candidatas))
    revertidas = []
    for rid, lin, v in candidatas[:40]:
        d = buscar_deal(rid)
        if d is None:
            log.warning("   linha %s (%s): nao consegui verificar no RD — mantida",
                        lin, str(v[0])[:34])
            continue
        if d == "404":
            revertidas.append((lin, v[0], "EXCLUIDO NO RD"))
            continue
        if d.get("win") is False:
            # perdido: nunca gravar o nome da etapa, que pode ser uma etapa de venda
            revertidas.append((lin, v[0], "PERDIDO"))
        elif not etapa_de_venda(d):
            revertidas.append((lin, v[0], (nome_etapa(d) or "REVERTIDO NO RD").upper()))
        # caso contrario continua sendo venda (data mudou, etc) — nao mexe

    if revertidas:
        log.warning("%s venda(s) deixaram de valer no RD:", len(revertidas))
        ws2 = f"{WB}/worksheets('{SHEET}')"
        for lin, nome, novo_st in revertidas:
            if DRYRUN or not MARCAR_REV:
                log.warning("   DRY linha %s (%s) -> Etapa '%s'", lin, str(nome)[:34], novo_st)
            else:
                g("PATCH", f"{ws2}/range(address='B{lin}')", tk, json={"values": [[novo_st]]})
                log.warning("   linha %s (%s) -> Etapa '%s' — fora do total", lin, str(nome)[:34], novo_st)

    log.info("Novas: %s | atualizacoes: %s", len(inserir_agora), len(atualizar))

    COLS = {0: "Nome", 1: "Etapa", 2: "Valor", 3: "Criacao", 4: "Fechamento", 5: "Fonte",
            6: "Responsavel", 7: "Produtos", 8: "Meio", 9: "Avaliador", 10: "MesAval",
            I_CIR: "DataCirurgia", I_MARC: "CirurgiaMarcada"}
    if I_FEE >= 0:
        COLS[I_FEE] = "IDFeegow"
    ws = f"{WB}/worksheets('{SHEET}')"
    def mostra(x):
        if x is None or x == "":
            return "(vazio)"
        if isinstance(x, (int, float)) and 20000 < float(x) < 80000:
            return (EPOCH + dt.timedelta(days=int(float(x)))).strftime("%d/%m/%Y")
        return str(x)[:30]

    for linha, l, dif, falta_id, origem, antes in atualizar:
        campos = ", ".join(f"{COLS.get(i, 'col'+str(i))}: {mostra(antes[i] if i < len(antes) else None)} -> {mostra(l[i])}"
                           for i in dif) or "-"
        if DRYRUN:
            log.info("  DRY linha %s [%s] (%s): %s%s", linha, origem, str(l[0])[:26], campos,
                     " +ID" if falta_id else "")
            continue
        if any(i <= 10 for i in dif):
            g("PATCH", f"{ws}/range(address='A{linha}:K{linha}')", tk,
              json={"values": [mesclar(antes, l, [i for i in dif if i <= 10], 11)]})
        if any(i in (I_CIR, I_MARC) for i in dif):
            cir = [antes[I_CIR] if I_CIR < len(antes) else "",
                   antes[I_MARC] if I_MARC < len(antes) else ""]
            if I_CIR in dif:  cir[0] = l[I_CIR]
            if I_MARC in dif: cir[1] = l[I_MARC]
            g("PATCH", f"{ws}/range(address='{COL_CIR_L}{linha}:{COL_MARC_L}{linha}')", tk,
              json={"values": [cir]})
        if falta_id:
            g("PATCH", f"{ws}/range(address='{COL_ID_L}{linha}')", tk, json={"values": [[l[I_ID]]]})
        if I_FEE >= 0 and I_FEE in dif:
            g("PATCH", f"{ws}/range(address='{COL_FEE_L}{linha}')", tk,
              json={"values": [[l[I_FEE]]]})
        log.info("  linha %s [%s] (%s): %s%s", linha, origem, str(l[0])[:26], campos,
                 " +ID" if falta_id else "")

    if not inserir_agora:
        fechar_sessao(tk)
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
    if not SEM_PV:
        try: sincronizar_prevendas(tk)
        except Exception as e: log.error("Pre-vendas falhou: %s", str(e)[:200])
    if FG_TOKEN:
        try: sincronizar_marcacoes(tk)
        except Exception as e: log.error("Marcacoes/Feegow falhou: %s", str(e)[:200])
    if FG_TOKEN and ARQ_CIR:
        try: sincronizar_cirurgias(tk)
        except Exception as e: log.error("Cirurgias falhou: %s", str(e)[:200])
    fechar_sessao(tk)
    log.info("OK — %s linhas inseridas", len(inserir_agora))


if __name__ == "__main__":
    main()
