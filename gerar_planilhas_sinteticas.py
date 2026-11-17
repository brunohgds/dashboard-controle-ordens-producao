"""
Gera as 4 planilhas do projeto "Controle de Ordens de Produção" com dados 100% SINTÉTICOS.

As planilhas mantêm os mesmos nomes de arquivo, nomes de abas, cabeçalhos e tipos
esperados pelo relatório Power BI, então basta colocá-las na pasta e atualizar.
Nenhum valor vem de empresa, cliente, produto ou sistema real (seed fixa => reproduzível).

Uso:   pip install pandas numpy openpyxl
       python gerar_planilhas_sinteticas.py
"""
import datetime as dt
import os
import shutil
import zipfile
import numpy as np
import openpyxl

SEED = 2026
REF = dt.date(2026, 1, 8)            # data de referência dos agings (equivale ao "hoje")
TOL = 0.10                           # ±10% => balanceada (OK)
OUT = "."
rng = np.random.default_rng(SEED)

# --------------------------------------------------------------------- utilidades
def d(days_from_ref):
    return REF - dt.timedelta(days=int(days_from_ref))

def rand_date(a, b):
    """data aleatória entre a e b (inclusive)"""
    return a + dt.timedelta(days=int(rng.integers(0, (b - a).days + 1)))

def clip_date(x, lo=None, hi=REF):
    if lo and x < lo: x = lo
    if hi and x > hi: x = hi
    return x

def faixa(n):
    if n <= 7: return "000 a 007"
    if n <= 14: return "008 a 014"
    if n <= 21: return "015 a 021"
    if n <= 28: return "022 a 028"
    return "Acima de 028"

def pick(options, probs):
    return options[int(rng.choice(len(options), p=probs))]

# ----------------------------------------------------------------------- catálogos
FASES = ["CORTE", "IMPRESSÃO", "LAMINAÇÃO", "ESTAMPARIA", "EXTRUSÃO"]
P_FASE = [0.55, 0.20, 0.12, 0.08, 0.05]
PERDA_FASE = {"CORTE": 0.035, "IMPRESSÃO": 0.07, "LAMINAÇÃO": 0.055,
              "ESTAMPARIA": 0.05, "EXTRUSÃO": 0.04}
MAQ = {"CORTE": ["COR-01", "COR-02", "COR-03", "COR-04", "COR-05", "COR-06"],
       "IMPRESSÃO": ["IMP-01", "IMP-02", "IMP-03"],
       "LAMINAÇÃO": ["LAM-01", "LAM-02", "LAM-03"],
       "ESTAMPARIA": ["EST-01", "EST-02"],
       "EXTRUSÃO": ["EXT-01", "EXT-02"]}
FAMILIAS = ["BLISTER", "TRILAMINADO", "COLDFORM", "MONOCAMADA", "BILAMINADO",
            "SELO PRE CORTADO", "MIX PAPER", ""]
P_FAM = [0.30, 0.22, 0.15, 0.08, 0.08, 0.06, 0.04, 0.07]
ABREV = {"BLISTER": "BLT", "TRILAMINADO": "TRI", "COLDFORM": "CFM", "MONOCAMADA": "MON",
         "BILAMINADO": "BIL", "SELO PRE CORTADO": "SPC", "MIX PAPER": "MXP", "": "GEN"}
TIPOS_OP = ["ZPP8-Make to Stock", "ZPP7-Make to Order", "ZPP9-Ordem de Retrabalho",
            "ZP11-Ordem combinada"]
P_TIPO = [0.51, 0.41, 0.05, 0.03]
STATUS_SIS = {
    "Aberta": ["ABER CAPC ERRD MatC NOLQ ORSL", "ABER MATF CAPC NOLQ ORSL",
               "ABER CAPC MatC NOLQ ORSL", "LIB  MATF CAPC LOAT NOLQ ORSL",
               "LIB  CAPC LOAT MatC NOLQ ORSL"],
    "Encerrada": ["ENTE MATF CNPA FORN CAPC DSVA LOAT MOME*", "ENTE CNPA FORN CAPC DSVA LOAT MOME MatC*"],
    "Processo concluído": ["LIB  MATF CNPA CAPC FRNP LOAT MOME NOLQ*", "LIB  MATF CNPA FORN CAPC DSVA LOAT MOME*"],
}

used_codes, used_ops = set(), set()

def new_code():
    while True:
        c = int("9" + "".join(str(x) for x in rng.integers(0, 10, 11)))
        if c not in used_codes:
            used_codes.add(c); return c

def new_op(lo=71_000_000, hi=71_999_999):
    while True:
        o = int(rng.integers(lo, hi))
        if o not in used_ops:
            used_ops.add(o); return o

# catálogo de itens fictícios
CATALOGO = []
for _ in range(240):
    fam = pick(FAMILIAS, P_FAM)
    esp = int(rng.choice([12, 15, 20, 21, 25, 30, 40, 50, 60]))
    larg = int(rng.choice(range(180, 960, 10)))
    var = "".join(rng.choice(list("ABCDEFGH"), 1))
    CATALOGO.append({
        "codigo": new_code(),
        "desc": f"{ABREV[fam]}{esp}u PADRAO {larg}MM {var}",
        "familia": fam,
        "tipo_mat": pick(["FERT", "HALB", "ROH"], [0.69, 0.30, 0.01]),
    })

# ------------------------------------------------------------------ estados das ordens
# (status geral, tem início, tem fim, apontamento, quantidade)
ESTADOS = [
    ("Aberta", False, False, "Não produziu", 1170),
    ("Encerrada", True, True, "Em produção", 205),
    ("Processo concluído", True, True, "Em produção", 118),
    ("Aberta", True, False, "Não produziu", 72),
    ("Encerrada", False, False, "Não produziu", 11),
    ("Encerrada", True, False, "Não produziu", 3),
    ("Aberta", True, False, "Em produção", 1),
]
INI_MIN, FIM_MIN, ABR_MIN = dt.date(2025, 10, 14), dt.date(2025, 12, 12), dt.date(2025, 7, 2)

ordens = []
for est in ESTADOS:
    for _ in range(est[4]):
        ordens.append({"geral": est[0], "tem_ini": est[1], "tem_fim": est[2], "apont": est[3]})
rng.shuffle(ordens)
ops = sorted(new_op() for _ in ordens)

for o, op in zip(ordens, ops):
    it = CATALOGO[int(rng.integers(len(CATALOGO)))]
    o.update(op=op, item=it)
    o["tipo_op"] = pick(TIPOS_OP, P_TIPO)
    o["unidade"] = pick(["M", "MIL", "KG"], [0.965, 0.03, 0.005])
    ini = fim = lib = mod = None
    if o["tem_fim"]:
        fim = rand_date(FIM_MIN, REF)
        ini = clip_date(fim - dt.timedelta(days=int(rng.integers(1, 10))), INI_MIN)
        abr = clip_date(ini - dt.timedelta(days=int(rng.integers(3, 60))), ABR_MIN)
        lib = clip_date(abr + dt.timedelta(days=int(rng.integers(0, 6))), None, ini)
        mod = clip_date(fim + dt.timedelta(days=int(rng.integers(0, 9))))
    elif o["tem_ini"]:
        ini = rand_date(INI_MIN, REF)
        abr = clip_date(ini - dt.timedelta(days=int(rng.integers(3, 60))), ABR_MIN)
        lib = clip_date(abr + dt.timedelta(days=int(rng.integers(0, 6))), None, ini)
        mod = clip_date(ini + dt.timedelta(days=int(rng.integers(0, 11))))
    else:
        abr = rand_date(ABR_MIN, REF)
        if rng.random() < 0.30:
            lib = clip_date(abr + dt.timedelta(days=int(rng.integers(0, 8))))
        if rng.random() < 0.45:
            base = lib or abr
            mod = clip_date(base + dt.timedelta(days=int(rng.integers(0, 31))))
    o.update(abr=abr, lib=lib, ini=ini, fim=fim, mod=mod)
    prog = float(rng.lognormal(np.log(16000), 0.8))
    o["prog"] = round(prog, 1) if rng.random() < 0.8 else float(round(prog))
    if o["apont"] == "Em produção":
        frac = rng.uniform(0.85, 1.05) if o["tem_fim"] else rng.uniform(0.2, 0.7)
        o["forn"] = int(round(prog * frac))
    else:
        o["forn"] = 0
    # status auxiliares
    if o["apont"] == "Não produziu":
        o["st_proc"], o["st_prod"] = "Pendente", "Pendente"
        o["st_horas"] = "Pendente" if not o["tem_ini"] else "Parcial"
    else:
        o["st_proc"] = "Concluído" if o["tem_fim"] else "Pendente"
        o["st_prod"] = pick(["Concluído", "Parcial"], [0.61, 0.39]) if o["tem_fim"] else "Parcial"
        o["st_horas"] = pick(["Concluído", "Parcial"], [0.25, 0.75])
    o["st_sis"] = STATUS_SIS[o["geral"]][int(rng.integers(len(STATUS_SIS[o["geral"]])))]

# ---------------------------------------------------------------------- aba BI
def linha_bi(op, item, fase, st_bi, ordem_tipo, produzida):
    prod = float(rng.lognormal(np.log(330), 1.5)) if produzida else 0.0
    cons = perdas = 0.0
    if produzida:
        if rng.random() > 0.08:                          # 92% com consumo apontado
            cons = prod * (1 + PERDA_FASE[fase] * rng.lognormal(0, 0.35))
            real = cons - prod
            r = rng.random()
            if r < 0.72:   perdas = max(0.0, real * (1 + rng.normal(0, 0.08)))
            elif r < 0.92: perdas = real + rng.uniform(0.12, 0.40) * prod      # perda apontada acima do balanço
            else:          perdas = 0.0
        else:
            perdas = float(rng.uniform(0, 30))
    mp = cons * 0.88
    quim = cons * 0.25
    quim_sol = cons * 0.12
    cons_sol = mp + quim_sol
    abs_bm = cons_sol - prod
    perc = (abs_bm / cons_sol) if cons_sol > 0 else None
    return [op, item["codigo"], item["desc"][:18], item["tipo_mat"], st_bi, ordem_tipo,
            fase, item["familia"], round(perdas, 2), round(prod, 2), round(mp, 2),
            round(quim, 2), round(quim_sol, 2), round(cons_sol, 2), round(abs_bm, 2),
            None if perc is None else round(perc, 4)]

bi_por_op = {}
for o in ordens:
    produziu = o["apont"] == "Em produção"
    extra = (not produziu) and rng.random() < 0.035          # algumas ordens sem produção também aparecem
    if produziu or extra:
        fase = pick(FASES, P_FASE)
        st_bi = "ABERTO" if o["geral"] == "Aberta" else ("LIBERADO" if o["lib"] else "ENCERRADO")
        bi_por_op[o["op"]] = linha_bi(o["op"], o["item"], fase, st_bi, o["tipo_op"][:4], produziu)
        o["fase"] = fase

extras_bi = []
for _ in range(46):
    op = new_op(70_000_000, 70_999_999)                      # faixa abaixo => é a menor, descartada pelo Skip(1) do Power Query
    it = CATALOGO[int(rng.integers(len(CATALOGO)))]
    fase = pick(FASES, P_FASE)
    extras_bi.append(linha_bi(op, it, fase, "LIBERADO", pick(["ZPP7", "ZPP8", "ZP11", "ZPP9"], [0.5, 0.33, 0.1, 0.07]),
                              rng.random() < 0.6))
BI_HDR = ["GRUPO_ORDEM", "COD_SISTEMA", "DESCRICAO", "TIPO_MATERIAL", "DESCRICAO", "ORDEM_TIPO", "FASE",
          "FAMILIA", "Kgs Perdas", "Kgs Producao", "Kgs Consumo MP", "Kgs Consumo Quim",
          "Kgs Quim Solido", "Consumo Solido KG", "Perdas BM Abs", "Perdas BM Perc"]
bi_rows = list(bi_por_op.values()) + extras_bi
bi_rows.sort(key=lambda r: r[0])

# -------------------------------------------------------------------- aba COOIS
COOIS_HDR = ["OP", "Código", "Item", "Tipo de OP", "Status geral", "Status - apontamento",
             "Status - processo", "Status - produção", "Status - horas", "Status do sistema",
             "Unidade", "Quantidade programada", "Quantidade fornecida", "Dt abertura",
             "Dt liberação", "Dt início", "Dt fim", "Dt última modificação",
             "Aging início de produção", "Aging fim de produção", "Faixa aging início",
             "Faixa aging fim", "Aging encerramento início", "Aging encerramento fim", "Fase",
             "Perda apontada", "Produção", "Consumo ", "Perda BM", "BM (%)", "Diferença (%)",
             "Status Balanceamento", "-"]
coois_rows = []
for o in ordens:
    ag_ini = (REF - o["ini"]).days if o["ini"] else "-"
    ag_fim = (REF - o["fim"]).days if o["fim"] else "-"
    f_ini = faixa(ag_ini) if o["ini"] else "Sem movimento"
    f_fim = faixa(ag_fim) if o["fim"] else "Sem movimento"
    enc_ini = enc_fim = "-"
    if o["geral"] == "Encerrada" and o["ini"] and o["mod"]:
        enc_ini = (o["mod"] - o["ini"]).days
    if o["geral"] == "Encerrada" and o["fim"] and o["mod"]:
        v = (o["mod"] - o["fim"]).days
        enc_fim = v if v <= 100 else "-"
    bi = bi_por_op.get(o["op"])
    if bi:
        fase, perda_ap, prod, cons = bi[6], bi[8], bi[9], bi[13]
        perda_bm = round(cons - prod, 2)
        bm = (perda_bm / cons) if cons else 0
        dif = ((perda_bm - perda_ap) / prod) if prod else 0
        st_bal = "OK" if abs(dif) <= TOL else "NOK"
    else:
        fase, perda_ap, prod, cons, perda_bm, bm, dif, st_bal = "-", "-", "-", "-", "-", 0, 0, "OK"
    coois_rows.append([
        o["op"], str(o["item"]["codigo"]), o["item"]["desc"], o["tipo_op"], o["geral"], o["apont"],
        o["st_proc"], o["st_prod"], o["st_horas"], o["st_sis"], o["unidade"], o["prog"], o["forn"],
        o["abr"], o["lib"], o["ini"], o["fim"], o["mod"], ag_ini, ag_fim, f_ini, f_fim,
        enc_ini, enc_fim, fase, perda_ap, prod, cons, perda_bm,
        round(bm, 4), round(dif, 4), st_bal, None])

# ------------------------------------------------------------- Ordens pendentes
PEND_HDR = ["OP", "Código", "Data da última produção", "Dias em atraso", "Item", "Status",
            "Pendencia", "Responsável", None]
pend_rows = []
em_coois = [o for o in ordens if o["geral"] != "Aberta"]
ops_pend = [o["op"] for o in rng.choice(em_coois, 14, replace=False)]
pend_items = {o["op"]: o["item"] for o in ordens}
for op in ops_pend + [new_op(72_000_000, 72_999_999) for _ in range(731)]:
    it = pend_items.get(op) or CATALOGO[int(rng.integers(len(CATALOGO)))]
    dia = rand_date(dt.date(2025, 7, 16), dt.date(2025, 12, 23))
    stat = pick(["Resolvido", "aguardando retorno", None], [0.965, 0.032, 0.003])
    pend_rows.append([op, it["codigo"], dia, (REF - dia).days, it["desc"], stat,
                      pick(["Revisão", "Embalar", "Refugar ou paletizar"], [0.47, 0.39, 0.14]),
                      "Produção", None])
rng.shuffle(pend_rows)

nao_enc = [r for r in coois_rows if r[4] != "Encerrada" and isinstance(r[26], (int, float)) and r[26] != "-"]
kg_cons = round(sum(r[27] for r in nao_enc), 2)
kg_prod = round(sum(r[26] for r in nao_enc), 2)
KG_ROWS = [["Total kgs Pendentes ", "QTD (KG)"], ["Consumo", kg_cons], ["Produção ", kg_prod]]

# --------------------------------------------------------------------- Consolidado
meses = [dt.date(2024 + (7 + i) // 12, (7 + i) % 12 + 1, 1) for i in range(18)]
tend = np.linspace(0.64, 0.93, 18)
cons_rows = []
for i, m in enumerate(meses):
    if i < 17:
        cons_rows.append([m, int(rng.integers(1000, 1500)), int(rng.integers(1100, 1900)),
                          int(rng.integers(400, 900)), int(rng.integers(1000, 1400)),
                          int(max(2, 580 - i * 33 + rng.normal(0, 25))),
                          round(float(np.clip(tend[i] + rng.normal(0, 0.02), 0.55, 0.97)), 4)])
    else:
        ab = sum(r[4] == "Aberta" for r in coois_rows)
        en = sum(r[4] == "Encerrada" for r in coois_rows)
        pc = sum(r[4] == "Processo concluído" for r in coois_rows)
        sm = sum(r[20] == "Sem movimento" for r in coois_rows)
        cons_rows.append([m, ab, en, pc, sm, 5, 0.9312])
CONS_HDR = ["Mês", "Abertas", "Encerradas", "Processo concluido", "Sem movimento ",
            "Qtd de ordens abertas acima de 07 dias", "Assertividade "]

# ----------------------------------------------------------------------- Máquinas
MAQ_HDR = ["Maq", "Data", "Turno", "Ordem", "CodItem", "Desc. Item", "Cor", "Qtde. Ordem",
           "Unid_1", "Qtde. Kilos", "Unid_2", "Qtde. Apontada", "Unid_3", "Aparas/Refugos",
           "Perc (%)", "Rp.", "Qtde. Gerencial", "Unid", "Hs Prod", "Hs Parada", "Hs Disp.", "Hs N/G"]
maq_rows = []

def linhas_maq(op, item, fase, prog, kilos_total, d0, unidade):
    maq = MAQ[fase][int(rng.integers(len(MAQ[fase])))]
    n = int(rng.integers(1, 4))
    for s in range(n):
        kilos = max(1.0, kilos_total / n)
        perc = max(0.05, rng.normal(PERDA_FASE[fase] * 100, 1.5))
        aparas = round(kilos * perc / 100, 1)
        maq_rows.append([maq, d0 + dt.timedelta(days=s // 3), int(rng.integers(1, 4)),
                         f"{op} {(s + 1) * 10:04d}", item["codigo"], item["desc"], None,
                         round(prog, 1), "KG", round(kilos, 1), "KG",
                         int(round(kilos * (1 - perc / 100))),
                         unidade, aparas, round(perc, 2), int(rng.integers(1, 4)),
                         round(prog / n, 1), unidade, 0, 0, 0, 0])

for o in ordens:
    if o["apont"] == "Em produção":
        bi = bi_por_op[o["op"]]
        linhas_maq(o["op"], o["item"], bi[6], o["prog"], max(bi[9], 5.0),
                   o["ini"] or o["abr"], pick(["M", "TS", "MIL", "KG"], [0.70, 0.20, 0.08, 0.02]))
for _ in range(2600):
    op = new_op(73_000_000, 73_999_999)
    it = CATALOGO[int(rng.integers(len(CATALOGO)))]
    fase = pick(FASES, P_FASE)
    linhas_maq(op, it, fase, float(rng.lognormal(np.log(16000), 0.8)),
               float(rng.lognormal(np.log(260), 1.2)),
               rand_date(dt.date(2025, 1, 2), dt.date(2026, 1, 7)), "M")

# ------------------------------------------------------------------------ gravação
DATE_FMT = "yyyy-mm-dd"

def gravar(nome_arquivo, abas):
    """abas: lista de (nome, cabeçalho, linhas, {coluna: formato})"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for nome, hdr, linhas, fmts in abas:
        ws = wb.create_sheet(nome)
        ws.append(hdr)
        for r in linhas:
            ws.append(r)
        for col, fmt in (fmts or {}).items():
            idx = hdr.index(col) + 1
            for cell in ws.iter_cols(min_col=idx, max_col=idx, min_row=2):
                for c in cell:
                    c.number_format = fmt
    destino = os.path.join(OUT, nome_arquivo)
    if nome_arquivo.endswith(".xlsm"):
        tmp = destino + ".tmp.xlsx"
        wb.save(tmp)
        # mesmo nome/extensão da planilha original, sem macros: ajusta o tipo do pacote
        with zipfile.ZipFile(tmp) as zin, zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                dados = zin.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    dados = dados.replace(
                        b"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                        b"application/vnd.ms-excel.sheet.macroEnabled.main+xml")
                zout.writestr(item, dados)
        os.remove(tmp)
    else:
        wb.save(destino)

datas_coois = {c: DATE_FMT for c in ["Dt abertura", "Dt liberação", "Dt início", "Dt fim", "Dt última modificação"]}
pct_coois = {"BM (%)": "0.00%", "Diferença (%)": "0.00%"}
gravar("Controle de encerramento de OPs.xlsm", [
    ("COOIS", COOIS_HDR, coois_rows, {**datas_coois, **pct_coois}),
    ("BI", BI_HDR, bi_rows, {}),
])
gravar("Controle de produção_Maq.xlsx", [("Máquinas ", MAQ_HDR, maq_rows, {"Data": DATE_FMT})])
gravar("Ordens pendentes.xlsx", [
    ("Pendencia", PEND_HDR, pend_rows, {"Data da última produção": DATE_FMT}),
    ("kg total pendente ", KG_ROWS[0], KG_ROWS[1:], {}),
])
gravar("Controle de encerramento Consolidado.xlsx", [
    ("Ence. OPs (Consolidado)", CONS_HDR, cons_rows, {"Mês": DATE_FMT, "Assertividade ": "0.0%"})])

print(f"COOIS: {len(coois_rows)} ordens | BI: {len(bi_rows)} | Pendências: {len(pend_rows)} | "
      f"Máquinas: {len(maq_rows)} linhas | Consolidado: {len(cons_rows)} meses")
