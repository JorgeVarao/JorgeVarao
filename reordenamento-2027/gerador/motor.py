# -*- coding: utf-8 -*-
"""
MOTOR DE CÁLCULO — Reordenamento 2027 (SEDUC-PI / SUPEX / UGERF / GDI)

Espelha em Python a lógica do Apps Script para que o .xlsx entregue já saia
com os valores calculados. As mesmas regras estão no arquivo .gs.
"""
import re, math, unicodedata
from collections import defaultdict

MAX_T = 40          # capacidade máxima por turma
CAP_1SERIE = 40     # capacidade usada para converter alunos de 9º ano em turmas

# ---------------------------------------------------------------- utilidades

def texto_(v):
    if v is None: return ""
    if isinstance(v, float) and v == int(v): return str(int(v))
    return str(v).strip()

def numero_(v):
    try:
        n = float(v)
        return 0.0 if math.isnan(n) else n
    except (TypeError, ValueError):
        return 0.0

def inteiro_(v):
    return int(round(numero_(v)))

def norm(s):
    s = texto_(s).upper()
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn")

def esp_(v):
    return re.sub(r"\s+", " ", texto_(v))

def p2_(n):
    n = int(n)
    return ("0" if n < 10 else "") + str(n)

def abrev_turno(t):
    t = texto_(t).upper()
    return {"INTEGRAL": "I", "MANHÃ": "M", "MANHA": "M",
            "TARDE": "T", "NOITE": "N"}.get(t, t[:1])

def suf_turno(t):
    t = texto_(t).upper()
    return {"INTEGRAL": "I", "MANHÃ": "M", "MANHA": "M",
            "TARDE": "T", "NOITE": "N"}.get(t, "")

def inep_(v):
    t = texto_(v)
    if not t: return ""
    try: return str(int(float(t)))
    except ValueError: return t

# ------------------------------------------------- classificação de anexos

def classificar_anexo(nome):
    """Diz de que natureza é o anexo / sala externa. Usado pela UEJA."""
    n = norm(nome)
    if not n: return "MATRIZ"
    if any(k in n for k in ("SISTEMA PRISIONAL", "PENITENCIARIA", "PRESIDIO",
                            "CADEIA", "CUSTODIA", "FEMININA")):
        return "PRIVACAO DE LIBERDADE"
    if any(k in n for k in ("SOCIO EDUCATIVO", "SOCIOEDUCATIVO", "CEIP", "CEM ", "CEF ")):
        return "SOCIOEDUCATIVO"
    if any(k in n for k in ("SALA EXTERNA", "TURMA EXTERNA", "EXTERNA")):
        return "SALA EXTERNA"
    if any(k in n for k in ("ESCOLA ", "ESC.", "E. M.", "E.M.", "U. E.", "U.E",
                            "U E ", "UEMA", "ESC ", "U. MARCELINA")):
        return "CEDIDA EM OUTRA ESCOLA"
    if any(k in n for k in ("POVOADO", "POV.", "ASSENTAMENTO", "LOCALIDADE",
                            "QUILOMBO", "COMUNIDADE", "ASSOCIACAO", "SITIO")):
        return "COMUNIDADE / POVOADO"
    return "ANEXO"

# ------------------------------------------------------ estrutura da unidade

CAMPOS_ZERO = """totalTurmas totalMatriculas emTotal efTotal totalEJA totalEJAentm
nI nM nT nN efI efM efT efN s1 s2 s3
s1I s1M s1T s1N s2I s2M s2T s2N s3I s3M s3T s3N
outrosEM outrosI outrosM outrosT outrosN ejI ejM ejT ejN
ef9Turmas ef9Matriculas ef8Turmas ef8Matriculas
aeeTurmas aeeMatriculas ejaAnexoTurmas ejaAnexoMatriculas
turmasZeradas ejaMatrizTurmas ejaMatrizEntm
ejI ejM ejT ejN ejMatrizI ejMatrizM ejMatrizT ejMatrizN""".split()

class Unidade(object):
    def __init__(self, inep, escola, anexo, meta):
        self.inep, self.escola, self.anexo = inep, escola, anexo
        self.gre = meta.get("gre", "")
        self.municipio = meta.get("municipio", "")
        self.salasGDI = meta.get("salas", "")
        self.escolaProxima = meta.get("escolaProxima", "")
        self.tipoAnexo = classificar_anexo(anexo)
        self.emLinhas, self.ejaLinhas, self.efLinhas = [], [], []
        self.ejaMatrizLinhas, self.linhasZeradas = [], []
        self.outrasLinhas, self.parciais = [], []
        self.cursos = {}          # fusão dentro da própria unidade
        self.cursos1a = {}        # cursos da 1ª série (para fusão entre escolas)
        self.ejaPorLocal = {}     # UEJA: matriz x anexo/sala externa
        for c in CAMPOS_ZERO: setattr(self, c, 0)

    def soma_turno(self, prefixo, turno, qtd):
        suf = suf_turno(turno)
        if not suf: return
        campo = prefixo + suf
        setattr(self, campo, getattr(self, campo, 0) + qtd)

# ------------------------------------------------------------ acumular turma

RE_EF   = re.compile(r"EF\s*-\s*\d+º", re.I)
RE_ANO  = re.compile(r"\d+º\s*ANO", re.I)
RE_AEE  = re.compile(r"(^|\s)AEE(\s|$)", re.I)
RE_S1   = re.compile(r"1ª\s*S[ÉE]RIE", re.I)
RE_S2   = re.compile(r"2ª\s*S[ÉE]RIE", re.I)
RE_S3   = re.compile(r"3ª\s*S[ÉE]RIE", re.I)
RE_EF9  = re.compile(r"EF\s*-\s*9º", re.I)
RE_EF8  = re.compile(r"EF\s*-\s*8º", re.I)

def acumular_turma(e, r, anexo=""):
    """r = linha da aba Turmas: 0 INEP 1 Escola 2 Anexo 3 Curso 4 Etapa
           5 Organização 6 Período 7 Turno 8 Enturmados 9 Turmas"""
    curso   = texto_(r[3]); etapa = texto_(r[4])
    organiz = texto_(r[5]) if len(r) > 5 else ""
    periodo = texto_(r[6]) if len(r) > 6 else ""
    turno   = texto_(r[7]) if len(r) > 7 else ""
    entm    = numero_(r[8] if len(r) > 8 else 0)
    turmas  = numero_(r[9] if len(r) > 9 else 0)
    if turmas == 0 and entm == 0 and not curso and not etapa: return

    # Turma declarada sem nenhuma matrícula não é turma: não aparece na oferta
    # nem entra em contagem alguma. Fica registrada só para conferência.
    if entm == 0 and turmas > 0:
        e.turmasZeradas += turmas
        e.linhasZeradas.append("%s - %s (%s) - %d turma(s) sem matrícula"
                               % (etapa, curso, turno, turmas))
        return

    cU, eU, tU = curso.upper(), etapa.upper(), turno.upper()
    ehEJA = ("EJA" in eU) or ("EJA" in cU)
    ehAEE = bool(RE_AEE.search(etapa) or RE_AEE.search(curso))
    ehEF  = False
    if not ehEJA:
        if "ANOS FINAIS" in cU or "ANOS INICIAIS" in cU: ehEF = True
        if RE_EF.search(etapa) or RE_ANO.search(etapa):  ehEF = True

    ab = abrev_turno(tU)
    txt = ("• %s - %s - %d matrículas %s - %d %s"
           % (etapa, curso, entm, ab, turmas, "turma" if turmas == 1 else "turmas"))

    e.totalTurmas += turmas
    e.totalMatriculas += entm

    # ---------------------------------------------------------------- EJA
    if ehEJA:
        local = classificar_anexo(anexo)
        e.ejaLinhas.append(txt)
        e.totalEJA += turmas
        e.totalEJAentm += entm
        e.soma_turno("ej", tU, turmas)
        # a oferta do prédio matriz é contada à parte da que está em anexo
        if local == "MATRIZ":
            e.ejaMatrizLinhas.append(txt)
            e.ejaMatrizTurmas += turmas
            e.ejaMatrizEntm += entm
            e.soma_turno("ejMatriz", tU, turmas)
        d = e.ejaPorLocal.setdefault(local, {"turmas": 0, "entm": 0, "cursos": {}, "nomes": set()})
        d["turmas"] += turmas; d["entm"] += entm
        d["cursos"][curso] = d["cursos"].get(curso, 0) + turmas
        if anexo: d["nomes"].add(esp_(anexo).upper())
        if local != "MATRIZ":
            e.ejaAnexoTurmas += turmas; e.ejaAnexoMatriculas += entm
        return

    # ---------------------------------------------------------------- AEE
    if ehAEE:
        e.outrasLinhas.append(txt)
        e.aeeTurmas += turmas; e.aeeMatriculas += entm
        return

    # -------------------------------------------------------- FUNDAMENTAL
    if ehEF:
        e.efLinhas.append(txt)
        e.efTotal += turmas
        e.soma_turno("ef", tU, turmas)
        if RE_EF9.search(etapa) or "9º ANO" in eU:
            e.ef9Turmas += turmas; e.ef9Matriculas += entm
        if RE_EF8.search(etapa) or "8º ANO" in eU:
            e.ef8Turmas += turmas; e.ef8Matriculas += entm
        return

    # ------------------------------------------------- ENSINO MÉDIO / EPT
    e.emLinhas.append(txt)
    e.emTotal += turmas
    e.soma_turno("n", tU, turmas)

    if "PARCIAL" in cU:
        e.parciais.append("%s | %s | %dT (%d al.)" % (etapa, turno, turmas, entm))

    ehSerie = False
    if RE_S1.search(etapa):
        e.s1 += turmas; e.soma_turno("s1", tU, turmas); ehSerie = True
        ck = curso
        c1 = e.cursos1a.setdefault(ck, {"curso": curso, "turmas": 0, "entm": 0, "turno": turno})
        c1["turmas"] += turmas; c1["entm"] += entm
    elif RE_S2.search(etapa):
        e.s2 += turmas; e.soma_turno("s2", tU, turmas); ehSerie = True
    elif RE_S3.search(etapa):
        e.s3 += turmas; e.soma_turno("s3", tU, turmas); ehSerie = True

    if not ehSerie:
        e.outrosEM += turmas
        e.soma_turno("outros", tU, turmas)

    chave = "|".join([curso, etapa, tU, organiz, periodo])
    d = e.cursos.setdefault(chave, {"curso": curso, "etapa": etapa, "turno": turno,
                                    "organizacao": organiz, "periodo": periodo,
                                    "turmas": 0, "entm": 0})
    d["turmas"] += turmas; d["entm"] += entm

# ------------------------------------------------------------- projeção 2027

def so_eja(e):
    """Unidade cuja oferta inteira é EJA — uma CEJA. Aí a EJA define a sala."""
    return e.totalEJA > 0 and e.emTotal == 0 and e.efTotal == 0


def projetar_2027(e, uetep_turmas=0):
    """1ª 2027 = 1ª atual · 2ª 2027 = 1ª atual · 3ª 2027 = 2ª atual.
       Sala: integral ocupa o dia todo, manhã e tarde dividem, noite à parte.
       EJA fica de fora. UETEP entra como turma nova de 1ª série (integral)."""
    pro1, pro2, pro3 = e.s1, e.s1, e.s2
    i27 = e.s1I + e.s1I + e.s2I + e.outrosI + e.efI + uetep_turmas
    m27 = e.s1M + e.s1M + e.s2M + e.outrosM + e.efM
    t27 = e.s1T + e.s1T + e.s2T + e.outrosT + e.efT
    n27 = e.s1N + e.s1N + e.s2N + e.outrosN + e.efN
    if so_eja(e):
        # CEJA: não há oferta diurna para dividir sala, então a EJA é a sala.
        i27 += e.ejMatrizI
        m27 += e.ejMatrizM
        t27 += e.ejMatrizT
        n27 += e.ejMatrizN
    salas = i27 + max(m27, t27) + n27
    total27 = pro1 + pro2 + pro3 + e.outrosEM + uetep_turmas
    return {"pro1": pro1, "pro2": pro2, "pro3": pro3,
            "integral": i27, "manha": m27, "tarde": t27, "noite": n27,
            "salasNec": int(salas), "delta": int(total27 - e.emTotal),
            "uetep": uetep_turmas}

# ------------------------------------------------------------ textos de saída

def montar_oferta_em(e):
    if not e.emLinhas: return "—"
    res = []
    for q, r in ((e.nI, "Integral"), (e.nM, "Manhã"), (e.nT, "Tarde"), (e.nN, "Noite")):
        if q > 0: res.append("%d %s" % (q, r))
    return "\n".join(sorted(e.emLinhas)) + "\nTurmas: " + (" | ".join(res) if res else "0")

def montar_oferta_eja(e):
    """Só a EJA do prédio matriz. O que está em anexo aparece na coluna da UEJA."""
    if not e.ejaMatrizLinhas: return "—"
    res = []
    for q, r in ((e.ejMatrizI, "Integral"), (e.ejMatrizM, "Manhã"),
                 (e.ejMatrizT, "Tarde"), (e.ejMatrizN, "Noite")):
        if q > 0: res.append("%d %s" % (q, r))
    return ("\n".join(sorted(e.ejaMatrizLinhas)) + "\nTurmas EJA no prédio matriz: "
            + (" | ".join(res) if res else "0"))

def montar_fusoes(e):
    """Fusão apenas dentro da própria unidade (mesmo curso/etapa/turno/org/período)."""
    out = []
    for d in e.cursos.values():
        if d["turmas"] <= 1 or d["entm"] <= 0: continue
        minimo = max(1, math.ceil(d["entm"] / MAX_T))
        if minimo < d["turmas"]:
            libera = int(d["turmas"] - minimo)
            out.append("%s - %s | %s: %dT / %d alunos → %dT (possível liberar %d %s)"
                       % (d["etapa"], d["curso"], d["turno"], d["turmas"], d["entm"],
                          minimo, libera, "turma" if libera == 1 else "turmas"))
    return "\n".join(out) if out else "—"

def resumo_hoje(e):
    t = []
    intT = e.s1I + e.s2I + e.s3I + e.outrosI
    if intT > 0:
        t.append(p2_(intT) + " TURMA(S) MÉDIO INTEGRAL")
        for q, r in ((e.s1I, "1ª"), (e.s2I, "2ª"), (e.s3I, "3ª")):
            if q: t.append(p2_(q) + " turma(s) - %s série" % r)
    for tot, lbl, trio in ((e.s1M + e.s2M + e.s3M + e.outrosM, "MANHÃ", (e.s1M, e.s2M, e.s3M)),
                           (e.s1T + e.s2T + e.s3T + e.outrosT, "TARDE", (e.s1T, e.s2T, e.s3T))):
        if tot > 0:
            if t: t.append("")
            t.append("%s - %s TURMA(S) MÉDIO" % (lbl, p2_(tot)))
            for q, r in zip(trio, ("1ª", "2ª", "3ª")):
                if q: t.append(p2_(q) + " turma(s) - %s série" % r)
    nEM = e.s1N + e.s2N + e.s3N + e.outrosN
    if nEM > 0:
        if t: t.append("")
        t.append("NOITE - %s TURMA(S) MÉDIO REGULAR" % p2_(nEM))
    if e.efTotal > 0:
        if t: t.append("")
        if e.efI: t.append(p2_(e.efI) + " TURMA(S) FUNDAMENTAL INTEGRAL")
        for q, lbl in ((e.efM, "MANHÃ"), (e.efT, "TARDE"), (e.efN, "NOITE")):
            if q: t.append("%s - %s TURMA(S) FUNDAMENTAL" % (lbl, p2_(q)))
    if e.totalEJA > 0:
        if t: t.append("")
        t.append("EJA — FORA DO CÁLCULO DE SALAS")
        for q, lbl in ((e.ejI, "INTEGRAL"), (e.ejM, "MANHÃ"), (e.ejT, "TARDE"), (e.ejN, "NOITE")):
            if q: t.append("%s - %s TURMA(S) DE EJA" % (lbl, p2_(q)))
    return "\n".join(t) if t else "—"
