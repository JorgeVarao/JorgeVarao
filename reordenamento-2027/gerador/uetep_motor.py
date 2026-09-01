# -*- coding: utf-8 -*-
"""
Lê a OFERTA_2027_BASE.xlsx — a base real da UETEP — e devolve, por escola,
a oferta de 2027 já decidida: 1ª série, continuidade (2ª/3ª) e subsequente.
"""
import math
from collections import defaultdict

def inep_(v):
    try: return str(int(float(v)))
    except (TypeError, ValueError):
        s = "" if v is None else str(v).strip()
        return s

def num_(v):
    try:
        n = float(v)
        return 0.0 if math.isnan(n) else n
    except (TypeError, ValueError):
        return 0.0

def txt_(v):
    return "" if v is None else str(v).strip()


def ler_oferta(linhas_oi, linhas_co, linhas_su):
    """
    OFERTA INTEGRAL 2027   A GRE · B Município · C INEP · D Entidade
                           E TURMAS 2027 · F CURSO · G PREVISÃO DE ALUNOS
    CONTINUIDADE 2027      A GRE · B Município · C INEP · D Entidade
                           E Curso · F Etapa · G Qtd Alunos      (1 linha = 1 turma)
    SUBSEQUENTE 2027       A GRE · B Município · C INEP · D Entidade
                           E PROJEÇÃO 1ª SÉRIE · F CURSO · G EIXO · H ALUNOS

    As linhas de total no rodapé (sem INEP) são descartadas.
    """
    esc = defaultdict(lambda: {
        "of1_turmas": 0, "of1_alunos": 0, "of1_cursos": {},
        "co2_turmas": 0, "co2_alunos": 0, "co2_cursos": {},
        "co3_turmas": 0, "co3_alunos": 0, "co3_cursos": {},
        "sub_turmas": 0, "sub_alunos": 0, "sub_cursos": {},
        "gre": "", "municipio": "", "escola": "",
    })
    descartadas = {"oi": 0, "co": 0, "su": 0}

    for r in linhas_oi[1:]:
        i = inep_(r[2])
        if not i or i == "None": descartadas["oi"] += 1; continue
        e = esc[i]
        e["gre"], e["municipio"], e["escola"] = txt_(r[0]), txt_(r[1]), txt_(r[3])
        t, curso, al = num_(r[4]), txt_(r[5]), num_(r[6])
        e["of1_turmas"] += t; e["of1_alunos"] += al
        e["of1_cursos"][curso] = e["of1_cursos"].get(curso, 0) + t

    for r in linhas_co[1:]:
        i = inep_(r[2])
        if not i or i == "None": descartadas["co"] += 1; continue
        e = esc[i]
        if not e["escola"]:
            e["gre"], e["municipio"], e["escola"] = txt_(r[0]), txt_(r[1]), txt_(r[3])
        curso, etapa, al = txt_(r[4]), txt_(r[5]).upper(), num_(r[6])
        alvo = "co2" if etapa.startswith("2") else "co3"
        e[alvo + "_turmas"] += 1            # uma linha = uma turma
        e[alvo + "_alunos"] += al
        e[alvo + "_cursos"][curso] = e[alvo + "_cursos"].get(curso, 0) + 1

    for r in linhas_su[1:]:
        i = inep_(r[2])
        if not i or i == "None": descartadas["su"] += 1; continue
        e = esc[i]
        if not e["escola"]:
            e["gre"], e["municipio"], e["escola"] = txt_(r[0]), txt_(r[1]), txt_(r[3])
        t, curso, al = num_(r[4]), txt_(r[5]), num_(r[7])
        e["sub_turmas"] += t; e["sub_alunos"] += al
        e["sub_cursos"][curso] = e["sub_cursos"].get(curso, 0) + t

    for i, e in esc.items():
        for p in ("of1", "co2", "co3", "sub"):
            itens = sorted(e[p + "_cursos"].items(), key=lambda x: (-x[1], x[0]))
            e[p + "_detalhe"] = " · ".join("%s (%dT)" % (c, t) for c, t in itens) or "—"
        e["total_turmas"] = e["of1_turmas"] + e["co2_turmas"] + e["co3_turmas"] + e["sub_turmas"]
        e["total_alunos"] = e["co2_alunos"] + e["co3_alunos"]     # alunos reais
    return dict(esc), descartadas


def salas_2027(unidade, of):
    """
    Salas de 2027 com a oferta real.

      integral      = 1ª + 2ª + 3ª + subsequente (a oferta da UETEP é integral)
                      + fundamental integral + outras turmas de EM integrais
      manhã / tarde = fundamental e outras turmas de EM parciais — dividem sala
      noite         = fundamental e EM noturno regular
      EJA           = fora do cálculo

    'unidade' é a escola já lida da aba Turmas (fundamental, outras ofertas
    e turnos de 2026); 'of' é a oferta de 2027 desta escola.
    """
    of = of or {}
    integral = (of.get("of1_turmas", 0) + of.get("co2_turmas", 0)
                + of.get("co3_turmas", 0) + of.get("sub_turmas", 0)
                + unidade.efI + unidade.outrosI)
    manha  = unidade.efM + unidade.outrosM
    tarde  = unidade.efT + unidade.outrosT
    noite  = unidade.efN + unidade.outrosN
    return {
        "integral": int(integral), "manha": int(manha),
        "tarde": int(tarde), "noite": int(noite),
        "salas": int(integral + max(manha, tarde) + noite),
    }
