# -*- coding: utf-8 -*-
"""Constrói as linhas de todas as abas de apoio."""
import math
from collections import defaultdict
from motor import *
from agrega import *

# ------------------------------------------------------------- BASE TRATADA

H_BASE_TRATADA = [
    "INEP", "Escola", "Oferta EM 2026", "1ª Série 2026", "2ª Série 2027",
    "3ª Série 2027", "Oferta EJA 2026", "Turmas EJA", "Oferta Fundamental",
    "Parciais 2026", "Possíveis Fusões", "Salas Necessárias 2027", "Crescimento",
    "Resumo Hoje", "Resumo 2027", "Matrículas e Turmas EJA",
    # ---- colunas novas (V2 / V3) ----
    "9º Ano nesta escola · turmas", "9º Ano nesta escola · matrículas",
    "UETEP · novas matrículas 2027", "UETEP · turmas previstas",
    "UETEP · detalhe por curso",
    "EJA fora do prédio matriz", "EJA · turmas em anexo/sala externa",
    "EJA · matrículas em anexo/sala externa",
    "Fusão entre escolas (mesmo município)",
    "Anexos desta escola", "Qtd. de anexos",
    "Salas necessárias 2027 (com UETEP)", "Município", "GRE",
]

def base_tratada(matrizes, ordem_m, anexos, ordem_a, uetep_idx, ueja_resumo, fusao_resumo):
    anexos_por_inep = defaultdict(list)
    for ch in ordem_a:
        a = anexos[ch]
        anexos_por_inep[a.inep].append(a)

    linhas = []
    for i in ordem_m:
        e = matrizes[i]
        u = uetep_idx.get(i, {})
        u_turmas = u.get("turmas_previstas", 0)
        p_sem = projetar_2027(e, 0)
        p_com = projetar_2027(e, u_turmas)

        oferta_em  = montar_oferta_em(e)
        oferta_eja = montar_oferta_eja(e)
        ef_texto   = "\n".join(sorted(e.efLinhas)) if e.efLinhas else "—"
        parciais   = "\n".join(e.parciais) if e.parciais else "—"
        fusoes     = montar_fusoes(e)
        hoje       = resumo_hoje(e)

        turmas_eja = ("%d turmas · %d matrículas" % (e.totalEJA, e.totalEJAentm)
                      if e.totalEJA > 0 else "—")
        resumo_eja = ("%d matrículas | %d turmas" % (e.totalEJAentm, e.totalEJA)
                      if e.totalEJA > 0 else "—")

        resumo27 = ("Turmas EM 2027 (base): %d de 1ª, %d de 2ª, %d de 3ª%s. "
                    "Salas estimadas na MATRIZ: %d. EJA não entra neste cálculo."
                    % (p_sem["pro1"], p_sem["pro2"], p_sem["pro3"],
                       (", %d outra(s) turma(s) EM" % e.outrosEM) if e.outrosEM > 0 else "",
                       p_sem["salasNec"]))

        lista_anexos = anexos_por_inep.get(i, [])
        txt_anexos = "\n".join(
            "• %s [%s] — %d turma(s) · %d matrícula(s)"
            % (a.anexo, ROTULO_LOCAL.get(a.tipoAnexo, a.tipoAnexo), a.totalTurmas, a.totalMatriculas)
            for a in lista_anexos) or "—"

        linhas.append([
            int(e.inep), e.escola, oferta_em, e.s1, p_sem["pro2"], p_sem["pro3"],
            oferta_eja, turmas_eja, ef_texto, parciais, fusoes,
            p_sem["salasNec"], p_sem["delta"], hoje, resumo27, resumo_eja,
            int(e.ef9Turmas), int(e.ef9Matriculas),
            int(u.get("pre", 0)), int(u_turmas), u.get("resumo", "—") or "—",
            ueja_resumo.get(i, "Sem oferta de EJA"),
            int(e.ejaAnexoTurmas), int(e.ejaAnexoMatriculas),
            fusao_resumo.get(i, "—") or "—",
            txt_anexos, len(lista_anexos),
            p_com["salasNec"], e.municipio, e.gre,
        ])
    return linhas

# --------------------------------------------------------------- BASE ANEXOS

H_BASE_ANEXOS = [
    "INEP", "GRE", "Município", "Escola Matriz", "Anexo / Sala Externa",
    "Natureza do anexo", "Oferta EM 2026", "1ª Série 2026", "2ª Série 2026",
    "3ª Série 2026", "Oferta Fundamental", "Oferta EJA 2026",
    "Matrículas / Turmas EJA", "Parciais 2026",
    "Possíveis Fusões no próprio anexo", "Matrículas totais", "Turmas totais",
    "Resumo Hoje", "Salas estimadas 2027 (sem EJA)",
    # ---- decisões (V3) ----
    "★ MANTER O ANEXO?", "★ ENCERRAR O ANEXO?",
    "★ REMANEJAR A OFERTA DO ANEXO?", "★ Para onde vai a oferta",
    "★ Justificativa da decisão",
]

def base_anexos(anexos, ordem_a):
    linhas = []
    for ch in ordem_a:
        e = anexos[ch]
        p = projetar_2027(e, 0)
        linhas.append([
            int(e.inep), e.gre, e.municipio, e.escola, e.anexo,
            ROTULO_LOCAL.get(e.tipoAnexo, e.tipoAnexo),
            montar_oferta_em(e), e.s1, e.s2, e.s3,
            "\n".join(sorted(e.efLinhas)) if e.efLinhas else "—",
            montar_oferta_eja(e),
            ("%d matrículas | %d turmas" % (e.totalEJAentm, e.totalEJA)) if e.totalEJA > 0 else "—",
            "\n".join(e.parciais) if e.parciais else "—",
            montar_fusoes(e), int(e.totalMatriculas), int(e.totalTurmas),
            resumo_hoje(e), p["salasNec"],
            "", "", "", "", "",
        ])
    return linhas

# ------------------------------------------------------- FUSÃO ENTRE ESCOLAS

H_FUSAO_ESCOLAS = [
    "★ Aplicar?", "Município", "GRE", "Curso", "Série 2027",
    "INEP Origem", "Escola Origem", "Turmas na origem", "Matrículas na origem",
    "Fundamental na origem (turmas)", "9º ano na origem (turmas)",
    "9º ano na origem (matrículas)", "IDEB do município",
    "INEP Destino", "Escola Destino", "Turmas no destino", "Matrículas no destino",
    "Fundamental no destino (turmas)", "Salas do destino",
    "Alunos somados no município", "Turmas hoje no município",
    "Turmas mínimas necessárias", "Turmas que podem ser liberadas", "Tipo",
    "★ Justificativa / decisão",
]

# ------------------------------------------------------------------ UETEP

H_UETEP = [
    "INEP", "GRE", "Município", "Escola", "Etapa", "Curso", "Turma",
    "Turno", "Pré-matrículas (oferta nova)", "Enturmados hoje", "Cursando",
    "Salas previstas para esta oferta",
]

# ------------------------------------------------------------------- UEJA

H_UEJA = [
    "INEP", "GRE", "Município", "Escola", "Onde a EJA acontece",
    "Nome do anexo / sala externa", "Turmas", "Matrículas",
    "Cursos ofertados neste local", "Fora do prédio matriz?",
    "★ Decisão 2027",
]

# ------------------------------------------------------- PANORAMA MUNICIPAL

H_PANORAMA = [
    "Município", "9º Ano · todas as redes (Censo)", "Turmas necessárias (referência)",
    "9º Ano na rede estadual (2026)", "9º Ano estadual · turmas",
    "9º Ano · outras redes (estimado)", "Demanda total de 1ª série 2027 (alunos)",
    "Turmas necessárias 2027 (recalculado)", "Fonte do 9º Ano",
    "1ª série decidida — V2", "Saldo — V2",
    "1ª série decidida — V3", "Saldo — V3", "Cobertura — V3",
]

# ------------------------------------------------------------------- IDEB

H_IDEB = ["Município", "IDEB Anos Finais (rede municipal)",
          "IDEB Anos Finais (rede estadual)", "IDEB Ensino Médio",
          "Ano de referência", "Observação"]


# ─────────────────────────────────────────── BASE TRATADA COMPLETA (44 colunas)

H_BT44 = H_BASE_TRATADA[:18] + [
    "Oferta 2027 · 1ª série (alunos previstos)", "Oferta 2027 · 1ª série (turmas)",
    "Oferta 2027 · 1ª série (cursos)"] + H_BASE_TRATADA[21:28] + [
    "Salas necessárias 2027 (oferta real)"] + H_BASE_TRATADA[29:] + [
    "Oferta 2027 · 2ª série (turmas)", "Oferta 2027 · 2ª série (alunos)",
    "Oferta 2027 · 3ª série (turmas)", "Oferta 2027 · 3ª série (alunos)",
    "Oferta 2027 · subsequente (turmas)", "Oferta 2027 · subsequente (alunos)",
    "Oferta 2027 · 2ª e 3ª série (cursos)", "Salas 2027 · composição",
    "Oferta 2027 · tem oferta?",
    "Fundamental 2026 · turmas", "Só EJA (CEJA)?",
    "EJA 2026 · turmas no prédio matriz", "EJA 2026 · matrículas no prédio matriz",
    "Turmas zeradas descartadas"]


def base_tratada_44(matrizes, ordem_m, anexos, ordem_a, oferta, ueja_resumo,
                    fusao_resumo, salas_fn):
    """As 44 colunas da Base Tratada, na mesma ordem que o Apps Script grava."""
    por_inep = defaultdict(list)
    for ch in ordem_a:
        por_inep[anexos[ch].inep].append(anexos[ch])

    linhas = []
    for i in ordem_m:
        e = matrizes[i]
        of = oferta.get(i)
        p = projetar_2027(e, 0)
        s = salas_fn(i)
        comp = ("integral %d + máx(manhã %d, tarde %d) + noite %d"
                % (s["integral"], s["manha"], s["tarde"], s["noite"]))

        lista = por_inep.get(i, [])
        txt_anexos = "\n".join(
            "• %s [%s] — %d turma(s) · %d matrícula(s)"
            % (a.anexo, ROTULO_LOCAL.get(a.tipoAnexo, a.tipoAnexo),
               a.totalTurmas, a.totalMatriculas) for a in lista) or "—"

        resumo27 = ("Turmas EM 2027 (base): %d de 1ª, %d de 2ª, %d de 3ª%s. "
                    "Salas estimadas na MATRIZ: %d. EJA não entra neste cálculo."
                    % (p["pro1"], p["pro2"], p["pro3"],
                       (", %d outra(s) turma(s) EM" % e.outrosEM) if e.outrosEM > 0 else "",
                       p["salasNec"]))

        linhas.append([
            int(e.inep), e.escola, montar_oferta_em(e), e.s1, p["pro2"], p["pro3"],
            montar_oferta_eja(e),
            ("%d turmas · %d matrículas" % (e.ejaMatrizTurmas, e.ejaMatrizEntm))
            if e.ejaMatrizTurmas > 0 else "—",
            "\n".join(sorted(e.efLinhas)) if e.efLinhas else "—",
            "\n".join(e.parciais) if e.parciais else "—",
            montar_fusoes(e), p["salasNec"], p["delta"], resumo_hoje(e), resumo27,
            ("%d matrículas | %d turmas" % (e.ejaMatrizEntm, e.ejaMatrizTurmas))
            if e.ejaMatrizTurmas > 0 else "—",
            int(e.ef9Turmas), int(e.ef9Matriculas),
            int(of["of1_alunos"]) if of else 0,
            int(of["of1_turmas"]) if of else 0,
            of["of1_detalhe"] if of else "— sem oferta de EM em 2027",
            ueja_resumo.get(i, "Sem oferta de EJA"),
            int(e.ejaAnexoTurmas), int(e.ejaAnexoMatriculas),
            fusao_resumo.get(i, "—") or "—",
            txt_anexos, len(lista),
            s["salas"], norm(e.municipio), e.gre,
            int(of["co2_turmas"]) if of else 0, int(of["co2_alunos"]) if of else 0,
            int(of["co3_turmas"]) if of else 0, int(of["co3_alunos"]) if of else 0,
            int(of["sub_turmas"]) if of else 0, int(of["sub_alunos"]) if of else 0,
            (of["co2_detalhe"] + " || " + of["co3_detalhe"]) if of else "—",
            comp, "SIM" if of else "NÃO",
            int(e.efTotal), "SIM" if so_eja(e) else "NÃO",
            int(e.ejaMatrizTurmas), int(e.ejaMatrizEntm), int(e.turmasZeradas),
        ])
    return linhas
