# -*- coding: utf-8 -*-
"""Layout da Reordenamento 2027 V3 e as listas dos menus com quantidade."""

MAX_QTD = 8      # até quantas turmas o menu oferece por curso / etapa

FUNDAMENTAL = ["EF Inicial - 1º ano", "EF Inicial - 2º ano", "EF Inicial - 3º ano",
               "EF Inicial - 4º ano", "EF Inicial - 5º ano", "EF Final - 6º ano",
               "EF Final - 7º ano", "EF Final - 8º ano", "EF Final - 9º ano"]

DECISAO_ANEXO = [
    "Manter o anexo como está",
    "Manter o anexo, com ajuste de oferta",
    "Encerrar o anexo — oferta vai para o prédio matriz",
    "Encerrar o anexo — oferta vai para outra escola estadual",
    "Encerrar o anexo — oferta vai para a rede municipal",
    "Encerrar o anexo — oferta encerrada",
    "Remanejar a oferta, mantendo o anexo aberto",
    "Em análise",
]

def com_quantidade(itens, maxq=MAX_QTD):
    """'Logística' vira 'Logística (1)' … 'Logística (8)'. Ninguém digita nada."""
    out = []
    for it in itens:
        for q in range(1, maxq + 1):
            out.append("%s (%d)" % (it, q))
    return out

# ─────────────────────────────────────────────── colunas da V3

V3 = [
    ("INEP",                                        "id",      11),
    ("GRE",                                         "id",      13),
    ("Município",                                   "id",      20),
    ("Escola",                                      "id",      30),
    ("Turmas hoje\n(2026)",                         "hoje",    34),
    ("Turmas 2027\n(projeção)",                     "hoje",    34),
    ("Oferta\nFundamental",                         "fund",    26),
    ("★ FUNDAMENTAL 2027\nDECIDA AQUI",             "estrela", 22),
    ("★ TURMAS FUNDAMENTAL\n2027 — DECIDA AQUI",    "estrela", 26),
    ("Fundamental 2027 ·\ntotal de turmas",         "fund",    14),
    ("9º Ano NESTA escola\n(rede estadual)",        "fund",    22),
    ("9º Ano do município\n(todas as redes)",       "fund",    34),
    ("Demanda total de\n1ª série 2027",             "fund",    26),
    ("Ainda faltam\nno município",                  "fund",    26),
    ("IDEB DA ESCOLA",                              "fund",    13),
    ("1ª Série\n2026",                              "em",       9),
    ("★ 1ª SÉRIE 2027\nDECIDA AQUI",                "estrela", 13),
    ("Cursos EM 2026\n(sem EJA)",                   "em",      48),
    ("★ CURSOS 1ª SÉRIE 2027\nDECIDA AQUI",         "estrela", 28),
    ("★ TURMAS 1ª SÉRIE 2027\nDECIDA AQUI",         "estrela", 30),
    ("1ª série 2027 ·\ntotal pelos cursos",         "em",      15),
    ("★ ALTERAÇÃO DE\nCURSOS EMI",                  "estrela", 22),
    ("★ 2ª SÉRIE 2027\nDECIDA AQUI",                "estrela", 12),
    ("★ 3ª SÉRIE 2027\nDECIDA AQUI",                "estrela", 12),
    ("Oferta 2027 ·\n1ª série (turmas)",            "uetep",   13),
    ("Oferta 2027 ·\ncursos da 1ª série",           "uetep",   34),
    ("Oferta 2027 ·\n2ª / 3ª / subseq.",            "uetep",   22),
    ("Parciais 2026",                               "hoje",    26),
    ("★ Parciais 2026\njustificativa",              "estrela", 22),
    ("Possíveis fusões\nNA PRÓPRIA ESCOLA",         "fusao",   38),
    ("Fusão ENTRE ESCOLAS\n(mesmo município)",      "fusao",   42),
    ("★ Validação\nda fusão",                       "estrela", 20),
    ("Salas\nexistentes",                           "sala",    10),
    ("Salas necessárias\n2027 (calculado)",         "sala",    14),
    ("★ SALAS NECESSÁRIAS\n2027 — DECIDA AQUI",     "estrela", 16),
    ("Situação\nda sala",                           "sala",    20),
    ("Oferta EJA 2026\n(só prédio matriz)",         "eja",     44),
    ("Matrículas / Turmas EJA\n(só prédio matriz)", "eja",     22),
    ("EJA fora do prédio\nmatriz (UEJA)",           "eja",     36),
    ("★ EJA TURMAS 2027\nDECIDA AQUI",              "estrela", 14),
    ("Anexos desta\nescola",                        "anexo",   36),
    ("Qtd. de\nanexos",                             "anexo",    9),
    ("★ DECISÃO SOBRE\nO ANEXO",                    "estrela", 34),
    ("Resumo 2027",                                 "saida",   40),
    ("Escola Próxima",                              "saida",   26),
    ("★ Reordenamento\n2027",                       "estrela", 18),
    ("★ Justificativa",                             "estrela", 26),
    ("★ Observação",                                "estrela", 26),
    ("✓ Pronto",                                    "estrela",  9),
    ("Oferta 2027 ·\nalunos reais 2ª/3ª",           "uetep",   16),
    ("★ x Oferta 2027 ·\ndivergência",              "fusao",   40),
]

CHAVES = ['inep', 'gre', 'municipio', 'escola', 'turmas_hoje', 'turmas_2027', 'oferta_fund', 'fund_etapas', 'fund_turmas', 'fund_total', 'ano9_escola', 'ano9_municipio', 'demanda', 'faltam', 'ideb', 's1_2026', 's1_2027', 'cursos_2026', 'cursos_2027', 'cursos_turmas', 'cursos_total', 'alteracao_emi', 's2_2027', 's3_2027', 'of_1a', 'of_cursos', 'of_23sub', 'parciais', 'parciais_just', 'fusao_propria', 'fusao_entre', 'fusao_valid', 'salas_existentes', 'salas_calc', 'salas_decide', 'salas_situacao', 'eja_oferta', 'eja_matriculas', 'eja_fora', 'eja_turmas', 'anexos_lista', 'anexos_qtd', 'anexos_decisao', 'resumo', 'escola_proxima', 'reord', 'justificativa', 'observacao', 'pronto', 'of_alunos23', 'divergencia']

assert len(CHAVES) == len(V3), "cada coluna precisa de uma chave"

COL = {k: i + 1 for i, k in enumerate(CHAVES)}

def c(chave):
    """Número da coluna pela chave curta — sem índice mágico espalhado no código."""
    return COL[chave]

def titulo(chave):
    return V3[COL[chave] - 1][0]
