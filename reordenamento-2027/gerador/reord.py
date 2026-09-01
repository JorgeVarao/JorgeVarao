# -*- coding: utf-8 -*-
"""Definição das três versões da aba Reordenamento 2027."""
from motor import norm, texto_, numero_, inep_

BT = "'Base Tratada'"          # A=1 ... AD=30
PM = "'Panorama Municipal'"    # A=1 ... N=14
GD = "'Base GDI'"              # A=1 ... AL=38
ID = "'IDEB Municípios'"

def vl(chave, aba, col, ult, err='""'):
    return 'IFERROR(VLOOKUP(%s,%s!$A$2:$%s$%d,%d,FALSE),%s)' % (
        chave, aba[0], aba[1], ult, col, err)

def bt(chave, col, n, err='""'):
    return 'IFERROR(VLOOKUP(%s,%s!$A$2:$AD$%d,%d,FALSE),%s)' % (chave, BT, n + 1, col, err)

def pm(chave, col, n, err='""'):
    return 'IFERROR(VLOOKUP(%s,%s!$A$2:$N$%d,%d,FALSE),%s)' % (chave, PM, n + 1, col, err)

def gd(chave, col, n, err='""'):
    return 'IFERROR(VLOOKUP(%s,%s!$A$2:$AL$%d,%d,FALSE),%s)' % (chave, GD, n + 1, col, err)

def ib(chave, col, n):
    """IDEB ainda não cadastrado (célula vazia) devolve texto, não zero."""
    v = 'VLOOKUP(%s,%s!$A$2:$F$%d,%d,FALSE)' % (chave, ID, n + 1, col)
    return 'IFERROR(IF(N(%s)=0,"— sem IDEB cadastrado",%s),"— sem IDEB cadastrado")' % (v, v)

def falta(chave, col_saldo, n):
    """Texto do 'Ainda faltam no município', reagindo ao que for preenchido."""
    v = pm(chave, col_saldo, n, '""')
    return ('IF(%s="","",IF(%s="","—",IF(%s>0,"FALTAM DISTRIBUIR "&%s&" TURMA(S)",'
            'IF(%s=0,"DEMANDA MUNICIPAL ATENDIDA",'
            '"EXCEDE A DEMANDA EM "&ABS(%s)&" TURMA(S)"))))'
            % (chave, v, v, v, v, v))

def situacao_sala(col_exist, col_nec, r):
    return ('=IF(${n}${r}="","",IF(${e}${r}>${n}${r},(${e}${r}-${n}${r})&" SALA(S) OCIOSA(S)",'
            'IF(${e}${r}<${n}${r},"CONSTRUIR "&(${n}${r}-${e}${r})&" SALA(S)",'
            '"QUANTIDADE ADEQUADA")))'
            ).replace("${e}", col_exist).replace("${n}", col_nec).replace("${r}", str(r))

# =====================================================================
#  V1 — retrato exato da aba que existe hoje
# =====================================================================

H_V1 = ["INEP", "GRE", "Município", "Escola", "Turmas hoje\n(2026)",
        "Turmas 2027\n(projeção)", "Oferta\nFundamental", "9º Ano do município",
        "Ainda faltam no município", "★ FUNDAMENTAL\n2027 — DECIDA AQUI", "IDEB",
        "1ª Série\n2026", "★ 1ª SÉRIE 2027\nDECIDA AQUI", "Cursos EM 2026\n(sem EJA)",
        "★ CURSOS 1ª SÉRIE\n2027 — DECIDA AQUI", "ALTERAÇÃO DE CURSOS EMI",
        "2ª Série\n2027", "3ª Série\n2027", "Parciais 2026",
        "Parciais 2026 justificativa", "Possíveis\nFusões de Turmas",
        "Validações Fusões Turmas", "Salas\nexistentes",
        "★ SALAS NECESSÁRIAS 2027 —  DECIDA AQUI", "Situação\nda sala",
        "Oferta EJA\n2026", "Matrículas / Turmas\nEJA", "★ EJA Turmas\n2027",
        "★ EJA Cursos\n2027", "ANEXOS", "CURSOS", "", "Resumo 2027",
        "Escola Próxima", "Reordenamento\n2027", "Justificativa", "Observação",
        "✓ Pronto"]

B_V1 = (["id"] * 4 + ["hoje"] * 2 + ["fund"] * 2 + ["fund"] + ["estrela"] + ["fund"] +
        ["em", "estrela", "em", "estrela", "estrela", "em", "em"] +
        ["hoje", "estrela"] + ["fusao", "estrela"] +
        ["sala", "estrela", "sala"] +
        ["eja"] * 2 + ["estrela"] * 2 + ["anexo", "anexo", "saida"] +
        ["saida"] * 2 + ["estrela"] * 4)

# =====================================================================
#  V2 — mesma estrutura, com o que faltava funcionando
# =====================================================================

H_V2 = list(H_V1)
H_V2[31] = "(reservado)"
B_V2 = list(B_V1)

def linhas_v2(ineps, n_bt, n_pm, n_gd, n_id, aba_fusao="'Fusão Turmas'", n_fus=62):
    """Devolve, para cada linha, um dict {coluna: valor ou fórmula}."""
    out = []
    for k, i in enumerate(ineps):
        r = k + 2
        A = "$A%d" % r; C = "$C%d" % r
        f = {}
        f[2]  = "=" + gd(A, 3, n_gd)                       # GRE
        f[4]  = "=" + gd(A, 5, n_gd)                       # Escola
        f[5]  = "=" + gd(A, 14, n_gd)                      # Turmas hoje (GDI)
        f[6]  = "=" + gd(A, 15, n_gd)                      # Turmas 2027 (GDI)
        f[7]  = "=" + bt(A, 9, n_bt, '"—"')                # Oferta Fundamental
        f[8]  = ('=IF(%s="","",%s&" alunos · "&%s&" turmas  (estadual nesta escola: "&%s&")")'
                 % (C, pm(C, 7, n_pm, '"—"'), pm(C, 8, n_pm, '"—"'), bt(A, 18, n_bt, '0')))
        f[9]  = "=" + falta(C, 11, n_pm)                   # Ainda faltam (lê saldo V2)
        f[11] = "=" + ib(C, 2, n_id)                       # IDEB
        f[14] = "=" + bt(A, 3, n_bt, '"—"')                # Cursos EM 2026
        f[19] = "=" + bt(A, 10, n_bt, '"—"')               # Parciais
        f[21] = "=" + bt(A, 11, n_bt, '"—"')               # Possíveis fusões
        f[22] = ('=IFERROR(IF(COUNTIF(%s!$J$2:$J$%d,%s)+COUNTIF(%s!$N$2:$N$%d,%s)=0,'
                 '"Sem fusão turma a turma",'
                 'COUNTIF(%s!$J$2:$J$%d,%s)&" turma(s) saem · "&'
                 'COUNTIF(%s!$N$2:$N$%d,%s)&" turma(s) recebem"),"—")'
                 % (aba_fusao, n_fus + 1, A, aba_fusao, n_fus + 1, A,
                    aba_fusao, n_fus + 1, A, aba_fusao, n_fus + 1, A))
        f[23] = "=" + gd(A, 6, n_gd)                       # Salas existentes
        f[25] = situacao_sala("$W", "$X", r)               # Situação da sala
        f[26] = "=" + bt(A, 7, n_bt, '"—"')                # Oferta EJA
        f[27] = "=" + bt(A, 16, n_bt, '"—"')               # Matrículas/Turmas EJA
        f[30] = "=" + bt(A, 26, n_bt, '"—"')               # ANEXOS
        f[31] = "=" + bt(A, 21, n_bt, '"—"')               # CURSOS (UETEP)
        f[33] = ('=IF(%s="","",$M%d&" de 1ª | "&$Q%d&" de 2ª | "&$R%d&" de 3ª | '
                 'Necessidade: "&$X%d&" salas | "&$Y%d)' % (A, r, r, r, r, r))
        f[34] = "=" + gd(A, 18, n_gd, '"—"')               # Escola próxima
        out.append(f)
    return out

# =====================================================================
#  V3 — tudo o que a V2 tem + os pedidos novos
# =====================================================================

H_V3 = [
    "INEP", "GRE", "Município", "Escola",
    "Turmas hoje\n(2026)", "Turmas 2027\n(projeção)",
    "Oferta\nFundamental", "9º Ano NESTA escola\n(rede estadual)",
    "9º Ano do município\n(todas as redes)", "Demanda total de\n1ª série 2027",
    "Ainda faltam\nno município", "★ FUNDAMENTAL 2027\nDECIDA AQUI",
    "IDEB do\nmunicípio",
    "1ª Série\n2026", "★ 1ª SÉRIE 2027\nDECIDA AQUI",
    "Cursos EM 2026\n(sem EJA)", "★ CURSOS 1ª SÉRIE 2027\nDECIDA AQUI",
    "★ ALTERAÇÃO DE\nCURSOS EMI",
    "★ 2ª SÉRIE 2027\nDECIDA AQUI", "★ 3ª SÉRIE 2027\nDECIDA AQUI",
    "UETEP · novas\nmatrículas 2027", "UETEP · detalhe\npor curso",
    "UETEP · turmas\nprevistas",
    "Parciais 2026", "★ Parciais 2026\njustificativa",
    "Possíveis fusões\nNA PRÓPRIA ESCOLA", "Fusão ENTRE ESCOLAS\n(mesmo município)",
    "★ Validação\nda fusão",
    "Salas\nexistentes", "Salas necessárias\n2027 (calculado)",
    "★ SALAS NECESSÁRIAS\n2027 — DECIDA AQUI", "Situação\nda sala",
    "Oferta EJA\n2026", "Matrículas /\nTurmas EJA",
    "EJA fora do prédio\nmatriz (UEJA)", "★ EJA Turmas\n2027", "★ EJA Cursos\n2027",
    "Anexos desta\nescola", "Qtd. de\nanexos",
    "★ MANTER\nO ANEXO?", "★ ENCERRAR\nO ANEXO?", "★ REMANEJAR A\nOFERTA DO ANEXO?",
    "Resumo 2027", "Escola Próxima",
    "★ Reordenamento\n2027", "★ Justificativa", "★ Observação", "✓ Pronto",
]

B_V3 = (["id"] * 4 + ["hoje"] * 2 +
        ["fund"] * 4 + ["fund"] + ["estrela"] + ["fund"] +
        ["em", "estrela", "em", "estrela", "estrela", "estrela", "estrela"] +
        ["uetep"] * 3 +
        ["hoje", "estrela"] +
        ["fusao", "fusao", "estrela"] +
        ["sala", "sala", "estrela", "sala"] +
        ["eja"] * 3 + ["estrela"] * 2 +
        ["anexo", "anexo"] + ["estrela"] * 3 +
        ["saida"] * 2 + ["estrela"] * 4)

def linhas_v3(ineps, n_bt, n_pm, n_gd, n_id, n_fe, aba_fe="'Fusão Entre Escolas'"):
    out = []
    for k, i in enumerate(ineps):
        r = k + 2
        A = "$A%d" % r; C = "$C%d" % r
        f = {}
        f[2]  = "=" + gd(A, 3, n_gd)
        f[4]  = "=" + gd(A, 5, n_gd)
        f[5]  = "=" + gd(A, 14, n_gd)
        f[6]  = "=" + gd(A, 15, n_gd)
        f[7]  = "=" + bt(A, 9, n_bt, '"—"')
        # 9º ano nesta escola (rede estadual)
        f[8]  = ('=IF(%s="","",%s&" matrícula(s) · "&%s&" turma(s)")'
                 % (A, bt(A, 18, n_bt, '0'), bt(A, 17, n_bt, '0')))
        # 9º ano do município — todas as redes
        f[9]  = ('=IF(%s="","",%s&" alunos (estadual "&%s&" + outras redes "&%s&")")'
                 % (C, pm(C, 2, n_pm, '0'), pm(C, 4, n_pm, '0'), pm(C, 6, n_pm, '0')))
        # demanda total de 1ª série
        f[10] = ('=IF(%s="","",%s&" alunos → "&%s&" turma(s) necessária(s)")'
                 % (C, pm(C, 7, n_pm, '0'), pm(C, 8, n_pm, '0')))
        f[11] = "=" + falta(C, 13, n_pm)                    # lê o saldo da V3
        f[13] = "=" + ib(C, 2, n_id)
        f[16] = "=" + bt(A, 3, n_bt, '"—"')
        # UETEP
        f[21] = "=" + bt(A, 19, n_bt, '0')
        f[22] = "=" + bt(A, 21, n_bt, '"—"')
        f[23] = "=" + bt(A, 20, n_bt, '0')
        f[24] = "=" + bt(A, 10, n_bt, '"—"')                # Parciais
        f[26] = "=" + bt(A, 11, n_bt, '"—"')                # Fusão na própria escola
        f[27] = ('=IFERROR(IF(COUNTIF(%s!$F$2:$F$%d,%s)+COUNTIF(%s!$N$2:$N$%d,%s)=0,'
                 '"Nenhuma fusão sugerida",%s),"—")'
                 % (aba_fe, n_fe + 1, A, aba_fe, n_fe + 1, A, bt(A, 25, n_bt, '"—"')))
        f[29] = "=" + gd(A, 6, n_gd)                        # Salas existentes
        # Salas necessárias calculado: base do motor + o que o usuário mexeu nas séries
        f[30] = ('=IF(%s="","",MAX(0,%s+($O%d-%s)+($S%d-%s)+($T%d-%s)))'
                 % (A, bt(A, 28, n_bt, '0'), r, bt(A, 4, n_bt, '0'),
                    r, bt(A, 5, n_bt, '0'), r, bt(A, 6, n_bt, '0')))
        f[32] = situacao_sala("$AC", "$AE", r)
        f[33] = "=" + bt(A, 7, n_bt, '"—"')                 # Oferta EJA
        f[34] = "=" + bt(A, 16, n_bt, '"—"')                # Matrículas/Turmas EJA
        f[35] = "=" + bt(A, 22, n_bt, '"—"')                # UEJA
        f[38] = "=" + bt(A, 26, n_bt, '"—"')                # Anexos
        f[39] = "=" + bt(A, 27, n_bt, '0')                  # Qtd anexos
        f[43] = ('=IF(%s="","",$O%d&" de 1ª | "&$S%d&" de 2ª | "&$T%d&" de 3ª | UETEP "&$W%d'
                 '&" | Necessidade: "&$AE%d&" salas | "&$AF%d)' % (A, r, r, r, r, r, r))
        f[44] = "=" + gd(A, 18, n_gd, '"—"')
        out.append(f)
    return out
