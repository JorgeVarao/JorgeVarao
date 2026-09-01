# -*- coding: utf-8 -*-
"""
Camada de agregação: lê as bases brutas e produz todas as abas auxiliares.
"""
import math, pickle
from collections import defaultdict
from motor import *

# ============================================================ leitura bruta

def carregar(pkl="dados.pkl"):
    return pickle.load(open(pkl, "rb"))

def mapa_gdi(linhas):
    """Base GDI: A INEP · C GRE · D Município · E Escola · F Nº salas · R Escola mais próxima"""
    m = {}
    for r in linhas[1:]:
        i = inep_(r[0])
        if not i: continue
        m[i] = {"gre": texto_(r[2]), "municipio": texto_(r[3]), "escola": texto_(r[4]),
                "salas": r[5], "escolaProxima": texto_(r[17]) if len(r) > 17 else "",
                "situacao": texto_(r[1]), "obs": texto_(r[10]) if len(r) > 10 else "",
                "proj2026": texto_(r[13]) if len(r) > 13 else "",
                "proj2027": texto_(r[14]) if len(r) > 14 else ""}
    return m

def processar_turmas(linhas_turmas, gdi):
    """Separa MATRIZ x ANEXO, exatamente como o script atual."""
    matrizes, ordem_m = {}, []
    anexos, ordem_a = {}, []
    for r in linhas_turmas[1:]:
        i = inep_(r[0])
        if not i: continue
        escola = texto_(r[1])
        anexo = esp_(r[2] if len(r) > 2 else "")
        meta = gdi.get(i, {})
        if not anexo:
            if i not in matrizes:
                matrizes[i] = Unidade(i, escola or meta.get("escola", ""), "", meta)
                ordem_m.append(i)
            acumular_turma(matrizes[i], r, "")
        else:
            ch = i + "|||" + anexo.upper()
            if ch not in anexos:
                anexos[ch] = Unidade(i, escola or meta.get("escola", ""), anexo, meta)
                ordem_a.append(ch)
            acumular_turma(anexos[ch], r, anexo)
        # a EJA em anexo também precisa aparecer no consolidado da matriz (UEJA)
        if anexo:
            if i not in matrizes:
                matrizes[i] = Unidade(i, escola or meta.get("escola", ""), "", meta)
                ordem_m.append(i)
            et = texto_(r[4]).upper(); cu = texto_(r[3]).upper()
            if "EJA" in et or "EJA" in cu:
                local = classificar_anexo(anexo)
                d = matrizes[i].ejaPorLocal.setdefault(
                    local, {"turmas": 0, "entm": 0, "cursos": {}, "nomes": set()})
                d["turmas"] += numero_(r[9]); d["entm"] += numero_(r[8])
                d["cursos"][texto_(r[3])] = d["cursos"].get(texto_(r[3]), 0) + numero_(r[9])
                d["nomes"].add(anexo.upper())
                matrizes[i].ejaAnexoTurmas += numero_(r[9])
                matrizes[i].ejaAnexoMatriculas += numero_(r[8])
    return matrizes, ordem_m, anexos, ordem_a

# ============================================================ UETEP

def base_uetep(linhas_mpe, gdi):
    """
    'Matriculas por etapa' — base UETEP (Ensino Profissional, turma a turma).
      4 Inep · 7 Turno · 9 Abreviação Etapa · 12 Nome Turma · 14 Nome Curso
      15 Ativas · 16 Enturmados · 17 Cursando · 18 Pré Matrícula
    Oferta NOVA = turma sem enturmados e com pré-matrícula > 0.
    """
    linhas, por_inep = [], {}
    for r in linhas_mpe[1:]:
        i = inep_(r[4])
        if not i: continue
        turno = texto_(r[7]); etapa = texto_(r[9]); turma = texto_(r[12])
        curso = texto_(r[14])
        enturm = numero_(r[16]); pre = numero_(r[18]); cursando = numero_(r[17])
        nova = (enturm == 0 and pre > 0)
        if not nova: continue
        linhas.append([int(i), gdi.get(i, {}).get("gre", ""), gdi.get(i, {}).get("municipio", ""),
                       gdi.get(i, {}).get("escola", ""), etapa, curso, turma, turno,
                       int(pre), int(enturm), int(cursando)])
        d = por_inep.setdefault(i, {"pre": 0, "turmas_lancadas": 0, "cursos": {}, "turnos": {}})
        d["pre"] += pre
        d["turmas_lancadas"] += 1
        d["cursos"][curso] = d["cursos"].get(curso, 0) + pre
        d["turnos"][turno] = d["turnos"].get(turno, 0) + pre
    # turmas previstas = teto(pré-matrículas / MAX_T) por curso
    for i, d in por_inep.items():
        prev = 0
        for c, p in d["cursos"].items():
            prev += max(1, math.ceil(p / MAX_T)) if p > 0 else 0
        d["turmas_previstas"] = prev
        d["resumo"] = " · ".join("%s: %d al." % (c, p) for c, p in
                                 sorted(d["cursos"].items(), key=lambda x: -x[1]))
    for L in linhas:
        L.append(max(1, math.ceil(L[8] / MAX_T)))
    return linhas, por_inep

# ============================================================ UEJA

ROTULO_LOCAL = {
    "MATRIZ": "Prédio matriz",
    "ANEXO": "Anexo",
    "SALA EXTERNA": "Sala externa",
    "COMUNIDADE / POVOADO": "Anexo em comunidade/povoado",
    "CEDIDA EM OUTRA ESCOLA": "Sala cedida em outra escola",
    "PRIVACAO DE LIBERDADE": "Privação de liberdade",
    "SOCIOEDUCATIVO": "Socioeducativo",
}

def base_ueja(matrizes, ordem_m):
    """Uma linha por escola × local da oferta de EJA, classificada pelo curso/anexo."""
    linhas = []
    resumo = {}
    for i in ordem_m:
        e = matrizes[i]
        if not e.ejaPorLocal: continue
        partes_fora = []
        for local, d in sorted(e.ejaPorLocal.items()):
            cursos = " · ".join("%s (%dT)" % (c, t) for c, t in
                                sorted(d["cursos"].items(), key=lambda x: -x[1]))
            nomes = " / ".join(sorted(d["nomes"])) if d["nomes"] else "—"
            linhas.append([int(i), e.gre, e.municipio, e.escola, ROTULO_LOCAL.get(local, local),
                           nomes, int(d["turmas"]), int(d["entm"]), cursos,
                           "NÃO" if local == "MATRIZ" else "SIM"])
            if local != "MATRIZ":
                partes_fora.append("%s (%s): %d turmas · %d matrículas"
                                   % (ROTULO_LOCAL.get(local, local), nomes,
                                      int(d["turmas"]), int(d["entm"])))
        resumo[i] = "\n".join(partes_fora) if partes_fora else "Toda a EJA no prédio matriz"
    return linhas, resumo

# ============================================================ FUSÃO ENTRE ESCOLAS

def fusao_entre_escolas(matrizes, ordem_m, ideb=None):
    """
    Sugere fusão de turmas de 1ª série do MESMO CURSO entre escolas do MESMO
    MUNICÍPIO. Traz matrículas, turmas de fundamental e IDEB para a análise.
    """
    ideb = ideb or {}
    por_mun = defaultdict(list)
    for i in ordem_m:
        e = matrizes[i]
        if e.municipio: por_mun[norm(e.municipio)].append(e)

    linhas = []
    resumo = defaultdict(list)
    for mun, escolas in sorted(por_mun.items()):
        if len(escolas) < 2: continue
        cursos = defaultdict(list)
        for e in escolas:
            for ck, d in e.cursos1a.items():
                if d["turmas"] > 0:
                    cursos[norm(d["curso"])].append((e, d))
        for ck, itens in sorted(cursos.items()):
            if len(itens) < 2: continue
            total_al = sum(d["entm"] for _, d in itens)
            total_tu = sum(d["turmas"] for _, d in itens)
            minimo = max(1, math.ceil(total_al / MAX_T))
            if minimo >= total_tu: continue      # não há ganho
            libera = int(total_tu - minimo)
            # destino = escola com mais alunos no curso (empate: mais salas)
            itens_ord = sorted(itens, key=lambda x: (-x[1]["entm"], -numero_(x[0].salasGDI)))
            dest, ddest = itens_ord[0]
            nome_curso = ddest["curso"]
            for e, d in itens_ord[1:]:
                linhas.append([
                    "", dest.municipio, dest.gre, nome_curso, "1ª Série",
                    int(e.inep), e.escola, int(d["turmas"]), int(d["entm"]),
                    int(e.efTotal), int(e.ef9Turmas), int(e.ef9Matriculas),
                    ideb.get(norm(dest.municipio), ""),
                    int(dest.inep), dest.escola, int(ddest["turmas"]), int(ddest["entm"]),
                    int(dest.efTotal), int(numero_(dest.salasGDI)),
                    int(total_al), int(total_tu), int(minimo), libera,
                    "Mesmo município"])
                resumo[e.inep].append(
                    "→ %s: %d turma(s)/%d alunos poderiam migrar para %s (%d alunos lá). "
                    "No município: %d alunos em %d turmas → cabem em %d (libera %d)."
                    % (nome_curso, d["turmas"], d["entm"], dest.escola, ddest["entm"],
                       total_al, total_tu, minimo, libera))
                resumo[dest.inep].append(
                    "← %s: pode receber %d aluno(s) de %s." % (nome_curso, d["entm"], e.escola))
    return linhas, {k: "\n".join(v) for k, v in resumo.items()}

# ============================================================ PANORAMA MUNICIPAL

def panorama(linhas_pan, matrizes, ordem_m):
    """
    Recalcula a demanda de 1ª série somando o 9º ano da própria rede estadual
    com o que vem das demais redes do município.
    """
    est9 = defaultdict(lambda: [0, 0])     # município -> [alunos, turmas]
    for i in ordem_m:
        e = matrizes[i]
        if not e.municipio: continue
        k = norm(e.municipio)
        est9[k][0] += e.ef9Matriculas
        est9[k][1] += e.ef9Turmas

    saida, indice = [], {}
    for r in linhas_pan[1:]:
        m = texto_(r[0])
        if not m: continue
        k = norm(m)
        todas = numero_(r[1])
        ref_turmas = inteiro_(r[2])
        fonte = texto_(r[6]) if len(r) > 6 else "CENSO"
        est_al, est_tu = est9.get(k, [0, 0])
        outras = max(0.0, todas - est_al)
        demanda = est_al + outras
        nec = int(math.ceil(demanda / CAP_1SERIE)) if demanda > 0 else 0
        saida.append([m, int(todas), ref_turmas, int(est_al), int(est_tu),
                      int(outras), int(demanda), nec, fonte])
        indice[k] = {"linha": len(saida) + 1, "todas": int(todas), "ref": ref_turmas,
                     "est_al": int(est_al), "est_tu": int(est_tu),
                     "outras": int(outras), "demanda": int(demanda), "nec": nec}
    # municípios que só aparecem na rede estadual
    for k, (al, tu) in sorted(est9.items()):
        if k in indice: continue
        nec = int(math.ceil(al / CAP_1SERIE)) if al > 0 else 0
        nome = next((matrizes[i].municipio for i in ordem_m if norm(matrizes[i].municipio) == k), k)
        saida.append([nome, int(al), nec, int(al), int(tu), 0, int(al), nec, "REDE ESTADUAL"])
        indice[k] = {"linha": len(saida) + 1, "todas": int(al), "ref": nec,
                     "est_al": int(al), "est_tu": int(tu), "outras": 0,
                     "demanda": int(al), "nec": nec}
    return saida, indice
