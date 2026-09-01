/**
 * =====================================================================
 *  REORDENAMENTO 2027 — SCRIPT ÚNICO
 *  SEDUC-PI / SUPEX / UGERF / GDI
 * =====================================================================
 *
 *  Turmas ─┐
 *  Base GDI ├─► atualizarTudo() ─► Base Tratada ─┐
 *  Matriculas por etapa ─┘        Base Anexos    ├─► Reordenamento 2027 V2
 *                                 Base UETEP     │   Reordenamento 2027 V3
 *                                 Base UEJA      │
 *                                 Fusão Entre Escolas
 *                                 Panorama Municipal
 *
 *  REGRA DE OURO: o script NUNCA apaga coluna marcada com ★.
 *  Tudo o que estiver em amarelo com ★ é decisão humana e é preservado.
 *
 *  ---------------------------------------------------------------------
 *  COMO INSTALAR
 *    1. Abra a planilha no Google Sheets
 *    2. Extensões → Apps Script
 *    3. Apague o que estiver lá e cole este arquivo inteiro
 *    4. Salve, volte para a planilha e recarregue a página (F5)
 *    5. Use o menu "📊 GDI / UGERF"
 *
 *  ORDEM DE USO
 *    a) ⟳ Atualizar tudo            — sempre que a aba Turmas mudar
 *    b) ⚙ Criar/atualizar Validações — se as listas dos menus mudarem
 *
 *  ABAS DE ENTRADA QUE PRECISAM EXISTIR
 *    Turmas                 A INEP · B Escola · C Anexo · D Curso · E Etapa
 *                           F Organização · G Período · H Turno
 *                           I Enturmados · J Turmas
 *    Base GDI               A INEP · C GRE · D Município · E Escola
 *                           F Nº salas · N Projeção 2026 · O Projeção 2027
 *                           R Escola mais próxima
 *    Matriculas por etapa   E INEP · H Turno · J Etapa · M Turma
 *                           O Curso · Q Enturmados · S Pré-matrícula
 *    Panorama Municipal     A Município · B 9º ano todas as redes
 *                           C Turmas necessárias (referência) · G Fonte
 * =====================================================================
 */


// ═════════════════════════════════════════════════════════════════════
//  PARÂMETROS
// ═════════════════════════════════════════════════════════════════════

var MAX_T      = 40;   // capacidade máxima de uma turma
var CAP_1SERIE = 40;   // alunos por turma ao converter 9º ano em 1ª série

var ABA_TURMAS  = "Turmas";
var ABA_GDI     = "Base GDI";
var ABA_MPE     = "Matriculas por etapa";
var ABA_BT      = "Base Tratada";
var ABA_BA      = "Base Anexos";      // se existir "Base Anexos (1)", o script não mexe
var ABA_UETEP   = "Base UETEP";
var ABA_OF1     = "UETEP · Oferta Integral 2027";   // 1ª série de 2027, por escola x curso
var ABA_OFC     = "UETEP · Continuidade 2027";      // 2ª e 3ª série — uma linha por turma
var ABA_OFS     = "UETEP · Subsequente 2027";
var ABA_UEJA    = "Base UEJA";
var ABA_FE      = "Fusão Entre Escolas";
var ABA_FT      = "Fusão Turmas";
var ABA_PAN     = "Panorama Municipal";
var ABA_IDEB    = "IDEB Municípios";
var ABA_VAL     = "Validações";
var ABA_V1      = "Reordenamento 2027";
var ABA_V2      = "Reordenamento 2027 V2";
var ABA_V3      = "Reordenamento 2027 V3";


// ═════════════════════════════════════════════════════════════════════
//  MENU
// ═════════════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi().createMenu("📊 GDI / UGERF")
    .addItem("⟳ Atualizar tudo (rode sempre)", "atualizarTudo")
    .addSeparator()
    .addItem("① Atualizar só as bases", "atualizarBases")
    .addItem("② Reescrever só as fórmulas da V2 e V3", "reescreverFormulas")
    .addItem("③ Criar/atualizar Validações", "criarValidacoes")
    .addSeparator()
    .addItem("↩ Desfazer último curso da célula", "desfazerUltimoCurso")
    .addToUi();
}


function atualizarTudo() {
  var res = atualizarBases(true);
  if (!res) return;   // aba Turmas ausente — o alerta já foi dado
  reescreverFormulas(true);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    "✅ Atualização concluída\n\n" +
    "Base Tratada: " + res.matrizes + " prédio(s) matriz\n" +
    "Base Anexos: " + res.anexos + " anexo(s) / sala(s) externa(s)\n" +
    "Base UETEP: " + res.uetep + " escola(s) com oferta 2027\n" +
    "Base UEJA: " + res.ueja + " linha(s)\n" +
    "Fusão Entre Escolas: " + res.fusoes + " sugestão(ões)\n" +
    "Panorama Municipal: " + res.municipios + " município(s)\n\n" +
    "Nenhuma coluna ★ foi tocada.");
}


// ═════════════════════════════════════════════════════════════════════
//  1. BASES
// ═════════════════════════════════════════════════════════════════════

function atualizarBases(silencioso) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var src = ss.getSheetByName(ABA_TURMAS);
  if (!src) { erro_('Aba "' + ABA_TURMAS + '" não encontrada.'); return null; }

  var dados = src.getDataRange().getValues();
  var gdi   = mapaGDI_(ss);

  var matrizes = {}, ordemM = [];
  var anexos   = {}, ordemA = [];

  for (var i = 1; i < dados.length; i++) {

    var r = dados[i];
    var inep = inep_(r[0]);
    if (!inep) continue;

    var escola = texto_(r[1]);
    var anexo  = espacos_(r[2]);
    var meta   = gdi[inep] || {};

    if (!matrizes[inep]) {
      matrizes[inep] = novaUnidade_(inep, escola || meta.escola || "", "", meta);
      ordemM.push(inep);
    }

    if (!anexo) {
      acumularTurma_(matrizes[inep], r, "");
    } else {
      var chave = inep + "|||" + anexo.toUpperCase();
      if (!anexos[chave]) {
        anexos[chave] = novaUnidade_(inep, escola || meta.escola || "", anexo, meta);
        ordemA.push(chave);
      }
      acumularTurma_(anexos[chave], r, anexo);
      registrarEJAExterna_(matrizes[inep], r, anexo);
    }
  }

  var oferta = lerOferta2027_(ss);
  var uetep = baseUETEP_(ss, gdi);
  var ueja  = baseUEJA_(matrizes, ordemM);
  var fus   = fusaoEntreEscolas_(matrizes, ordemM);
  var pan   = panorama_(ss, matrizes, ordemM);

  gravarBaseTratada_(ss, matrizes, ordemM, anexos, ordemA, oferta, ueja.resumo, fus.resumo);

  // Base Anexos: se vocês criaram a sua própria versão, o script não encosta nela
  if (ss.getSheetByName("Base Anexos (1)")) {
    Logger.log("Base Anexos (1) existe — mantida como está, sem reescrita.");
  } else {
    gravarBaseAnexos_(ss, anexos, ordemA);
  }

  gravarBaseUETEP_(ss, oferta, matrizes);
  gravarUEJA_(ss, ueja.linhas);
  gravarPanorama_(ss, pan);
  gravarIDEB_(ss, pan);            // precisa vir antes: a fusão consulta o IDEB
  gravarFusaoEntreEscolas_(ss, fus.linhas);

  SpreadsheetApp.flush();

  var res = { matrizes: ordemM.length, anexos: ordemA.length,
              uetep: Object.keys(oferta).length, ueja: ueja.linhas.length,
              fusoes: fus.linhas.length, municipios: pan.length };

  if (!silencioso) {
    SpreadsheetApp.getUi().alert(
      "✅ Bases atualizadas\n\n" +
      "Matrizes: " + res.matrizes + "\nAnexos: " + res.anexos +
      "\nUETEP: " + res.uetep + "\nUEJA: " + res.ueja +
      "\nFusões entre escolas: " + res.fusoes);
  }
  return res;
}


// ─────────────────────────────────────────────────── mapa da Base GDI

function mapaGDI_(ss) {
  var sh = ss.getSheetByName(ABA_GDI);
  var mapa = {};
  if (!sh) return mapa;
  var d = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    var inep = inep_(d[i][0]);
    if (!inep) continue;
    mapa[inep] = {
      gre:           texto_(d[i][2]),
      municipio:     texto_(d[i][3]),
      escola:        texto_(d[i][4]),
      salas:         d[i][5],
      escolaProxima: texto_(d[i][17])
    };
  }
  return mapa;
}


// ─────────────────────────────────────────────── estrutura da unidade

function novaUnidade_(inep, escola, anexo, meta) {
  var e = {
    inep: inep, escola: escola, anexo: anexo,
    gre: meta.gre || "", municipio: meta.municipio || "",
    salasGDI: meta.salas, escolaProxima: meta.escolaProxima || "",
    tipoAnexo: classificarAnexo_(anexo),
    emLinhas: [], ejaLinhas: [], efLinhas: [], outrasLinhas: [], parciais: [],
    cursos: {}, cursos1a: {}, ejaPorLocal: {}
  };
  var zeros = ("totalTurmas totalMatriculas emTotal efTotal totalEJA totalEJAentm " +
    "nI nM nT nN efI efM efT efN s1 s2 s3 " +
    "s1I s1M s1T s1N s2I s2M s2T s2N s3I s3M s3T s3N " +
    "outrosEM outrosI outrosM outrosT outrosN ejI ejM ejT ejN " +
    "ef9Turmas ef9Matriculas ef8Turmas ef8Matriculas " +
    "aeeTurmas aeeMatriculas ejaAnexoTurmas ejaAnexoMatriculas").split(" ");
  for (var k = 0; k < zeros.length; k++) e[zeros[k]] = 0;
  return e;
}


// ─────────────────────────────────────────── classificação do anexo

function classificarAnexo_(nome) {
  var n = semAcento_(nome).toUpperCase();
  if (!n) return "MATRIZ";
  if (/SISTEMA PRISIONAL|PENITENCIARIA|PRESIDIO|CADEIA|CUSTODIA|FEMININA/.test(n))
    return "PRIVACAO DE LIBERDADE";
  if (/SOCIO ?EDUCATIVO|CEIP|CEM |CEF /.test(n)) return "SOCIOEDUCATIVO";
  if (/SALA EXTERNA|TURMA EXTERNA|EXTERNA/.test(n)) return "SALA EXTERNA";
  if (/ESCOLA |ESC\.|ESC |E\. ?M\.|U\. ?E\.|U\.E|U E |UEMA|U\. MARCELINA/.test(n))
    return "CEDIDA EM OUTRA ESCOLA";
  if (/POVOADO|POV\.|ASSENTAMENTO|LOCALIDADE|QUILOMBO|COMUNIDADE|ASSOCIACAO|SITIO/.test(n))
    return "COMUNIDADE / POVOADO";
  return "ANEXO";
}

function rotuloLocal_(t) {
  return ({ "MATRIZ": "Prédio matriz", "ANEXO": "Anexo",
            "SALA EXTERNA": "Sala externa",
            "COMUNIDADE / POVOADO": "Anexo em comunidade/povoado",
            "CEDIDA EM OUTRA ESCOLA": "Sala cedida em outra escola",
            "PRIVACAO DE LIBERDADE": "Privação de liberdade",
            "SOCIOEDUCATIVO": "Socioeducativo" })[t] || t;
}


// ──────────────────────────────────────────────────── acumular turma

function acumularTurma_(e, r, anexo) {

  var curso = texto_(r[3]), etapa = texto_(r[4]);
  var organizacao = texto_(r[5]), periodo = texto_(r[6]);
  var turno = texto_(r[7]);
  var entm = numero_(r[8]), turmas = numero_(r[9]);

  if (turmas === 0 && entm === 0 && !curso && !etapa) return;

  var cU = curso.toUpperCase(), eU = etapa.toUpperCase(), tU = turno.toUpperCase();

  var ehEJA = eU.indexOf("EJA") > -1 || cU.indexOf("EJA") > -1;
  var ehAEE = /(^|\s)AEE(\s|$)/i.test(etapa) || /(^|\s)AEE(\s|$)/i.test(curso);

  var ehEF = false;
  if (!ehEJA) {
    if (cU.indexOf("ANOS FINAIS") > -1 || cU.indexOf("ANOS INICIAIS") > -1) ehEF = true;
    if (/EF\s*-\s*\d+º/i.test(etapa) || /\d+º\s*ANO/i.test(etapa)) ehEF = true;
  }

  var ab = abreviarTurno_(tU);
  var txt = "• " + etapa + " - " + curso + " - " + entm + " matrículas " + ab +
            " - " + turmas + (turmas === 1 ? " turma" : " turmas");

  e.totalTurmas += turmas;
  e.totalMatriculas += entm;

  // ---- EJA: aparece, mas não entra no cálculo físico de salas
  if (ehEJA) {
    e.ejaLinhas.push(txt);
    e.totalEJA += turmas;
    e.totalEJAentm += entm;
    somarTurno_(e, "ej", tU, turmas);
    somarEJALocal_(e, anexo, curso, turmas, entm);
    return;
  }

  // ---- AEE: também fica fora da projeção de salas regulares
  if (ehAEE) {
    e.outrasLinhas.push(txt);
    e.aeeTurmas += turmas;
    e.aeeMatriculas += entm;
    return;
  }

  // ---- Fundamental
  if (ehEF) {
    e.efLinhas.push(txt);
    e.efTotal += turmas;
    somarTurno_(e, "ef", tU, turmas);
    if (/EF\s*-\s*9º/i.test(etapa) || eU.indexOf("9º ANO") > -1) {
      e.ef9Turmas += turmas; e.ef9Matriculas += entm;
    }
    if (/EF\s*-\s*8º/i.test(etapa) || eU.indexOf("8º ANO") > -1) {
      e.ef8Turmas += turmas; e.ef8Matriculas += entm;
    }
    return;
  }

  // ---- Ensino Médio regular / EPT EMI
  e.emLinhas.push(txt);
  e.emTotal += turmas;
  somarTurno_(e, "n", tU, turmas);

  if (cU.indexOf("PARCIAL") > -1)
    e.parciais.push(etapa + " | " + turno + " | " + turmas + "T (" + entm + " al.)");

  var ehSerie = false;
  if (/1ª\s*S[ÉE]RIE/i.test(etapa)) {
    e.s1 += turmas; somarTurno_(e, "s1", tU, turmas); ehSerie = true;
    if (!e.cursos1a[curso])
      e.cursos1a[curso] = { curso: curso, turmas: 0, entm: 0, turno: turno };
    e.cursos1a[curso].turmas += turmas;
    e.cursos1a[curso].entm   += entm;
  } else if (/2ª\s*S[ÉE]RIE/i.test(etapa)) {
    e.s2 += turmas; somarTurno_(e, "s2", tU, turmas); ehSerie = true;
  } else if (/3ª\s*S[ÉE]RIE/i.test(etapa)) {
    e.s3 += turmas; somarTurno_(e, "s3", tU, turmas); ehSerie = true;
  }

  if (!ehSerie) {
    e.outrosEM += turmas;
    somarTurno_(e, "outros", tU, turmas);
  }

  // ---- fusão dentro da própria unidade
  var ck = [curso, etapa, tU, organizacao, periodo].join("|");
  if (!e.cursos[ck])
    e.cursos[ck] = { curso: curso, etapa: etapa, turno: turno,
                     organizacao: organizacao, periodo: periodo, turmas: 0, entm: 0 };
  e.cursos[ck].turmas += turmas;
  e.cursos[ck].entm   += entm;
}


function somarEJALocal_(e, anexo, curso, turmas, entm) {
  var local = classificarAnexo_(anexo);
  if (!e.ejaPorLocal[local])
    e.ejaPorLocal[local] = { turmas: 0, entm: 0, cursos: {}, nomes: {} };
  var d = e.ejaPorLocal[local];
  d.turmas += turmas;
  d.entm   += entm;
  d.cursos[curso] = (d.cursos[curso] || 0) + turmas;
  if (anexo) d.nomes[espacos_(anexo).toUpperCase()] = true;
  if (local !== "MATRIZ") {
    e.ejaAnexoTurmas += turmas;
    e.ejaAnexoMatriculas += entm;
  }
}

/** A EJA que fica em anexo precisa aparecer também no consolidado da matriz. */
function registrarEJAExterna_(matriz, r, anexo) {
  var eU = texto_(r[4]).toUpperCase(), cU = texto_(r[3]).toUpperCase();
  if (eU.indexOf("EJA") === -1 && cU.indexOf("EJA") === -1) return;
  somarEJALocal_(matriz, anexo, texto_(r[3]), numero_(r[9]), numero_(r[8]));
}


// ═════════════════════════════════════════════════════════════════════
//  2. PROJEÇÃO 2027
// ═════════════════════════════════════════════════════════════════════

/**
 * 1ª de 2027 = 1ª de hoje · 2ª de 2027 = 1ª de hoje · 3ª de 2027 = 2ª de hoje.
 *
 * Regra física da sala:
 *   integral ocupa a sala o dia inteiro
 *   manhã e tarde dividem a mesma sala  →  vale o maior dos dois
 *   noturno regular fica à parte
 *   EJA não entra
 *   UETEP entra como turma nova de 1ª série integral
 */
function projetar2027_(e, uetepTurmas) {

  uetepTurmas = uetepTurmas || 0;

  var pro1 = e.s1, pro2 = e.s1, pro3 = e.s2;

  var i27 = e.s1I + e.s1I + e.s2I + e.outrosI + e.efI + uetepTurmas;
  var m27 = e.s1M + e.s1M + e.s2M + e.outrosM + e.efM;
  var t27 = e.s1T + e.s1T + e.s2T + e.outrosT + e.efT;
  var n27 = e.s1N + e.s1N + e.s2N + e.outrosN + e.efN;

  var salas = i27 + Math.max(m27, t27) + n27;
  var total27 = pro1 + pro2 + pro3 + e.outrosEM + uetepTurmas;

  return { pro1: pro1, pro2: pro2, pro3: pro3,
           integral: i27, manha: m27, tarde: t27, noite: n27,
           salasNec: salas, delta: total27 - e.emTotal, uetep: uetepTurmas };
}


// ═════════════════════════════════════════════════════════════════════
//  3. TEXTOS
// ═════════════════════════════════════════════════════════════════════

function montarOfertaEM_(e) {
  if (!e.emLinhas.length) return "—";
  var res = [];
  if (e.nI > 0) res.push(e.nI + " Integral");
  if (e.nM > 0) res.push(e.nM + " Manhã");
  if (e.nT > 0) res.push(e.nT + " Tarde");
  if (e.nN > 0) res.push(e.nN + " Noite");
  return e.emLinhas.slice().sort().join("\n") +
         "\nTurmas: " + (res.length ? res.join(" | ") : "0");
}

function montarOfertaEJA_(e) {
  if (!e.ejaLinhas.length) return "—";
  var res = [];
  if (e.ejI > 0) res.push(e.ejI + " Integral");
  if (e.ejM > 0) res.push(e.ejM + " Manhã");
  if (e.ejT > 0) res.push(e.ejT + " Tarde");
  if (e.ejN > 0) res.push(e.ejN + " Noite");
  return e.ejaLinhas.slice().sort().join("\n") +
         "\nTurmas EJA: " + (res.length ? res.join(" | ") : "0");
}

/** Fusão apenas dentro da própria unidade. */
function montarFusoes_(e) {
  var out = [], ks = Object.keys(e.cursos);
  for (var i = 0; i < ks.length; i++) {
    var d = e.cursos[ks[i]];
    if (d.turmas <= 1 || d.entm <= 0) continue;
    var minimo = Math.max(1, Math.ceil(d.entm / MAX_T));
    if (minimo >= d.turmas) continue;
    var libera = d.turmas - minimo;
    out.push(d.etapa + " - " + d.curso + " | " + d.turno + ": " +
             d.turmas + "T / " + d.entm + " alunos → " + minimo + "T" +
             " (possível liberar " + libera + (libera === 1 ? " turma)" : " turmas)"));
  }
  return out.length ? out.join("\n") : "—";
}

function resumoHoje_(e) {
  var t = [];
  var intT = e.s1I + e.s2I + e.s3I + e.outrosI;
  if (intT > 0) {
    t.push(p2_(intT) + " TURMA(S) MÉDIO INTEGRAL");
    if (e.s1I) t.push(p2_(e.s1I) + " turma(s) - 1ª série");
    if (e.s2I) t.push(p2_(e.s2I) + " turma(s) - 2ª série");
    if (e.s3I) t.push(p2_(e.s3I) + " turma(s) - 3ª série");
  }
  var blocos = [["MANHÃ", e.s1M, e.s2M, e.s3M, e.outrosM],
                ["TARDE", e.s1T, e.s2T, e.s3T, e.outrosT]];
  for (var b = 0; b < blocos.length; b++) {
    var x = blocos[b], tot = x[1] + x[2] + x[3] + x[4];
    if (tot <= 0) continue;
    if (t.length) t.push("");
    t.push(x[0] + " - " + p2_(tot) + " TURMA(S) MÉDIO");
    if (x[1]) t.push(p2_(x[1]) + " turma(s) - 1ª série");
    if (x[2]) t.push(p2_(x[2]) + " turma(s) - 2ª série");
    if (x[3]) t.push(p2_(x[3]) + " turma(s) - 3ª série");
  }
  var nEM = e.s1N + e.s2N + e.s3N + e.outrosN;
  if (nEM > 0) {
    if (t.length) t.push("");
    t.push("NOITE - " + p2_(nEM) + " TURMA(S) MÉDIO REGULAR");
  }
  if (e.efTotal > 0) {
    if (t.length) t.push("");
    if (e.efI) t.push(p2_(e.efI) + " TURMA(S) FUNDAMENTAL INTEGRAL");
    if (e.efM) t.push("MANHÃ - " + p2_(e.efM) + " TURMA(S) FUNDAMENTAL");
    if (e.efT) t.push("TARDE - " + p2_(e.efT) + " TURMA(S) FUNDAMENTAL");
    if (e.efN) t.push("NOITE - " + p2_(e.efN) + " TURMA(S) FUNDAMENTAL");
  }
  if (e.totalEJA > 0) {
    if (t.length) t.push("");
    t.push("EJA — FORA DO CÁLCULO DE SALAS");
    if (e.ejI) t.push("INTEGRAL - " + p2_(e.ejI) + " TURMA(S) DE EJA");
    if (e.ejM) t.push("MANHÃ - " + p2_(e.ejM) + " TURMA(S) DE EJA");
    if (e.ejT) t.push("TARDE - " + p2_(e.ejT) + " TURMA(S) DE EJA");
    if (e.ejN) t.push("NOITE - " + p2_(e.ejN) + " TURMA(S) DE EJA");
  }
  return t.length ? t.join("\n") : "—";
}


// ═════════════════════════════════════════════════════════════════════
//  4. UETEP  —  quantas matrículas novas e quantas salas elas pedem
// ═════════════════════════════════════════════════════════════════════

/**
 * Na aba "Matriculas por etapa" cada linha é uma turma do Ensino Profissional.
 * Turma SEM enturmados e COM pré-matrícula é oferta NOVA: ninguém está
 * matriculado ainda, mas já há gente inscrita. São essas matrículas que
 * precisam virar turma — e, portanto, sala — em 2027.
 */
function baseUETEP_(ss, gdi) {

  var sh = ss.getSheetByName(ABA_MPE);
  var linhas = [], porInep = {};
  if (!sh) return { linhas: linhas, porInep: porInep };

  var d = sh.getDataRange().getValues();

  for (var i = 1; i < d.length; i++) {

    var r = d[i];
    var inep = inep_(r[4]);
    if (!inep) continue;

    var enturm = numero_(r[16]);
    var pre    = numero_(r[18]);
    if (!(enturm === 0 && pre > 0)) continue;      // só oferta nova

    var meta  = gdi[inep] || {};
    var turno = texto_(r[7]), etapa = texto_(r[9]);
    var turma = texto_(r[12]), curso = texto_(r[14]);

    linhas.push([Number(inep), meta.gre || "", semAcento_(meta.municipio || ""),
                 meta.escola || "", etapa, curso, turma, turno,
                 pre, enturm, numero_(r[17]), Math.max(1, Math.ceil(pre / MAX_T))]);

    if (!porInep[inep]) porInep[inep] = { pre: 0, cursos: {} };
    porInep[inep].pre += pre;
    porInep[inep].cursos[curso] = (porInep[inep].cursos[curso] || 0) + pre;
  }

  // turmas previstas = teto(pré-matrículas / capacidade), curso a curso
  var ks = Object.keys(porInep);
  for (var k = 0; k < ks.length; k++) {
    var d2 = porInep[ks[k]], prev = 0, detalhe = [];
    var cs = Object.keys(d2.cursos);
    cs.sort(function (a, b) { return d2.cursos[b] - d2.cursos[a]; });
    for (var c = 0; c < cs.length; c++) {
      var p = d2.cursos[cs[c]];
      if (p > 0) prev += Math.max(1, Math.ceil(p / MAX_T));
      detalhe.push(cs[c] + ": " + p + " al.");
    }
    d2.turmasPrevistas = prev;
    d2.resumo = detalhe.length ? detalhe.join(" · ") : "—";
  }
  return { linhas: linhas, porInep: porInep };
}


/**
 * OFERTA 2027 — a base real da UETEP, em três abas:
 *
 *   UETEP · Oferta Integral 2027   A GRE · B Município · C INEP · D Entidade
 *                                  E TURMAS 2027 · F Curso · G Previsão de alunos
 *   UETEP · Continuidade 2027      A GRE · B Município · C INEP · D Entidade
 *                                  E Curso · F Etapa · G Qtd Alunos
 *                                  (uma linha = uma turma)
 *   UETEP · Subsequente 2027       A GRE · B Município · C INEP · D Entidade
 *                                  E Projeção 1ª série · F Curso · G Eixo · H Alunos
 *
 * Linhas de total no rodapé (sem INEP) são descartadas.
 */
function lerOferta2027_(ss) {

  var esc = {};

  function unidade(inep) {
    if (!esc[inep]) {
      esc[inep] = {
        inep: inep, gre: "", municipio: "", escola: "",
        of1Turmas: 0, of1Alunos: 0, of1Cursos: {},
        co2Turmas: 0, co2Alunos: 0, co2Cursos: {},
        co3Turmas: 0, co3Alunos: 0, co3Cursos: {},
        subTurmas: 0, subAlunos: 0, subCursos: {}
      };
    }
    return esc[inep];
  }

  function identifica(e, r) {
    if (!e.escola) {
      e.gre = texto_(r[0]); e.municipio = texto_(r[1]); e.escola = texto_(r[3]);
    }
  }

  // ---- 1ª série
  var sh = ss.getSheetByName(ABA_OF1);
  if (sh) {
    var d = sh.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      var k = inep_(d[i][2]); if (!k) continue;
      var e = unidade(k); identifica(e, d[i]);
      var t = numero_(d[i][4]), curso = texto_(d[i][5]);
      e.of1Turmas += t; e.of1Alunos += numero_(d[i][6]);
      e.of1Cursos[curso] = (e.of1Cursos[curso] || 0) + t;
    }
  }

  // ---- 2ª e 3ª série (continuidade): cada linha é uma turma
  sh = ss.getSheetByName(ABA_OFC);
  if (sh) {
    var d2 = sh.getDataRange().getValues();
    for (var j = 1; j < d2.length; j++) {
      var k2 = inep_(d2[j][2]); if (!k2) continue;
      var e2 = unidade(k2); identifica(e2, d2[j]);
      var curso2 = texto_(d2[j][4]);
      var etapa = texto_(d2[j][5]).toUpperCase();
      var pre = etapa.indexOf("2") === 0 ? "co2" : "co3";
      e2[pre + "Turmas"] += 1;
      e2[pre + "Alunos"] += numero_(d2[j][6]);
      e2[pre + "Cursos"][curso2] = (e2[pre + "Cursos"][curso2] || 0) + 1;
    }
  }

  // ---- subsequente
  sh = ss.getSheetByName(ABA_OFS);
  if (sh) {
    var d3 = sh.getDataRange().getValues();
    for (var m = 1; m < d3.length; m++) {
      var k3 = inep_(d3[m][2]); if (!k3) continue;
      var e3 = unidade(k3); identifica(e3, d3[m]);
      var t3 = numero_(d3[m][4]), curso3 = texto_(d3[m][5]);
      e3.subTurmas += t3; e3.subAlunos += numero_(d3[m][7]);
      e3.subCursos[curso3] = (e3.subCursos[curso3] || 0) + t3;
    }
  }

  // ---- detalhe por curso
  var ks = Object.keys(esc);
  for (var x = 0; x < ks.length; x++) {
    var u = esc[ks[x]];
    var blocos = ["of1", "co2", "co3", "sub"];
    for (var b = 0; b < blocos.length; b++) {
      var mapa = u[blocos[b] + "Cursos"];
      var cs = Object.keys(mapa);
      cs.sort(function (a, c) {
        if (mapa[c] !== mapa[a]) return mapa[c] - mapa[a];
        return a < c ? -1 : (a > c ? 1 : 0);
      });
      var lista = [];
      for (var y = 0; y < cs.length; y++) lista.push(cs[y] + " (" + mapa[cs[y]] + "T)");
      u[blocos[b] + "Detalhe"] = lista.length ? lista.join(" · ") : "—";
    }
    u.totalTurmas = u.of1Turmas + u.co2Turmas + u.co3Turmas + u.subTurmas;
  }
  return esc;
}


/**
 * Salas de 2027 com a oferta real.
 *
 *   integral = 1ª + 2ª + 3ª + subsequente (a oferta da UETEP é integral)
 *              + fundamental integral + outras turmas de EM integrais
 *   manhã e tarde dividem a mesma sala — vale o maior dos dois
 *   noite fica à parte · EJA não entra
 */
function salasOferta2027_(e, of) {
  var i27 = (of ? of.of1Turmas + of.co2Turmas + of.co3Turmas + of.subTurmas : 0)
            + e.efI + e.outrosI;
  var m27 = e.efM + e.outrosM;
  var t27 = e.efT + e.outrosT;
  var n27 = e.efN + e.outrosN;
  return { integral: i27, manha: m27, tarde: t27, noite: n27,
           salas: i27 + Math.max(m27, t27) + n27 };
}


// ═════════════════════════════════════════════════════════════════════
//  5. UEJA  —  onde a EJA acontece: matriz, anexo ou sala externa
// ═════════════════════════════════════════════════════════════════════

function baseUEJA_(matrizes, ordemM) {

  var linhas = [], resumo = {};

  for (var k = 0; k < ordemM.length; k++) {

    var e = matrizes[ordemM[k]];
    var locais = Object.keys(e.ejaPorLocal);
    if (!locais.length) continue;
    locais.sort();

    var fora = [];

    for (var i = 0; i < locais.length; i++) {

      var local = locais[i], d = e.ejaPorLocal[local];

      var cs = Object.keys(d.cursos);
      cs.sort(function (a, b) { return d.cursos[b] - d.cursos[a]; });
      var cursos = [];
      for (var c = 0; c < cs.length; c++) cursos.push(cs[c] + " (" + d.cursos[cs[c]] + "T)");

      var nomes = Object.keys(d.nomes).sort().join(" / ") || "—";

      linhas.push([Number(e.inep), e.gre, semAcento_(e.municipio), e.escola,
                   rotuloLocal_(local), nomes, d.turmas, d.entm,
                   cursos.join(" · "), local === "MATRIZ" ? "NÃO" : "SIM", ""]);

      if (local !== "MATRIZ")
        fora.push(rotuloLocal_(local) + " (" + nomes + "): " +
                  d.turmas + " turmas · " + d.entm + " matrículas");
    }
    resumo[e.inep] = fora.length ? fora.join("\n") : "Toda a EJA no prédio matriz";
  }
  return { linhas: linhas, resumo: resumo };
}


// ═════════════════════════════════════════════════════════════════════
//  6. FUSÃO ENTRE ESCOLAS DO MESMO MUNICÍPIO
// ═════════════════════════════════════════════════════════════════════

/**
 * Junta o MESMO curso de 1ª série ofertado por escolas diferentes do MESMO
 * município. Se os alunos somados couberem em menos turmas do que existem
 * hoje, há fusão possível. O destino é a escola com mais alunos no curso
 * (em caso de empate, a que tem mais salas).
 */
function fusaoEntreEscolas_(matrizes, ordemM) {

  var porMun = {}, linhas = [], resumo = {};

  for (var k = 0; k < ordemM.length; k++) {
    var e = matrizes[ordemM[k]];
    if (!e.municipio) continue;
    var m = semAcento_(e.municipio).toUpperCase();
    if (!porMun[m]) porMun[m] = [];
    porMun[m].push(e);
  }

  var muns = Object.keys(porMun).sort();

  for (var i = 0; i < muns.length; i++) {

    var escolas = porMun[muns[i]];
    if (escolas.length < 2) continue;

    var cursos = {};
    for (var j = 0; j < escolas.length; j++) {
      var ee = escolas[j], cks = Object.keys(ee.cursos1a);
      for (var c = 0; c < cks.length; c++) {
        var dd = ee.cursos1a[cks[c]];
        if (dd.turmas <= 0) continue;
        var chave = semAcento_(dd.curso).toUpperCase();
        if (!cursos[chave]) cursos[chave] = [];
        cursos[chave].push({ escola: ee, dados: dd });
      }
    }

    var chaves = Object.keys(cursos).sort();

    for (var q = 0; q < chaves.length; q++) {

      var itens = cursos[chaves[q]];
      if (itens.length < 2) continue;

      var totalAl = 0, totalTu = 0;
      for (var x = 0; x < itens.length; x++) {
        totalAl += itens[x].dados.entm;
        totalTu += itens[x].dados.turmas;
      }

      var minimo = Math.max(1, Math.ceil(totalAl / MAX_T));
      if (minimo >= totalTu) continue;               // não há ganho

      var libera = totalTu - minimo;

      itens.sort(function (a, b) {
        if (b.dados.entm !== a.dados.entm) return b.dados.entm - a.dados.entm;
        return numero_(b.escola.salasGDI) - numero_(a.escola.salasGDI);
      });

      var dest = itens[0].escola, ddest = itens[0].dados;

      for (var y = 1; y < itens.length; y++) {

        var org = itens[y].escola, dorg = itens[y].dados;

        linhas.push(["", semAcento_(dest.municipio), dest.gre, ddest.curso, "1ª Série",
                     Number(org.inep), org.escola, dorg.turmas, dorg.entm,
                     org.efTotal, org.ef9Turmas, org.ef9Matriculas, "",
                     Number(dest.inep), dest.escola, ddest.turmas, ddest.entm,
                     dest.efTotal, numero_(dest.salasGDI),
                     totalAl, totalTu, minimo, libera, "Mesmo município", ""]);

        if (!resumo[org.inep]) resumo[org.inep] = [];
        resumo[org.inep].push("→ " + ddest.curso + ": " + dorg.turmas +
          " turma(s)/" + dorg.entm + " alunos poderiam migrar para " + dest.escola +
          " (" + ddest.entm + " alunos lá). No município: " + totalAl + " alunos em " +
          totalTu + " turmas → cabem em " + minimo + " (libera " + libera + ").");

        if (!resumo[dest.inep]) resumo[dest.inep] = [];
        resumo[dest.inep].push("← " + ddest.curso + ": pode receber " +
                               dorg.entm + " aluno(s) de " + org.escola + ".");
      }
    }
  }

  var rk = Object.keys(resumo);
  for (var z = 0; z < rk.length; z++) resumo[rk[z]] = resumo[rk[z]].join("\n");

  return { linhas: linhas, resumo: resumo };
}


// ═════════════════════════════════════════════════════════════════════
//  7. PANORAMA MUNICIPAL  —  demanda real de 1ª série
// ═════════════════════════════════════════════════════════════════════

/**
 * O 9º ano que alimenta a 1ª série de 2027 vem de duas fontes:
 *   • o 9º ano que a própria rede estadual oferta no município
 *   • o 9º ano das demais redes (Censo, coluna B da aba)
 * A demanda total é a soma das duas, e as turmas necessárias saem dela.
 */
function panorama_(ss, matrizes, ordemM) {

  var est9 = {};
  for (var k = 0; k < ordemM.length; k++) {
    var e = matrizes[ordemM[k]];
    if (!e.municipio) continue;
    var m = semAcento_(e.municipio).toUpperCase();
    if (!est9[m]) est9[m] = { al: 0, tu: 0, nome: semAcento_(e.municipio) };
    est9[m].al += e.ef9Matriculas;
    est9[m].tu += e.ef9Turmas;
  }

  var sh = ss.getSheetByName(ABA_PAN);
  var saida = [], vistos = {};

  if (sh) {
    var d = sh.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      var nome = texto_(d[i][0]);
      if (!nome) continue;
      var chave = semAcento_(nome).toUpperCase();
      vistos[chave] = true;
      var todas = numero_(d[i][1]);
      var ref   = numero_(d[i][2]);
      var fonte = texto_(d[i][6]) || "CENSO";
      var est   = est9[chave] || { al: 0, tu: 0 };
      var outras  = Math.max(0, todas - est.al);
      var demanda = est.al + outras;
      var nec = demanda > 0 ? Math.ceil(demanda / CAP_1SERIE) : 0;
      saida.push([semAcento_(nome), todas, ref, est.al, est.tu, outras, demanda, nec, fonte]);
    }
  }

  // municípios que só existem na rede estadual
  var ks = Object.keys(est9).sort();
  for (var j = 0; j < ks.length; j++) {
    if (vistos[ks[j]]) continue;
    var x = est9[ks[j]];
    var n2 = x.al > 0 ? Math.ceil(x.al / CAP_1SERIE) : 0;
    saida.push([x.nome, x.al, n2, x.al, x.tu, 0, x.al, n2, "REDE ESTADUAL"]);
  }
  return saida;
}


// ═════════════════════════════════════════════════════════════════════
//  8. GRAVAÇÃO DAS BASES
// ═════════════════════════════════════════════════════════════════════

function gravarBaseTratada_(ss, matrizes, ordemM, anexos, ordemA, oferta, uejaResumo, fusaoResumo) {

  var porInep = {};
  for (var a = 0; a < ordemA.length; a++) {
    var an = anexos[ordemA[a]];
    if (!porInep[an.inep]) porInep[an.inep] = [];
    porInep[an.inep].push(an);
  }

  var h = ["INEP", "Escola", "Oferta EM 2026", "1ª Série 2026", "2ª Série 2027",
    "3ª Série 2027", "Oferta EJA 2026", "Turmas EJA", "Oferta Fundamental",
    "Parciais 2026", "Possíveis Fusões", "Salas Necessárias 2027", "Crescimento",
    "Resumo Hoje", "Resumo 2027", "Matrículas e Turmas EJA",
    "9º Ano nesta escola · turmas", "9º Ano nesta escola · matrículas",
    "Oferta 2027 · 1ª série (alunos previstos)", "Oferta 2027 · 1ª série (turmas)",
    "Oferta 2027 · 1ª série (cursos)", "EJA fora do prédio matriz",
    "EJA · turmas em anexo/sala externa", "EJA · matrículas em anexo/sala externa",
    "Fusão entre escolas (mesmo município)", "Anexos desta escola", "Qtd. de anexos",
    "Salas necessárias 2027 (oferta real)", "Município", "GRE",
    "Oferta 2027 · 2ª série (turmas)", "Oferta 2027 · 2ª série (alunos)",
    "Oferta 2027 · 3ª série (turmas)", "Oferta 2027 · 3ª série (alunos)",
    "Oferta 2027 · subsequente (turmas)", "Oferta 2027 · subsequente (alunos)",
    "Oferta 2027 · 2ª e 3ª série (cursos)", "Salas 2027 · composição",
    "Oferta 2027 · tem oferta?"];

  var linhas = [];

  for (var k = 0; k < ordemM.length; k++) {

    var e = matrizes[ordemM[k]];
    var of = oferta[e.inep] || null;

    var pSem = projetar2027_(e, 0);
    var sal = salasOferta2027_(e, of);
    var composicao = "integral " + sal.integral + " + máx(manhã " + sal.manha +
                     ", tarde " + sal.tarde + ") + noite " + sal.noite;

    var lista = porInep[e.inep] || [];
    var txtAnexos = [];
    for (var x = 0; x < lista.length; x++)
      txtAnexos.push("• " + lista[x].anexo + " [" + rotuloLocal_(lista[x].tipoAnexo) +
                     "] — " + lista[x].totalTurmas + " turma(s) · " +
                     lista[x].totalMatriculas + " matrícula(s)");

    var resumo27 = "Turmas EM 2027 (base): " + pSem.pro1 + " de 1ª, " + pSem.pro2 +
      " de 2ª, " + pSem.pro3 + " de 3ª" +
      (e.outrosEM > 0 ? ", " + e.outrosEM + " outra(s) turma(s) EM" : "") +
      ". Salas estimadas na MATRIZ: " + pSem.salasNec + ". EJA não entra neste cálculo.";

    linhas.push([Number(e.inep), e.escola, montarOfertaEM_(e), e.s1, pSem.pro2, pSem.pro3,
      montarOfertaEJA_(e),
      e.totalEJA > 0 ? (e.totalEJA + " turmas · " + e.totalEJAentm + " matrículas") : "—",
      e.efLinhas.length ? e.efLinhas.slice().sort().join("\n") : "—",
      e.parciais.length ? e.parciais.join("\n") : "—",
      montarFusoes_(e), pSem.salasNec, pSem.delta, resumoHoje_(e), resumo27,
      e.totalEJA > 0 ? (e.totalEJAentm + " matrículas | " + e.totalEJA + " turmas") : "—",
      e.ef9Turmas, e.ef9Matriculas,
      of ? of.of1Alunos : 0, of ? of.of1Turmas : 0,
      of ? of.of1Detalhe : "— sem oferta de EM em 2027",
      uejaResumo[e.inep] || "Sem oferta de EJA",
      e.ejaAnexoTurmas, e.ejaAnexoMatriculas,
      fusaoResumo[e.inep] || "—",
      txtAnexos.length ? txtAnexos.join("\n") : "—", lista.length,
      sal.salas, semAcento_(e.municipio), e.gre,
      of ? of.co2Turmas : 0, of ? of.co2Alunos : 0,
      of ? of.co3Turmas : 0, of ? of.co3Alunos : 0,
      of ? of.subTurmas : 0, of ? of.subAlunos : 0,
      of ? (of.co2Detalhe + " || " + of.co3Detalhe) : "—",
      composicao, of ? "SIM" : "NÃO"]);
  }

  gravar_(ss, ABA_BT, h, linhas,
    [100, 220, 400, 70, 70, 70, 380, 160, 280, 220, 360, 90, 80, 280, 300, 200,
     100, 110, 130, 110, 400, 300, 110, 120, 400, 320, 80, 130, 180, 110,
     110, 110, 110, 110, 120, 120, 400, 300, 100],
    [3, 7, 8, 9, 10, 11, 14, 15, 16, 21, 22, 25, 26, 37, 38]);
}


function gravarBaseAnexos_(ss, anexos, ordemA) {

  var h = ["INEP", "GRE", "Município", "Escola Matriz", "Anexo / Sala Externa",
    "Natureza do anexo", "Oferta EM 2026", "1ª Série 2026", "2ª Série 2026",
    "3ª Série 2026", "Oferta Fundamental", "Oferta EJA 2026",
    "Matrículas / Turmas EJA", "Parciais 2026", "Possíveis Fusões no próprio anexo",
    "Matrículas totais", "Turmas totais", "Resumo Hoje",
    "Salas estimadas 2027 (sem EJA)",
    "★ MANTER O ANEXO?", "★ ENCERRAR O ANEXO?", "★ REMANEJAR A OFERTA DO ANEXO?",
    "★ Para onde vai a oferta", "★ Justificativa da decisão"];

  // preserva as decisões já tomadas (colunas T..X), casando por INEP + nome do anexo
  var antigas = lerDecisoes_(ss, ABA_BA, [1, 5], 20, 5);

  var linhas = [];
  for (var k = 0; k < ordemA.length; k++) {
    var e = anexos[ordemA[k]];
    var p = projetar2027_(e, 0);
    var chave = e.inep + "|||" + e.anexo.toUpperCase().trim();
    var dec = antigas[chave] || ["", "", "", "", ""];
    linhas.push([Number(e.inep), e.gre, semAcento_(e.municipio), e.escola, e.anexo,
      rotuloLocal_(e.tipoAnexo), montarOfertaEM_(e), e.s1, e.s2, e.s3,
      e.efLinhas.length ? e.efLinhas.slice().sort().join("\n") : "—",
      montarOfertaEJA_(e),
      e.totalEJA > 0 ? (e.totalEJAentm + " matrículas | " + e.totalEJA + " turmas") : "—",
      e.parciais.length ? e.parciais.join("\n") : "—",
      montarFusoes_(e), e.totalMatriculas, e.totalTurmas, resumoHoje_(e), p.salasNec,
      dec[0], dec[1], dec[2], dec[3], dec[4]]);
  }

  var sh = gravar_(ss, ABA_BA, h, linhas,
    [100, 110, 150, 220, 260, 170, 400, 70, 70, 70, 280, 380, 180, 220, 360,
     100, 90, 300, 110, 130, 130, 190, 220, 260],
    [5, 7, 11, 12, 14, 15, 18, 23, 24]);

  if (linhas.length) {
    listaSimples_(sh, 20, linhas.length, ["SIM", "NÃO", "EM ANÁLISE"]);
    listaSimples_(sh, 21, linhas.length, ["SIM", "NÃO", "EM ANÁLISE"]);
    listaDaAba_(ss, sh, 22, linhas.length, ABA_VAL, 10);
  }
}


function gravarBaseUETEP_(ss, oferta, matrizes) {

  var h = ["INEP", "GRE", "Município", "Escola",
    "1ª série 2027 · turmas", "1ª série 2027 · cursos", "1ª série 2027 · alunos previstos",
    "2ª série 2027 · turmas", "2ª série 2027 · alunos", "2ª série 2027 · cursos",
    "3ª série 2027 · turmas", "3ª série 2027 · alunos", "3ª série 2027 · cursos",
    "Subsequente 2027 · turmas", "Subsequente 2027 · alunos", "Subsequente 2027 · cursos",
    "Total de turmas 2027", "Salas necessárias 2027 (oferta real)", "Composição das salas"];

  var ks = Object.keys(oferta);
  ks.sort(function (a, b) {
    var x = oferta[a], y = oferta[b];
    if (x.gre !== y.gre) return x.gre < y.gre ? -1 : 1;
    if (x.municipio !== y.municipio) return x.municipio < y.municipio ? -1 : 1;
    return x.escola < y.escola ? -1 : 1;
  });

  var linhas = [];
  for (var i = 0; i < ks.length; i++) {
    var o = oferta[ks[i]];
    var e = matrizes[ks[i]];
    var sal = e ? salasOferta2027_(e, o) : null;
    linhas.push([Number(o.inep), o.gre, semAcento_(o.municipio), o.escola,
      o.of1Turmas, o.of1Detalhe, o.of1Alunos,
      o.co2Turmas, o.co2Alunos, o.co2Detalhe,
      o.co3Turmas, o.co3Alunos, o.co3Detalhe,
      o.subTurmas, o.subAlunos, o.subDetalhe,
      o.totalTurmas,
      sal ? sal.salas : "",
      sal ? ("integral " + sal.integral + " + máx(manhã " + sal.manha +
             ", tarde " + sal.tarde + ") + noite " + sal.noite) : ""]);
  }

  gravar_(ss, ABA_UETEP, h, linhas,
    [100, 110, 160, 240, 110, 400, 130, 110, 110, 400, 110, 110, 400,
     110, 110, 280, 110, 130, 300],
    [6, 10, 13, 16, 19]);
}


function gravarUEJA_(ss, linhas) {
  var antigas = lerDecisoes_(ss, ABA_UEJA, [1, 5, 6], 11, 1);
  for (var i = 0; i < linhas.length; i++) {
    var ch = inep_(linhas[i][0]) + "|||" +
             String(linhas[i][4]).toUpperCase().trim() + "|||" +
             String(linhas[i][5]).toUpperCase().trim();
    linhas[i][10] = (antigas[ch] || [""])[0];
  }
  var sh = gravar_(ss, ABA_UEJA,
    ["INEP", "GRE", "Município", "Escola", "Onde a EJA acontece",
     "Nome do anexo / sala externa", "Turmas", "Matrículas",
     "Cursos ofertados neste local", "Fora do prédio matriz?", "★ Decisão 2027"],
    linhas, [100, 110, 160, 240, 200, 280, 80, 90, 420, 130, 240], [5, 6, 9, 11]);
  if (linhas.length) listaDaAba_(ss, sh, 11, linhas.length, ABA_VAL, 10);
}


function gravarFusaoEntreEscolas_(ss, linhas) {

  // ★ Aplicar? e ★ Justificativa, casados por INEP origem + INEP destino + curso
  var antigas = lerDecisoes_(ss, ABA_FE, [6, 14, 4], 1, 1);
  var just    = lerDecisoes_(ss, ABA_FE, [6, 14, 4], 25, 1);

  var n = ss.getSheetByName(ABA_IDEB) ? Math.max(2, ss.getSheetByName(ABA_IDEB).getLastRow()) : 2;

  for (var i = 0; i < linhas.length; i++) {
    var ch = inep_(linhas[i][5]) + "|||" + inep_(linhas[i][13]) + "|||" +
             String(linhas[i][3]).toUpperCase().trim();
    linhas[i][0]  = (antigas[ch] || [""])[0];
    linhas[i][24] = (just[ch] || [""])[0];
    linhas[i][12] = '=IFERROR(IF(N(VLOOKUP($B' + (i + 2) + ",'" + ABA_IDEB +
      "'!$A$2:$F$" + n + ',2,FALSE))=0,"— sem IDEB",VLOOKUP($B' + (i + 2) + ",'" +
      ABA_IDEB + "'!$A$2:$F$" + n + ',2,FALSE)),"— sem IDEB")';
  }

  var sh = gravar_(ss, ABA_FE,
    ["★ Aplicar?", "Município", "GRE", "Curso", "Série 2027",
     "INEP Origem", "Escola Origem", "Turmas na origem", "Matrículas na origem",
     "Fundamental na origem (turmas)", "9º ano na origem (turmas)",
     "9º ano na origem (matrículas)", "IDEB do município",
     "INEP Destino", "Escola Destino", "Turmas no destino", "Matrículas no destino",
     "Fundamental no destino (turmas)", "Salas do destino",
     "Alunos somados no município", "Turmas hoje no município",
     "Turmas mínimas necessárias", "Turmas que podem ser liberadas", "Tipo",
     "★ Justificativa / decisão"],
    linhas,
    [90, 170, 120, 400, 90, 100, 260, 100, 110, 130, 110, 130, 100, 100, 260,
     100, 120, 130, 90, 130, 120, 120, 130, 130, 280], [4, 7, 15, 25]);

  if (linhas.length) listaSimples_(sh, 1, linhas.length, ["SIM", "NÃO", "EM ANÁLISE"]);
}


function gravarPanorama_(ss, pan) {

  var h = ["Município", "9º Ano · todas as redes (Censo)",
    "Turmas necessárias (referência)", "9º Ano na rede estadual (2026)",
    "9º Ano estadual · turmas", "9º Ano · outras redes (estimado)",
    "Demanda total de 1ª série 2027 (alunos)",
    "Turmas necessárias 2027 (recalculado)", "Fonte do 9º Ano",
    "1ª série decidida — V2", "Saldo — V2",
    "1ª série decidida — V3", "Saldo — V3", "Cobertura — V3"];

  var nV2 = ultimaLinha_(ss, ABA_V2);
  var nV3 = ultimaLinha_(ss, ABA_V3);

  for (var i = 0; i < pan.length; i++) {
    var r = i + 2;
    pan[i][9]  = "=IF($A" + r + '="","",SUMIF(\'' + ABA_V2 + "'!$C$2:$C$" + nV2 +
                 ",$A" + r + ",'" + ABA_V2 + "'!$M$2:$M$" + nV2 + "))";
    pan[i][10] = "=IF($A" + r + '="","",$H' + r + "-$J" + r + ")";
    pan[i][11] = "=IF($A" + r + '="","",SUMIF(\'' + ABA_V3 + "'!$C$2:$C$" + nV3 +
                 ",$A" + r + ",'" + ABA_V3 + "'!$O$2:$O$" + nV3 + "))";
    pan[i][12] = "=IF($A" + r + '="","",$H' + r + "-$L" + r + ")";
    pan[i][13] = "=IF($A" + r + '="","",IF($H' + r + '=0,"",$L' + r + "/$H" + r + "))";
  }

  var sh = gravar_(ss, ABA_PAN, h, pan,
    [180, 140, 130, 140, 120, 140, 150, 150, 130, 130, 100, 130, 100, 100], []);

  if (pan.length) sh.getRange(2, 14, pan.length, 1).setNumberFormat("0%");
}


/** Cria a aba de IDEB com os municípios, sem apagar nota já digitada. */
function gravarIDEB_(ss, pan) {

  var sh = ss.getSheetByName(ABA_IDEB);
  var antigas = {};

  if (sh) {
    var d = sh.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      var m = semAcento_(d[i][0]).toUpperCase();
      if (m) antigas[m] = [d[i][1], d[i][2], d[i][3], d[i][4], d[i][5]];
    }
  }

  var linhas = [];
  for (var k = 0; k < pan.length; k++) {
    var nome = pan[k][0];
    var v = antigas[semAcento_(nome).toUpperCase()] || ["", "", "", "", ""];
    linhas.push([nome, v[0], v[1], v[2], v[3], v[4]]);
  }

  gravar_(ss, ABA_IDEB,
    ["Município", "IDEB Anos Finais (rede municipal)",
     "IDEB Anos Finais (rede estadual)", "IDEB Ensino Médio",
     "Ano de referência", "Observação"],
    linhas, [180, 160, 160, 130, 110, 300], [6]);
}


// ═════════════════════════════════════════════════════════════════════
//  9. FÓRMULAS DA V2 E DA V3
// ═════════════════════════════════════════════════════════════════════

/**
 * Reescreve SOMENTE as colunas calculadas. As colunas ★ nunca são tocadas.
 * A única exceção é a primeira carga: se uma coluna ★ numérica estiver
 * vazia, ela recebe a projeção como ponto de partida.
 */
function reescreverFormulas(silencioso) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var nBT = ultimaLinha_(ss, ABA_BT);
  var nPM = ultimaLinha_(ss, ABA_PAN);
  var nGD = ultimaLinha_(ss, ABA_GDI);
  var nID = ultimaLinha_(ss, ABA_IDEB);
  var nFE = ultimaLinha_(ss, ABA_FE);
  var nFT = ultimaLinha_(ss, ABA_FT);

  formulasV2_(ss, nBT, nPM, nGD, nID, nFT);
  formulasV3_(ss, nBT, nPM, nGD, nID, nFE);

  SpreadsheetApp.flush();
  if (!silencioso) SpreadsheetApp.getUi().alert("✅ Fórmulas da V2 e da V3 reescritas.");
}


function formulasV2_(ss, nBT, nPM, nGD, nID, nFT) {

  var sh = ss.getSheetByName(ABA_V2);
  if (!sh) return;

  var n = ultimaLinha_(ss, ABA_V2);
  if (n < 2) return;
  var qtd = n - 1;

  semear_(sh, nBT, [[13, 20], [17, 31], [18, 33], [24, 28]]);   // ★ vazias → oferta 2027
  municipios_(sh, nGD, qtd);

  var cols = {};
  cols[2]  = gd_("$A@", 3, nGD);
  cols[4]  = gd_("$A@", 5, nGD);
  cols[5]  = gd_("$A@", 14, nGD);
  cols[6]  = gd_("$A@", 15, nGD);
  cols[7]  = bt_("$A@", 9, nBT, '"—"');
  cols[8]  = 'IF($C@="","",' + pm_("$C@", 7, nPM, '"—"') + '&" alunos · "&' +
             pm_("$C@", 8, nPM, '"—"') + '&" turmas  (estadual nesta escola: "&' +
             bt_("$A@", 18, nBT, "0") + '&")")';
  cols[9]  = falta_("$C@", 11, nPM);
  cols[11] = ideb_("$C@", 2, nID);
  cols[14] = bt_("$A@", 3, nBT, '"—"');
  cols[19] = bt_("$A@", 10, nBT, '"—"');
  cols[21] = bt_("$A@", 11, nBT, '"—"');
  cols[22] = "IFERROR(IF(COUNTIF('" + ABA_FT + "'!$J$2:$J$" + nFT + ",$A@)+COUNTIF('" +
             ABA_FT + "'!$N$2:$N$" + nFT + ',$A@)=0,"Sem fusão turma a turma",' +
             "COUNTIF('" + ABA_FT + "'!$J$2:$J$" + nFT + ',$A@)&" turma(s) saem · "&' +
             "COUNTIF('" + ABA_FT + "'!$N$2:$N$" + nFT + ',$A@)&" turma(s) recebem"),"—")';
  cols[23] = gd_("$A@", 6, nGD);
  cols[25] = situacaoSala_("$W", "$X");
  cols[26] = bt_("$A@", 7, nBT, '"—"');
  cols[27] = bt_("$A@", 16, nBT, '"—"');
  cols[30] = bt_("$A@", 26, nBT, '"—"');
  cols[31] = bt_("$A@", 21, nBT, '"—"');
  cols[33] = 'IF($A@="","",$M@&" de 1ª | "&$Q@&" de 2ª | "&$R@&" de 3ª | ' +
             'Necessidade: "&$X@&" salas | "&$Y@)';
  cols[34] = gd_("$A@", 18, nGD, '"—"');
  cols[39] = bt_("$A@", 20, nBT, "0");
  cols[40] = bt_("$A@", 31, nBT, "0");
  cols[41] = bt_("$A@", 33, nBT, "0");
  cols[42] = 'IF($A@="","",IF(AND($M@=' + bt_("$A@", 20, nBT, "0") + ',$Q@=' +
             bt_("$A@", 31, nBT, "0") + ',$R@=' + bt_("$A@", 33, nBT, "0") +
             '),"igual à oferta","decidido "&$M@&"/"&$Q@&"/"&$R@&" · oferta "&' +
             bt_("$A@", 20, nBT, "0") + '&"/"&' + bt_("$A@", 31, nBT, "0") +
             '&"/"&' + bt_("$A@", 33, nBT, "0") + '))';

  aplicar_(sh, cols, qtd);

  listaDaAba_(ss, sh, 10, qtd, ABA_VAL, 1);    // ★ Fundamental
  listaDaAba_(ss, sh, 22, qtd, ABA_VAL, 8);    // ★ Validação da fusão
  listaDaAba_(ss, sh, 35, qtd, ABA_VAL, 9);    // ★ Reordenamento
}


function formulasV3_(ss, nBT, nPM, nGD, nID, nFE) {

  var sh = ss.getSheetByName(ABA_V3);
  if (!sh) return;

  var n = ultimaLinha_(ss, ABA_V3);
  if (n < 2) return;
  var qtd = n - 1;

  semear_(sh, nBT, [[15, 20], [19, 31], [20, 33], [31, 28]]);   // ★ vazias → oferta 2027
  municipios_(sh, nGD, qtd);

  var cols = {};
  cols[2]  = gd_("$A@", 3, nGD);
  cols[4]  = gd_("$A@", 5, nGD);
  cols[5]  = gd_("$A@", 14, nGD);
  cols[6]  = gd_("$A@", 15, nGD);
  cols[7]  = bt_("$A@", 9, nBT, '"—"');
  cols[8]  = 'IF($A@="","",' + bt_("$A@", 18, nBT, "0") + '&" matrícula(s) · "&' +
             bt_("$A@", 17, nBT, "0") + '&" turma(s)")';
  cols[9]  = 'IF($C@="","",' + pm_("$C@", 2, nPM, "0") + '&" alunos (estadual "&' +
             pm_("$C@", 4, nPM, "0") + '&" + outras redes "&' +
             pm_("$C@", 6, nPM, "0") + '&")")';
  cols[10] = 'IF($C@="","",' + pm_("$C@", 7, nPM, "0") + '&" alunos → "&' +
             pm_("$C@", 8, nPM, "0") + '&" turma(s) necessária(s)")';
  cols[11] = falta_("$C@", 13, nPM);
  cols[13] = ideb_("$C@", 2, nID);
  cols[16] = bt_("$A@", 3, nBT, '"—"');
  cols[21] = bt_("$A@", 20, nBT, "0");                       // 1ª série 2027 (turmas)
  cols[22] = bt_("$A@", 21, nBT, '"—"');                     // cursos da 1ª série
  cols[23] = 'IF($A@="","",' + bt_("$A@", 31, nBT, "0") + '&" de 2ª | "&' +
             bt_("$A@", 33, nBT, "0") + '&" de 3ª | "&' +
             bt_("$A@", 35, nBT, "0") + '&" subseq.")';
  cols[24] = bt_("$A@", 10, nBT, '"—"');
  cols[26] = bt_("$A@", 11, nBT, '"—"');
  cols[27] = "IFERROR(IF(COUNTIF('" + ABA_FE + "'!$F$2:$F$" + nFE + ",$A@)+COUNTIF('" +
             ABA_FE + "'!$N$2:$N$" + nFE + ',$A@)=0,"Nenhuma fusão sugerida",' +
             bt_("$A@", 25, nBT, '"—"') + '),"—")';
  cols[29] = gd_("$A@", 6, nGD);
  // cada turma que a reunião decidir a mais ou a menos move a sala na mesma medida
  cols[30] = 'IF($A@="","",MAX(0,' + bt_("$A@", 28, nBT, "0") +
             "+($O@-" + bt_("$A@", 20, nBT, "0") + ")" +
             "+($S@-" + bt_("$A@", 31, nBT, "0") + ")" +
             "+($T@-" + bt_("$A@", 33, nBT, "0") + ")))";
  cols[32] = situacaoSala_("$AC", "$AE");
  cols[33] = bt_("$A@", 7, nBT, '"—"');
  cols[34] = bt_("$A@", 16, nBT, '"—"');
  cols[35] = bt_("$A@", 22, nBT, '"—"');
  cols[38] = bt_("$A@", 26, nBT, '"—"');
  cols[39] = bt_("$A@", 27, nBT, "0");
  cols[43] = 'IF($A@="","",$O@&" de 1ª | "&$S@&" de 2ª | "&$T@&" de 3ª | UETEP "&$W@' +
             '&" | Necessidade: "&$AE@&" salas | "&$AF@)';
  cols[44] = gd_("$A@", 18, nGD, '"—"');
  cols[49] = 'IF($A@="","",' + bt_("$A@", 32, nBT, "0") + "+" +
             bt_("$A@", 34, nBT, "0") + '&" aluno(s)")';
  cols[50] = 'IF($A@="","",IF(AND($O@=' + bt_("$A@", 20, nBT, "0") + ',$S@=' +
             bt_("$A@", 31, nBT, "0") + ',$T@=' + bt_("$A@", 33, nBT, "0") +
             '),"igual à oferta","decidido "&$O@&"/"&$S@&"/"&$T@&" · oferta "&' +
             bt_("$A@", 20, nBT, "0") + '&"/"&' + bt_("$A@", 31, nBT, "0") +
             '&"/"&' + bt_("$A@", 33, nBT, "0") + '))';

  aplicar_(sh, cols, qtd);

  listaDaAba_(ss, sh, 12, qtd, ABA_VAL, 1);
  listaDaAba_(ss, sh, 28, qtd, ABA_VAL, 8);
  listaSimples_(sh, 40, qtd, ["SIM", "NÃO", "EM ANÁLISE"]);
  listaSimples_(sh, 41, qtd, ["SIM", "NÃO", "EM ANÁLISE"]);
  listaDaAba_(ss, sh, 42, qtd, ABA_VAL, 10);
  listaDaAba_(ss, sh, 45, qtd, ABA_VAL, 9);
}


// ─────────────────────────────────────────── montadores de fórmula

function bt_(chave, col, n, err) {
  return "IFERROR(VLOOKUP(" + chave + ",'" + ABA_BT + "'!$A$2:$AM$" + n + "," +
         col + ",FALSE)," + (err || '""') + ")";
}
function pm_(chave, col, n, err) {
  return "IFERROR(VLOOKUP(" + chave + ",'" + ABA_PAN + "'!$A$2:$N$" + n + "," +
         col + ",FALSE)," + (err || '""') + ")";
}
function gd_(chave, col, n, err) {
  return "IFERROR(VLOOKUP(" + chave + ",'" + ABA_GDI + "'!$A$2:$AL$" + n + "," +
         col + ",FALSE)," + (err || '""') + ")";
}
function ideb_(chave, col, n) {
  var v = "VLOOKUP(" + chave + ",'" + ABA_IDEB + "'!$A$2:$F$" + n + "," + col + ",FALSE)";
  return 'IFERROR(IF(N(' + v + ')=0,"— sem IDEB cadastrado",' + v +
         '),"— sem IDEB cadastrado")';
}
function falta_(chave, colSaldo, n) {
  var v = pm_(chave, colSaldo, n, '""');
  return 'IF(' + chave + '="","",IF(' + v + '="","—",IF(' + v + '>0,' +
         '"FALTAM DISTRIBUIR "&' + v + '&" TURMA(S)",IF(' + v + '=0,' +
         '"DEMANDA MUNICIPAL ATENDIDA","EXCEDE A DEMANDA EM "&ABS(' + v + ')&" TURMA(S)"))))';
}
function situacaoSala_(colExiste, colNec) {
  var e = colExiste + "@", x = colNec + "@";
  return 'IF(' + x + '="","",IF(' + e + '>' + x + ',(' + e + '-' + x +
         ')&" SALA(S) OCIOSA(S)",IF(' + e + '<' + x + ',"CONSTRUIR "&(' + x + '-' + e +
         ')&" SALA(S)","QUANTIDADE ADEQUADA")))';
}


/** Escreve uma fórmula por linha, trocando @ pelo número da linha. */
function aplicar_(sh, cols, qtd) {
  var chaves = Object.keys(cols);
  for (var c = 0; c < chaves.length; c++) {
    var col = Number(chaves[c]), base = cols[chaves[c]], f = [];
    for (var i = 0; i < qtd; i++)
      f.push(["=" + base.replace(/@/g, String(i + 2))]);
    sh.getRange(2, col, qtd, 1).setFormulas(f);
  }
}


/** Município sem acento — é a chave que casa com Panorama e IDEB. */
function municipios_(sh, nGD, qtd) {
  var ineps = sh.getRange(2, 1, qtd, 1).getValues();
  var gdi = mapaGDI_(SpreadsheetApp.getActiveSpreadsheet());
  var out = [];
  for (var i = 0; i < qtd; i++) {
    var k = inep_(ineps[i][0]);
    out.push([k && gdi[k] ? semAcento_(gdi[k].municipio) : ""]);
  }
  sh.getRange(2, 3, qtd, 1).setValues(out);
}


/** Preenche coluna ★ numérica só quando ela ainda estiver vazia. */
function semear_(sh, nBT, pares) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bt = ss.getSheetByName(ABA_BT);
  if (!bt) return;

  var d = bt.getDataRange().getValues();
  var mapa = {};
  for (var i = 1; i < d.length; i++) {
    var k = inep_(d[i][0]);
    if (k) mapa[k] = d[i];
  }

  var n = sh.getLastRow();
  if (n < 2) return;
  var qtd = n - 1;
  var ineps = sh.getRange(2, 1, qtd, 1).getValues();

  for (var p = 0; p < pares.length; p++) {
    var colDest = pares[p][0], colOrig = pares[p][1];
    var atual = sh.getRange(2, colDest, qtd, 1).getValues();
    var mudou = false;
    for (var r = 0; r < qtd; r++) {
      if (atual[r][0] !== "" && atual[r][0] !== null) continue;
      var k2 = inep_(ineps[r][0]);
      if (!k2 || !mapa[k2]) continue;
      atual[r][0] = mapa[k2][colOrig - 1];
      mudou = true;
    }
    if (mudou) sh.getRange(2, colDest, qtd, 1).setValues(atual);
  }
}


// ═════════════════════════════════════════════════════════════════════
//  10. VALIDAÇÕES  —  as listas dos menus suspensos
// ═════════════════════════════════════════════════════════════════════

function criarValidacoes() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var cols = [
    { t: "FUNDAMENTAL", cor: "#0891b2", bg: "#ecfeff", dados: [
      "EF Inicial - 1º ano", "EF Inicial - 2º ano", "EF Inicial - 3º ano",
      "EF Inicial - 4º ano", "EF Inicial - 5º ano", "EF Final - 6º ano",
      "EF Final - 7º ano", "EF Final - 8º ano", "EF Final - 9º ano",
      "Não oferta Fundamental", "Encerrar oferta de Fundamental"] },
    { t: "EJA · Seg I\nEF Inicial", cor: "#b45309", bg: "#fef3c7", dados: [
      "Módulo I — 1º ano (Alfabetiza PI) · Início",
      "Módulo II — 2º ano · Continuidade", "Módulo III — 3º ano · Continuidade",
      "Módulo IV — 4º ano · Continuidade", "Módulo V — 5º ano · Continuidade",
      "Módulo VI — 5º ano · Continuidade"] },
    { t: "EJA · Seg II\nEF Final", cor: "#92400e", bg: "#fefce8", dados: [
      "Módulo VII — 6º ano · Continuidade", "Módulo VIII — 7º ano · Continuidade",
      "Módulo IX — 8º ano · Continuidade", "Módulo X — 9º ano · Continuidade"] },
    { t: "EJA · Seg III\nEM Técnico", cor: "#7c2d12", bg: "#fff7ed", dados: [
      "Módulo I — 1ª série · Início", "Módulo II — 1ª série · Início",
      "Módulo III — 2ª série · Continuidade", "Módulo IV — 2ª série · Continuidade",
      "Módulo V — 3ª série · Finalização"] },
    { t: "EJA · Seg III\nEM FIC", cor: "#78350f", bg: "#fffbeb", dados: [
      "Módulo I — 1ª série · Início", "Módulo II — 2ª série · Continuidade",
      "Módulo III — 3ª série · Continuidade", "Módulo IV — 3ª série · Finalização"] },
    { t: "TURNOS", cor: "#6b7280", bg: "#f3f4f6",
      dados: ["Integral", "Manhã", "Tarde", "Noite"] },
    { t: "MOVIMENTO\nEJA", cor: "#7c3aed", bg: "#f5f3ff",
      dados: ["Em continuidade", "Finalizando", "A iniciar (entrada nova)"] },
    { t: "VALIDAÇÃO\nDA FUSÃO", cor: "#7c3aed", bg: "#f5f3ff", dados: [
      "Fusão confirmada", "Fusão recusada", "Fusão em análise",
      "Fusão apenas na própria escola", "Fusão entre escolas do município",
      "Depende de transporte escolar", "Sem fusão possível"] },
    { t: "REORDENAMENTO\n2027", cor: "#dc2626", bg: "#fef2f2", dados: [
      "Manter", "Ampliar oferta", "Reduzir oferta", "Fundir com outra escola",
      "Receber turmas de outra escola", "Mudar de turno", "Encerrar oferta"] },
    { t: "DESTINO DA OFERTA\nDO ANEXO", cor: "#9333ea", bg: "#faf5ff", dados: [
      "Manter no anexo", "Trazer para o prédio matriz",
      "Transferir para outra escola estadual", "Transferir para a rede municipal",
      "Encerrar a oferta", "Em análise"] }
  ];

  var sh = ss.getSheetByName(ABA_VAL);
  if (sh) { sh.clear(); } else { sh = ss.insertSheet(ABA_VAL); }

  var maxL = 0;
  for (var c = 0; c < cols.length; c++)
    if (cols[c].dados.length > maxL) maxL = cols[c].dados.length;

  var mat = [[]];
  for (var c2 = 0; c2 < cols.length; c2++) mat[0].push(cols[c2].t);
  for (var r = 0; r < maxL; r++) {
    var linha = [];
    for (var c3 = 0; c3 < cols.length; c3++)
      linha.push(r < cols[c3].dados.length ? cols[c3].dados[r] : "");
    mat.push(linha);
  }

  sh.getRange(1, 1, mat.length, cols.length).setValues(mat);
  sh.getRange(1, 1, mat.length, cols.length).setFontFamily("Arial").setFontSize(9);
  sh.getRange(1, 1, 1, cols.length).setFontWeight("bold").setWrap(true)
    .setHorizontalAlignment("center").setFontSize(8);

  for (var c4 = 0; c4 < cols.length; c4++) {
    sh.getRange(1, c4 + 1).setBackground(cols[c4].cor).setFontColor("#ffffff");
    if (cols[c4].dados.length)
      sh.getRange(2, c4 + 1, cols[c4].dados.length, 1).setBackground(cols[c4].bg);
    sh.setColumnWidth(c4 + 1, 240);
  }
  sh.setFrozenRows(1);

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("✅ Validações criadas/atualizadas.");
}


function listaDaAba_(ss, sh, col, qtd, abaLista, colLista) {
  if (qtd < 1) return;
  var origem = ss.getSheetByName(abaLista);
  if (!origem) return;
  var n = origem.getLastRow();
  if (n < 2) return;
  var regra = SpreadsheetApp.newDataValidation()
    .requireValueInRange(origem.getRange(2, colLista, n - 1, 1), true)
    .setAllowInvalid(true).build();
  sh.getRange(2, col, qtd, 1).setDataValidation(regra);
}

function listaSimples_(sh, col, qtd, itens) {
  if (qtd < 1) return;
  var regra = SpreadsheetApp.newDataValidation()
    .requireValueInList(itens, true).setAllowInvalid(true).build();
  sh.getRange(2, col, qtd, 1).setDataValidation(regra);
}


// ═════════════════════════════════════════════════════════════════════
//  11. MENU DE CURSOS ACUMULATIVO
// ═════════════════════════════════════════════════════════════════════

/**
 * Faz a coluna de cursos acumular as escolhas em vez de substituir:
 *
 *   clica → "Logística"      → célula fica "Logística"
 *   clica → "2"              → célula fica "Logística, 2"
 *   clica → "Agropecuária"   → célula fica "Logística, 2, Agropecuária"
 *
 * A vantagem sobre a seleção múltipla nativa é que a MESMA quantidade
 * pode aparecer duas vezes ("..., 1, ..., 1"), coisa que o menu de
 * caixinhas do Google não permite.
 */

var COL_CURSOS = { };
COL_CURSOS[ABA_V1] = [15];        // ★ CURSOS 1ª SÉRIE 2027
COL_CURSOS[ABA_V2] = [15];
COL_CURSOS[ABA_V3] = [17];
var SEPARADOR = ", ";

function onEdit(e) {

  if (!e || !e.range) return;

  var cel = e.range;
  var aba = cel.getSheet().getName();

  var cols = COL_CURSOS[aba];
  if (!cols) return;
  if (cols.indexOf(cel.getColumn()) === -1) return;
  if (cel.getRow() < 2) return;
  if (cel.getNumRows() !== 1 || cel.getNumColumns() !== 1) return;

  var escolhido = e.value;      // o que acabou de ser marcado no menu
  var anterior  = e.oldValue;   // o que já estava na célula

  if (escolhido === undefined) return;                       // apagou ou colou em bloco
  if (anterior === undefined || anterior === "") return;     // primeira escolha entra sozinha
  if (escolhido === anterior) return;                        // repetiu o conteúdo inteiro

  cel.setValue(anterior + SEPARADOR + escolhido);
}


/** Tira o último item da célula selecionada. */
function desfazerUltimoCurso() {

  var cel = SpreadsheetApp.getActiveRange();
  var aba = cel.getSheet().getName();
  var cols = COL_CURSOS[aba];

  if (!cols || cols.indexOf(cel.getColumn()) === -1) {
    SpreadsheetApp.getUi().alert("Selecione uma célula da coluna ★ CURSOS 1ª SÉRIE 2027.");
    return;
  }

  var partes = String(cel.getValue()).split(SEPARADOR);
  var limpo = [];
  for (var i = 0; i < partes.length; i++)
    if (String(partes[i]).trim() !== "") limpo.push(partes[i]);

  limpo.pop();
  cel.setValue(limpo.join(SEPARADOR));
}


// ═════════════════════════════════════════════════════════════════════
//  12. UTILITÁRIOS
// ═════════════════════════════════════════════════════════════════════

/** Grava uma aba inteira: limpa, escreve, formata e recria o filtro. */
function gravar_(ss, nome, headers, linhas, larguras, wrapCols) {

  var sh = ss.getSheetByName(nome);
  if (sh) {
    var filtro = sh.getFilter();
    if (filtro) filtro.remove();
    sh.clear();
  } else {
    sh = ss.insertSheet(nome);
  }

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (linhas.length) {
    var alvo = sh.getRange(2, 1, linhas.length, headers.length);
    alvo.setValues(linhas);   // string começando com "=" vira fórmula
    alvo.setFontFamily("Arial").setFontSize(9).setVerticalAlignment("top");
    for (var w = 0; w < (wrapCols || []).length; w++)
      sh.getRange(2, wrapCols[w], linhas.length, 1).setWrap(true);
  }

  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold").setBackground("#374151").setFontColor("#ffffff")
    .setFontFamily("Arial").setFontSize(8).setWrap(true)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  for (var c = 0; c < (larguras || []).length && c < headers.length; c++)
    sh.setColumnWidth(c + 1, larguras[c]);

  sh.setFrozenRows(1);
  if (linhas.length)
    sh.getRange(1, 1, linhas.length + 1, headers.length).createFilter();

  return sh;
}


/**
 * Lê as colunas de decisão de uma aba antes de reescrevê-la, para que a
 * atualização das bases não apague o que já foi decidido.
 *   chaveCols  = colunas que, juntas, identificam a linha
 *   primeiraCol / quantas = bloco de colunas de decisão a preservar
 */
function lerDecisoes_(ss, nome, chaveCols, primeiraCol, quantas) {

  var sh = ss.getSheetByName(nome);
  var mapa = {};
  if (!sh) return mapa;

  var n = sh.getLastRow();
  if (n < 2) return mapa;

  var d = sh.getRange(1, 1, n, sh.getLastColumn()).getValues();

  for (var i = 1; i < d.length; i++) {

    var partes = [];
    for (var c = 0; c < chaveCols.length; c++) {
      var v = d[i][chaveCols[c] - 1];
      partes.push(chaveCols[c] === 1 || typeof v === "number"
        ? inep_(v) : String(v === undefined || v === null ? "" : v).toUpperCase().trim());
    }
    var chave = partes.join("|||");
    if (chave.replace(/\|/g, "").trim() === "") continue;

    var vals = [];
    for (var q = 0; q < quantas; q++) {
      var x = d[i][primeiraCol - 1 + q];
      vals.push(x === undefined || x === null ? "" : x);
    }
    mapa[chave] = vals;
  }
  return mapa;
}


function ultimaLinha_(ss, nome) {
  var sh = ss.getSheetByName(nome);
  if (!sh) return 2;
  var n = sh.getLastRow();
  return n < 2 ? 2 : n;
}


function somarTurno_(e, prefixo, turno, qtd) {
  var suf = { "INTEGRAL": "I", "MANHÃ": "M", "MANHA": "M",
              "TARDE": "T", "NOITE": "N" }[turno];
  if (!suf) return;
  var campo = prefixo + suf;
  if (typeof e[campo] !== "number") e[campo] = 0;
  e[campo] += qtd;
}

function abreviarTurno_(t) {
  return ({ "INTEGRAL": "I", "MANHÃ": "M", "MANHA": "M",
            "TARDE": "T", "NOITE": "N" })[t] || (t ? t.charAt(0) : "");
}

function texto_(v) {
  return String(v === null || v === undefined ? "" : v).trim();
}

function espacos_(v) {
  return texto_(v).replace(/\s+/g, " ");
}

function numero_(v) {
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

function inep_(v) {
  var t = texto_(v);
  if (!t) return "";
  var n = Number(t);
  return isNaN(n) ? t : String(Math.round(n));
}

function semAcento_(v) {
  return texto_(v).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
}

function p2_(n) {
  n = Math.round(n);
  return (n < 10 ? "0" : "") + n;
}

function erro_(msg) {
  SpreadsheetApp.getUi().alert("❌ " + msg);
}
