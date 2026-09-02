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
//  0. COLUNAS POR NOME
// ═════════════════════════════════════════════════════════════════════

/*
 * O script não usa letra nem número de coluna: procura pelo NOME que está
 * escrito na linha 1 da aba. Assim, mover ou inserir coluna não quebra nada.
 *
 * Cada entrada abaixo é [nome do cabeçalho, posição de reserva]. A posição só
 * é usada quando o nome não aparece na linha 1 — o que acontece nas abas que
 * chegam por IMPORTRANGE e às vezes vêm sem cabeçalho legível. Quando isso
 * ocorre, fica um aviso no registro de execução.
 *
 * Para renomear uma coluna na planilha, mude o nome aqui também. É o único
 * lugar do arquivo que precisa saber como as colunas se chamam.
 */

var NOMES_COLUNAS = {};

NOMES_COLUNAS[ABA_TURMAS] = {
  inep:["INEP",1], escola:["Escola",2], anexo:["Anexo",3], curso:["Curso",4],
  etapa:["Etapa",5], organizacao:["Organização da turma",6], periodo:["Período",7],
  turno:["Turno",8], enturmados:["Enturmados",9], turmas:["Turmas",10]
};

NOMES_COLUNAS[ABA_GDI] = {
  inep:["INEP",1], gre:["GRE",3], municipio:["MUNICIPIO",4], escola:["NOME ENTIDADE",5],
  salas:["Nº SLS DE AULA",6], proj2026:["PROJEÇÃO TURMA(S) 2026",14],
  proj2027:["PROJEÇÃO TURMA(S) 2027",15], escolaProxima:["ESCOLA MAIS PRÓXIMA",18]
};

NOMES_COLUNAS[ABA_MPE] = {
  inep:["Inep Escola",5], turno:["Turno",8], etapa:["Abreviação Etapa",10],
  turma:["Nome Turma",13], curso:["Nome Curso",15], ativas:["Ativas",16],
  enturmados:["Enturmados",17], cursando:["Cursando",18], pre:["Pré Matrícula",19]
};

NOMES_COLUNAS[ABA_OF1] = {
  gre:["GRE",1], municipio:["Municipio",2], inep:["INEP da Entidade",3],
  escola:["Entidade",4], turmas:["TURMAS 2027",5],
  curso:["CURSOS PRE-DEFINIDOS 2027",6], alunos:["PREVISÃO DE ALUNOS",7]
};

NOMES_COLUNAS[ABA_OFC] = {
  gre:["GRE",1], municipio:["Municipio",2], inep:["INEP da Entidade",3],
  escola:["Entidade",4], curso:["Curso",5], etapa:["Etapa",6], alunos:["Qtd Alunos",7]
};

NOMES_COLUNAS[ABA_OFS] = {
  gre:["GRE",1], municipio:["Municipio",2], inep:["INEP da Entidade",3],
  escola:["Entidade",4], turmas:["PROJEÇÃO 1ª SÉRIE",5],
  curso:["CURSOS PRE-DEFINIDOS 2027",6], eixo:["EIXO",7], alunos:["ALUNOS",8]
};

NOMES_COLUNAS[ABA_PAN] = {
  municipio:["Município",1], todasRedes:["9º Ano · todas as redes (Censo)",2],
  refTurmas:["Turmas necessárias (referência)",3],
  est9Alunos:["9º Ano na rede estadual (2026)",4],
  est9Turmas:["9º Ano estadual · turmas",5],
  outrasRedes:["9º Ano · outras redes (estimado)",6],
  demanda:["Demanda total de 1ª série 2027 (alunos)",7],
  necessarias:["Turmas necessárias 2027 (recalculado)",8],
  fonte:["Fonte do 9º Ano",9],
  decididaV2:["1ª série decidida — V2",10], saldoV2:["Saldo — V2",11],
  decididaV3:["1ª série decidida — V3",12], saldoV3:["Saldo — V3",13],
  cobertura:["Cobertura — V3",14]
};

NOMES_COLUNAS[ABA_BT] = {
  inep:["INEP",1], escola:["Escola",2], ofertaEM:["Oferta EM 2026",3],
  s1_2026:["1ª Série 2026",4], s2_proj:["2ª Série 2027",5], s3_proj:["3ª Série 2027",6],
  ofertaEJA:["Oferta EJA 2026",7], turmasEJA:["Turmas EJA",8],
  ofertaFund:["Oferta Fundamental",9], parciais:["Parciais 2026",10],
  fusoes:["Possíveis Fusões",11], salasProj:["Salas Necessárias 2027",12],
  crescimento:["Crescimento",13], resumoHoje:["Resumo Hoje",14],
  resumo2027:["Resumo 2027",15], matrTurmasEJA:["Matrículas e Turmas EJA",16],
  ano9Turmas:["9º Ano nesta escola · turmas",17],
  ano9Matr:["9º Ano nesta escola · matrículas",18],
  of1Alunos:["Oferta 2027 · 1ª série (alunos previstos)",19],
  of1Turmas:["Oferta 2027 · 1ª série (turmas)",20],
  of1Cursos:["Oferta 2027 · 1ª série (cursos)",21],
  ejaFora:["EJA fora do prédio matriz",22],
  ejaAnexoTurmas:["EJA · turmas em anexo/sala externa",23],
  ejaAnexoMatr:["EJA · matrículas em anexo/sala externa",24],
  fusaoEntre:["Fusão entre escolas (mesmo município)",25],
  anexos:["Anexos desta escola",26], qtdAnexos:["Qtd. de anexos",27],
  salasReal:["Salas necessárias 2027 (oferta real)",28],
  municipio:["Município",29], gre:["GRE",30],
  co2Turmas:["Oferta 2027 · 2ª série (turmas)",31],
  co2Alunos:["Oferta 2027 · 2ª série (alunos)",32],
  co3Turmas:["Oferta 2027 · 3ª série (turmas)",33],
  co3Alunos:["Oferta 2027 · 3ª série (alunos)",34],
  subTurmas:["Oferta 2027 · subsequente (turmas)",35],
  subAlunos:["Oferta 2027 · subsequente (alunos)",36],
  co23Cursos:["Oferta 2027 · 2ª e 3ª série (cursos)",37],
  composicao:["Salas 2027 · composição",38], temOferta:["Oferta 2027 · tem oferta?",39],
  fund2026:["Fundamental 2026 · turmas",40], soEJA:["Só EJA (CEJA)?",41],
  ejaMatrizTurmas:["EJA 2026 · turmas no prédio matriz",42],
  ejaMatrizMatr:["EJA 2026 · matrículas no prédio matriz",43],
  zeradas:["Turmas zeradas descartadas",44]
};

NOMES_COLUNAS[ABA_IDEB] = {
  municipio:["Município",1], ideb:["IDEB Ensino Médio",2],
  ano:["Ano de referência",3], obs:["Observação",4]
};

NOMES_COLUNAS[ABA_VAL] = {
  fundamental:["FUNDAMENTAL",1], ejaSeg1:["EJA · Seg I",2], ejaSeg2:["EJA · Seg II",3],
  ejaSeg3Tec:["EJA · Seg III\nEM Técnico",4], ejaSeg3Fic:["EJA · Seg III\nEM FIC",5],
  turnos:["TURNOS",6], movEJA:["MOVIMENTO",7], validFusao:["VALIDAÇÃO",8],
  reordenamento:["REORDENAMENTO",9], destinoAnexo:["DESTINO DA OFERTA",10],
  cursosQtd:["CURSOS 1ª SÉRIE",11], fundQtd:["FUNDAMENTAL\ncom quantidade",12],
  decisaoAnexo:["DECISÃO SOBRE",13]
};

NOMES_COLUNAS[ABA_V2] = {
  inep:["INEP",1], gre:["GRE",2], municipio:["Município",3], escola:["Escola",4],
  turmasHoje:["Turmas hoje",5], turmas2027:["Turmas 2027",6],
  ofertaFund:["Oferta\nFundamental",7], ano9:["9º Ano do município",8],
  faltam:["Ainda faltam no município",9], estrelaFund:["★ FUNDAMENTAL",10],
  ideb:["IDEB",11], s1_2026:["1ª Série\n2026",12], s1_2027:["★ 1ª SÉRIE 2027",13],
  cursos2026:["Cursos EM 2026",14], cursos2027:["★ CURSOS 1ª SÉRIE",15],
  alteracaoEMI:["ALTERAÇÃO DE CURSOS EMI",16], s2_2027:["2ª Série\n2027",17],
  s3_2027:["3ª Série\n2027",18], parciais:["Parciais 2026",19],
  parciaisJust:["Parciais 2026 justificativa",20], fusoes:["Possíveis\nFusões",21],
  validFusao:["Validações Fusões Turmas",22], salasExist:["Salas\nexistentes",23],
  salasDecide:["★ SALAS NECESSÁRIAS",24], situacao:["Situação\nda sala",25],
  ofertaEJA:["Oferta EJA",26], matrEJA:["Matrículas / Turmas",27],
  ejaTurmas:["★ EJA Turmas",28], ejaCursos:["★ EJA Cursos",29],
  anexos:["ANEXOS",30], cursos:["CURSOS",31], reservado:["(reservado)",32],
  resumo:["Resumo 2027",33], escolaProxima:["Escola Próxima",34],
  reord:["Reordenamento\n2027",35], justificativa:["Justificativa",36],
  observacao:["Observação",37], pronto:["✓ Pronto",38],
  of1:["Oferta 2027 ·\n1ª série (turmas)",39], of2:["Oferta 2027 ·\n2ª série (turmas)",40],
  of3:["Oferta 2027 ·\n3ª série (turmas)",41], divergencia:["★ x Oferta 2027",42]
};

NOMES_COLUNAS[ABA_V3] = {
  inep:["INEP",1], gre:["GRE",2], municipio:["Município",3], escola:["Escola",4],
  turmasHoje:["Turmas hoje",5], turmas2027:["Turmas 2027",6],
  ofertaFund:["Oferta\nFundamental",7],
  fundEtapas:["★ FUNDAMENTAL 2027",8], fundTurmas:["★ TURMAS FUNDAMENTAL",9],
  fundTotal:["Fundamental 2027 ·",10],
  ano9Escola:["9º Ano NESTA escola",11], ano9Municipio:["9º Ano do município",12],
  demanda:["Demanda total de",13], faltam:["Ainda faltam",14], ideb:["IDEB DA ESCOLA",15],
  s1_2026:["1ª Série\n2026",16], s1_2027:["★ 1ª SÉRIE 2027",17],
  cursos2026:["Cursos EM 2026",18], cursos2027:["★ CURSOS 1ª SÉRIE 2027",19],
  cursosTurmas:["★ TURMAS 1ª SÉRIE 2027",20], cursosTotal:["1ª série 2027 ·",21],
  alteracaoEMI:["★ ALTERAÇÃO DE",22], s2_2027:["★ 2ª SÉRIE 2027",23],
  s3_2027:["★ 3ª SÉRIE 2027",24],
  of1:["Oferta 2027 ·\n1ª série (turmas)",25], ofCursos:["Oferta 2027 ·\ncursos da 1ª série",26],
  of23sub:["Oferta 2027 ·\n2ª / 3ª / subseq.",27],
  parciais:["Parciais 2026",28], parciaisJust:["★ Parciais 2026",29],
  fusaoPropria:["Possíveis fusões",30], fusaoEntre:["Fusão ENTRE ESCOLAS",31],
  fusaoValid:["★ Validação",32],
  salasExist:["Salas\nexistentes",33], salasCalc:["Salas necessárias\n2027 (calculado)",34],
  salasDecide:["★ SALAS NECESSÁRIAS",35], salasSituacao:["Situação\nda sala",36],
  ejaOferta:["Oferta EJA 2026",37], ejaMatriculas:["Matrículas / Turmas EJA",38],
  ejaFora:["EJA fora do prédio",39], ejaTurmas:["★ EJA TURMAS 2027",40],
  anexosLista:["Anexos desta",41], anexosQtd:["Qtd. de\nanexos",42],
  anexosDecisao:["★ DECISÃO SOBRE",43],
  resumo:["Resumo 2027",44], escolaProxima:["Escola Próxima",45],
  reord:["★ Reordenamento",46], justificativa:["★ Justificativa",47],
  observacao:["★ Observação",48], pronto:["✓ Pronto",49],
  ofAlunos23:["Oferta 2027 ·\nalunos reais 2ª/3ª",50], divergencia:["★ x Oferta 2027",51]
};


var _cacheColunas = {};

/** Limpa o cache de cabeçalhos. Chamado no começo de cada rodada. */
function limparCacheColunas_() { _cacheColunas = {}; }


/** Cabeçalho normalizado: sem acento, sem quebra de linha, maiúsculo. */
function chaveCabecalho_(v) {
  return semAcento_(String(v === null || v === undefined ? "" : v))
           .replace(/\s+/g, " ").trim();
}


/**
 * Resolve todas as colunas de uma aba pelo nome escrito na linha 1.
 * Devolve { apelido: número da coluna }, com base 1.
 */
function colunasDe_(ss, aba) {

  if (_cacheColunas[aba]) return _cacheColunas[aba];

  var mapa = {};
  var esperados = NOMES_COLUNAS[aba] || {};
  var sh = ss.getSheetByName(aba);

  var presentes = {};
  if (sh && sh.getLastColumn() > 0) {
    var cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    for (var c = 0; c < cab.length; c++) {
      var k = chaveCabecalho_(cab[c]);
      if (k && presentes[k] === undefined) presentes[k] = c + 1;
    }
  }

  var apelidos = Object.keys(esperados);
  for (var i = 0; i < apelidos.length; i++) {

    var nome = esperados[apelidos[i]][0];
    var reserva = esperados[apelidos[i]][1];
    var alvo = chaveCabecalho_(nome);

    if (presentes[alvo] !== undefined) {          // nome bate exatamente
      mapa[apelidos[i]] = presentes[alvo];
      continue;
    }

    var achou = 0;                                 // senão, procura por começo
    var chaves = Object.keys(presentes);
    for (var j = 0; j < chaves.length; j++) {
      if (chaves[j].indexOf(alvo) === 0) { achou = presentes[chaves[j]]; break; }
    }

    if (achou) {
      mapa[apelidos[i]] = achou;
    } else {
      mapa[apelidos[i]] = reserva;
      if (sh) Logger.log('Coluna "' + nome + '" não encontrada em "' + aba +
                         '". Usando a posição de reserva ' + letra_(reserva) + '.');
    }
  }

  _cacheColunas[aba] = mapa;
  return mapa;
}


/** Número da coluna (base 1) pelo apelido. */
function col_(ss, aba, apelido) {
  var n = colunasDe_(ss, aba)[apelido];
  if (!n) throw new Error('Coluna "' + apelido + '" não está declarada para a aba "' + aba + '".');
  return n;
}

/** Índice base 0, para ler de uma matriz de getValues(). */
function ix_(ss, aba, apelido) { return col_(ss, aba, apelido) - 1; }

/** Letra da coluna, para montar fórmula: 1 → A, 28 → AB. */
function letra_(n) {
  var s = "";
  while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - r - 1) / 26; }
  return s;
}

/** Referência absoluta de coluna para dentro de fórmula: "$AB@". */
function ref_(ss, aba, apelido) { return "$" + letra_(col_(ss, aba, apelido)) + "@"; }


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
  limparCacheColunas_();          // as abas podem ter mudado desde a última vez

  var src = ss.getSheetByName(ABA_TURMAS);
  if (!src) { erro_('Aba "' + ABA_TURMAS + '" não encontrada.'); return null; }

  var dados = src.getDataRange().getValues();
  var T     = colunasDe_(ss, ABA_TURMAS);
  var gdi   = mapaGDI_(ss);

  var matrizes = {}, ordemM = [];
  var anexos   = {}, ordemA = [];

  for (var i = 1; i < dados.length; i++) {

    var r = dados[i];
    var inep = inep_(r[T.inep - 1]);
    if (!inep) continue;

    var escola = texto_(r[T.escola - 1]);
    var anexo  = espacos_(r[T.anexo - 1]);
    var meta   = gdi[inep] || {};

    if (!matrizes[inep]) {
      matrizes[inep] = novaUnidade_(inep, escola || meta.escola || "", "", meta);
      ordemM.push(inep);
    }

    if (!anexo) {
      acumularTurma_(matrizes[inep], r, "", T);
    } else {
      var chave = inep + "|||" + anexo.toUpperCase();
      if (!anexos[chave]) {
        anexos[chave] = novaUnidade_(inep, escola || meta.escola || "", anexo, meta);
        ordemA.push(chave);
      }
      acumularTurma_(anexos[chave], r, anexo, T);
      registrarEJAExterna_(matrizes[inep], r, anexo, T);
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
  var C = colunasDe_(ss, ABA_GDI);
  for (var i = 1; i < d.length; i++) {
    var inep = inep_(d[i][C.inep - 1]);
    if (!inep) continue;
    mapa[inep] = {
      gre:           texto_(d[i][C.gre - 1]),
      municipio:     texto_(d[i][C.municipio - 1]),
      escola:        texto_(d[i][C.escola - 1]),
      salas:         d[i][C.salas - 1],
      escolaProxima: texto_(d[i][C.escolaProxima - 1])
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
    ejaMatrizLinhas: [],
    cursos: {}, cursos1a: {}, ejaPorLocal: {}
  };
  var zeros = ("totalTurmas totalMatriculas emTotal efTotal totalEJA totalEJAentm " +
    "nI nM nT nN efI efM efT efN s1 s2 s3 " +
    "s1I s1M s1T s1N s2I s2M s2T s2N s3I s3M s3T s3N " +
    "outrosEM outrosI outrosM outrosT outrosN ejI ejM ejT ejN " +
    "ef9Turmas ef9Matriculas ef8Turmas ef8Matriculas " +
    "aeeTurmas aeeMatriculas ejaAnexoTurmas ejaAnexoMatriculas " +
    "turmasZeradas ejaMatrizTurmas ejaMatrizEntm " +
    "ejMatrizI ejMatrizM ejMatrizT ejMatrizN").split(" ");
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

function acumularTurma_(e, r, anexo, T) {

  var curso = texto_(r[T.curso - 1]), etapa = texto_(r[T.etapa - 1]);
  var organizacao = texto_(r[T.organizacao - 1]), periodo = texto_(r[T.periodo - 1]);
  var turno = texto_(r[T.turno - 1]);
  var entm = numero_(r[T.enturmados - 1]), turmas = numero_(r[T.turmas - 1]);

  if (turmas === 0 && entm === 0 && !curso && !etapa) return;

  // Turma declarada sem nenhuma matrícula não é turma: não aparece na oferta
  // nem entra em contagem alguma. Fica só registrada para conferência.
  if (entm === 0 && turmas > 0) {
    e.turmasZeradas += turmas;
    return;
  }

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
    if (classificarAnexo_(anexo) === "MATRIZ") {
      e.ejaMatrizLinhas.push(txt);
      e.ejaMatrizTurmas += turmas;
      e.ejaMatrizEntm += entm;
      somarTurno_(e, "ejMatriz", tU, turmas);
    }
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
function registrarEJAExterna_(matriz, r, anexo, T) {
  var eU = texto_(r[T.etapa - 1]).toUpperCase();
  var cU = texto_(r[T.curso - 1]).toUpperCase();
  if (eU.indexOf("EJA") === -1 && cU.indexOf("EJA") === -1) return;
  somarEJALocal_(matriz, anexo, texto_(r[T.curso - 1]),
                 numero_(r[T.turmas - 1]), numero_(r[T.enturmados - 1]));
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

  // CEJA: sem oferta diurna para dividir sala, a própria EJA é a sala.
  if (soEJA_(e)) {
    i27 += e.ejMatrizI; m27 += e.ejMatrizM;
    t27 += e.ejMatrizT; n27 += e.ejMatrizN;
  }

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

/** Só a EJA do prédio matriz. O que está em anexo aparece na coluna da UEJA. */
function montarOfertaEJA_(e) {
  if (!e.ejaMatrizLinhas.length) return "—";
  var res = [];
  if (e.ejMatrizI > 0) res.push(e.ejMatrizI + " Integral");
  if (e.ejMatrizM > 0) res.push(e.ejMatrizM + " Manhã");
  if (e.ejMatrizT > 0) res.push(e.ejMatrizT + " Tarde");
  if (e.ejMatrizN > 0) res.push(e.ejMatrizN + " Noite");
  return e.ejaMatrizLinhas.slice().sort().join("\n") +
         "\nTurmas EJA no prédio matriz: " + (res.length ? res.join(" | ") : "0");
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
  var M = colunasDe_(ss, ABA_MPE);

  for (var i = 1; i < d.length; i++) {

    var r = d[i];
    var inep = inep_(r[M.inep - 1]);
    if (!inep) continue;

    var enturm = numero_(r[M.enturmados - 1]);
    var pre    = numero_(r[M.pre - 1]);
    if (!(enturm === 0 && pre > 0)) continue;      // só oferta nova

    var meta  = gdi[inep] || {};
    var turno = texto_(r[M.turno - 1]), etapa = texto_(r[M.etapa - 1]);
    var turma = texto_(r[M.turma - 1]), curso = texto_(r[M.curso - 1]);

    linhas.push([Number(inep), meta.gre || "", semAcento_(meta.municipio || ""),
                 meta.escola || "", etapa, curso, turma, turno,
                 pre, enturm, numero_(r[M.cursando - 1]),
                 Math.max(1, Math.ceil(pre / MAX_T))]);

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

  function identifica(e, r, C) {
    if (!e.escola) {
      e.gre = texto_(r[C.gre - 1]);
      e.municipio = texto_(r[C.municipio - 1]);
      e.escola = texto_(r[C.escola - 1]);
    }
  }

  // ---- 1ª série
  var sh = ss.getSheetByName(ABA_OF1);
  if (sh) {
    var d = sh.getDataRange().getValues();
    var C1 = colunasDe_(ss, ABA_OF1);
    for (var i = 1; i < d.length; i++) {
      var k = inep_(d[i][C1.inep - 1]); if (!k) continue;
      var e = unidade(k); identifica(e, d[i], C1);
      var t = numero_(d[i][C1.turmas - 1]), curso = texto_(d[i][C1.curso - 1]);
      e.of1Turmas += t; e.of1Alunos += numero_(d[i][C1.alunos - 1]);
      e.of1Cursos[curso] = (e.of1Cursos[curso] || 0) + t;
    }
  }

  // ---- 2ª e 3ª série (continuidade): cada linha é uma turma
  sh = ss.getSheetByName(ABA_OFC);
  if (sh) {
    var d2 = sh.getDataRange().getValues();
    var C2 = colunasDe_(ss, ABA_OFC);
    for (var j = 1; j < d2.length; j++) {
      var k2 = inep_(d2[j][C2.inep - 1]); if (!k2) continue;
      var e2 = unidade(k2); identifica(e2, d2[j], C2);
      var curso2 = texto_(d2[j][C2.curso - 1]);
      var etapa = texto_(d2[j][C2.etapa - 1]).toUpperCase();
      var pre = etapa.indexOf("2") === 0 ? "co2" : "co3";
      e2[pre + "Turmas"] += 1;
      e2[pre + "Alunos"] += numero_(d2[j][C2.alunos - 1]);
      e2[pre + "Cursos"][curso2] = (e2[pre + "Cursos"][curso2] || 0) + 1;
    }
  }

  // ---- subsequente
  sh = ss.getSheetByName(ABA_OFS);
  if (sh) {
    var d3 = sh.getDataRange().getValues();
    var C3 = colunasDe_(ss, ABA_OFS);
    for (var m = 1; m < d3.length; m++) {
      var k3 = inep_(d3[m][C3.inep - 1]); if (!k3) continue;
      var e3 = unidade(k3); identifica(e3, d3[m], C3);
      var t3 = numero_(d3[m][C3.turmas - 1]), curso3 = texto_(d3[m][C3.curso - 1]);
      e3.subTurmas += t3; e3.subAlunos += numero_(d3[m][C3.alunos - 1]);
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
function soEJA_(e) {
  return e.totalEJA > 0 && e.emTotal === 0 && e.efTotal === 0;
}

function salasOferta2027_(e, of) {
  var i27 = (of ? of.of1Turmas + of.co2Turmas + of.co3Turmas + of.subTurmas : 0)
            + e.efI + e.outrosI;
  var m27 = e.efM + e.outrosM;
  var t27 = e.efT + e.outrosT;
  var n27 = e.efN + e.outrosN;

  // CEJA: não há oferta diurna para dividir sala, então a EJA é a sala.
  if (soEJA_(e)) {
    i27 = e.ejMatrizI; m27 = e.ejMatrizM; t27 = e.ejMatrizT; n27 = e.ejMatrizN;
  }
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
    var P = colunasDe_(ss, ABA_PAN);
    for (var i = 1; i < d.length; i++) {
      var nome = texto_(d[i][P.municipio - 1]);
      if (!nome) continue;
      var chave = semAcento_(nome).toUpperCase();
      vistos[chave] = true;
      var todas = numero_(d[i][P.todasRedes - 1]);
      var ref   = numero_(d[i][P.refTurmas - 1]);
      var fonte = texto_(d[i][P.fonte - 1]) || "CENSO";
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
    "Oferta 2027 · tem oferta?",
    "Fundamental 2026 · turmas", "Só EJA (CEJA)?",
    "EJA 2026 · turmas no prédio matriz", "EJA 2026 · matrículas no prédio matriz",
    "Turmas zeradas descartadas"];

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
      e.ejaMatrizTurmas > 0 ? (e.ejaMatrizTurmas + " turmas · " + e.ejaMatrizEntm + " matrículas") : "—",
      e.efLinhas.length ? e.efLinhas.slice().sort().join("\n") : "—",
      e.parciais.length ? e.parciais.join("\n") : "—",
      montarFusoes_(e), pSem.salasNec, pSem.delta, resumoHoje_(e), resumo27,
      e.ejaMatrizTurmas > 0 ? (e.ejaMatrizEntm + " matrículas | " + e.ejaMatrizTurmas + " turmas") : "—",
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
      composicao, of ? "SIM" : "NÃO",
      e.efTotal, soEJA_(e) ? "SIM" : "NÃO",
      e.ejaMatrizTurmas, e.ejaMatrizEntm, e.turmasZeradas]);
  }

  gravar_(ss, ABA_BT, h, linhas,
    [100, 220, 400, 70, 70, 70, 380, 160, 280, 220, 360, 90, 80, 280, 300, 200,
     100, 110, 130, 110, 400, 300, 110, 120, 400, 320, 80, 130, 180, 110,
     110, 110, 110, 110, 120, 120, 400, 300, 100,
     120, 100, 120, 130, 120],
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
    listaDaAba_(ss, sh, 22, linhas.length, ABA_VAL, col_(ss, ABA_VAL, "destinoAnexo"));
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
  if (linhas.length) listaDaAba_(ss, sh, 11, linhas.length, ABA_VAL,
                                 col_(ss, ABA_VAL, "destinoAnexo"));
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

  // colunas de origem e destino resolvidas pelo nome, não pela letra
  var P  = colunasDe_(ss, ABA_PAN);
  var C2 = colunasDe_(ss, ABA_V2);
  var C3 = colunasDe_(ss, ABA_V3);

  var lMun = letra_(P.municipio), lNec = letra_(P.necessarias);
  var lDecV2 = letra_(P.decididaV2), lDecV3 = letra_(P.decididaV3);

  function somaDecidida(aba, C, n) {
    return "SUMIF('" + aba + "'!$" + letra_(C.municipio) + "$2:$" + letra_(C.municipio) +
           "$" + n + ",$" + lMun + "R,'" + aba + "'!$" + letra_(C.s1_2027) + "$2:$" +
           letra_(C.s1_2027) + "$" + n + ")";
  }

  for (var i = 0; i < pan.length; i++) {
    var r = i + 2;
    var mun = "$" + lMun + r;
    function aqui(f) { return f.replace(/R/g, r); }
    pan[i][P.decididaV2 - 1] = "=IF(" + mun + '="","",' + aqui(somaDecidida(ABA_V2, C2, nV2)) + ")";
    pan[i][P.saldoV2 - 1]    = "=IF(" + mun + '="","",$' + lNec + r + "-$" + lDecV2 + r + ")";
    pan[i][P.decididaV3 - 1] = "=IF(" + mun + '="","",' + aqui(somaDecidida(ABA_V3, C3, nV3)) + ")";
    pan[i][P.saldoV3 - 1]    = "=IF(" + mun + '="","",$' + lNec + r + "-$" + lDecV3 + r + ")";
    pan[i][P.cobertura - 1]  = "=IF(" + mun + '="","",IF($' + lNec + r + '=0,"",$' +
                               lDecV3 + r + "/$" + lNec + r + "))";
  }

  var sh = gravar_(ss, ABA_PAN, h, pan,
    [180, 140, 130, 140, 120, 140, 150, 150, 130, 130, 100, 130, 100, 100], []);

  if (pan.length) sh.getRange(2, P.cobertura, pan.length, 1).setNumberFormat("0%");
}


/** Cria a aba de IDEB com os municípios, sem apagar nota já digitada. */
function gravarIDEB_(ss, pan) {

  var sh = ss.getSheetByName(ABA_IDEB);
  var antigas = {};

  if (sh) {
    var d = sh.getDataRange().getValues();
    var I = colunasDe_(ss, ABA_IDEB);
    for (var i = 1; i < d.length; i++) {
      var m = semAcento_(d[i][I.municipio - 1]).toUpperCase();
      if (m) antigas[m] = [d[i][I.ideb - 1], d[i][I.ano - 1], d[i][I.obs - 1]];
    }
  }

  var linhas = [];
  for (var k = 0; k < pan.length; k++) {
    var nome = pan[k][0];
    var v = antigas[semAcento_(nome).toUpperCase()] || ["", "", ""];
    linhas.push([nome, v[0], v[1], v[2]]);
  }

  gravar_(ss, ABA_IDEB,
    ["Município", "IDEB Ensino Médio", "Ano de referência", "Observação"],
    linhas, [180, 160, 110, 300], [4]);
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
  limparCacheColunas_();          // as abas podem ter mudado desde a última vez

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

  var V = colunasDe_(ss, ABA_V2);
  var A = "$" + letra_(V.inep) + "@";
  var C = "$" + letra_(V.municipio) + "@";
  function r(apelido) { return "$" + letra_(V[apelido]) + "@"; }

  semear_(ss, sh, [["s1_2027", "of1Turmas"], ["s2_2027", "co2Turmas"],
                   ["s3_2027", "co3Turmas"], ["salasDecide", "salasReal"]]);
  municipios_(ss, sh, qtd);

  var cols = {};
  cols[V.gre]        = gd_(ss, A, "gre", nGD);
  cols[V.escola]     = gd_(ss, A, "escola", nGD);
  cols[V.turmasHoje] = gd_(ss, A, "proj2026", nGD);
  cols[V.turmas2027] = gd_(ss, A, "proj2027", nGD);
  cols[V.ofertaFund] = bt_(ss, A, "ofertaFund", nBT, '"—"');
  cols[V.ano9]       = 'IF(' + C + '="","",' + pm_(ss, C, "demanda", nPM, '"—"') +
                       '&" alunos · "&' + pm_(ss, C, "necessarias", nPM, '"—"') +
                       '&" turmas  (estadual nesta escola: "&' +
                       bt_(ss, A, "ano9Matr", nBT, "0") + '&")")';
  cols[V.faltam]     = falta_(ss, C, "saldoV2", nPM);
  cols[V.ideb]       = ideb_(ss, C, "ideb", nID);
  cols[V.cursos2026] = bt_(ss, A, "ofertaEM", nBT, '"—"');
  cols[V.parciais]   = bt_(ss, A, "parciais", nBT, '"—"');
  cols[V.fusoes]     = bt_(ss, A, "fusoes", nBT, '"—"');

  var fJ = "'" + ABA_FT + "'!$J$2:$J$" + nFT;
  var fN = "'" + ABA_FT + "'!$N$2:$N$" + nFT;
  cols[V.validFusao] = "IFERROR(IF(COUNTIF(" + fJ + "," + A + ")+COUNTIF(" + fN + "," + A +
                       ')=0,"Sem fusão turma a turma",COUNTIF(' + fJ + "," + A +
                       ')&" turma(s) saem · "&COUNTIF(' + fN + "," + A +
                       ')&" turma(s) recebem"),"—")';

  cols[V.salasExist] = gd_(ss, A, "salas", nGD);
  cols[V.situacao]   = situacaoSala_(r("salasExist"), r("salasDecide"));
  cols[V.ofertaEJA]  = bt_(ss, A, "ofertaEJA", nBT, '"—"');
  cols[V.matrEJA]    = bt_(ss, A, "matrTurmasEJA", nBT, '"—"');
  cols[V.anexos]     = bt_(ss, A, "anexos", nBT, '"—"');
  cols[V.cursos]     = bt_(ss, A, "of1Cursos", nBT, '"—"');
  cols[V.resumo]     = 'IF(' + A + '="","",' + r("s1_2027") + '&" de 1ª | "&' +
                       r("s2_2027") + '&" de 2ª | "&' + r("s3_2027") +
                       '&" de 3ª | Necessidade: "&' + r("salasDecide") +
                       '&" salas | "&' + r("situacao") + ')';
  cols[V.escolaProxima] = gd_(ss, A, "escolaProxima", nGD, '"—"');
  cols[V.of1] = bt_(ss, A, "of1Turmas", nBT, "0");
  cols[V.of2] = bt_(ss, A, "co2Turmas", nBT, "0");
  cols[V.of3] = bt_(ss, A, "co3Turmas", nBT, "0");
  cols[V.divergencia] =
    'IF(' + A + '="","",IF(AND(' + r("s1_2027") + '=' + bt_(ss, A, "of1Turmas", nBT, "0") +
    ',' + r("s2_2027") + '=' + bt_(ss, A, "co2Turmas", nBT, "0") +
    ',' + r("s3_2027") + '=' + bt_(ss, A, "co3Turmas", nBT, "0") +
    '),"igual à oferta","decidido "&' + r("s1_2027") + '&"/"&' + r("s2_2027") +
    '&"/"&' + r("s3_2027") + '&" · oferta "&' + bt_(ss, A, "of1Turmas", nBT, "0") +
    '&"/"&' + bt_(ss, A, "co2Turmas", nBT, "0") + '&"/"&' +
    bt_(ss, A, "co3Turmas", nBT, "0") + '))';

  aplicar_(sh, cols, qtd);

  lista_(ss, sh, "estrelaFund", ABA_V2, "fundamental", qtd);
  lista_(ss, sh, "validFusao",  ABA_V2, "validFusao",  qtd);
  lista_(ss, sh, "reord",       ABA_V2, "reordenamento", qtd);
}


/**
 * V3 — as colunas ★ NUNCA recebem fórmula: quem decide digita ou escolhe no
 * menu. As contas ficam em colunas próprias, ao lado, que leem a ★.
 */
function formulasV3_(ss, nBT, nPM, nGD, nID, nFE) {

  var sh = ss.getSheetByName(ABA_V3);
  if (!sh) return;

  var n = ultimaLinha_(ss, ABA_V3);
  if (n < 2) return;
  var qtd = n - 1;

  var V = colunasDe_(ss, ABA_V3);
  var A = "$" + letra_(V.inep) + "@";
  var C = "$" + letra_(V.municipio) + "@";
  function r(apelido) { return "$" + letra_(V[apelido]) + "@"; }

  semear_(ss, sh, [["s1_2027", "of1Turmas"], ["s2_2027", "co2Turmas"],
                   ["s3_2027", "co3Turmas"], ["salasDecide", "salasReal"],
                   ["ejaTurmas", "ejaMatrizTurmas"]]);
  municipios_(ss, sh, qtd);

  // soma o que estiver entre parênteses em "Curso (2), Outro (1)"
  function somaParenteses(ref) {
    return 'IFERROR(SUM(ARRAYFORMULA(IFERROR(VALUE(REGEXEXTRACT(' +
           'SPLIT(' + ref + ',","),"\\((\\d+)\\)")),0))),0)';
  }
  // curso a curso quando houver; senão, o total digitado
  var turmas1a = "(IF(" + r("cursosTotal") + ">0," + r("cursosTotal") + "," + r("s1_2027") + "))";

  var cols = {};
  cols[V.gre]        = gd_(ss, A, "gre", nGD);
  cols[V.escola]     = gd_(ss, A, "escola", nGD);
  cols[V.turmasHoje] = gd_(ss, A, "proj2026", nGD);
  cols[V.turmas2027] = gd_(ss, A, "proj2027", nGD);
  cols[V.ofertaFund] = bt_(ss, A, "ofertaFund", nBT, '"—"');
  cols[V.fundTotal]  = somaParenteses(r("fundTurmas"));
  cols[V.ano9Escola] = 'IF(' + A + '="","",' + bt_(ss, A, "ano9Matr", nBT, "0") +
                       '&" matrícula(s) · "&' + bt_(ss, A, "ano9Turmas", nBT, "0") +
                       '&" turma(s)")';
  cols[V.ano9Municipio] = 'IF(' + C + '="","",' + pm_(ss, C, "todasRedes", nPM, "0") +
                       '&" alunos (estadual "&' + pm_(ss, C, "est9Alunos", nPM, "0") +
                       '&" + outras redes "&' + pm_(ss, C, "outrasRedes", nPM, "0") + '&")")';
  cols[V.demanda]    = 'IF(' + C + '="","",' + pm_(ss, C, "demanda", nPM, "0") +
                       '&" alunos → "&' + pm_(ss, C, "necessarias", nPM, "0") +
                       '&" turma(s) necessária(s)")';
  cols[V.faltam]     = falta_(ss, C, "saldoV3", nPM);
  cols[V.ideb]       = "IFERROR(VLOOKUP(" + A + ",'IDEB - ESCOLAS'!C:F,4,0),\"\")";
  cols[V.cursos2026] = bt_(ss, A, "ofertaEM", nBT, '"—"');
  cols[V.cursosTotal] = somaParenteses(r("cursosTurmas"));
  cols[V.of1]        = bt_(ss, A, "of1Turmas", nBT, "0");
  cols[V.ofCursos]   = bt_(ss, A, "of1Cursos", nBT, '"—"');
  cols[V.of23sub]    = 'IF(' + A + '="","",' + bt_(ss, A, "co2Turmas", nBT, "0") +
                       '&" de 2ª | "&' + bt_(ss, A, "co3Turmas", nBT, "0") +
                       '&" de 3ª | "&' + bt_(ss, A, "subTurmas", nBT, "0") + '&" subseq.")';
  cols[V.parciais]     = bt_(ss, A, "parciais", nBT, '"—"');
  cols[V.fusaoPropria] = bt_(ss, A, "fusoes", nBT, '"—"');
  cols[V.fusaoEntre]   = bt_(ss, A, "fusaoEntre", nBT, '"—"');
  cols[V.salasExist]   = gd_(ss, A, "salas", nGD);

  // a base do motor, movida por tudo que for decidido nas ★
  cols[V.salasCalc] = 'IF(' + A + '="","",MAX(0,' + bt_(ss, A, "salasReal", nBT, "0") +
    "+(" + turmas1a + "-" + bt_(ss, A, "of1Turmas", nBT, "0") + ")" +
    "+(" + r("s2_2027") + "-" + bt_(ss, A, "co2Turmas", nBT, "0") + ")" +
    "+(" + r("s3_2027") + "-" + bt_(ss, A, "co3Turmas", nBT, "0") + ")" +
    "+(" + r("fundTotal") + "-" + bt_(ss, A, "fund2026", nBT, "0") + ")" +
    "+IF(" + bt_(ss, A, "soEJA", nBT, '"NÃO"') + '="SIM",' + r("ejaTurmas") + "-" +
    bt_(ss, A, "ejaMatrizTurmas", nBT, "0") + ",0)))";

  cols[V.salasSituacao] = situacaoSala_(r("salasExist"), r("salasDecide"));
  cols[V.ejaOferta]     = bt_(ss, A, "ofertaEJA", nBT, '"—"');
  cols[V.ejaMatriculas] = bt_(ss, A, "matrTurmasEJA", nBT, '"—"');
  cols[V.ejaFora]       = bt_(ss, A, "ejaFora", nBT, '"—"');
  cols[V.anexosLista]   = bt_(ss, A, "anexos", nBT, '"—"');
  cols[V.anexosQtd]     = bt_(ss, A, "qtdAnexos", nBT, "0");
  cols[V.resumo] = 'IF(' + A + '="","",' + turmas1a + '&" de 1ª | "&' + r("s2_2027") +
    '&" de 2ª | "&' + r("s3_2027") + '&" de 3ª | Fund "&' + r("fundTotal") +
    '&" | EJA "&' + r("ejaTurmas") + '&" | Necessidade: "&' + r("salasDecide") +
    '&" salas | "&' + r("salasSituacao") + ')';
  cols[V.escolaProxima] = gd_(ss, A, "escolaProxima", nGD, '"—"');
  cols[V.ofAlunos23] = 'IF(' + A + '="","",' + bt_(ss, A, "co2Alunos", nBT, "0") + "+" +
                       bt_(ss, A, "co3Alunos", nBT, "0") + '&" aluno(s)")';
  cols[V.divergencia] =
    'IF(' + A + '="","",IF(AND(' + turmas1a + '=' + bt_(ss, A, "of1Turmas", nBT, "0") +
    ',' + r("s2_2027") + '=' + bt_(ss, A, "co2Turmas", nBT, "0") +
    ',' + r("s3_2027") + '=' + bt_(ss, A, "co3Turmas", nBT, "0") +
    '),"igual à oferta","decidido "&' + turmas1a + '&"/"&' + r("s2_2027") +
    '&"/"&' + r("s3_2027") + '&" · oferta "&' + bt_(ss, A, "of1Turmas", nBT, "0") +
    '&"/"&' + bt_(ss, A, "co2Turmas", nBT, "0") + '&"/"&' +
    bt_(ss, A, "co3Turmas", nBT, "0") + '))';

  aplicar_(sh, cols, qtd);

  lista_(ss, sh, "fundEtapas",    ABA_V3, "fundamental",   qtd);
  lista_(ss, sh, "fundTurmas",    ABA_V3, "fundQtd",       qtd);
  lista_(ss, sh, "cursos2027",    ABA_V3, "cursosQtd",     qtd);
  lista_(ss, sh, "cursosTurmas",  ABA_V3, "cursosQtd",     qtd);
  lista_(ss, sh, "fusaoValid",    ABA_V3, "validFusao",    qtd);
  lista_(ss, sh, "anexosDecisao", ABA_V3, "decisaoAnexo",  qtd);
  lista_(ss, sh, "reord",         ABA_V3, "reordenamento", qtd);
}


/*
 * Os montadores abaixo recebem o APELIDO da coluna, não o índice. O índice do
 * PROCV sai do nome escrito na linha 1 da aba de origem — todos os intervalos
 * começam em A, então o índice é o próprio número da coluna. A largura do
 * intervalo também acompanha a aba, em vez de ficar presa a uma letra.
 */

function faixa_(ss, aba, n) {
  var sh = ss.getSheetByName(aba);
  var larg = sh ? Math.max(1, sh.getLastColumn()) : 1;
  return "'" + aba + "'!$A$2:$" + letra_(larg) + "$" + n;
}

function proc_(ss, aba, chave, apelido, n, err) {
  return "IFERROR(VLOOKUP(" + chave + "," + faixa_(ss, aba, n) + "," +
         col_(ss, aba, apelido) + ",FALSE)," + (err || '""') + ")";
}

function bt_(ss, chave, apelido, n, err) { return proc_(ss, ABA_BT,  chave, apelido, n, err); }
function pm_(ss, chave, apelido, n, err) { return proc_(ss, ABA_PAN, chave, apelido, n, err); }
function gd_(ss, chave, apelido, n, err) { return proc_(ss, ABA_GDI, chave, apelido, n, err); }

function ideb_(ss, chave, apelido, n) {
  var v = "VLOOKUP(" + chave + "," + faixa_(ss, ABA_IDEB, n) + "," +
          col_(ss, ABA_IDEB, apelido) + ",FALSE)";
  return 'IFERROR(IF(N(' + v + ')=0,"— sem IDEB cadastrado",' + v +
         '),"— sem IDEB cadastrado")';
}
function falta_(ss, chave, apelidoSaldo, n) {
  var v = pm_(ss, chave, apelidoSaldo, n, '""');
  return 'IF(' + chave + '="","",IF(' + v + '="","—",IF(' + v + '>0,' +
         '"FALTAM DISTRIBUIR "&' + v + '&" TURMA(S)",IF(' + v + '=0,' +
         '"DEMANDA MUNICIPAL ATENDIDA","EXCEDE A DEMANDA EM "&ABS(' + v + ')&" TURMA(S)"))))';
}
function situacaoSala_(refExiste, refNec) {
  var e = refExiste, x = refNec;
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
function municipios_(ss, sh, qtd) {
  var V = colunasDe_(ss, sh.getName());
  var ineps = sh.getRange(2, V.inep, qtd, 1).getValues();
  var gdi = mapaGDI_(ss);
  var out = [];
  for (var i = 0; i < qtd; i++) {
    var k = inep_(ineps[i][0]);
    out.push([k && gdi[k] ? semAcento_(gdi[k].municipio) : ""]);
  }
  sh.getRange(2, V.municipio, qtd, 1).setValues(out);
}


/** Liga um menu suspenso: coluna de destino e lista de origem, ambas por nome. */
function lista_(ss, sh, apelidoDest, abaDest, apelidoLista, qtd) {
  if (qtd < 1) return;
  var origem = ss.getSheetByName(ABA_VAL);
  if (!origem) return;
  var n = origem.getLastRow();
  if (n < 2) return;
  var colLista = col_(ss, ABA_VAL, apelidoLista);
  var colDest = col_(ss, abaDest, apelidoDest);
  var regra = SpreadsheetApp.newDataValidation()
    .requireValueInRange(origem.getRange(2, colLista, n - 1, 1), true)
    .setAllowInvalid(true).build();
  sh.getRange(2, colDest, qtd, 1).setDataValidation(regra);
}


/**
 * Semente x decisão.
 *
 * O script precisa saber se um número numa coluna ★ foi ele mesmo que plantou
 * ou se alguém digitou. Isso fica gravado na aba "★ Sementes (controle)", que
 * ele reescreve a cada rodada. A regra:
 *
 *   célula vazia                        → recebe a semente
 *   célula igual à última semente       → ainda é semente, pode ser refeita
 *   qualquer outro valor                → é decisão de gente, não se toca
 */
var ABA_SEMENTES = "★ Sementes (controle)";

function semear_(ss, sh, pares) {

  var bt = ss.getSheetByName(ABA_BT);
  if (!bt) return;

  var d = bt.getDataRange().getValues();
  var B = colunasDe_(ss, ABA_BT);
  var mapa = {};
  for (var i = 1; i < d.length; i++) {
    var k = inep_(d[i][B.inep - 1]);
    if (k) mapa[k] = d[i];
  }

  var aba = sh.getName();
  var V = colunasDe_(ss, aba);

  var n = sh.getLastRow();
  if (n < 2) return;
  var qtd = n - 1;
  var ineps = sh.getRange(2, V.inep, qtd, 1).getValues();

  var anterior = lerSementes_(ss, aba);
  var agora = {};

  for (var p = 0; p < pares.length; p++) {

    var apelidoDest = pares[p][0], apelidoOrig = pares[p][1];
    var colDest = V[apelidoDest];
    var colOrig = B[apelidoOrig];
    if (!colDest || !colOrig) continue;

    var atual = sh.getRange(2, colDest, qtd, 1).getValues();
    var mudou = false;

    for (var r = 0; r < qtd; r++) {

      var k2 = inep_(ineps[r][0]);
      if (!k2 || !mapa[k2]) continue;

      var semente = mapa[k2][colOrig - 1];
      if (!agora[k2]) agora[k2] = {};
      agora[k2][apelidoDest] = semente;

      var v = atual[r][0];
      var vazio = (v === "" || v === null || v === undefined);
      var ant = (anterior[k2] || {})[apelidoDest];
      var aindaSemente = !vazio && ant !== undefined && ant !== null &&
                         String(ant) !== "" && Number(v) === Number(ant);

      if (vazio || aindaSemente) {
        if (String(atual[r][0]) !== String(semente)) {
          atual[r][0] = semente;
          mudou = true;
        }
      }
    }
    if (mudou) sh.getRange(2, colDest, qtd, 1).setValues(atual);
  }

  gravarSementes_(ss, aba, ineps, pares, agora);
}


function lerSementes_(ss, aba) {
  var sh = ss.getSheetByName(ABA_SEMENTES);
  var out = {};
  if (!sh) return out;
  var n = sh.getLastRow();
  if (n < 2) return out;
  var d = sh.getRange(1, 1, n, sh.getLastColumn()).getValues();
  var cols = {};
  for (var c = 1; c < d[0].length; c++) {
    var t = String(d[0][c] || "");
    if (t.indexOf(aba + "|") === 0) cols[t.substring(aba.length + 1)] = c;
  }
  for (var i = 1; i < d.length; i++) {
    var k = inep_(d[i][0]);
    if (!k) continue;
    out[k] = {};
    for (var chave in cols) out[k][chave] = d[i][cols[chave]];
  }
  return out;
}


function gravarSementes_(ss, aba, ineps, pares, agora) {

  var sh = ss.getSheetByName(ABA_SEMENTES);
  if (!sh) {
    sh = ss.insertSheet(ABA_SEMENTES);
    sh.getRange(1, 1).setValue("INEP");
    sh.hideSheet();
  }

  var n = Math.max(2, sh.getLastRow());
  var largura = Math.max(1, sh.getLastColumn());
  var d = sh.getRange(1, 1, n, largura).getValues();

  var linhaDe = {};
  for (var i = 1; i < d.length; i++) {
    var k = inep_(d[i][0]);
    if (k) linhaDe[k] = i;
  }

  var colDe = {};
  for (var c = 1; c < d[0].length; c++) colDe[String(d[0][c])] = c;

  for (var p = 0; p < pares.length; p++) {
    var titulo = aba + "|" + pares[p][0];
    if (colDe[titulo] === undefined) {
      colDe[titulo] = d[0].length;
      d[0].push(titulo);
      for (var r = 1; r < d.length; r++) d[r].push("");
    }
  }
  var largura2 = d[0].length;
  for (var r2 = 0; r2 < d.length; r2++)
    while (d[r2].length < largura2) d[r2].push("");

  for (var q = 0; q < ineps.length; q++) {
    var k2 = inep_(ineps[q][0]);
    if (!k2) continue;
    if (linhaDe[k2] === undefined) {
      var nova = [];
      for (var z = 0; z < largura2; z++) nova.push("");
      nova[0] = Number(k2);
      d.push(nova);
      linhaDe[k2] = d.length - 1;
    }
    var vals = agora[k2] || {};
    for (var p2 = 0; p2 < pares.length; p2++) {
      var t2 = aba + "|" + pares[p2][0];
      var v2 = vals[pares[p2][0]];
      d[linhaDe[k2]][colDe[t2]] = (v2 === undefined ? "" : v2);
    }
  }

  sh.clear();
  sh.getRange(1, 1, d.length, largura2).setValues(d);
  sh.setFrozenRows(1);
}


// ═════════════════════════════════════════════════════════════════════
//  10. VALIDAÇÕES  —  as listas dos menus suspensos
// ═════════════════════════════════════════════════════════════════════

var MAX_QTD = 8;          // até quantas turmas o menu oferece por curso / etapa

var FUNDAMENTAL_ = [
  "EF Inicial - 1º ano", "EF Inicial - 2º ano", "EF Inicial - 3º ano",
  "EF Inicial - 4º ano", "EF Inicial - 5º ano", "EF Final - 6º ano",
  "EF Final - 7º ano", "EF Final - 8º ano", "EF Final - 9º ano"];


/**
 * "Logística" vira "Logística (1)" … "Logística (8)".
 *
 * É isso que faz ninguém precisar digitar nome de curso: a pessoa escolhe o
 * item pronto e a quantidade já vem junto. As colunas de total leem o número
 * de dentro dos parênteses.
 */
function comQuantidade_(itens) {
  var out = [];
  for (var i = 0; i < itens.length; i++)
    for (var q = 1; q <= MAX_QTD; q++)
      out.push(itens[i] + " (" + q + ")");
  return out;
}


/** Os cursos marcados como ofertados em 2027 na aba Cursos. */
function cursosOfertados_(ss) {
  var sh = ss.getSheetByName("Cursos");
  if (!sh) return [];
  var d = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < d.length; i++) {
    var nome = texto_(d[i][2]);
    var sit = texto_(d[i][6]).toUpperCase();
    if (nome && sit.indexOf("OFERTADO") === 0) out.push(nome);
  }
  return out;
}


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
      "Encerrar a oferta", "Em análise"] },
    { t: "CURSOS 1ª SÉRIE\ncom quantidade", cor: "#059669", bg: "#ecfdf5",
      dados: comQuantidade_(cursosOfertados_(ss)) },
    { t: "FUNDAMENTAL\ncom quantidade", cor: "#0891b2", bg: "#ecfeff",
      dados: comQuantidade_(FUNDAMENTAL_) },
    { t: "DECISÃO SOBRE\nO ANEXO", cor: "#9333ea", bg: "#faf5ff", dados: [
      "Manter o anexo como está",
      "Manter o anexo, com ajuste de oferta",
      "Encerrar o anexo — oferta vai para o prédio matriz",
      "Encerrar o anexo — oferta vai para outra escola estadual",
      "Encerrar o anexo — oferta vai para a rede municipal",
      "Encerrar o anexo — oferta encerrada",
      "Remanejar a oferta, mantendo o anexo aberto",
      "Em análise"] }
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
    sh.setColumnWidth(c4 + 1, 260);
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

// Quais colunas acumulam, por aba — por apelido, resolvido na hora do clique.
var COL_CURSOS = { };
COL_CURSOS[ABA_V2] = ["cursos2027"];
COL_CURSOS[ABA_V3] = ["fundTurmas", "cursos2027", "cursosTurmas"];
var SEPARADOR = ", ";


/** Números das colunas que acumulam nesta aba. */
function colunasAcumulativas_(ss, aba) {
  var apelidos = COL_CURSOS[aba];
  if (!apelidos) return null;
  var V = colunasDe_(ss, aba);
  var out = [];
  for (var i = 0; i < apelidos.length; i++)
    if (V[apelidos[i]]) out.push(V[apelidos[i]]);
  return out;
}

function onEdit(e) {

  if (!e || !e.range) return;

  var cel = e.range;
  var aba = cel.getSheet().getName();

  if (!COL_CURSOS[aba]) return;
  var cols = colunasAcumulativas_(cel.getSheet().getParent(), aba);
  if (!cols || cols.indexOf(cel.getColumn()) === -1) return;
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
  var cols = COL_CURSOS[aba] ?
             colunasAcumulativas_(SpreadsheetApp.getActiveSpreadsheet(), aba) : null;

  if (!cols || cols.indexOf(cel.getColumn()) === -1) {
    SpreadsheetApp.getUi().alert(
      "Selecione uma célula de uma coluna ★ de cursos ou de turmas por curso.");
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
