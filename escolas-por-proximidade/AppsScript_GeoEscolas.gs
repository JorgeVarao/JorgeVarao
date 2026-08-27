/**
 * GEO-ESCOLAS — funções de proximidade direto na planilha
 * GDI/UGERF · SEDUC-PI
 *
 * Nenhuma API externa, nenhuma chave, nenhum custo. O Plus Code que já existe
 * na base é decodificado localmente em latitude/longitude.
 *
 * ---------------------------------------------------------------------------
 * COMO USAR — 3 passos
 * ---------------------------------------------------------------------------
 * 1. Extensões > Apps Script, cole este arquivo, salve.
 *
 * 2. Na base, crie duas colunas ao lado de "google PlusCode" (digamos que ela
 *    seja a coluna U). Na primeira linha de dados:
 *
 *       =PLUSCODE_LAT(U2)      e      =PLUSCODE_LON(U2)
 *
 *    Arraste para baixo. Pronto: toda a base geolocalizada.
 *    (Prefere valores fixos em vez de fórmula? Rode GEOCODIFICAR_BASE no menu
 *     "Geo-Escolas" que aparece na barra depois de recarregar a planilha.)
 *
 * 3. Numa aba nova, a resposta da pergunta "escolas perto do bairro X":
 *
 *       =ESCOLAS_PROXIMAS("MOCAMBINHO"; 3)
 *
 *    E a resposta da pergunta "escolas a X km da escola Y":
 *
 *       =ESCOLAS_PROXIMAS_DA_ESCOLA("CETI FIRMINA SOBREIRA"; 2)
 *
 *    Ajuste as constantes do bloco CONFIG abaixo para casar com a sua base.
 * ---------------------------------------------------------------------------
 */

// ============================ CONFIG =======================================
// Ajuste os nomes de coluna conforme o cabeçalho real da sua aba.
var CFG = {
  ABA:         'Dados/inep_teresina', // nome da aba com a base
  LINHA_CAB:   1,                      // linha do cabeçalho
  COL_NOME:    'NOME ENTIDADE',
  COL_PLUS:    'google PlusCode',
  COL_BAIRRO:  'BAIRROS',
  COL_ENDER:   'ENDEREÇOS',
  COL_TIPO:    'TIPOLOGIA',
  COL_SIT:     'SITUAÇÃO EDUCACIONAL',
  COL_GRE:     'GRE',
  COL_MAT:     'MATRÍCULAS 2026',
  SIT_ATIVA:   'EM ATIVIDADE',
  REF_LAT:     -5.0892,   // âncora para recuperar Plus Code curto (centro de Teresina)
  REF_LON:     -42.8019
};

/**
 * Colunas novas (contato e nome do gestor). Como o cabeçalho pode ser escrito
 * de mais de um jeito, procuramos por qualquer um dos apelidos abaixo — e, se
 * nenhum aparecer, caímos na letra da coluna informada pela GDI (AM e AN).
 * Ou seja: se um dia o cabeçalho mudar de texto, a leitura continua de pé.
 */
var CFG_EXTRA = {
  CONTATO: { apelidos: ['CONTATO', 'CONTATO DIRETOR', 'CONTATO DO DIRETOR',
                        'CONTATO DO GESTOR', 'CONTATO GESTOR', 'TELEFONE',
                        'TELEFONE DIRETOR'], letra: 'AM' },
  GESTOR:  { apelidos: ['NOME DO GESTOR', 'NOME GESTOR', 'GESTOR', 'DIRETOR',
                        'NOME DO DIRETOR', 'NOME DIRETOR'], letra: 'AN' }
};
// ===========================================================================


/* ---------- Open Location Code: decodificação local ---------- */
var _OA = '23456789CFGHJMPQRVWX';

function _olcDecode(code) {
  code = String(code).replace(/\+/g, '').replace(/0+$/, '').toUpperCase();
  var lat = -90, lon = -180, rLat = 400, rLon = 400, i = 0;
  while (i < code.length && i < 10) {
    rLat /= 20; rLon /= 20;
    lat += _OA.indexOf(code.charAt(i)) * rLat;
    lon += _OA.indexOf(code.charAt(i + 1)) * rLon;
    i += 2;
  }
  while (i < code.length) {
    rLat /= 5; rLon /= 4;
    var d = _OA.indexOf(code.charAt(i));
    lat += Math.floor(d / 4) * rLat;
    lon += (d % 4) * rLon;
    i++;
  }
  return [lat + rLat / 2, lon + rLon / 2];
}

function _olcEncode(lat, lon, len) {
  var a = lat + 90, o = lon + 180, res = 20, out = '';
  for (var i = 0; i < Math.ceil(len / 2); i++) {
    var la = Math.floor(a / res), lo = Math.floor(o / res);
    out += _OA.charAt(la) + _OA.charAt(lo);
    a -= la * res; o -= lo * res; res /= 20;
  }
  return out.substring(0, len);
}

function _plus(code) {
  if (code === null || code === undefined) return null;
  var m = String(code).trim().toUpperCase()
          .match(/^([2-9CFGHJMPQRVWX]{4,8}\+[2-9CFGHJMPQRVWX]{2,3})/);
  if (!m) return null;
  var c = m[1], sep = c.indexOf('+'), falta = 8 - sep;
  if (falta <= 0) return _olcDecode(c);
  var res = Math.pow(20, 2 - falta / 2), meio = res / 2;
  var p = _olcDecode(_olcEncode(CFG.REF_LAT, CFG.REF_LON, falta) + c);
  if (CFG.REF_LAT + meio < p[0]) p[0] -= res; else if (CFG.REF_LAT - meio > p[0]) p[0] += res;
  if (CFG.REF_LON + meio < p[1]) p[1] -= res; else if (CFG.REF_LON - meio > p[1]) p[1] += res;
  return p;
}


/* ---------- localizar colunas com tolerância ---------- */

/** "AM" -> 38 (índice 0). Serve de rede de segurança para as colunas novas. */
function _letraParaIndice_(letra) {
  var s = String(letra).toUpperCase().replace(/[^A-Z]/g, ''), n = 0;
  for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n - 1;
}

/** 38 -> "AM". Só para relatório legível. */
function _letraCol_(i) {
  var s = '', n = i + 1;
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/** Acha a coluna por qualquer um dos apelidos; se falhar, usa a letra. */
function _colFlex_(cabNorm, cfg) {
  for (var i = 0; i < cfg.apelidos.length; i++) {
    var k = _norm_(cfg.apelidos[i]), j = cabNorm.indexOf(k);
    if (j >= 0) return j;
  }
  var l = _letraParaIndice_(cfg.letra);
  return (l >= 0 && l < cabNorm.length) ? l : -1;
}

/** Telefone legível: 86999998888 -> (86) 99999-8888. */
function _fone_(v) {
  if (v === null || v === undefined) return '';
  var t = String(v).trim();
  if (!t) return '';
  var d = t.replace(/\D+/g, '');
  if (d.length === 11) return '(' + d.substr(0, 2) + ') ' + d.substr(2, 5) + '-' + d.substr(7);
  if (d.length === 10) return '(' + d.substr(0, 2) + ') ' + d.substr(2, 4) + '-' + d.substr(6);
  if (d.length === 9)  return d.substr(0, 5) + '-' + d.substr(5);
  if (d.length === 8)  return d.substr(0, 4) + '-' + d.substr(4);
  return t;
}


/* =========================== FUNÇÕES DE PLANILHA =========================== */

/**
 * Latitude a partir de um Plus Code. Aceita célula única ou intervalo.
 * @param {string|Array} codigo Ex.: "W55P+3W" ou "W55P+3W Centro, Teresina - PI"
 * @return {number} Latitude em graus decimais.
 * @customfunction
 */
function PLUSCODE_LAT(codigo) {
  if (codigo && codigo.map) return codigo.map(function (r) { return r.map(function (c) { var p = _plus(c); return p ? p[0] : ''; }); });
  var p = _plus(codigo); return p ? p[0] : '';
}

/**
 * Longitude a partir de um Plus Code. Aceita célula única ou intervalo.
 * @param {string|Array} codigo Plus Code.
 * @return {number} Longitude em graus decimais.
 * @customfunction
 */
function PLUSCODE_LON(codigo) {
  if (codigo && codigo.map) return codigo.map(function (r) { return r.map(function (c) { var p = _plus(c); return p ? p[1] : ''; }); });
  var p = _plus(codigo); return p ? p[1] : '';
}

/**
 * Distância em linha reta entre dois pontos, em quilômetros (Haversine).
 * @param {number} lat1 Latitude do ponto de origem.
 * @param {number} lon1 Longitude do ponto de origem.
 * @param {number} lat2 Latitude do ponto de destino.
 * @param {number} lon2 Longitude do ponto de destino.
 * @return {number} Distância em km.
 * @customfunction
 */
function DISTKM(lat1, lon1, lat2, lon2) {
  if (lat1 === '' || lat2 === '' || lon1 === '' || lon2 === '') return '';
  var R = 6371, r = Math.PI / 180;
  var dLa = (lat2 - lat1) * r, dLo = (lon2 - lon1) * r;
  var h = Math.pow(Math.sin(dLa / 2), 2) +
          Math.cos(lat1 * r) * Math.cos(lat2 * r) * Math.pow(Math.sin(dLo / 2), 2);
  return 2 * R * Math.asin(Math.sqrt(h));
}

/**
 * Lista as unidades da rede mais próximas de um ponto, ordenadas por distância.
 *
 * A origem pode ser: nome de um bairro (usa o centro geométrico das unidades
 * daquele bairro), um Plus Code, "lat, lon", ou o nome de uma escola da base.
 *
 * @param {string} origem Bairro, Plus Code, "lat, lon" ou nome de escola.
 * @param {number} raioKm Raio máximo em quilômetros. Padrão 3.
 * @param {boolean} somenteAtivas Se VERDADEIRO, ignora unidades sem atividade. Padrão VERDADEIRO.
 * @param {number} limite Máximo de linhas retornadas. Padrão 30.
 * @return {Array} Tabela: distância, escola, tipologia, bairro, endereço, GRE, matrículas, gestor, contato.
 * @customfunction
 */
function ESCOLAS_PROXIMAS(origem, raioKm, somenteAtivas, limite) {
  return _proximas_(origem, raioKm, somenteAtivas, limite, null);
}

/**
 * Escolas a X km de uma escola específica — a pergunta "quais unidades estão
 * a até 2 km do CETI Firmina Sobreira?" respondida direto na planilha.
 *
 * A própria escola de partida não entra na lista (ela seria sempre 0 km).
 *
 * @param {string} escola Nome (ou parte do nome) da escola de partida.
 * @param {number} raioKm Raio máximo em quilômetros. Padrão 3.
 * @param {boolean} somenteAtivas Se VERDADEIRO, ignora unidades sem atividade. Padrão VERDADEIRO.
 * @param {number} limite Máximo de linhas retornadas. Padrão 30.
 * @return {Array} Tabela: distância, escola, tipologia, bairro, endereço, GRE, matrículas, gestor, contato.
 * @customfunction
 */
function ESCOLAS_PROXIMAS_DA_ESCOLA(escola, raioKm, somenteAtivas, limite) {
  var base = _lerBase_();
  if (!base.linhas.length) return [['Base não encontrada. Confira CFG.ABA no script.']];
  var org = _acharEscola_(escola, base);
  if (!org) return [['Não encontrei uma escola com coordenada para "' + escola + '".']];
  return _proximas_(escola, raioKm, somenteAtivas, limite, org);
}

/**
 * Distância em linha reta entre duas escolas da base, em quilômetros.
 * @param {string} escolaA Nome (ou parte do nome) da primeira escola.
 * @param {string} escolaB Nome (ou parte do nome) da segunda escola.
 * @return {number} Distância em km.
 * @customfunction
 */
function DISTANCIA_ENTRE_ESCOLAS(escolaA, escolaB) {
  var base = _lerBase_();
  var a = _acharEscola_(escolaA, base), b = _acharEscola_(escolaB, base);
  if (!a) return 'Não encontrei "' + escolaA + '"';
  if (!b) return 'Não encontrei "' + escolaB + '"';
  return Math.round(DISTKM(a.lat, a.lon, b.lat, b.lon) * 100) / 100;
}

/**
 * Nome do gestor de uma escola da base (coluna AN).
 * @param {string} escola Nome (ou parte do nome) da escola.
 * @return {string} Nome do gestor.
 * @customfunction
 */
function GESTOR_DA_ESCOLA(escola) {
  var e = _acharEscola_(escola, _lerBase_(), true);
  return e ? e.gestor : '';
}

/**
 * Contato do diretor de uma escola da base (coluna AM), já formatado.
 * @param {string} escola Nome (ou parte do nome) da escola.
 * @return {string} Telefone de contato.
 * @customfunction
 */
function CONTATO_DA_ESCOLA(escola) {
  var e = _acharEscola_(escola, _lerBase_(), true);
  return e ? _fone_(e.contato) : '';
}


/* ---------- motor comum das listas de proximidade ---------- */
function _proximas_(origem, raioKm, somenteAtivas, limite, escolaOrigem) {
  raioKm = raioKm || 3;
  somenteAtivas = (somenteAtivas === undefined || somenteAtivas === '') ? true : somenteAtivas;
  limite = limite || 30;

  var base = _lerBase_();
  if (!base.linhas.length) return [['Base não encontrada. Confira CFG.ABA no script.']];

  var o = escolaOrigem
    ? { lat: escolaOrigem.lat, lon: escolaOrigem.lon, rotulo: 'escola ' + escolaOrigem.nome }
    : _resolverOrigem_(origem, base);
  if (!o) return [['Não reconheci "' + origem + '" como bairro, Plus Code, coordenada ou escola.']];

  var res = [];
  for (var i = 0; i < base.linhas.length; i++) {
    var e = base.linhas[i];
    if (e.lat === null) continue;
    if (escolaOrigem && e === escolaOrigem) continue;      // a origem não entra na lista
    if (somenteAtivas && e.sit !== CFG.SIT_ATIVA) continue;
    var d = DISTKM(o.lat, o.lon, e.lat, e.lon);
    if (d <= raioKm) res.push([d, e.nome, e.tipo, e.bairro, e.ender, e.gre, e.mat, e.gestor, _fone_(e.contato)]);
  }
  res.sort(function (a, b) { return a[0] - b[0]; });
  res = res.slice(0, limite);

  var cab = [['Distância (km)', 'Escola', 'Tipologia', 'Bairro', 'Endereço', 'GRE',
              'Matrículas', 'Nome do gestor', 'Contato']];
  if (!res.length) return cab.concat([['—', 'Nenhuma unidade em até ' + raioKm + ' km de ' + o.rotulo, '', '', '', '', '', '', '']]);
  for (var j = 0; j < res.length; j++) res[j][0] = Math.round(res[j][0] * 100) / 100;
  return cab.concat(res);
}


/* ---------- leitura da base ---------- */
function _lerBase_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(CFG.ABA);
  if (!sh) return { linhas: [] };
  var vals = sh.getDataRange().getValues();
  var cab = vals[CFG.LINHA_CAB - 1].map(function (c) { return String(c).trim(); });
  var cabNorm = cab.map(function (c) { return _norm_(c); });
  function col(nome) { return cab.indexOf(nome); }
  var iNome = col(CFG.COL_NOME), iPlus = col(CFG.COL_PLUS), iBai = col(CFG.COL_BAIRRO),
      iEnd = col(CFG.COL_ENDER), iTip = col(CFG.COL_TIPO), iSit = col(CFG.COL_SIT),
      iGre = col(CFG.COL_GRE), iMat = col(CFG.COL_MAT),
      iCon = _colFlex_(cabNorm, CFG_EXTRA.CONTATO), iGes = _colFlex_(cabNorm, CFG_EXTRA.GESTOR);

  var linhas = [];
  for (var r = CFG.LINHA_CAB; r < vals.length; r++) {
    var v = vals[r];
    if (!String(v[iNome]).trim()) continue;
    var p = iPlus >= 0 ? _plus(v[iPlus]) : null;
    linhas.push({
      nome: String(v[iNome]).trim(),
      tipo: iTip >= 0 ? String(v[iTip]).trim() : '',
      sit:  iSit >= 0 ? String(v[iSit]).trim() : '',
      bairro: iBai >= 0 ? String(v[iBai]).trim() : '',
      ender:  iEnd >= 0 ? String(v[iEnd]).trim() : '',
      gre:    iGre >= 0 ? String(v[iGre]).trim() : '',
      mat:    iMat >= 0 ? v[iMat] : '',
      gestor: iGes >= 0 ? String(v[iGes]).trim() : '',
      contato: iCon >= 0 ? String(v[iCon]).trim() : '',
      lat: p ? p[0] : null, lon: p ? p[1] : null
    });
  }
  return { linhas: linhas };
}

function _norm_(t) {
  return String(t).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
         .toUpperCase().replace(/[^A-Z0-9]+/g, ' ').trim();
}

/** Acha uma escola pelo nome. semCoord=true aceita unidade sem Plus Code. */
function _acharEscola_(nome, base, semCoord) {
  var k = _norm_(nome);
  if (k.length < 3) return null;
  var cand = [];
  for (var i = 0; i < base.linhas.length; i++) {
    var e = base.linhas[i];
    if (!semCoord && e.lat === null) continue;
    if (_norm_(e.nome) === k) return e;                    // acerto exato ganha
    if (_norm_(e.nome).indexOf(k) >= 0) cand.push(e);
  }
  if (cand.length) {                                       // o nome mais curto é o mais específico
    cand.sort(function (a, b) { return a.nome.length - b.nome.length; });
    return cand[0];
  }
  return null;
}

function _resolverOrigem_(origem, base) {
  var t = String(origem).trim();

  // 1) "lat, lon"
  var m = t.match(/^(-?\d+[.,]\d+)\s*[,;\s]\s*(-?\d+[.,]\d+)$/);
  if (m) return { lat: parseFloat(m[1].replace(',', '.')), lon: parseFloat(m[2].replace(',', '.')), rotulo: t };

  // 2) Plus Code
  var p = _plus(t);
  if (p) return { lat: p[0], lon: p[1], rotulo: 'Plus Code ' + t };

  var k = _norm_(t);

  // 3) bairro — centro geométrico das unidades daquele bairro
  var sLat = 0, sLon = 0, n = 0;
  for (var i = 0; i < base.linhas.length; i++) {
    var e = base.linhas[i];
    if (e.lat !== null && _norm_(e.bairro) === k) { sLat += e.lat; sLon += e.lon; n++; }
  }
  if (n) return { lat: sLat / n, lon: sLon / n, rotulo: 'bairro ' + t };

  // 4) nome de escola
  var esc = _acharEscola_(t, base);
  if (esc) return { lat: esc.lat, lon: esc.lon, rotulo: 'escola ' + esc.nome };
  return null;
}


/* ---------- menu: gravar LAT/LON como valores fixos ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Geo-Escolas')
    .addItem('Gravar LAT/LON na base', 'GEOCODIFICAR_BASE')
    .addItem('Diagnóstico da base', 'DIAGNOSTICO')
    .addItem('Atualizar página web agora', 'LIMPAR_CACHE')
    .addToUi();
}

function GEOCODIFICAR_BASE() {
  var ui = SpreadsheetApp.getUi();
  var sh = SpreadsheetApp.getActive().getSheetByName(CFG.ABA);
  if (!sh) { ui.alert('Aba "' + CFG.ABA + '" não encontrada. Ajuste CFG.ABA no script.'); return; }

  var vals = sh.getDataRange().getValues();
  var cab = vals[CFG.LINHA_CAB - 1].map(function (c) { return String(c).trim(); });
  var iPlus = cab.indexOf(CFG.COL_PLUS);
  if (iPlus < 0) { ui.alert('Coluna "' + CFG.COL_PLUS + '" não encontrada.'); return; }

  var cLat = cab.indexOf('LAT'), cLon = cab.indexOf('LON');
  if (cLat < 0) { cLat = cab.length;     sh.getRange(CFG.LINHA_CAB, cLat + 1).setValue('LAT'); }
  if (cLon < 0) { cLon = cab.length + 1; sh.getRange(CFG.LINHA_CAB, cLon + 1).setValue('LON'); }

  var out = [], ok = 0;
  for (var r = CFG.LINHA_CAB; r < vals.length; r++) {
    var p = _plus(vals[r][iPlus]);
    if (p) ok++;
    out.push([p ? p[0] : '', p ? p[1] : '']);
  }
  if (!out.length) { ui.alert('Nenhuma linha de dados encontrada.'); return; }

  sh.getRange(CFG.LINHA_CAB + 1, cLat + 1, out.length, 1).setValues(out.map(function (x) { return [x[0]]; }));
  sh.getRange(CFG.LINHA_CAB + 1, cLon + 1, out.length, 1).setValues(out.map(function (x) { return [x[1]]; }));
  ui.alert(ok + ' de ' + out.length + ' linhas geolocalizadas a partir do Plus Code.');
}

/**
 * PRIMEIRA COISA A RODAR NUM AMBIENTE NOVO.
 *
 * Diz se o script achou a aba, o cabeçalho, as colunas obrigatórias, as duas
 * colunas novas (contato e gestor) e quantas linhas têm coordenada. Não depende
 * de menu: a saída vai para o registro de execução do editor e, se der, também
 * aparece numa caixa de diálogo na planilha.
 */
function DIAGNOSTICO() {
  var L = [], ss = SpreadsheetApp.getActive();
  L.push('PLANILHA: ' + ss.getName());
  L.push('ABAS: ' + ss.getSheets().map(function (x) { return x.getName(); }).join(' | '));
  L.push('CFG.ABA = "' + CFG.ABA + '"');
  L.push('');

  var sh = ss.getSheetByName(CFG.ABA);
  if (!sh) {
    L.push('>> PARE AQUI: não existe aba com esse nome nesta planilha.');
    L.push('   Copie o nome certo da lista ABAS acima para CFG.ABA, lá no topo');
    L.push('   deste arquivo, salve e rode DIAGNOSTICO de novo.');
    return _saidaDiag_(L);
  }

  var vals = sh.getDataRange().getValues();
  if (vals.length <= CFG.LINHA_CAB) {
    L.push('>> A aba existe mas não tem linha de dados abaixo do cabeçalho.');
    return _saidaDiag_(L);
  }
  var cab = vals[CFG.LINHA_CAB - 1].map(function (c) { return String(c).trim(); });
  var cabNorm = cab.map(function (c) { return _norm_(c); });
  L.push('LINHAS DE DADOS: ' + (vals.length - CFG.LINHA_CAB));
  L.push('');

  function cheias(i) {
    if (i < 0) return 0;
    var c = 0;
    for (var r = CFG.LINHA_CAB; r < vals.length; r++) if (String(vals[r][i]).trim()) c++;
    return c;
  }
  function linha(rotulo, i, nome) {
    while (rotulo.length < 14) rotulo += ' ';
    return '  ' + rotulo + (i < 0
      ? 'NÃO ENCONTRADA  (procurei por "' + nome + '")'
      : 'coluna ' + _letraCol_(i) + '  ·  "' + cab[i] + '"  ·  ' + cheias(i) + ' preenchidas');
  }

  var obrig = [['Nome', CFG.COL_NOME], ['Plus Code', CFG.COL_PLUS], ['Bairro', CFG.COL_BAIRRO],
               ['Endereço', CFG.COL_ENDER], ['Tipologia', CFG.COL_TIPO], ['Situação', CFG.COL_SIT],
               ['GRE', CFG.COL_GRE], ['Matrículas', CFG.COL_MAT]];
  var faltam = [];
  L.push('COLUNAS DA BASE');
  for (var k = 0; k < obrig.length; k++) {
    var i = cab.indexOf(obrig[k][1]);
    if (i < 0) faltam.push(obrig[k][1]);
    L.push(linha(obrig[k][0], i, obrig[k][1]));
  }
  L.push('');
  L.push('COLUNAS NOVAS (procuradas pelo nome; se falhar, pela letra)');
  var iCon = _colFlex_(cabNorm, CFG_EXTRA.CONTATO), iGes = _colFlex_(cabNorm, CFG_EXTRA.GESTOR);
  L.push(linha('Contato', iCon, CFG_EXTRA.CONTATO.apelidos[0] + ' / col. ' + CFG_EXTRA.CONTATO.letra));
  L.push(linha('Gestor', iGes, CFG_EXTRA.GESTOR.apelidos[0] + ' / col. ' + CFG_EXTRA.GESTOR.letra));
  L.push('');

  var iPlus = cab.indexOf(CFG.COL_PLUS), geo = 0, comNome = 0;
  var iNome = cab.indexOf(CFG.COL_NOME);
  for (var r = CFG.LINHA_CAB; r < vals.length; r++) {
    if (iNome >= 0 && !String(vals[r][iNome]).trim()) continue;
    comNome++;
    if (iPlus >= 0 && _plus(vals[r][iPlus])) geo++;
  }
  L.push('COORDENADAS: ' + geo + ' de ' + comNome + ' unidades com Plus Code utilizável');
  L.push('  (só essas entram no mapa da página)');
  L.push('');

  if (faltam.length) {
    L.push('>> AJUSTE ANTES DE PUBLICAR: não achei ' + faltam.join(', ') + '.');
    L.push('   Corrija os nomes no bloco CONFIG, no topo deste arquivo.');
  } else if (!geo) {
    L.push('>> A base está legível, mas nenhuma linha tem Plus Code válido.');
    L.push('   Sem isso a página abre vazia. Confira a coluna "' + CFG.COL_PLUS + '".');
  } else {
    L.push('>> TUDO CERTO. Pode implantar o app da web.');
    if (iCon < 0 || iGes < 0) L.push('   (contato/gestor não foram achados: a página funciona, só não mostra esses campos)');
  }
  return _saidaDiag_(L);
}

/** Manda o relatório para o registro e, quando dá, para uma caixa na planilha. */
function _saidaDiag_(L) {
  var txt = L.join('\n');
  Logger.log(txt);
  try { SpreadsheetApp.getUi().alert(txt); } catch (e) {}   // sem UI disponível: fica só no registro
  return txt;
}
