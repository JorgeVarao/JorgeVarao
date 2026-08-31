/**
 * ESCOLAS POR PROXIMIDADE — Web App
 * GDI/UGERF · SEDUC-PI
 *
 * Publica a ferramenta como página web hospedada pelo Google, lendo os dados
 * AO VIVO da planilha. Nada de exportar e regerar arquivo: mudou a planilha,
 * mudou a página no próximo carregamento.
 *
 * ---------------------------------------------------------------------------
 * INSTALAÇÃO — 4 passos
 * ---------------------------------------------------------------------------
 * 1. Extensões > Apps Script. Cole ESTE arquivo junto com AppsScript_GeoEscolas.gs
 *    (ele reaproveita o decodificador de Plus Code e a busca de colunas que já
 *    estão lá — inclusive CFG_EXTRA, que localiza CONTATO e NOME DO GESTOR).
 *
 * 2. No editor, clique em "+" ao lado de Arquivos > HTML. Nomeie exatamente
 *    "Index" (sem .html) e cole o conteúdo do arquivo Index.html entregue junto.
 *    Ele já vem com o scriptlet <?!= dadosApp() ?> no lugar dos dados estáticos.
 *
 * 3. Implantar > Nova implantação > tipo "App da Web".
 *      Executar como.......: Eu
 *      Quem pode acessar...: "Qualquer pessoa da SEDUC-PI"  (recomendado)
 *                            ou "Qualquer pessoa" para link público
 *
 * 4. Copie a URL /exec e distribua. Quem abrir NÃO precisa de acesso à planilha
 *    — o script roda com a sua permissão e devolve só os campos abaixo.
 *
 * ATENÇÃO ao publicar contato de gestor: a página passa a expor telefone e nome
 * de servidor. Se a implantação for pública ("Qualquer pessoa"), prefira
 * restringir ao domínio da SEDUC — ou zere CFG_EXTRA.CONTATO no outro arquivo.
 * ---------------------------------------------------------------------------
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Escolas por proximidade — GDI/SEDUC-PI')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Cache de 10 min: evita reler a planilha inteira a cada visita. */
function dadosApp() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('dadosApp_v3');            // v3: passou a carregar a relação de bairros
  if (hit) return hit;
  var js = JSON.stringify(_montarDados_());
  try { cache.put('dadosApp_v3', js, 600); } catch (e) {}  // >100KB não cabe: segue sem cache
  return js;
}

/** Força atualização imediata depois de editar a planilha. */
function LIMPAR_CACHE() {
  var c = CacheService.getScriptCache();
  c.remove('dadosApp_v3');
  c.remove('dadosApp_v2');
  c.remove('dadosApp_v1');                       // restos das versões anteriores
  SpreadsheetApp.getUi().alert('Cache limpo. A página já reflete a planilha atual.');
}

// ---------------------------------------------------------------------------

var _CORR = {
  'TRESN ANDARES': 'TRES ANDARES', 'CEREMICA CIL': 'CERAMICA CIL',
  'DIRCEU I': 'DIRCEU', 'DIRCEU ARCOVERDE I': 'DIRCEU ARCOVERDE',
  'DIRCEU ARCOVERDE II': 'DIRCEU ARCOVERDE', 'RENASCENCA II': 'RENASCENCA',
  'POTY VELHO': 'POTI VELHO'
};
var _LIXO = { 'TERESINA': 1, 'PI, 64017-772': 1 };

/**
 * Aba com a relação oficial dos bairros de Teresina e suas coordenadas. É ela
 * que permite partir de um bairro onde não existe escola nenhuma.
 */
var CFG_BAIRROS = {
  ABA:        'bairros_teresina_coordenadas',
  COL_NOME:   'bairro',
  COL_CHAVE:  'bairro_chave',
  COL_ZONA:   'zona',
  COL_LAT:    'latitude',
  COL_LON:    'longitude',
  COL_LATLON: 'lat_lon'
};

/** Aceita só coordenada dentro da faixa possível do planeta. */
function _faixaOk_(lat, lon) {
  return !isNaN(lat) && !isNaN(lon) && Math.abs(lat) <= 90 && Math.abs(lon) <= 180;
}

/**
 * Coordenada de um bairro. Lemos primeiro a coluna "lat_lon", que é texto
 * "lat, lon" e chega inteira. As colunas "latitude"/"longitude" separadas
 * vieram com o ponto decimal perdido — -5.0824259 virou -50824259, porque o
 * Sheets leu o ponto como separador de milhar — então elas só são usadas se
 * estiverem dentro da faixa válida, o que já cobre o dia em que forem
 * corrigidas na planilha.
 */
function _coordBairro_(latLon, lat, lon) {
  var m = String(latLon).match(/(-?\d+[.,]\d+)\s*[,;]\s*(-?\d+[.,]\d+)/);
  if (m) {
    var a = parseFloat(m[1].replace(',', '.')), o = parseFloat(m[2].replace(',', '.'));
    if (_faixaOk_(a, o)) return [a, o];
  }
  var a2 = parseFloat(String(lat).replace(',', '.'));
  var o2 = parseFloat(String(lon).replace(',', '.'));
  return _faixaOk_(a2, o2) ? [a2, o2] : null;
}

/** Chave sem acento -> { bl: nome com acento, zona, lat, lon }. */
function _bairrosOficiais_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(CFG_BAIRROS.ABA);
  if (!sh) return {};                                  // aba ausente: segue sem ela
  var v = sh.getDataRange().getValues();
  if (v.length < 2) return {};
  var cab = v[0].map(function (c) { return _norm_(c); });
  function col(nome) { return cab.indexOf(_norm_(nome)); }
  var iNome  = col(CFG_BAIRROS.COL_NOME),  iChave = col(CFG_BAIRROS.COL_CHAVE),
      iZona  = col(CFG_BAIRROS.COL_ZONA),  iLat   = col(CFG_BAIRROS.COL_LAT),
      iLon   = col(CFG_BAIRROS.COL_LON),   iLL    = col(CFG_BAIRROS.COL_LATLON);
  if (iNome < 0) return {};

  var out = {};
  for (var r = 1; r < v.length; r++) {
    var nome = String(v[r][iNome]).replace(/\s+/g, ' ').trim();
    if (!nome) continue;
    var par = _coordBairro_(iLL  >= 0 ? v[r][iLL]  : '',
                            iLat >= 0 ? v[r][iLat] : '',
                            iLon >= 0 ? v[r][iLon] : '');
    if (!par) continue;
    var k = _chave_((iChave >= 0 && String(v[r][iChave]).trim()) ? v[r][iChave] : nome);
    if (_CORR[k]) k = _CORR[k];
    if (out[k]) continue;                              // primeira grafia ganha
    out[k] = {
      bl: nome,
      zona: iZona >= 0 ? String(v[r][iZona]).trim() : '',
      lat: Math.round(par[0] * 1e5) / 1e5,
      lon: Math.round(par[1] * 1e5) / 1e5
    };
  }
  return out;
}

function _chave_(t) {
  return String(t).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase().replace(/\s+/g, ' ').trim();
}

function _bairroDoEndereco_(e) {
  e = String(e);
  if (!e.trim()) return '';
  e = e.split(/,\s*Teresina/)[0];
  var p = e.split(' - ').filter(function (x) { return x.trim(); });
  return p.length > 1 ? p[p.length - 1].trim() : '';
}

/**
 * A planilha guarda "FASE TEMPO INTEGRAL" como ano, mas o Sheets costuma
 * interpretar 2022 como data. Nesse caso o Apps Script devolve um objeto Date,
 * e String(date) vira "Wed Jan 01 2022 00:00:00 GMT-0300...". Aqui extraímos
 * só o ano, venha ele como Date, número ou texto.
 */
function _ano_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return String(v.getFullYear());
  var s = String(v).trim();
  var m = s.match(/\b(19|20)\d{2}\b/);
  return m ? m[0] : '';
}

/**
 * Contato pode chegar como número (o Sheets come o zero à esquerda e o
 * parêntese) ou como texto com vários telefones. Devolvemos só dígitos e
 * separadores, deixando a formatação para a página.
 */
function _contatoTxt_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') return String(Math.round(v));
  return String(v).replace(/\s+/g, ' ').trim();
}

function _etapas_(t) {
  t = String(t).toUpperCase(); var e = [];
  if (/\bEF\s*-/.test(t)) e.push('Fundamental');
  if (/\bEM\s*-/.test(t)) e.push('Médio');
  if (t.indexOf('EJA') >= 0) e.push('EJA');
  if (t.indexOf('EPT') >= 0 || t.indexOf('CURSO TECNICO') >= 0) e.push('Técnico');
  return e;
}

function _turnos_(t) {
  var m = String(t).match(/Turmas:\s*([^\n\r]+)/);
  if (!m) return [];
  var vis = {}, out = [];
  m[1].split('|').forEach(function (x) {
    var p = x.trim().split(/\s+/), u = p[p.length - 1];
    if (u && !vis[u]) { vis[u] = 1; out.push(u); }
  });
  return out.sort();
}

function _montarDados_() {
  // CFG, CFG_EXTRA, _plus, _norm_ e _colFlex_ moram no AppsScript_GeoEscolas.gs.
  // Sem aquele arquivo no projeto, o erro que aparece é um "CFG is not defined"
  // apontando para cá, que não diz o que fazer. Este aviso diz.
  if (typeof CFG === 'undefined') {
    throw new Error('Falta o arquivo AppsScript_GeoEscolas.gs no projeto (ou ele não foi salvo). ' +
      'É nele que ficam CFG, CFG_EXTRA e as funções de Plus Code que este arquivo usa. ' +
      'Cole-o como um arquivo de script à parte, salve, e implante de novo.');
  }
  var sh = SpreadsheetApp.getActive().getSheetByName(CFG.ABA);
  if (!sh) return { escolas: [], bairros: [] };
  var v = sh.getDataRange().getValues();
  var cab = v[CFG.LINHA_CAB - 1].map(function (c) { return String(c).trim(); });
  var cabNorm = cab.map(function (c) { return _norm_(c); });
  var C = function (n) { return cab.indexOf(n); };

  var iN = C(CFG.COL_NOME), iP = C(CFG.COL_PLUS), iB = C(CFG.COL_BAIRRO), iE = C(CFG.COL_ENDER),
      iT = C(CFG.COL_TIPO), iS = C(CFG.COL_SIT), iG = C(CFG.COL_GRE), iM = C(CFG.COL_MAT),
      iI = C('INEP'), iZ = C('ZONA'), iF = C('FASE TEMPO INTEGRAL'), iO = C('OFERTAS 2026'),
      iC = _colFlex_(cabNorm, CFG_EXTRA.CONTATO),   // AM — contato do diretor
      iX = _colFlex_(cabNorm, CFG_EXTRA.GESTOR);    // AN — nome do gestor

  var esc = [], soma = {}, rotulos = {};   // rotulos: chave sem acento -> grafia com acento
  for (var r = CFG.LINHA_CAB; r < v.length; r++) {
    var lin = v[r];
    if (!String(lin[iN]).trim()) continue;
    var p = _plus(lin[iP]);
    if (!p) continue;                                  // sem coordenada, fora do mapa

    var bOrig = String(iB >= 0 ? lin[iB] : '').replace(/\s+/g, ' ').trim() ||
                _bairroDoEndereco_(iE >= 0 ? lin[iE] : '');
    var b = _chave_(bOrig);
    if (_CORR[b]) b = _CORR[b];
    if (_LIXO[b] || /^(RUA|AV|AVENIDA)\b/.test(b)) b = '';
    // só serve de rótulo a grafia que não passou por correção — senão "DIRCEU
    // ARCOVERDE II" viraria o nome do grupo "DIRCEU ARCOVERDE"
    if (b && !rotulos[b] && _chave_(bOrig) === b) rotulos[b] = bOrig.trim();

    var of = iO >= 0 ? String(lin[iO]) : '';
    var mat = iM >= 0 ? parseInt(lin[iM], 10) : NaN;

    esc.push({
      n: String(lin[iN]).trim(),
      inep: iI >= 0 ? String(lin[iI]).trim().replace(/\.0+$/, '') : '',
      tip: iT >= 0 ? String(lin[iT]).trim() : '',
      sit: iS >= 0 ? String(lin[iS]).trim() : '',
      b: b,
      end: iE >= 0 ? String(lin[iE]).replace(/,\s*Teresina.*$/, '').trim() : '',
      gre: iG >= 0 ? String(lin[iG]).trim() : '',
      zona: iZ >= 0 ? String(lin[iZ]).trim() : '',
      mat: isNaN(mat) ? null : mat,
      ti: iF >= 0 ? _ano_(lin[iF]) : '',
      gest: iX >= 0 ? String(lin[iX]).replace(/\s+/g, ' ').trim() : '',   // NOME DO GESTOR
      ct: iC >= 0 ? _contatoTxt_(lin[iC]) : '',                           // CONTATO
      lat: Math.round(p[0] * 1e5) / 1e5,
      lon: Math.round(p[1] * 1e5) / 1e5,
      pc: String(lin[iP]).trim().split(/\s/)[0],
      et: _etapas_(of),
      tu: _turnos_(of)
    });

    if (b) {
      if (!soma[b]) soma[b] = { la: 0, lo: 0, n: 0 };
      soma[b].la += p[0]; soma[b].lo += p[1]; soma[b].n++;
    }
  }

  var oficiais = _bairrosOficiais_();

  // rótulo com acento: o nome da relação oficial ganha do que veio da base
  esc.forEach(function (e) {
    e.bl = (oficiais[e.b] && oficiais[e.b].bl) || rotulos[e.b] || e.b;
  });

  // A lista de bairros junta duas fontes. Quem tem escola entra pelo centro
  // geométrico das unidades; a relação oficial traz a cidade inteira. Estando
  // nas duas, a coordenada oficial ganha — ela é o centro do bairro, não a
  // média das escolas. Bairro sem escola entra com n = 0, e é exatamente para
  // ele que a busca por raio serve.
  var mapaB = {};
  Object.keys(soma).forEach(function (k) {
    mapaB[k] = {
      b: k, bl: rotulos[k] || k, zona: '', n: soma[k].n, of: false,
      lat: Math.round(soma[k].la / soma[k].n * 1e5) / 1e5,
      lon: Math.round(soma[k].lo / soma[k].n * 1e5) / 1e5
    };
  });
  Object.keys(oficiais).forEach(function (k) {
    var o = oficiais[k];
    if (mapaB[k]) {
      mapaB[k].lat = o.lat; mapaB[k].lon = o.lon;
      mapaB[k].bl = o.bl;   mapaB[k].zona = o.zona; mapaB[k].of = true;
    } else {
      mapaB[k] = { b: k, bl: o.bl, zona: o.zona, n: 0, of: true, lat: o.lat, lon: o.lon };
    }
  });
  var bairros = Object.keys(mapaB).sort().map(function (k) { return mapaB[k]; });

  return { escolas: esc, bairros: bairros };
}
