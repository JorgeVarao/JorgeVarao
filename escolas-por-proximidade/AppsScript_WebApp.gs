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
  var hit = cache.get('dadosApp_v2');            // v2: passou a carregar gestor e contato
  if (hit) return hit;
  var js = JSON.stringify(_montarDados_());
  try { cache.put('dadosApp_v2', js, 600); } catch (e) {}  // >100KB não cabe: segue sem cache
  return js;
}

/** Força atualização imediata depois de editar a planilha. */
function LIMPAR_CACHE() {
  var c = CacheService.getScriptCache();
  c.remove('dadosApp_v2');
  c.remove('dadosApp_v1');                       // resto da versão anterior, se houver
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

  esc.forEach(function (e) { e.bl = rotulos[e.b] || e.b; });   // rótulo com acento

  var bairros = Object.keys(soma).sort().map(function (k) {
    return {
      b: k,
      bl: rotulos[k] || k,
      lat: Math.round(soma[k].la / soma[k].n * 1e5) / 1e5,
      lon: Math.round(soma[k].lo / soma[k].n * 1e5) / 1e5,
      n: soma[k].n
    };
  });

  return { escolas: esc, bairros: bairros };
}
