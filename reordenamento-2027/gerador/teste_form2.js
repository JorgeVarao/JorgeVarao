// Confere se as fórmulas que o Apps Script escreve batem com as do .xlsx
const fs = require("fs");
const dados = JSON.parse(fs.readFileSync("dados_gs.json", "utf8"));
const gsSaida = JSON.parse(fs.readFileSync("gs_saida.json", "utf8"));
const escritas = {};   // aba -> col -> [formulas]

function faixa(vals, sh, r, c, nr, nc) {
  return {
    getValues: () => vals,
    setValues: (v) => { if (sh && sh._set) sh._set(r, c, v); return faixa(vals); },
    setFormulas: (v) => {
      if (sh) {
        escritas[sh._nome] = escritas[sh._nome] || {};
        escritas[sh._nome][c] = v.map(x => x[0]);
      }
      return faixa(vals);
    },
    setValue: () => {}, setFormula: () => {},
    setFontFamily: function(){return this}, setFontSize: function(){return this},
    setFontWeight: function(){return this}, setBackground: function(){return this},
    setFontColor: function(){return this}, setWrap: function(){return this},
    setHorizontalAlignment: function(){return this}, setVerticalAlignment: function(){return this},
    setNumberFormat: function(){return this}, setDataValidation: function(){return this},
    createFilter: function(){return this},
  };
}
function aba(nome, linhas) {
  const sh = {
    _nome: nome, _d: linhas || [],
    _set: (r, c, v) => { for (let i=0;i<v.length;i++){ sh._d[r-1+i] = sh._d[r-1+i] || [];
                          for (let j=0;j<v[i].length;j++) sh._d[r-1+i][c-1+j]=v[i][j]; } },
    getName: () => nome,
    getDataRange: () => faixa(sh._d, sh, 1, 1),
    getLastRow: () => sh._d.length,
    getLastColumn: () => Math.max.apply(null, sh._d.map(r => r.length).concat([0])),
    getRange: function (r, c, nr, nc) {
      nr = nr === undefined ? 1 : nr; nc = nc === undefined ? 1 : nc;
      const bloco = [];
      for (let i = 0; i < nr; i++) { const L=[];
        for (let j = 0; j < nc; j++) { const src = sh._d[r-1+i];
          L.push(src ? (src[c-1+j] === undefined ? "" : src[c-1+j]) : ""); }
        bloco.push(L); }
      return faixa(bloco, sh, r, c, nr, nc);
    },
    clear: () => { sh._d = []; return sh; },
    hideSheet: () => sh, showSheet: () => sh,
    getFilter: () => null, setFrozenRows: () => sh, setColumnWidth: () => sh,
  };
  return sh;
}

const abas = {};
for (const k of Object.keys(dados)) abas[k] = aba(k, dados[k]);
// as bases que atualizarBases teria criado
for (const k of Object.keys(gsSaida))
  abas[k] = aba(k, [gsSaida[k].headers].concat(gsSaida[k].linhas));
// V2 e V3 com os mesmos INEPs do .xlsx
const ineps = gsSaida["Base Tratada"].linhas.map(r => r[0]);
abas["Reordenamento 2027 V2"] = aba("Reordenamento 2027 V2",
  [new Array(38).fill("")].concat(ineps.map(i => { const L=new Array(38).fill(""); L[0]=i; return L; })));
abas["Reordenamento 2027 V3"] = aba("Reordenamento 2027 V3",
  [new Array(51).fill("")].concat(ineps.map(i => { const L=new Array(51).fill(""); L[0]=i; return L; })));
abas["Fusão Turmas"] = aba("Fusão Turmas", new Array(63).fill(0).map(()=>new Array(18).fill("")));
abas["IDEB - ESCOLAS"] = aba("IDEB - ESCOLAS", new Array(531).fill(0).map(()=>new Array(6).fill("")));

const ss = { getSheetByName: (n) => abas[n] || null,
             insertSheet: (n) => { abas[n] = aba(n, []); return abas[n]; } };
global.SpreadsheetApp = {
  getActiveSpreadsheet: () => ss, getActiveRange: () => null, flush: () => {},
  getUi: () => ({ alert: () => {} }),
  newDataValidation: () => ({ requireValueInRange: function(){return this},
    requireValueInList: function(){return this}, setAllowInvalid: function(){return this},
    build: () => ({}) }),
};
const grav={};
let fonte=fs.readFileSync("/home/user/JorgeVarao/reordenamento-2027/Reordenamento_2027.gs","utf8");
fonte=fonte.replace("function gravar_(ss, nome, headers, linhas, larguras, wrapCols) {",
  "function gravar_(ss, nome, headers, linhas, larguras, wrapCols) {\n  grav[nome]={headers:headers,linhas:linhas};");
eval(fonte);
atualizarBases(true);
reescreverFormulas(true);
fs.writeFileSync("gs_grav2.json", JSON.stringify(grav));
fs.writeFileSync("gs_formulas.json", JSON.stringify(escritas));
for (const a of Object.keys(escritas))
  console.log(a, "→ colunas:", Object.keys(escritas[a]).join(","));
