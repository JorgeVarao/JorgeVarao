// Roda a lógica do Apps Script em Node, com um SpreadsheetApp de mentira,
// e compara o resultado com o que o gerador Python produziu.
const fs = require("fs");
const dados = JSON.parse(fs.readFileSync("dados_gs.json", "utf8"));
const gravado = {};

function faixa(vals) {
  return {
    getValues: () => vals,
    setValues: () => faixa(vals), setFormulas: () => faixa(vals),
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
    getName: () => nome,
    getDataRange: () => faixa(sh._d),
    getLastRow: () => sh._d.length,
    getLastColumn: () => (sh._d[0] ? sh._d[0].length : 0),
    getRange: function (r, c, nr, nc) {
      nr = nr === undefined ? 1 : nr; nc = nc === undefined ? 1 : nc;
      const bloco = [];
      for (let i = 0; i < nr; i++) {
        const linha = [];
        for (let j = 0; j < nc; j++) {
          const src = sh._d[r - 1 + i];
          linha.push(src ? (src[c - 1 + j] === undefined ? "" : src[c - 1 + j]) : "");
        }
        bloco.push(linha);
      }
      return faixa(bloco);
    },
    clear: () => { sh._d = []; return sh; },
    hideSheet: () => sh, showSheet: () => sh,
    getFilter: () => null,
    setFrozenRows: () => sh, setColumnWidth: () => sh,
  };
  return sh;
}
const abas = {};
for (const k of Object.keys(dados)) abas[k] = aba(k, dados[k]);

const ss = {
  getSheetByName: (n) => abas[n] || null,
  insertSheet: (n) => { abas[n] = aba(n, []); return abas[n]; },
};
global.SpreadsheetApp = {
  getActiveSpreadsheet: () => ss,
  getActiveRange: () => null,
  flush: () => {},
  getUi: () => ({ alert: (m) => console.log("[alert]", String(m).split("\n")[0]) }),
  newDataValidation: () => ({
    requireValueInRange: function(){return this}, requireValueInList: function(){return this},
    setAllowInvalid: function(){return this}, build: () => ({}),
  }),
};

let fonte = fs.readFileSync("/home/user/JorgeVarao/reordenamento-2027/Reordenamento_2027.gs", "utf8");
// captura o que seria gravado
fonte = fonte.replace("function gravar_(ss, nome, headers, linhas, larguras, wrapCols) {",
  "function gravar_(ss, nome, headers, linhas, larguras, wrapCols) {\n  gravado[nome] = { headers: headers, linhas: linhas };");
eval(fonte);

atualizarBases(true);

for (const k of Object.keys(gravado))
  console.log(`${k.padEnd(24)} ${String(gravado[k].linhas.length).padStart(5)} linhas x ${gravado[k].headers.length} cols`);
fs.writeFileSync("gs_saida.json", JSON.stringify(gravado));
