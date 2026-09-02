// Confere se TODA coluna declarada é achada pelo nome, sem cair na reserva.
const fs=require("fs");
const avisos=[];
global.Logger={log:m=>avisos.push(String(m))};
global.SpreadsheetApp={getActiveSpreadsheet:()=>ss,getUi:()=>({alert:()=>{}}),flush:()=>{},
  newDataValidation:()=>({requireValueInRange(){return this},requireValueInList(){return this},
    setAllowInvalid(){return this},build:()=>({})})};
const cabs=JSON.parse(fs.readFileSync("cabecalhos_reais.json","utf8"));
function aba(nome){const c=cabs[nome]||[];return{
  getName:()=>nome,getLastColumn:()=>c.length,getLastRow:()=>2,
  getRange:(r,cc,nr,nc)=>({getValues:()=>[c.slice(cc-1,cc-1+(nc||1))]}),
  getDataRange:()=>({getValues:()=>[c]})};}
const abas={}; for(const k of Object.keys(cabs)) abas[k]=aba(k);
const ss={getSheetByName:n=>abas[n]||null,insertSheet:n=>abas[n]=aba(n)};
eval(fs.readFileSync("/home/user/JorgeVarao/reordenamento-2027/Reordenamento_2027.gs","utf8"));
limparCacheColunas_();
let erros=0, total=0;
for(const nomeAba of Object.keys(NOMES_COLUNAS)){
  if(!cabs[nomeAba]){console.log(`  (sem cabeçalho de referência: ${nomeAba})`);continue;}
  const m=colunasDe_(ss,nomeAba);
  const usados={};
  for(const ap of Object.keys(m)){
    total++;
    const c=m[ap];
    if(usados[c]){console.log(`  ✗ ${nomeAba}: "${ap}" e "${usados[c]}" apontam para a mesma coluna ${letra_(c)}`);erros++;}
    usados[c]=ap;
  }
}
console.log(`\n${total} colunas declaradas`);
if(avisos.length){console.log(`${avisos.length} caíram na posição de reserva:`);avisos.forEach(a=>console.log("   "+a));erros+=avisos.length;}
else console.log("nenhuma caiu na reserva ✓");
console.log(erros?"\nATENÇÃO":"\nTODAS RESOLVIDAS PELO NOME ✓");
