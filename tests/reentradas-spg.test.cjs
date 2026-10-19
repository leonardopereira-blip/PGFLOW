const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const path=require('node:path');
const code=fs.readFileSync(path.join(__dirname,'..','ValidaReentradasSPG.js'),'utf8');
const context={};
vm.createContext(context);
vm.runInContext(code,context);
test('script isolado nao usa Google Sheets nem envio de email',()=>{
  assert.doesNotMatch(code, /SpreadsheetApp|MailApp|GmailApp|gravarNaFila|ValidaPrincipal/);
  assert.match(code,/function gerarValidacaoReentradasSPG\(/);
  assert.match(code,/\/api\/tiragem\/.*\/alocacao/);
});
test('normalizacao numerica com virgula e ponto',()=>{
  assert.equal(context._reSpQuant_('4.146'),4146);
  assert.equal(context._reSpQuant_('4146'),4146);
  assert.equal(context._reSpQuant_('4.146,5'),4146.5);
  assert.equal(context._reSpQuant_('2.5'),2.5);
  assert.throws(()=>context._reSpQuant_(''),/vazia/);
  assert.throws(()=>context._reSpQuant_('-10'),/invalida/);
});
test('formato de referencias Excel e protecao criada no final',()=>{
  assert.equal(context._reSpCol_(1),'A');
  assert.equal(context._reSpCol_(81),'CC');
  assert.equal(context._reSpIdx_('CC'),80);
  const hash=context._reSpSenhaHash_('PGFLOW2026');
  assert.match(hash,/^[0-9A-F]{4}$/);
  const xml='<worksheet><sheetData><row r="1"></row></sheetData><extLst/></worksheet>';
  assert.match(context._reSpProteger_(xml),new RegExp('sheetProtection password="'+hash+'"'));
  assert.match(context._reSpFullCalc_('<workbook></workbook>'),/fullCalcOnLoad="1"/);
});
test('mapa de especificacao tem 79 campos e nao replica valores do chamado 2740',()=>{
  const cab=['1. ID_Código SKU','1. ID_Descrição','4. CAP_Tipo de capa','3. DIM_BXA (Fechado)'];
  const raw={headers:cab,linhas:[['SKU-A','LIVRO','SIMPLES','205X275']]};
  const ar={headers:['1. ID_Código Kit','13. MAN_Total de itens colecionados'],linhas:[['SKU-A','1']]};
  const tr={headers:['1. ID_Código Kit','1. ID_Código SKU','16. TIR_SKU'],linhas:[['SKU-A','SKU-A','100']]};
  const r=context._reSpConstrucaoNC_(raw,ar,tr,[]);
  assert.equal(r.especificacoes.length,1);
  assert.equal(r.especificacoes[0].length,79);
  assert.equal(r.especificacoes[0][0],'SKU-A');
  assert.equal(r.calc[0].quant,100);
});
test('interface usa fluxo separado e retorna arquivos',()=>{
  const ui=fs.readFileSync(path.join(__dirname,'..','ValidaInterface.html'),'utf8');
  assert.match(ui,/gerarValidacaoReentradasSPG/);
  assert.match(ui,/validarReentradasSPG/);
  assert.match(ui,/buscarResumoReentradasSPG/);
  assert.match(ui,/Abrir De-Para/);
  assert.doesNotMatch(ui,/\.finalizarProcessamentoLoteDepara\(/);
});
