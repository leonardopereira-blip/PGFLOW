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
  assert.doesNotMatch(code, /\bSpreadsheetApp\.|\bMailApp\.|\bGmailApp\.|\bgravarNaFila\(/);
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
  const cab=Array(80).fill('');
  cab[1]='1. ID_Código SKU';cab[2]='1. ID_Descrição';
  cab[23]='3. DIM_BXA (Fechado)';cab[27]='4. CAP_Tipo de capa';
  const cells=Array(80).fill('');cells[1]='SKU-A';cells[2]='LIVRO';cells[23]='205X275';cells[27]='SIMPLES';
  const raw={headers:cab,linhas:[cells]};
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

test('CALC Geral preserva formulas de todas as linhas e remove entradas do chamado anterior',()=>{
 const xml='<worksheet><sheetData>'+
 '<row r="5"><c r="C5"><v>header</v></c></row>'+
 '<row r="6"><c r="C6" t="inlineStr"><is><t>Chamado 2910</t></is></c>'+
 '<c r="D6" t="inlineStr"><is><t>ANTIGO</t></is></c>'+
 '<c r="E6"><f>XLOOKUP(D6,D:D,E:E)</f><v>200</v></c>'+
 '<c r="F6"><v>20</v></c><c r="Y6" t="inlineStr"><is><t>DIGITAL</t></is></c></row>'+
 '<row r="7"><c r="C7" t="inlineStr"><is><t>Chamado 2910</t></is></c>'+
 '<c r="D7" t="inlineStr"><is><t>ANTIGO2</t></is></c>'+
 '<c r="E7"><f>XLOOKUP(D7,D:D,E:E)</f><v>300</v></c>'+
 '<c r="F7"><v>10</v></c></row></sheetData></worksheet>';
 const res=context._reSpPreencherCalcGeral_(xml,['SKU-NOVA'],
  {'SKU-NOVA':{avulso:100,kit:0,maleta:0}},'2964',
  {tecnologia:'PLANA',fornecedor:'GRÁFICA',elementos:{}});
 assert.match(res,/XLOOKUP\(D6,D:D,E:E\)/);
 assert.match(res,/XLOOKUP\(D7,D:D,E:E\)/);
 assert.match(res,/Chamado 2964/);
 assert.match(res,/SKU-NOVA/);
 assert.doesNotMatch(res,/ANTIGO/);
 assert.doesNotMatch(res,/Chamado 2910/);
 assert.doesNotMatch(res,/<c r="F7"><v>10/);
});
test('calcPr respeita ordem do XML do Excel',()=>{
 const workbook='<workbook><definedNames></definedNames><calcPr calcId="191029"/>'+
   '<pivotCaches></pivotCaches><extLst></extLst></workbook>';
 const res=context._reSpFullCalc_(workbook);
 assert.ok(res.indexOf('<calcPr')<res.indexOf('<pivotCaches'));
 assert.match(res,/fullCalcOnLoad="1"/);
});

test('NC especificacoes replica regras do exemplo real 2740: 0 numerico onde nao existe operacao',()=>{
 const cab=Array(80).fill('');
 cab[1]='1. ID_Código SKU';
 const linha=Array(80).fill('-');
 linha[1]='SKU-2740'; linha[2]='LIVRO TESTE';linha[27]='Simples';
 linha[29]='Cartão 250g';linha[30]='4x0';
 linha[33]='Laminação BOPP brilho';linha[37]='248';
 linha[38]='Offset 70g';linha[39]='4x4';linha[64]='PUR';
 const ar={headers:['1. ID_Código Kit','13. MAN_Total de itens colecionados','13. MAN_Embalagem'],
  linhas:[['SKU-2740',1,'Caixa']]};
 const tr={headers:['1. ID_Código Kit','1. ID_Código SKU','16. TIR_SKU'],
  linhas:[['SKU-2740','SKU-2740',1000]]};
 const result=context._reSpConstrucaoNC_({headers:cab,linhas:[linha]},ar,tr,[]).especificacoes[0];
 const zeradas=[31,32,34,36,40,41,51,55,61,67,68,75];
 zeradas.forEach(function(col){
   assert.equal(result[col-3],0,'Coluna do De-Para '+col+' deveria ser 0 numerico');
 });
 assert.equal(result[21-3],'SIMPLES');
 assert.equal(result[24-3],'CARTÃO TRIPLEX 250G');
 assert.equal(result[30-3],'LAMINAÇÃO BRILHO');
 assert.equal(result[44-3],'OFFSET 70G');
});

test('resumo nao oculta falhas de API como chamado com tiragem zero',()=>{
  const ui=fs.readFileSync(path.join(__dirname,'..','ValidaInterface.html'),'utf8');
  assert.match(ui,/var resumoComFalha = resumosArr\.some/);
  assert.match(ui,/!!r\.erro/);
  assert.match(ui,/escHtml\(resumo\.erro/);
  assert.match(ui,/#botoesResumo \.btn-continuar/);
  assert.match(ui,/Falha ao carregar dados do SPG/);
  assert.match(ui,/Erro ao consultar o SPG/);
});
