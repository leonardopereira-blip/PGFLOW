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
  assert.doesNotMatch(code,/\/api\/tiragem\/.*\/alocacao/);
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

test('lista de graficas de reentradas nao depende de Sheets nem inclui graficas sem calculadora',()=>{
  const graficas=context.listarPastasGraficasReentradasSPG();
  assert.deepEqual(Array.from(graficas,g=>g.id),['LOGPRINT','CASADAGRAFICA','FORMACERTA','META','REPROSET']);
  assert.equal(new Set(graficas.map(g=>g.id)).size,5);
  const ui=fs.readFileSync(path.join(__dirname,'..','ValidaInterface.html'),'utf8');
  assert.match(ui,/listarPastasGraficasReentradasSPG/);
  assert.doesNotMatch(ui,/\.obterPastasGraficas\(/);
});
test('resumo traz dados reais da OP de producao obtida pela API',()=>{
  const orig=context._reSpPrepararDados_;
  context._reSpPrepararDados_=(id,grafica)=>{
    assert.equal(id,'2964');
    assert.equal(grafica,'LOGPRINT');
    return {
      espec:{headers:['1. ID_Código SKU','2. INF_Marca','2. INF_Série','4. CAP_Papel da capa','5. MIO_PRIN_Papel'],linhas:[
        ['SKU-1','SAS','3ª Série','Couché 250g','Offset 75g'],
        ['SKU-2','SAS','3ª Série','Couché 250g','Offset 75g'],
        ['SKU-3','Geekie','2ª Série','Cartão 250g','Pólen 80g']
      ]},
      tiragem:{headers:['16. TIR_SKU','15. DES_CD envio'],linhas:[['100','FOR'],['200','JDI']]}
    };
  };
  try{
    const res=context.buscarResumoReentradasSPG('2964');
    assert.equal(res[0].totalSkus,3);
    assert.equal(res[0].tiragemTotal,300);
    assert.equal(res[0].marcas,'SAS, Geekie');
    assert.equal(res[0].series,'3ª Série, 2ª Série');
    assert.equal(res[0].papeisCapa,'Couché 250g, Cartão 250g');
    assert.equal(res[0].papeisMiolo,'Offset 75g, Pólen 80g');
    assert.equal(res[0].cdsDestino,'FOR, JDI');
  }finally{context._reSpPrepararDados_=orig;}
});

test('tela de custos SPG nao simula a esteira do Office e so conclui com arquivos confirmados',()=>{
  const ui=fs.readFileSync(path.join(__dirname,'..','ValidaInterface.html'),'utf8');
  assert.match(ui,/Gerando De-Para e calculadora com dados do SPG/);
  assert.match(ui,/De-Para e calculadora gerados no Drive/);
  assert.match(ui,/Array\.isArray\(res\.arquivos\)/);
  assert.match(ui,/concluirTracker\(\)/);
  assert.doesNotMatch(ui,/pollStatus|checarStatusProcessamento|_hashLote|_jobs|Dados em tr[aâ]nsito para o office/i);
  assert.doesNotMatch(ui,/<div id="step2"|<div id="step3"/);
});

test('custos nao oferecem selecao de versao legada e rejeitam cliente antigo',()=>{
 const ui=fs.readFileSync(path.join(__dirname,'..','ValidaInterface.html'),'utf8');
 assert.doesNotMatch(ui,/opcaoVersao|txtVersoes|alternarCampoManual/);
 const r=context.gerarValidacaoReentradasSPG('LOGPRINT',['2964'],'selecionar','1',[],{});
 assert.equal(r.sucesso,false);
 assert.match(r.erro,/versão vinculada/);
});
test('entradas antigas de custos falham sem ler Sheets ou enviar Office',()=>{
 const legacy=fs.readFileSync(path.join(__dirname,'..','ValidaReentradas.js'),'utf8');
 const isolated={}; vm.createContext(isolated); vm.runInContext(legacy,isolated);
 ['obterChamadosLogadosMaster','validarFase1Multiplos','finalizarProcessamentoLoteDepara','checarStatusProcessamento','buscarResumoChamados'].forEach(name=>{
   assert.throws(()=>isolated[name](),/fluxo antigo.*desativado/);
 });
 assert.doesNotMatch(legacy,/MailApp|GmailApp|gravarNaFila|ID_PLANILHA_LOG/);
});

function fixtureOp2964(){
 const skus=['PG25LP295SDI0','PG26LP276SDI0','SAE27ARACL21102BRDO200','SAE27LSPEN38101BRPDR00'];
 const matrizes={
  'Esp. Técnica (legado)':{headers:['1. ID_Código SKU'],linhas:skus.map(sku=>[sku])},
  'Árvore (legado)':{headers:['1. ID_Código Kit','1. ID_Código SKU','13. MAN_Total de itens colecionados'],
   linhas:skus.flatMap(sku=>[[sku,'',1],[sku,sku,'']])},
  'Tiragem (legado)':{headers:['1. ID_Código Kit','1. ID_Código SKU','15. DES_CD envio','16. TIR_SKU','16. TIR_Item','16. TIR_Miolo','GRAFICA'],
   linhas:[[skus[0],'CITA',200],[skus[0],'CJUN',200],[skus[1],'CITA',150],
    [skus[2],'CITA',250],[skus[3],'CITA',140],[skus[3],'CJUN',60]]
    .map(([sku,cd,qtd])=>[sku,sku,cd,qtd,0,0,'OUTRA GRAFICA'])}
 };
 const pacote={id:'2964',versaoId:'dac7d511-83a1-4670-b9e9-cccffe06942e',blob:{},statusChamado:'SOLICITADO',statusProducao:'PUBLICADA'};
 const isolated={};vm.createContext(isolated);vm.runInContext(code,isolated);
 isolated._emailSpgPacote_=(id,forcar,tipo)=>{
  assert.equal(id,'2964');assert.equal(tipo,'reentrada');return pacote;
 };
 isolated._reSpPartes_=()=>({});
 isolated._reSpMatriz_=(_,nome)=>matrizes[nome];
 isolated._emailSpgLerJson_=()=>{throw Error('Consulta adicional de alocacao nao permitida');};
 return {isolated,pacote,matrizes};
}

test('modelo escolhido usa OP integral do chamado sem exigir alocacao ou alterar dados',()=>{
 const {isolated,matrizes}=fixtureOp2964(),antes=JSON.stringify(matrizes);
 const res=isolated._reSpPrepararDados_('2964','LOGPRINT');
 assert.equal(res.espec.linhas.length,4);
 assert.equal(res.arvore.linhas.length,8);
 assert.equal(res.tiragem.linhas.length,6);
 assert.equal(res.tiragem.linhas.reduce((soma,linha)=>soma+Number(linha[3]),0),1000);
 assert.ok(res.tiragem.linhas.every(linha=>linha[4]===0&&linha[5]===0&&linha[6]==='OUTRA GRAFICA'));
 assert.equal(JSON.stringify(matrizes),antes);
});

test('OP integral mantem bloqueio de chamado cancelado',()=>{
 const {isolated,pacote}=fixtureOp2964();pacote.statusChamado='CANCELADO';
 assert.throws(()=>isolated._reSpPrepararDados_('2964','LOGPRINT'),/cancelad/i);
});

test('OP integral rejeita tiragem sem especificacao ou sem par kit e SKU na arvore',()=>{
 const semEspec=fixtureOp2964();semEspec.matrizes['Esp. Técnica (legado)'].linhas.shift();
 assert.throws(()=>semEspec.isolated._reSpPrepararDados_('2964','LOGPRINT'),/especifica/i);
 const semArvore=fixtureOp2964();semArvore.matrizes['Árvore (legado)'].linhas.splice(1,1);
 assert.throws(()=>semArvore.isolated._reSpPrepararDados_('2964','LOGPRINT'),/árvore|arvore/i);
});
