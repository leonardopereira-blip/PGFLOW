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
test('API de Reentradas aplica a normalizacao de dimensoes da aba Espec sem alterar outros campos',()=>{
  const cab=['1. ID_Código SKU','3, DIM_B (Aberto)(mm)','3, DIM_A (Aberto)(mm)',
    '3,DIM_BxA(Aberto)(mm)','3, DIM_B (Fechado)(mm)','3, DIM_A (Fechado)(mm)',
    '3. DIM_E (mm)','3.DIM_BXA(Fechado)','3.DIM_Dimensãoacabada(BxAxE)','2. INF_Marca'];
  const orig=['SKU-1','425,9','275,8','425,9x275,8','205,99',275.7,'15,79',
    '205,99 x 275,7','205,99x275,7x15.79','SAE'];
  const semDecimais=['SKU-2',310.9,440.4,'',310.9,220.6,'7,09','','','SAS'];
  const input={headers:cab,linhas:[orig,semDecimais]};
  const a=context._reSpNormalizarDimensoesAPI_(input);
  assert.deepEqual(Array.from(a.linhas[0]).slice(1,9),
    [425,275,'425X275',205,275,15.7,'205X275','205X275X15,7']);
  assert.deepEqual(Array.from(a.linhas[1]).slice(1,9),
    [310,440,'310X440',310,220,7,'310X220','310X220X7,0']);
  assert.equal(a.linhas[0][9],'SAE');
  assert.equal(orig[1],'425,9','a matriz original nao pode ser modificada');
  const numericos=context._reSpMapearCab_(cab,cab,a.linhas)[0];
  assert.equal(typeof numericos[1],'number');
  assert.equal(typeof numericos[6],'number');
  assert.equal(numericos[8],'205X275X15,7');
});
test('API de Reentradas normaliza escala do peso sem eliminar casas decimais',()=>{
  const cab=['1. ID_Código SKU','3. DIM_Peso (Kg)','2. INF_Marca'];
  const orig={headers:cab,linhas:[
    ['SKU-A','145.900','SAE'],['SKU-B',98100,'SAS'],
    ['SKU-C','1,89','COC'],['SKU-D',0.99,'SAS'],['SKU-E','-','SAE']
  ]};
  const out=context._reSpNormalizarDimensoesAPI_(orig);
  assert.deepEqual(Array.from(out.linhas,x=>x[1]),[1.459,0.981,1.89,0.99,'-']);
  assert.deepEqual(Array.from(out.linhas,x=>x[2]),['SAE','SAS','COC','SAS','SAE']);
  assert.equal(orig.linhas[0][1],'145.900','nao muta dados da API');
});
test('Reentradas usa as referencias da Tiragem do proprio SKU sem travar duplicatas',()=>{
  const espec={headers:['1. ID_Código SKU','11. REF_Referência troca de chapa',
    '11. REF_Referência personalização','1. ID_Descrição'],
    linhas:[['SKU-A','ANTIGA','ANTIGA','Livro A'],['SKU-B','ANTIGA','ANTIGA','Livro B']]};
  const cab=['1. ID_Código Kit','1. ID_Código SKU',
    '11. REF_Referência troca de chapa','11. REF_Referência personalização'];
  const tiragem={headers:cab,linhas:[
    ['KIT-A','SKU-A','CHAPA-A','PERSON-A'],
    ['KIT-A','SKU-A','CHAPA-A','PERSON-A'],
    ['KIT-B','SKU-B','CHAPA-B','']
  ]};
  const atual=context._reSpAplicarReferenciasTiragemAPI_(espec,tiragem);
  assert.deepEqual(Array.from(atual.linhas[0]),['SKU-A','CHAPA-A','PERSON-A','Livro A']);
  assert.deepEqual(Array.from(atual.linhas[1]),['SKU-B','CHAPA-B','','Livro B']);
  assert.equal(espec.linhas[0][1],'ANTIGA','nao altera matriz recebida');
  const semColunas=context._reSpAplicarReferenciasTiragemAPI_(espec,
    {headers:['1. ID_Código SKU'],linhas:[['SKU-A']]});
  assert.equal(semColunas,espec,'OP antiga sem campos de referencia preserva a origem');
  const incompleta=context._reSpAplicarReferenciasTiragemAPI_(
    {headers:['1. ID_Código SKU'],linhas:[['SKU-A']]},tiragem);
  assert.equal(incompleta.linhas[0][1],'CHAPA-A');
  assert.equal(incompleta.linhas[0][2],'PERSON-A');
  const divergente={headers:cab,linhas:[['KIT-1','SKU-A','CHAPA-1',''],['KIT-2','SKU-A','CHAPA-2','']]};
  const repetido=context._reSpAplicarReferenciasTiragemAPI_(espec,divergente);
  assert.equal(repetido.linhas[0][1],'CHAPA-1');
});

test('somente a leitura de chamado da API aplica a normalizacao de dimensoes',()=>{
  const original={
    pacote:context._emailSpgPacote_,partes:context._reSpPartes_,matriz:context._reSpMatriz_
  };
  context._emailSpgPacote_=()=>({id:'2632',versaoId:'v1',blob:{},statusChamado:'Em producao',statusProducao:'Em producao'});
  context._reSpPartes_=()=>({});
  const espec={headers:['1. ID_Código SKU','3, DIM_B (Fechado)(mm)',
    '3, DIM_A (Fechado)(mm)','3. DIM_E (mm)','3.DIM_BXA(Fechado)',
    '3.DIM_Dimensãoacabada(BxAxE)'],
    linhas:[['SKU-1','205,99','275,9','2,78','origem','origem']]};
  const arvore={headers:['1. ID_Código Kit','1. ID_Código SKU'],linhas:[['KIT-1','SKU-1']]};
  const tiragem={headers:['1. ID_Código Kit','1. ID_Código SKU','15. DES_CD envio','16. TIR_SKU',
    '11. REF_Referência troca de chapa','11. REF_Referência personalização'],
    linhas:[['KIT-1','SKU-1','JDI',100,'CHAPA-2632','PERSON-2632']]};
  context._reSpMatriz_=(_,nome)=>nome==='Esp. Técnica (legado)'?espec:
    nome==='Árvore (legado)'?arvore:tiragem;
  try{
    const result=context._reSpPrepararDados_('2632','LOGPRINT');
    assert.equal(result.espec.linhas[0][4],'205X275');
    assert.equal(result.espec.linhas[0][5],'205X275X2,7');
    assert.equal(result.espec.linhas[0][result.espec.headers.indexOf('11. REF_Referência troca de chapa')],'CHAPA-2632');
    assert.equal(result.espec.linhas[0][result.espec.headers.indexOf('11. REF_Referência personalização')],'PERSON-2632');
    assert.equal(espec.linhas[0][4],'origem','dados extraidos nao sao mutados');
  }finally{
    context._emailSpgPacote_=original.pacote;
    context._reSpPartes_=original.partes;
    context._reSpMatriz_=original.matriz;
  }
});

test('valores numericos do SPG e do legado entram como numero sem alterar SKU, ISBN ou texto',()=>{
  const cab=['1. ID_Código SKU','3, DIM_B (Aberto)(mm)','3. DIM_Peso (Kg)',
    '4.Cap_orelha','16. TIR_SKU','2. INF_Série','1. ID_ISBN'];
  const linha=['000123','210','0,125','80,5','1.234','3','9780000000123'];
  const mapped=context._reSpMapearCab_(cab,cab,[linha])[0];
  assert.equal(mapped[0],'000123');
  assert.equal(mapped[1],210);
  assert.equal(mapped[2],0.125);
  assert.equal(mapped[3],80.5);
  assert.equal(mapped[4],1234);
  assert.equal(mapped[5],'3');
  assert.equal(mapped[6],'9780000000123');
  assert.match(context._reSpCelula_('B2',mapped[1]),/<v>210<\/v>/);
  assert.match(context._reSpCelula_('A2',mapped[0]),/t="inlineStr"/);
});
test('leitor de XLSX preserva tipo numerico original e nao converte texto de codigo',()=>{
  const xml='<worksheet><sheetData><row r="1">'+
    '<c r="A1"><v>210</v></c><c r="B1" t="inlineStr"><is><t>00123</t></is></c>'+
    '</row></sheetData></worksheet>';
  const x={'xl/worksheets/sheet1.xml':{getDataAsString:()=>xml}};
  const row=context._reSpLerAba_(x,1,[])[0].cells;
  assert.equal(row[0],210);assert.equal(row[1],'00123');
});

test('SKU e kit numericos usam o tipo das consultas sem perder codigos de texto',()=>{
  for(const [entrada,esperado] of [['200713648',200713648],['0',0],['999999999999999',999999999999999],
    ['MS26VF331P1SP','MS26VF331P1SP'],['000123','000123'],['1000000000000000','1000000000000000'],
    ['20E3','20E3'],['200.123','200.123'],[' 200713648 ',' 200713648 '],['','']]){
    assert.equal(context._reSpCodigoSku_(entrada),esperado);
    assert.equal(context._reSpNumeroPorCabecalho_('1. ID_Código SKU',entrada),esperado);
    assert.equal(context._reSpNumeroPorCabecalho_('1. ID_Código Kit',entrada),esperado);
  }
  assert.equal(context._reSpNumeroPorCabecalho_('1. ID_ISBN','9780000000123'),'9780000000123');
  assert.equal(context._reSpNumeroPorCabecalho_('2. INF_Série','200713648'),'200713648');
  assert.match(context._reSpCelula_('A1','200713648'),/t="inlineStr"/,'serializador generico preserva texto');
  const h=['1. ID_Código Kit','1. ID_Código SKU'];
  const dados=context._reSpMapearCab_(h,h,[['200713648',200713648],[200713648,'200713648']]);
  for(const row of dados)assert.deepEqual(Array.from(row),[200713648,200713648]);
});
test('formato de referencias Excel e protecao original preservada',()=>{
  assert.equal(context._reSpCol_(1),'A');
  assert.equal(context._reSpCol_(81),'CC');
  assert.equal(context._reSpIdx_('CC'),80);
  const hash=context._reSpSenhaHash_('PGFLOW2026');
  assert.match(hash,/^[0-9A-F]{4}$/);
  const xml='<worksheet><sheetData><row r="1"></row></sheetData><extLst/></worksheet>';
  assert.equal(context._reSpProteger_(xml),xml);
  const protegido=xml.replace('</sheetData>','</sheetData><sheetProtection password="'+hash+'" sheet="1" autoFilter="0"/>');
  assert.equal(context._reSpProteger_(protegido),protegido);
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

test('orelha aceita Sim/Nao sem caixa e preserva largura numerica sem interpretar Nao como largura',()=>{
  const cab=Array(80).fill('');cab[1]='1. ID_Código SKU';
  const ar={headers:['1. ID_Código Kit'],linhas:[['SKU-A']]};
  const tr={headers:['1. ID_Código Kit','1. ID_Código SKU','16. TIR_SKU'],linhas:[['SKU-A','SKU-A',100]]};
  for(const [entrada,esperado,largura] of [[false,'Não',0],['FALSE','Não',0],['nÃo','Não',0],
    ['nao','Não',0],['Sim','Sim',0],['TrUe','Sim',0],[0,'Não',0],[80,'Sim',80],['80,5','Sim',80.5]]){
    const row=Array(80).fill('');row[1]='SKU-A';row[5]='Maralto';row[28]=entrada;
    const r=context._reSpConstrucaoNC_({headers:cab,linhas:[row]},ar,tr,[]).especificacoes[0];
    assert.equal(r[27-3],esperado,String(entrada));assert.equal(r[28-3],largura,String(entrada));
  }
});
test('interface usa fluxo separado e retorna arquivos',()=>{
  const ui=fs.readFileSync(path.join(__dirname,'..','ValidaInterface.html'),'utf8');
  assert.match(ui,/gerarValidacaoReentradasSPG/);
  assert.match(ui,/validarReentradasSPG/);
  assert.match(ui,/buscarResumoReentradasSPG/);
  assert.match(ui,/Abrir De-Para/);
  assert.doesNotMatch(ui,/\.finalizarProcessamentoLoteDepara\(/);
});

test('CALC Geral preserva formulas ate o ultimo SKU e remove linhas sobrando',()=>{
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
 assert.doesNotMatch(res,/XLOOKUP\(D7,D:D,E:E\)/);
 assert.match(res,/Chamado 2964/);
 assert.match(res,/SKU-NOVA/);
 assert.doesNotMatch(res,/ANTIGO/);
 assert.doesNotMatch(res,/Chamado 2910/);
 assert.doesNotMatch(res,/<c r="F7"><v>10/);
 assert.doesNotMatch(res,/<row r="7"/);
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

test('CALC Geral deixa vazias as tecnologias nao selecionadas e inicia em Digital',()=>{
  const cols={
    DIGITAL:['CF','CK','CP','CU','CZ','DE','DI'],
    PLANA:['DN','DS','DX','EC','EH','EM','EQ'],
    ROTATIVA:['EV','FA','FF','FK','FP','FU','FY']
  };
  const antigas=Object.values(cols).flat().map(c=>'<c r="'+c+'6" t="inlineStr"><is><t>ANTIGO</t></is></c>').join('');
  const xml='<worksheet><sheetData><row r="6">'+antigas+'<c r="E6"><f>1+1</f></c></row>'+
    '<row r="7"><c r="E7"><f>1+1</f></c></row></sheetData></worksheet>';
  const entrada={'SKU-1':{avulso:10,kit:20,maleta:0}};
  for(const escolha of ['DIGITAL','PLANA','ROTATIVA',undefined]){
    const f={fornecedor:'GRÁFICA',elementos:{
      DIGITAL:{capa:'GRÁFICA',forro:'ARCO',miolo_principal:'GRÁFICA'},
      PLANA:{capa:'ARCO',forro:'ARCO',miolo_principal:'GRÁFICA'},
      ROTATIVA:{capa:'ARCO TRIANG',forro:'ARCO',miolo_principal:'GRÁFICA'}
    }};
    if(escolha)f.tecnologia=escolha;
    const out=context._reSpPreencherCalcGeral_(xml,['SKU-1'],entrada,'1234',f);
    for(const [tec,all] of Object.entries(cols)){
      for(const [i,col] of all.entries()){
        assert.equal(new RegExp('<c r="'+col+'6"[^>]*><is>').test(out),tec===(escolha||'DIGITAL')&&i<3,
          (escolha||'DIGITAL')+' '+col);
      }
    }
    assert.match(out,/<c r="E6"><f>1\+1<\/f>/);
    assert.doesNotMatch(out,/<row r="7"/);
    assert.doesNotMatch(out,/ANTIGO/);
  }
});

test('novas entradas materializam o estilo herdado e celulas vazias conservam seu estilo',()=>{
  const xml='<worksheet><cols><col min="84" max="85" style="584"/></cols><sheetData><row r="6">'+
    '<c r="D6"><v>10</v></c><c r="CG6" s="12"><v>antigo</v></c>'+
    '<c r="DN6" s="583"><v>antigo</v></c><c r="E6" s="9"><f>1+1</f></c></row></sheetData></worksheet>';
  const out=context._reSpPreencherCalcGeral_(xml,['200713648'],{'200713648':{avulso:10,kit:0,maleta:0}},'2932',{});
  assert.match(out,/<c r="CF6" s="584" t="inlineStr">/,'entrada nova precisa do estilo da coluna para continuar editavel');
  assert.match(out,/<c r="D6" s="0"><v>200713648<\/v>/,'celula original sem s conserva XF0');
  assert.match(out,/<c r="DN6" s="583"\/>/,'tecnologia vazia conserva bloqueio e formato');
  assert.match(out,/<c r="E6" s="9"><f>1\+1<\/f>/,'formula intacta');
  const styles=context._reSpEstilos_(xml,6);
  assert.equal(styles[83],'584');
  assert.equal(styles[3],'0');
});
