const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const root=path.join(__dirname,'..');
const b=fs.readFileSync(path.join(root,'ReentradasExcel.js'),'utf8');
const gerador=fs.readFileSync(path.join(root,'ValidaReentradasSPG.js'),'utf8');
const interfaceHtml=fs.readFileSync(path.join(root,'ValidaInterface.html'),'utf8');
const context={};vm.createContext(context);
vm.runInContext(gerador,context);vm.runInContext(b,context);
const xml='<worksheet><sheetData><row r="2">'+
  '<c r="A2" s="7"><v>1</v></c>'+
  '<c r="B2" s="8" t="inlineStr"><is><t>Texto antigo</t></is></c>'+
  '<c r="C2" s="4"><f>A2*2</f><v>2</v></c>'+
  '</row></sheetData></worksheet>';
test('prévia não modifica planilha nem consulta fluxos Principal/Email',()=>{
  assert.match(b,/function obterPreviaReentradasExcel\(/);
  assert.match(b,/function salvarEdicoesReentradasExcel\(/);
  assert.match(b,/function listarArquivosReentradasExcel\(/);
  assert.doesNotMatch(b,/GmailApp|MailApp|ValidaPrincipal|SpreadsheetApp/);
  assert.match(interfaceHtml,/id="rxOverlay"/);
  assert.match(interfaceHtml,/Salvar alterações no arquivo/);
  assert.match(interfaceHtml,/rxVoltarAjustes\(\)/);
  assert.match(interfaceHtml,/rxFechar\(\)/);
  assert.match(interfaceHtml,/rxAbrirArquivo\(rxUltimo\)/);
  assert.match(gerador,/calculadoraId:fc\.getId\(\)/);
  assert.match(gerador,/deparaId:fd\.getId\(\)/);
});
test('edição preserva fórmulas, estilos e demais células',()=>{
  const novo=context._rxEditarAba_(xml,[
    {aba:0,linha:2,coluna:1,valor:'1,5'},
    {aba:0,linha:2,coluna:2,valor:'Novo & ok'},
    {aba:0,linha:2,coluna:4,valor:'42'}
  ]);
  assert.match(novo,/<c r="A2" s="7"><v>1\.5<\/v><\/c>/);
  assert.match(novo,/<c r="B2" s="8" t="inlineStr"><is><t>Novo &amp; ok<\/t><\/is><\/c>/);
  assert.match(novo,/<c r="D2"><v>42<\/v><\/c>/);
  assert.match(novo,/<c r="C2" s="4"><f>A2\*2<\/f><v>2<\/v><\/c>/);
  assert.ok(novo.indexOf('A2')<novo.indexOf('B2')&&novo.indexOf('B2')<novo.indexOf('C2')&&novo.indexOf('C2')<novo.indexOf('D2'));
  assert.match(xml,/Texto antigo/);
});
test('não substitui fórmulas nem células especiais',()=>{
  assert.throws(()=>context._rxEditarAba_(xml,[{aba:0,linha:2,coluna:3,valor:'100'}]),/fórmula/);
  assert.throws(()=>context._rxEditarAba_(xml,[{aba:0,linha:4,coluna:1,valor:'100'}]),/não encontrada/);
  assert.throws(()=>context._rxNovaCelula_('A1','<c r="A1" cm="3"><v>1</v></c>','2'),/metadados/);
});
test('zero inicial não vira número e limpeza mantém estilo',()=>{
  assert.match(context._rxNovaCelula_('A2','<c r="A2" s="7"><v>1</v></c>','00123'),/t="inlineStr"/);
  assert.equal(context._rxNovaCelula_('A2','<c r="A2" s="7"><v>1</v></c>',''),'<c r="A2" s="7"/>');
});

test('salvamento integra no mesmo ID e confirma nova revisão sem mudar fórmulas',()=>{
  const refs={};
  const mk=(name,contents)=>({getName(){return name},getDataAsString(){return contents}});
  let partes=[
    mk('xl/workbook.xml','<workbook/>'),
    mk('xl/_rels/workbook.xml.rels','<Relationships/>'),
    mk('xl/worksheets/sheet1.xml',xml),
    mk('xl/styles.xml','<styles intact="sim"/>')
  ];
  const id='ABCD_1234567890_abcdefghij';
  const nome='CALCULADORA FORNECEDOR LOGPRINT - CHAMADO 123.xlsx';
  let atualizadoEm=1700000000000,quantidade=0;
  const blob=(pp)=>({parts:pp,copyBlob(){return this},setContentType(){return this}});
  const arquivo={
    getName(){return nome},getId(){return id},getUrl(){return 'https://drive.google.com/file/d/'+id},
    getMimeType(){return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'},
    isTrashed(){return false},getParents(){let i=0;return{hasNext(){return i<1},next(){i++;return{getId(){return '1DKie1AHE7EFHa6roJY2XxE2ZlFhC_lCy'}}}}},
    getLastUpdated(){return new Date(atualizadoEm)},getBlob(){return blob(partes)}
  };
  const mkHash=(b)=>JSON.stringify(b.parts.map(p=>[p.getName(),p.getDataAsString()]));
  const env={
    DriveApp:{getFileById(fileId){assert.equal(fileId,id);return arquivo}},
    Drive:{Files:{update(meta,fileId,data,opts){
      assert.equal(fileId,id);assert.equal(meta.title,nome);assert.equal(opts.newRevision,true);
      partes=data.parts;atualizadoEm+=1000;quantidade++;
    }}},
    Utilities:{
      unzip(b){return b.parts},
      newBlob(data,type,name){return mk(name,data)},
      zip(parts){return blob(parts)}
    },
    LockService:{getScriptLock(){return{tryLock(){return true},releaseLock(){}}}},
    _reLpAbas_(){return[{nome:'CALC Geral',caminho:'xl/worksheets/sheet1.xml'}]},
    _bcDigest_:mkHash
  };
  vm.createContext(env);vm.runInContext(gerador,env);vm.runInContext(b,env);
  const result=env.salvarEdicoesReentradasExcel(id,1700000000000,[
    {aba:0,linha:2,coluna:2,valor:'Alterado'}
  ]);
  assert.equal(result.sucesso,true,result.erro);
  assert.equal(result.id,id);
  assert.equal(quantidade,1);
  const changed=partes.find(p=>p.getName()==='xl/worksheets/sheet1.xml').getDataAsString();
  assert.match(changed,/Alterado/);
  assert.match(changed,/<f>A2\*2<\/f>/);
  assert.equal(partes.find(p=>p.getName()==='xl/styles.xml').getDataAsString(),'<styles intact="sim"/>');
  const outdated=env.salvarEdicoesReentradasExcel(id,1700000000000,[{aba:0,linha:2,coluna:1,valor:'3'}]);
  assert.match(outdated.erro,/alterado/);
  assert.equal(quantidade,1);
  const badFormula=env.salvarEdicoesReentradasExcel(id,1700000001000,[{aba:0,linha:2,coluna:3,valor:'3'}]);
  assert.match(badFormula.erro,/fórmula/);
  assert.equal(quantidade,1);
});
