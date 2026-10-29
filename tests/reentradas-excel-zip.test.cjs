const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const crypto=require('node:crypto');
const mx=require('../tools/molde-node.cjs');

function ambiente(){
  const app=mx.carregarApps(['MoldeXlsx.js','MoldeMetaCalc.js','ValidaReentradasSPG.js',
    'MacroProtecao.js','ProtecaoLogprint.js','ProtecaoCalculadoras.js','ReentradasExcel.js']);
  const u=app.Utilities;
  u.zip=(partes,nome)=>u.newBlob(mx.escreverZip(partes.map(p=>({nome:p.getName(),dados:Buffer.from(p.getBytes())}))),app.RX_MIME,nome);
  u.unzip=b=>mx.lerZip(Buffer.from(b.getBytes())).map(p=>u.newBlob(p.dados,'application/xml',p.nome));
  u.DigestAlgorithm={SHA_256:'sha256'};
  u.computeDigest=(alg,bytes)=>crypto.createHash(alg).update(Buffer.from(bytes)).digest();
  return app;
}
const partes=b=>new Map(mx.lerZip(Buffer.from(b.getBytes())).map(p=>[p.nome,p.dados]));
const formulas=xml=>Array.from(xml.matchAll(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g),m=>m[0]).filter(c=>/<f\b/.test(c));
const celula=(xml,ref)=>new RegExp('<c\\b[^>]*r="'+ref+'"[^>]*(?:/>|>[\\s\\S]*?</c>)').exec(xml)?.[0];

test('XLSX META real: edição substitui mesmo ID, mantém todas partes, fórmulas, estilos e proteção',()=>{
  const app=ambiente(),u=app.Utilities,molde=app._moldeMetaCalc_(),p=app._mxPartes_(molde);
  const geral=app._reSpAbaPorNome_(p,'CALC Geral');
  p.partes=p.partes.map(parte=>parte.getName()==='xl/workbook.xml'?
    u.newBlob(app._reSpFullCalc_(parte.getDataAsString()),'application/xml',parte.getName()):parte);
  const inicial=u.zip(p.partes,'CALCULADORA FORNECEDOR META - CHAMADO TESTE.xlsx');
  const id='teste_integridade_excel_20261029',nome=inicial.getName();
  let atual=inicial,mtime=1800000000000,updates=0;
  const file={getName:()=>nome,getId:()=>id,getUrl:()=>`https://drive.google.com/file/d/${id}`,
    getMimeType:()=>app.RX_MIME,isTrashed:()=>false,getLastUpdated:()=>new Date(mtime),getBlob:()=>atual,
    getParents:()=>{let i=0;return {hasNext:()=>i===0,next:()=>{i++;return {getId:()=>app.RX_PASTA};}};}};
  app.DriveApp={getFileById:n=>{assert.equal(n,id);return file;}};
  app.Drive={Files:{update(meta,n,blob,opts){assert.equal(n,id);assert.equal(meta.title,nome);assert.equal(opts.newRevision,true);atual=blob;mtime++;updates++;}}};
  app.LockService={getScriptLock:()=>({tryLock:()=>true,releaseLock(){}})};
  const aba=app._reLpAbas_(p.mapa['xl/workbook.xml'].getDataAsString(),p.mapa['xl/_rels/workbook.xml.rels'].getDataAsString()).findIndex(s=>s.nome==='CALC Geral');
  const edicao={aba,linha:6,coluna:7,valor:'12,5'};
  const r=app.salvarEdicoesReentradasExcel(id,mtime,[edicao]);
  assert.equal(r.sucesso,true,r.erro);assert.equal(r.id,id);assert.equal(updates,1);
  const antes=partes(inicial),depois=partes(atual);
  assert.deepEqual([...depois.keys()],[...antes.keys()]);
  let quantidadeFormulas=0;
  for(const [n,bytes] of antes){
    if(n!==geral){assert.ok(bytes.equals(depois.get(n)),n+' mudou sem ser editada');continue;}
    const a=bytes.toString('utf8'),b=depois.get(n).toString('utf8');
    assert.deepEqual(formulas(b),formulas(a),'alguma fórmula/cache mudou');
    assert.equal(a.replace(celula(a,'G6'),'CELULA_EDITADA'),b.replace(celula(b,'G6'),'CELULA_EDITADA'));
    assert.equal(celula(b,'G6'),'<c r="G6" s="582"><v>12.5</v></c>');
  }
  for(const [n,bytes] of depois)if(/^xl\/worksheets\//.test(n))quantidadeFormulas+=formulas(bytes.toString('utf8')).length;
  assert.ok(quantidadeFormulas>375000,'não conferiu fórmulas reais suficientes: '+quantidadeFormulas);
  const desprotegido=u.zip(u.unzip(atual).map(parte=>/^xl\/worksheets\//.test(parte.getName())?
    u.newBlob(parte.getDataAsString().replace(/<sheetProtection\b[^>]*?(?:\/>|>[\s\S]*?<\/sheetProtection>)/g,''),'application/xml',parte.getName()):parte),nome);
  const bloqueado=app._bcMontarXlsx_(desprotegido,'META',nome).blob;
  const restauradas=partes(bloqueado);
  for(const [n,bytes] of depois){
    if(/^xl\/worksheets\//.test(n)){
      assert.equal(app._bcSemMudancaDeConteudo_(restauradas.get(n).toString('utf8')),app._bcSemMudancaDeConteudo_(bytes.toString('utf8')),n+' mudou conteúdo ao restaurar');
      assert.deepEqual(formulas(restauradas.get(n).toString('utf8')),formulas(bytes.toString('utf8')));
    }else if(n!=='xl/styles.xml'&&n!=='[Content_Types].xml'&&n!=='xl/workbook.xml')assert.ok(bytes.equals(restauradas.get(n)),n+' mudou ao restaurar');
  }
  if(process.env.PGFLOW_XLSX_TEST_OUTPUT_DIR){
    const destino=path.resolve(process.env.PGFLOW_XLSX_TEST_OUTPUT_DIR);fs.mkdirSync(destino,{recursive:true});
    fs.writeFileSync(path.join(destino,'before.xlsx'),Buffer.from(inicial.getBytes()));
    fs.writeFileSync(path.join(destino,'edited.xlsx'),Buffer.from(atual.getBytes()));
    fs.writeFileSync(path.join(destino,'unprotected.xlsx'),Buffer.from(desprotegido.getBytes()));
    fs.writeFileSync(path.join(destino,'blocked.xlsx'),Buffer.from(bloqueado.getBytes()));
    fs.writeFileSync(path.join(destino,'metadata.json'),JSON.stringify({nome,aba,parte:geral,edicao,partes:antes.size,formulas:quantidadeFormulas,
      hashAntes:mx.sha256(Buffer.from(inicial.getBytes())),hashEditado:mx.sha256(Buffer.from(atual.getBytes())),hashDesprotegido:mx.sha256(Buffer.from(desprotegido.getBytes())),hashBloqueado:mx.sha256(Buffer.from(bloqueado.getBytes()))},null,2));
  }
});

test('valor XML inválido e membro de fórmula matricial são rejeitados antes de gerar substituição',()=>{
  const app=ambiente();
  for(const valor of ['a\u0000b','a\u0001b','\uD800','\uDC00','\uFFFE','\uFFFF'])
    assert.throws(()=>app._rxNovaCelula_('A1','<c r="A1"/>',valor),/caractere inválido/);
  assert.match(app._rxNovaCelula_('A1','<c r="A1"/>','Olá 😀\t\n'),/Olá 😀/);
  for(const tipo of ['array','shared','dataTable']){
    const xml='<worksheet><sheetData><row r="2"><c r="B2"><f t="'+tipo+'" ref="B2:C3">A1</f><v>2</v></c><c r="C2"><v>2</v></c></row><row r="3"><c r="C3"><v>2</v></c></row></sheetData></worksheet>';
    assert.throws(()=>app._rxEditarAba_(xml,[{linha:3,coluna:3,valor:'9'}]),/intervalo de fórmula/);
    assert.throws(()=>app._rxEditarAba_(xml,[{linha:2,coluna:2,valor:'9'}]),/fórmula/);
  }
  assert.equal(app._rxNovaCelula_('A1','<c r="A1" s="7" t="b"><v>1</v></c>','FALSO'),'<c r="A1" s="7" t="b"><v>0</v></c>');
  assert.equal(app._rxNovaCelula_('A1','<c r="A1" t="b"><v>0</v></c>','VERDADEIRO'),'<c r="A1" t="b"><v>1</v></c>');
  assert.throws(()=>app._rxNovaCelula_('A1','<c r="A1" t="b"><v>0</v></c>','talvez'),/VERDADEIRO ou FALSO/);
  const mesclada='<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"/></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>';
  assert.throws(()=>app._rxEditarAba_(mesclada,[{linha:1,coluna:2,valor:'9'}]),/primeira célula/);
  assert.match(app._rxEditarAba_(mesclada,[{linha:1,coluna:1,valor:'9'}]),/<c r="A1"><v>9<\/v><\/c>/);
});
