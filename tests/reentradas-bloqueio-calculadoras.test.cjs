const test=require('node:test');
const assert=require('node:assert/strict');
const crypto=require('node:crypto');
const fs=require('node:fs');
const path=require('node:path');
const mx=require('../tools/molde-node.cjs');
const FONTES=['MoldeXlsx.js','ValidaReentradasSPG.js','MacroProtecao.js','ProtecaoLogprint.js','ProtecaoCalculadoras.js'];
const at=(x,n)=>new RegExp('\\b'+n+'="([^"]+)"').exec(x||'')?.[1];
const iterator=itens=>{let i=0;return {hasNext:()=>i<itens.length,next:()=>itens[i++]};};
function carregar(extra={}){
  const app=mx.carregarApps(FONTES,extra),u=app.Utilities;
  u.zip=(partes,nome)=>u.newBlob(mx.escreverZip(partes.map(p=>({nome:p.getName(),dados:Buffer.from(p.getBytes())}))), 'application/zip',nome);
  u.unzip=blob=>mx.lerZip(Buffer.from(blob.getBytes())).map(p=>u.newBlob(p.dados,'application/xml',p.nome));
  u.DigestAlgorithm={SHA_256:'sha256'};u.computeDigest=(alg,bytes)=>crypto.createHash(alg).update(Buffer.from(bytes)).digest();
  return app;
}
function ambiente(){
  const arquivos={},chamadas=[],t0=1800000000000;
  const app=carregar({
    DriveApp:{getFolderById(id){assert.equal(id,'1DKie1AHE7EFHa6roJY2XxE2ZlFhC_lCy');return {getFiles:()=>iterator(Object.values(arquivos).map(a=>a.file))};},
      getFileById(id){if(!arquivos[id])throw Error('Arquivo indisponivel');return arquivos[id].file;}},
    LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock:()=>chamadas.push('release')})},
    Drive:{Files:{update(meta,id,blob,opcoes){
      chamadas.push({meta,id,blob,opcoes});const a=arquivos[id];a.nome=meta.title;a.tipo=meta.mimeType;a.blob=a.corromper?app.Utilities.newBlob('invalido'):blob;a.atualizadoEm++;
      return {id};
    }}}
  });
  function adicionar(id,opts={}){
    const a=Object.assign({id,nome:'CALCULADORA FORNECEDOR META - CHAMADO 2932.xlsx',tipo:app.BC_XLSX,
      pasta:app.BC_PASTA,excluido:false,atualizadoEm:t0,blob:app.Utilities.newBlob('original'),leituras:0},opts);
    a.file={getId:()=>a.id,getName:()=>a.nome,getMimeType:()=>a.tipo,isTrashed:()=>a.excluido,getLastUpdated:()=>new Date(a.atualizadoEm),
      getParents:()=>iterator([{getId:()=>a.pasta}]),getUrl:()=>`https://drive.google.com/file/d/${a.id}/view`,getBlob:()=>{a.leituras++;return a.blob;}};
    arquivos[id]=a;return a;
  }
  app._bcMontarXlsx_=(blob,grafica,nome)=>({blob:app.Utilities.newBlob('pronto',app.BC_XLSX,nome.replace(/\.xlsm$/i,'.xlsx')),nome:nome.replace(/\.xlsm$/i,'.xlsx'),abas:12});
  return {app,adicionar,chamadas,t0,arquivos};
}

test('lista somente Calculadora Excel na pasta; reconhece MIME XLSM lowercase e grafica do nome',()=>{
  const {app,adicionar}=ambiente();
  adicionar('file_valido_0001');
  adicionar('file_macro_0002',{nome:'Calculadora LOGPRINT.xlsm',tipo:'application/vnd.ms-excel.sheet.macroenabled.12'});
  adicionar('file_camel_0003',{nome:'Calculadora FORMA CERTA.xlsm',tipo:'application/vnd.ms-excel.sheet.macroEnabled.12'});
  adicionar('file_outra_0004',{nome:'Calculadora grafica desconhecida.xlsx'});
  adicionar('file_nome_0005',{nome:'De Para META.xlsx'});
  adicionar('file_tipo_0006',{tipo:'application/vnd.google-apps.spreadsheet'});
  adicionar('file_pasta_0007',{pasta:'outra'});
  adicionar('file_lixo_0008',{excluido:true});
  adicionar('file_ext_0009',{nome:'Calculadora META.xlsx',tipo:app.BC_XLSM});
  const res=app.listarCalculadorasParaBloqueio();assert.equal(res.erro,undefined);assert.equal(res.arquivos.length,4);
  assert.equal(res.arquivos.find(a=>a.id==='file_macro_0002').grafica,'LOGPRINT');
  assert.equal(res.arquivos.find(a=>a.id==='file_camel_0003').grafica,'FORMACERTA');
  assert.equal(res.arquivos.find(a=>a.id==='file_outra_0004').disponivel,false);
  for(const [nome,grafica] of [['Calculadora Casa da Gráfica.xlsx','CASADAGRAFICA'],['Calculadora Reproset.xlsx','REPROSET']])assert.equal(app._bcGrafica_(nome),grafica);
  assert.throws(()=>app._bcGrafica_('Calculadora META LOGPRINT.xlsx'),/unico modelo/);
});

test('bloqueio selecionado atualiza o mesmo ID com revisao, sem mudar pais/permissoes, e confere bytes',()=>{
  const {app,adicionar,chamadas,t0}=ambiente(),a=adicionar('file_valido_0001',{nome:'Calculadora META.xlsm',tipo:app.BC_XLSM});
  const res=app.bloquearCalculadoraSelecionada(a.id,t0);assert.equal(res.sucesso,true);assert.equal(res.id,a.id);assert.equal(res.nome,'Calculadora META.xlsx');
  const escritas=chamadas.filter(x=>typeof x==='object');assert.equal(escritas.length,1);
  assert.deepEqual(Object.keys(escritas[0].meta).sort(),['mimeType','title']);assert.equal(escritas[0].id,a.id);
  assert.equal(escritas[0].opcoes.newRevision,true);assert.equal(escritas[0].meta.mimeType,app.BC_XLSX);assert.equal(chamadas.at(-1),'release');
  assert.equal(a.pasta,app.BC_PASTA);assert.ok(a.leituras>=2,'nao releu os bytes apos salvar');
});

test('pasta, nome, MIME, lixeira, grafica ou mtime invalidos impedem leitura de blob e qualquer update',()=>{
  for(const opts of [{pasta:'outra'},{nome:'De Para META.xlsx'},{tipo:'application/octet-stream'},{excluido:true},
    {nome:'Calculadora desconhecida.xlsx'},{nome:'Calculadora META.xlsm',tipo:'application/vnd.ms-excel'},{atualizadoEm:1800000000001}]){
    const {app,adicionar,chamadas,t0}=ambiente(),a=adicionar('file_valido_0001',opts),res=app.bloquearCalculadoraSelecionada(a.id,t0);
    assert.ok(res.erro,JSON.stringify(opts));assert.equal(a.leituras,0);assert.equal(chamadas.filter(x=>typeof x==='object').length,0);
  }
});

test('alteracao concorrente durante build impede sobrescrita; divergencia apos update nao declara sucesso',()=>{
  const {app,adicionar,chamadas,t0}=ambiente(),a=adicionar('file_valido_0001'),montar=app._bcMontarXlsx_;
  app._bcMontarXlsx_=(...args)=>{const r=montar(...args);a.atualizadoEm++;return r;};
  const stale=app.bloquearCalculadoraSelecionada(a.id,t0);assert.match(stale.erro,/mudou durante/);assert.equal(chamadas.filter(x=>typeof x==='object').length,0);
  app._bcMontarXlsx_=montar;a.corromper=true;
  const corrompido=app.bloquearCalculadoraSelecionada(a.id,a.atualizadoEm);assert.equal(corrompido.sucesso,undefined);assert.match(corrompido.erro,/Drive recebeu.*verificacao final/);
});

function fixture(app){
  const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const estilos=locked=>`<styleSheet xmlns="${ns}"><cellStyleXfs count="1"><xf numFmtId="0" fontId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0"/><xf numFmtId="0" fontId="0" applyProtection="1"><protection locked="0"/></xf><xf numFmtId="4" fontId="1" applyProtection="1"><alignment horizontal="right"/><protection locked="${locked}"/></xf></cellXfs></styleSheet>`;
  const wb=`<workbook xmlns="${ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr codeName="ThisWorkbook"/><sheets><sheet name="CALC Geral" sheetId="1" r:id="rId1"/><sheet name="CALC Digital" sheetId="2" r:id="rId2"/><sheet name="Livre" sheetId="3" r:id="rId3"/></sheets></workbook>`;
  const rbase='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>';
  const proteger='<sheetProtection sheet="1" objects="1" scenarios="1" autoFilter="0"/>';
  const xml=(body,prot=proteger)=>`<worksheet xmlns="${ns}">${body}${prot}</worksheet>`;
  const geral='<cols><col min="4" max="4" width="15" customWidth="1" style="1"/><col min="84" max="84" style="1"/></cols><sheetData><row r="1" s="0" customFormat="1"><c r="D1" s="0" t="inlineStr"><is><t>SKU</t></is></c></row><row r="6"><c r="D6" s="1"><v>200713648</v></c><c r="E6" s="2"><f>SUM(1,2)</f><v>3</v></c></row></sheetData>';
  const digital='<cols><col min="4" max="4" style="2" width="14"/></cols><sheetData><row r="8"><c r="D8" s="2"><f>5+6</f><v>11</v></c></row></sheetData>';
  const livre='<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>livre</t></is></c></row></sheetData>';
  const tipos='<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.main+xml"/></Types>';
  const base=[{nome:'xl/workbook.xml',texto:wb},{nome:'xl/_rels/workbook.xml.rels',texto:rbase+'</Relationships>'},
    {nome:'xl/styles.xml',texto:estilos(1)},{nome:'[Content_Types].xml',texto:tipos},
    {nome:'xl/worksheets/sheet1.xml',texto:xml(geral)},{nome:'xl/worksheets/sheet2.xml',texto:xml(digital)},
    {nome:'xl/worksheets/sheet3.xml',texto:xml(livre,'')},{nome:'docProps/core.xml',texto:'<properties>inalterado</properties>'},
    {nome:'xl/printerSettings/printerSettings1.bin',base64:Buffer.from([0,255,128,42]).toString('base64')}];
  const ref=app._mxMapear_(base,{}),entrada=base.map(p=>({...p}));
  entrada.find(p=>p.nome==='xl/styles.xml').texto=estilos(0); // Usuario liberou uma formula na aba de calculo.
  entrada.find(p=>p.nome==='xl/worksheets/sheet1.xml').texto=xml(geral.replace('</row></sheetData>','<c r="CF6" t="inlineStr"><is><t>GRAFICA</t></is></c></row></sheetData>'),'');
  entrada.find(p=>p.nome==='xl/worksheets/sheet2.xml').texto=xml(digital,'');
  entrada.find(p=>p.nome==='xl/worksheets/sheet3.xml').texto=xml(livre,proteger);
  entrada.find(p=>p.nome==='[Content_Types].xml').texto=tipos.replace('application/vnd.ms-excel.sheet.main+xml','application/vnd.ms-excel.sheet.macroEnabled.main+xml').replace('</Types>',
    '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/customXml/pgflowProtectionProps.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>');
  entrada.find(p=>p.nome==='xl/_rels/workbook.xml.rels').texto=rbase+'<Relationship Id="vba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/><Relationship Id="map" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/pgflowProtection.xml"/></Relationships>';
  for(const n of ['xl/vbaProject.bin','xl/vbaProjectSignature.bin','xl/_rels/vbaProject.bin.rels','customXml/pgflowProtection.xml','customXml/pgflowProtectionProps.xml','customXml/_rels/pgflowProtection.xml.rels'])entrada.push({nome:n,texto:n.endsWith('.bin')?'binario-vba':'<antigo/>'});
  const blob=app.Utilities.zip(entrada.map(p=>app.Utilities.newBlob(p.base64?Buffer.from(p.base64,'base64'):p.texto,'application/xml',p.nome)),'Calculadora META.xlsm');
  return {ref,entrada,blob};
}

test('engine restaura CF6 e formula de Digital, conserva valores/formatos/binarios e remove VBA/metadata em XLSX',()=>{
  const app=carregar(),f=fixture(app);app._bcPerfil_=()=>({molde:f.ref,protecao:f.ref});
  const pronto=app._bcMontarXlsx_(f.blob,'META','Calculadora META.xlsm'),saida=app._reSpPartes_(pronto.blob);
  assert.equal(pronto.nome,'Calculadora META.xlsx');assert.equal(pronto.blob.getContentType(),app.BC_XLSX);
  assert.equal(Object.keys(saida.mapa).some(n=>/vbaProject|pgflowProtection/.test(n)),false);
  const get=n=>saida.mapa[n].getDataAsString();
  assert.match(get('[Content_Types].xml'),/application\/vnd\.openxmlformats-officedocument.spreadsheetml.sheet.main\+xml/);
  assert.doesNotMatch(get('xl/_rels/workbook.xml.rels'),/vbaProject|pgflowProtection/);assert.doesNotMatch(get('xl/workbook.xml'),/codeName/);
  assert.equal(get('docProps/core.xml'),'<properties>inalterado</properties>');
  assert.deepEqual(Buffer.from(saida.mapa['xl/printerSettings/printerSettings1.bin'].getBytes()),Buffer.from([0,255,128,42]));
  const est=app._reMpEstilos_(get('xl/styles.xml')),tags=Array.from(app._reMpXfs_(get('xl/styles.xml'),'cellXfs'));
  const cel=(n,ref)=>new RegExp('<c\\b[^>]*r="'+ref+'"[^>]*>[\\s\\S]*?<\\/c>').exec(get(n))[0];
  const cf=cel('xl/worksheets/sheet1.xml','CF6'),d8=cel('xl/worksheets/sheet2.xml','D8');
  assert.ok(at(cf,'s'),'CF6 sem s voltaria bloqueada no Excel');assert.equal(est[+at(cf,'s')].locked,0);
  assert.equal(est[+at(d8,'s')].locked,1,'nao restaurou protecao na aba de calculo sem entradas');
  assert.equal(at(tags[+at(d8,'s')],'numFmtId'),'4');assert.equal(at(tags[+at(d8,'s')],'fontId'),'1');
  assert.equal(est[2].locked,0,'alterou estilo compartilhado em vez de clonar');
  assert.match(d8,/<f>5\+6<\/f><v>11<\/v>/);assert.match(cel('xl/worksheets/sheet1.xml','D6'),/<v>200713648<\/v>/);
  assert.doesNotMatch(get('xl/worksheets/sheet3.xml'),/sheetProtection/,'ativou protecao na aba livre do perfil');
  for(const n of ['xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml'])assert.match(get(n),new RegExp('password="'+app._reSpSenhaHash_('PGFLOW2026')+'"'));
});

test('layout alterado e XF0 modificado falham antes de gerar uma substituicao',()=>{
  const app=carregar(),f=fixture(app);app._bcPerfil_=()=>({molde:f.ref,protecao:f.ref});
  const partes=app._reSpPartes_(f.blob),aba='xl/worksheets/sheet1.xml';
  partes.mapa[aba]=app.Utilities.newBlob(partes.mapa[aba].getDataAsString().replace('</worksheet>','<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>'),'application/xml',aba);
  assert.throws(()=>app._bcValidarLayout_(partes,f.ref),/layout.*mudou/);
  const estilos=partes.mapa['xl/styles.xml'].getDataAsString().replace('<xf numFmtId="0" fontId="0"/>','<xf numFmtId="0" fontId="0" applyProtection="1"><protection locked="0"/></xf>');
  const semProtecao=estilos.replace('<cellXfs count="3"><xf numFmtId="0" fontId="0"/>','<cellXfs count="3"><xf numFmtId="0" fontId="0" applyProtection="1"><protection locked="0"/></xf>');
  partes.mapa['xl/styles.xml']=app.Utilities.newBlob(semProtecao,'application/xml','xl/styles.xml');
  assert.throws(()=>app._reSpAplicarProtecaoLogprint_(partes,{},f.ref,true),/estilo padrao de protecao.*alterado/);
});

test('perfil META real conserva CALC Digital e TAB Papel livres; modal usa apenas os endpoints de selecao',()=>{
  const app=mx.carregarApps([...FONTES,'MoldeMetaCalc.js']),ref=app._moldeMetaCalc_(),estilos=ref.textos['xl/styles.xml'];
  const aba='<worksheet><sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData><sheetProtection sheet="1" password="1234"/></worksheet>';
  const wb='<workbook><sheets><sheet name="CALC Digital" r:id="r1"/><sheet name="TAB Papel" r:id="r2"/></sheets></workbook>',rels='<Relationships><Relationship Id="r1" Target="worksheets/sheet1.xml"/><Relationship Id="r2" Target="worksheets/sheet2.xml"/></Relationships>';
  const mapa=Object.fromEntries([['xl/workbook.xml',wb],['xl/_rels/workbook.xml.rels',rels],['xl/styles.xml',estilos],['xl/worksheets/sheet1.xml',aba],['xl/worksheets/sheet2.xml',aba]].map(([n,x])=>[n,app.Utilities.newBlob(x,'application/xml',n)]));
  const alter={};app._reSpAplicarProtecaoLogprint_({mapa},alter,ref,true);
  for(const n of ['xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml']){
    const tag=/<sheetProtection\b[^>]*>/.exec(alter[n])?.[0];assert.notEqual(attrSheet(tag),1);assert.match(alter[n],/<v>42<\/v>/);
  }
  const html=fs.readFileSync(path.join(mx.RAIZ,'BloquearCalculadoras.html'),'utf8');
  assert.match(html,/listarCalculadorasParaBloqueio\(\)/);assert.match(html,/bloquearCalculadoraSelecionada\(arquivo.id,arquivo.atualizadoEm\)/);
  assert.match(html,/textContent=a.nome/);assert.doesNotMatch(html,/innerHTML\s*=/);
  let exibido=false;app.HtmlService={createHtmlOutputFromFile:n=>({setWidth(){assert.equal(n,'BloquearCalculadoras');return this;},setHeight(){return this;}})};
  app.SpreadsheetApp={getUi:()=>({showModalDialog:()=>{exibido=true;}})};app.abrirBloqueioCalculadoras();assert.equal(exibido,true);
});
function attrSheet(tag){return /^(?:1|true)$/.test(at(tag,'sheet')||'0')?1:0;}
