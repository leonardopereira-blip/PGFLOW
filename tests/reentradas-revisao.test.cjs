const test=require('node:test');
const assert=require('node:assert/strict');
const mx=require('../tools/molde-node.cjs');
const plain=v=>JSON.parse(JSON.stringify(v));
const WB='<workbook><sheets><sheet name="Dados" r:id="s1"/><sheet name="Custos" r:id="s2"/></sheets></workbook>';
const REL='<Relationships>'+[1,2].map(n=>'<Relationship Id="s'+n+'" Target="worksheets/sheet'+n+'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>').join('')+'</Relationships>';
const XML='<worksheet><dimension ref="A1:D6"/><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row>'+
 '<row r="5"><c r="A5" t="inlineStr"><is><t>SKU &amp; 001</t></is></c><c r="B5"><f>A5</f><v>7</v></c><c r="C5"><f t="array" ref="C5:D6">1+1</f><v>2</v></c></row>'+
 '<row r="6"><c r="B6"><f t="shared" si="2"/><v>8</v></c><c r="C6"><f>1+1</f><v>2</v></c></row></sheetData><mergeCells><mergeCell ref="A1:B1"/></mergeCells></worksheet>';
function ambiente(){
  let mtime=1800000000000,sourceReads=0,imports=0,snapshotWrites=0,folder=null;
  const files=new Map(),id='SOURCE_FILE_1234567890',nome='CALCULADORA FORNECEDOR META - CHAMADO 123.xlsx';
  const app=mx.carregarApps(['ValidaReentradasSPG.js','MacroProtecao.js','PreviaExcelEmail.js','ReentradasExcel.js','ReentradasRevisao.js']);
  const u=app.Utilities,blob=(data,type,name)=>{const b=u.newBlob(data,type,name);b.setName=n=>{name=n;return b};b.getName=()=>name;return b};
  const gzip=u.gzip;u.gzip=b=>{const z=gzip(b);return blob(z.getBytes(),'application/gzip','snapshot.gz')};
  const partes=[['xl/workbook.xml',WB],['xl/_rels/workbook.xml.rels',REL],['xl/sharedStrings.xml','<sst><si><t>Titulo</t></si></sst>'],['xl/worksheets/sheet1.xml',XML],['xl/worksheets/sheet2.xml',XML]].map(([n,t])=>blob(t,'application/xml',n));
  function snapshotFile(b){const file={id:'snapshot',blob:b,getId(){return this.id},getBlob(){return this.blob},setName(){return this}};files.set(b.getName(),file);return file}
  const technical={getFilesByName:n=>{const f=files.get(n);let taken=false;return {hasNext:()=>!!f&&!taken,next:()=>{taken=true;return f}}},createFile:b=>{snapshotWrites++;return snapshotFile(b)}};
  app.DriveApp={getFolderById:()=>({getFoldersByName:()=>{let taken=false;return{hasNext:()=>!!folder&&!taken,next:()=>{taken=true;return folder}}},createFolder:()=>folder=technical})};
  app.Drive={Files:{update(meta,fileId,b){snapshotWrites++;for(const f of files.values())if(f.id===fileId)f.blob=b;}}};
  const arq=()=>({id,nome,url:'https://drive.google.com/file/d/'+id,atualizadoEm:mtime,calculadora:true,
    arquivo:{getBlob(){sourceReads++;throw Error('nao pode abrir XLSX na navegacao');}}});
  app._rxArquivo_=(n,stamp)=>{assert.equal(n,id);if(stamp!=null&&Number(stamp)!==mtime)throw Error('O arquivo foi alterado no Drive.');return arq()};
  app.LockService={getScriptLock:()=>({tryLock:()=>true,releaseLock(){}})};
  return {app,partes,arq,id,files,reads:()=>sourceReads,writes:()=>snapshotWrites,
    time:v=>mtime=v,prepare(){const model=app._rxRevisaoMontar_(partes);app._rxRevisaoSalvar_(arq(),model);return model},
    allowImport(){app._rxArquivo_=(n,stamp)=>{assert.equal(n,id);if(stamp!=null&&Number(stamp)!==mtime)throw Error('O arquivo foi alterado no Drive.');const a=arq();a.arquivo.getBlob=()=>{sourceReads++;return {}};return a};app._reSpPartes_=()=>{imports++;return {partes,mapa:Object.fromEntries(partes.map(p=>[p.getName(),p]))}}},imports:()=>imports};
}
test('modelo nasce das partes finais em memoria e compacta protecao sem modificar o Excel',()=>{
  const s=ambiente(),m=s.app._rxRevisaoMontar_(s.partes,{'xl/worksheets/sheet1.xml':XML.replace('SKU &amp; 001','SKU NOVO')});
  assert.equal(s.reads(),0);assert.equal(m.abas[0].linhas[1].valores[0],'SKU NOVO');
  assert.equal(m.abas[1].linhas[1].valores[0],'SKU & 001');
  assert.equal(m.abas[0].linhas[0].valores[0],'Titulo');
  assert(m.faixasFormulas.some(f=>f.aba===0&&f.x1===2&&f.x2===3&&f.y1===5&&f.y2===6));
  assert(m.faixasFormulas.some(f=>f.aba===0&&f.x1===3&&f.x2===4&&f.y1===5&&f.y2===6));
  assert.deepEqual(plain(m.mesclagens[0]),{x1:1,x2:2,y1:1,y2:1,aba:0});
});
test('consulta e troca de aba leem apenas o JSON preparado; ausencia e versao externa nao abrem XLSX',()=>{
  const s=ambiente(),a=s.app;
  assert.equal(a.obterPreviaReentradasExcel({id:s.id,porAba:true}).codigo,'REVISAO_AUSENTE');assert.equal(s.reads(),0);
  s.prepare();const r=a.obterPreviaReentradasExcel({id:s.id,porAba:true});
  assert.equal(r.revisaoPreparada,true);assert.equal(r.abaCompleta.nome,'Dados');
  assert.equal(a.obterPreviaReentradasExcel({id:s.id,aba:1,porAba:true}).abaCompleta.nome,'Custos');
  a.PF_PREVIA_CONJUNTO_MAX_JSON=1;const p=a.obterPreviaReentradasExcel({id:s.id,porAba:true,busca:'SKU &'});
  assert.equal(p.abaCompleta,null);assert.equal(p.pagina.totalFiltradas,1);
  s.time(1800000001000);assert.equal(a.obterPreviaReentradasExcel({id:s.id}).codigo,'REVISAO_DESATUALIZADA');
  assert.equal(s.reads(),0);assert.equal(s.writes(),1);
});
test('preparo explicito importa uma vez; atualizar snapshot preserva o arquivo e substitui o JSON tecnico',()=>{
  const s=ambiente(),a=s.app;s.allowImport();
  const r=a.prepararRevisaoReentradasExcel(s.id,1800000000000);
  assert.equal(r.sucesso,true,r.erro);assert.equal(s.reads(),1);assert.equal(s.imports(),1);
  const before=s.files.size;s.time(1800000002000);s.prepare();
  assert.equal(s.files.size,before);assert.equal(s.writes(),2);
  assert.equal(a.obterPreviaReentradasExcel({id:s.id,porAba:true}).arquivo.atualizadoEm,1800000002000);
  assert.equal(s.reads(),1);
});
test('geracao prepara modelo antes de criar XLSX e nunca o reabre para criar revisao',()=>{
  const s=ambiente(),a=s.app,events=[];
  const model=a._rxRevisaoMontar_(s.partes);
  a._rxRevisaoMontar_=p=>{events.push('modelo');assert.equal(p[3].getDataAsString(),XML.replace('SKU &amp; 001','SKU NOVO'));return model};
  const file={setName(){return this},getBlob(){throw Error('nao pode reabrir XLSX')}};
  a.Utilities.zip=()=>({setContentType(){return this}});
  a.DriveApp.getFolderById=()=>({createFile(){events.push('xlsx');return file}});
  a._rxRevisaoSalvar_=(f,m)=>{events.push('revisao');assert.equal(f,file);assert.equal(m,model)};
  assert.equal(a._reSpSalvarZip_(null,{partes:s.partes},{'xl/worksheets/sheet1.xml':XML.replace('SKU &amp; 001','SKU NOVO')},'arquivo.xlsx'),file);
  assert.deepEqual(events,['modelo','xlsx','revisao']);
  const avisos=[];a._rxRevisaoMontar_=()=>model;a._rxRevisaoSalvar_=()=>{throw Error('falha tecnica simulada')};
  assert.equal(a._reSpSalvarZip_(null,{partes:s.partes},{},'arquivo.xlsx',avisos),file);
  assert.equal(avisos.length,1);assert.match(avisos[0],/arquivo.xlsx criado/);assert.match(avisos[0],/Preparar revisao/);
});
test('persistencia serializa criacao tecnica, libera lock em falha e nao readquire lock do salvamento',()=>{
  const s=ambiente(),a=s.app;const modelo=s.prepare();let aquisicoes=0,liberacoes=0;
  a.LockService={getScriptLock:()=>({tryLock(){aquisicoes++;return true},releaseLock(){liberacoes++}})};
  const atualizacao=a.Drive.Files.update;
  a.Drive.Files.update=(meta,id,blob,opt)=>{
    assert.equal(meta.mimeType,undefined,'atualizacao do snapshot nao pode tentar mudar o MIME no Drive');
    assert.equal(meta.title,blob.getName());
    return atualizacao(meta,id,blob,opt);
  };
  a._rxRevisaoSalvar_(s.arq(),modelo);
  assert.equal(aquisicoes,1);assert.equal(liberacoes,1);
  a._rxRevisaoSalvar_(s.arq(),modelo,true);
  assert.equal(aquisicoes,1,'save/import ja possuem o lock');assert.equal(liberacoes,1);
  a.Drive.Files.update=()=>{throw Error('Falha controlada ao gravar JSON')};
  assert.throws(()=>a._rxRevisaoSalvar_(s.arq(),modelo),/Falha controlada/);
  assert.equal(aquisicoes,2);assert.equal(liberacoes,2,'falha nao pode prender a fila de revisoes');
});

test('quadro de validação oculta margens e linhas só de fórmulas preservando coordenadas',()=>{
  const a=ambiente().app;
  const aba={nome:'CALC Geral',totalColunas:185,linhas:[
    {numero:1,valores:{2:'Papel',50:'Miolo',150:'Acabamento'}},
    {numero:2,valores:{2:'Dados gerais',80:'Custos'}},
    {numero:3,valores:{}},{numero:4,valores:{}},
    {numero:5,valores:{2:'SKU',3:'Custo unitário'}},
    {numero:6,valores:{2:'SKU-01',3:'83.85',180:'Aplicado'}},
    {numero:7,valores:{3:'=IF(C7="",0,D7*2)'}},
    {numero:999,valores:{3:'=IF(C999="",0,D999*2)'}}
  ]};
  const modelo={abas:[aba],formulas:{},faixasFormulas:[],mesclagens:[{aba:0,x1:3,x2:180,y1:1,y2:1}]};
  const q=a._rxRevisaoLayout_(aba,modelo,0);
  assert.equal(q.cabecalhoAte,5);assert.equal(q.ultimaLinha,6);
  assert.equal(q.indices[0],2);assert.equal(q.indices.at(-1),180);
  assert.equal(q.indices.includes(182),false,'GA/GB/GC vazias não devem aparecer');
  a.PF_PREVIA_CONJUNTO_MAX_JSON=1;
  const r=a._rxRevisaoResposta_({id:'VALIDACAO_TESTE',nome:'Calculadora',
    url:'https://drive.google.com/',atualizadoEm:1,calculadora:true},modelo,{aba:0,porAba:true});
  assert.equal(r.abaCompleta,null);
  assert.equal(r.pagina.linhas.length,1);
  assert.equal(r.pagina.linhas[0].numero,6);
  assert.equal(r.pagina.indices[0],2);
  assert.equal(r.pagina.colunas.at(-1),'FY');
  assert.equal(r.pagina.linhas[0].celulas[0],'SKU-01');
});

test('Especificacoes usa titulos da linha 3 e subtitulos da 4; colunas de total com formula ficam visiveis',()=>{
  const a=ambiente().app;
  const espec={nome:'PPG - Especificações',totalColunas:8,linhas:[
    {numero:1,valores:{2:'ESPECIFICACOES'}},{numero:2,valores:{2:'DADOS GERAIS'}},
    {numero:3,valores:{2:'CÓDIGO SKU',3:'CLASSIFICAÇÃO DO MATERIAL',4:'FORMATO'}},
    {numero:4,valores:{4:'mm'}},{numero:5,valores:{2:'SKU-01',3:'Não editorial',4:'300'}},
    {numero:200,valores:{4:'=A200'}}]};
  const geral={nome:'CALC Geral',totalColunas:185,linhas:[
    {numero:1,valores:{2:'CALCULADORA FORNECEDOR'}},{numero:2,valores:{2:'IDENTIFICACAO',5:'TIRAGEM'}},
    {numero:3,valores:{2:'CÓDIGO SKU',5:'AVULSO',9:'TOTAL'}},
    {numero:4,valores:{5:'Exemplares',9:'Exemplares'}},{numero:5,valores:{}},
    {numero:6,valores:{2:'SKU-01',5:'800',9:'=SUM(F6:I6)',182:'=XFD6'}}]};
  const modelo={abas:[espec,geral],formulas:{},faixasFormulas:[],mesclagens:[]};
  const q1=a._rxRevisaoLayout_(espec,modelo,0),q2=a._rxRevisaoLayout_(geral,modelo,1);
  assert.equal(q1.cabecalhoAte,4);assert.equal(q1.ultimaLinha,5);
  assert.equal(q1.topo.find(x=>x.numero===3).valores[2],'CÓDIGO SKU');
  assert.equal(q2.cabecalhoAte,5);assert.equal(q2.ultimaLinha,6);
  assert(q2.indices.includes(9),'TOTAL precisa permanecer, mesmo sem valor em cache');
  assert(!q2.indices.includes(182),'coluna de reserva sem titulo nao entra');
  const digital={nome:'CALC Digital',totalColunas:20,linhas:[
    {numero:3,valores:{2:'CÓDIGO SKU',3:'VALOR UNITÁRIO'}},
    {numero:4,valores:{3:'BRL / Unitário'}},{numero:5,valores:{3:'FORN-D-PU'}},
    {numero:6,valores:{2:'SKU-01',3:'1.25'}}]};
  assert.equal(a._rxRevisaoLayout_(digital,{abas:[digital],mesclagens:[]},0).cabecalhoAte,5);
});

test('Total da CALC Geral ignora cache antigo do molde e soma a tiragem atual',()=>{
  const s=ambiente(),a=s.app,geral={nome:'CALC Geral',totalColunas:11,linhas:[
    {numero:3,valores:{3:'CÓDIGO SKU',5:'AVULSO ACAB',6:'AVULSO GRAF',7:'KIT S/ MALETA',8:'KIT C/ MALETA',9:'TOTAL'}},
    {numero:4,valores:{9:'Exemplares'}},
    {numero:6,valores:{3:'SKU-01',5:'800',6:'0',7:'0',8:'0',9:'30'}},
    {numero:7,valores:{3:'SKU-02',5:'0',6:'0',7:'230',8:'0',9:'10'}}]};
  const modelo={schema:1,abas:[geral],formulas:{},faixasFormulas:[{aba:0,x1:10,x2:10,y1:6,y2:7}],mesclagens:[]};
  a._rxRevisaoSalvar_(s.arq(),modelo);
  const r=a.obterPreviaReentradasExcel({id:s.id,aba:0,porAba:true});
  assert.equal(r.abaCompleta.linhas[2].valores[9],'800');
  assert.equal(r.abaCompleta.linhas[3].valores[9],'230');
  const total=a.recalcularTotaisReentradasExcel(s.id,1800000000000);
  assert.equal(total.sucesso,true,total.erro);assert.equal(total.totaisRecalculados,2);
  assert.equal(s.reads(),0,'o recalculo nao pode abrir XLSX nem modificar formulas');
  assert.equal(s.writes(),1,'o botao de recalculo nao deve sobrescrever o XLSX');
});

test('separacao Não editorial fica na revisao preparada, sem tocar o XLSX',()=>{
  const s=ambiente(),a=s.app;
  const modelo={schema:1,abas:[{nome:'PPG - Especificações',totalColunas:4,linhas:[
    {numero:3,valores:{1:'CÓDIGO SKU',2:'CLASSIFICAÇÃO DO MATERIAL'}},
    {numero:4,valores:{}},{numero:5,valores:{1:'SKU-A',2:'Não editorial'}},
    {numero:6,valores:{1:'SKU-B',2:'Editorial'}}]}],
    formulas:{},faixasFormulas:[],mesclagens:[]};
  a._rxRevisaoSalvar_(s.arq(),modelo);
  assert.equal(a.separarNaoEditorialReentradas(s.id,1800000000000,0,6).erro.includes('Somente itens'),true);
  const r=a.separarNaoEditorialReentradas(s.id,1800000000000,0,5);
  assert.equal(r.sucesso,true,r.erro);
  assert.deepEqual(plain(r.naoEditoriais),[{aba:0,linha:5}]);
  assert.deepEqual(plain(a._rxRevisaoLer_(s.arq()).naoEditoriais),[{aba:0,linha:5}]);
  assert.equal(s.reads(),0);
});
