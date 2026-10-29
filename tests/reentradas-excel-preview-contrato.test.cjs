const test=require('node:test');
const assert=require('node:assert/strict');
const mx=require('../tools/molde-node.cjs');

// Mesmo adaptador simples de XmlService utilizado pela prova da previa Email.
function parseXml(xml){
  const stack=[],roots=[];
  for(const m of xml.matchAll(/<[^>]+>|[^<]+/g)){
    const p=m[0];if(p.startsWith('<?')||p.startsWith('<!'))continue;
    if(p.startsWith('</')){stack.pop();continue;}
    if(p.startsWith('<')){
      const name=/^<([^\s/>]+)/.exec(p)[1],attrs={};
      for(const a of p.matchAll(/([^\s=<>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g))attrs[a[1]]=a[2]??a[3];
      const e={name,attrs,children:[],texts:[],getName(){return name.split(':').pop();},
        getText(){return this.texts.join('');},getChildren(n){return n===undefined?this.children:this.children.filter(c=>c.getName()===n);},
        getChild(n){return this.getChildren(n)[0]||null;},getAttribute(n){const k=Object.keys(attrs).find(k=>k===n||k.split(':').pop()===n);return k===undefined?null:{getValue:()=>attrs[k]};},getNamespace(){return {};}};
      (stack.at(-1)?.children||roots).push(e);if(!p.endsWith('/>'))stack.push(e);
    }else if(stack.length)stack.at(-1).texts.push(p);
  }
  return roots[0];
}
function ambiente(){
  const app=mx.carregarApps(['ValidaReentradasSPG.js','MacroProtecao.js','PreviaExcelEmail.js','ReentradasExcel.js'],{
    _emailSpgXlsxRaiz_:parseXml,_emailSpgXlsxTexto_:e=>e.getChild('t')?.getText()||''
  });
  const xml='<worksheet><dimension ref="A1:CF8"/><sheetData>'+
    '<row r="1"><c r="A1" t="inlineStr"><is><t>SKU</t></is></c><c r="C1"><f>1+1</f><v>2</v></c></row>'+
    '<row r="2"><c r="A2" t="s"><v>0</v></c><c r="D2" t="b"><v>1</v></c></row>'+
    '<row r="5"><c r="A5" t="inlineStr"><is><t>achado</t></is></c><c r="B5"><f t="array" ref="B5:C6">1+1</f><v>2</v></c><c r="C5"><v>2</v></c></row>'+
    '<row r="6"><c r="A6" t="inlineStr"><is><t>achado</t></is></c><c r="C6"><v>2</v></c></row>'+
    '<row r="8"><c r="A8"><f>1+2</f></c><c r="CF8"><v>8</v></c></row></sheetData></worksheet>';
  const livro={nomes:['Geral','Outra'],caminhos:['xl/sheet1.xml','xl/sheet2.xml'],arquivos:{'xl/sheet1.xml':xml,'xl/sheet2.xml':'<worksheet><sheetData><row r="1"><c r="A1"><v>3</v></c></row></sheetData></worksheet>'},compartilhadas:['Texto compartilhado']};
  let leituras=0;app._pfPreviaLivroXlsx_=()=>{leituras++;return livro;};
  const blob={getName:()=> 'Calculadora.xlsx'};
  return {app,livro,blob,leituras:()=>leituras};
}
const plain=x=>JSON.parse(JSON.stringify(x));

test('prévia RX retorna mesmos dados que Email para conjunto, página, busca e limites',()=>{
  const {app,livro,blob}=ambiente();
  assert.deepEqual(plain(app._rxPreviaConjunto_(livro,blob.getName(),{})),plain(app._pfPreviaConjuntoXlsx_(blob)));
  for(const pedido of [{},{aba:1},{busca:'achado'},{linha:1,coluna:2},{linha:999,coluna:10000},{busca:'inexistente'},{coluna:81}])
    assert.deepEqual(plain(app._rxPreviaPagina_(livro,blob.getName(),pedido,{})),plain(app._pfPreviaXlsx_(blob,pedido)),JSON.stringify(pedido));
  for(const [limite,valor] of [['PF_PREVIA_CONJUNTO_MAX_XML',5],['PF_PREVIA_CONJUNTO_MAX_CELULAS',2],['PF_PREVIA_CONJUNTO_MAX_JSON',10]]){
    const original=app[limite];app[limite]=valor;
    assert.equal(app._rxPreviaConjunto_(livro,blob.getName(),{}),null);assert.equal(app._pfPreviaConjuntoXlsx_(blob),null);app[limite]=original;
  }
});
test('fórmula cacheada tem metadado independente de seu valor e ranges não são expandidos',()=>{
  const {app,livro}=ambiente(),formulas={};
  const conjunto=app._rxPreviaConjunto_(livro,'Calculadora.xlsx',formulas);
  assert.equal(conjunto.abas[0].linhas[0].valores[2],'2');assert.equal(formulas['0:1:3'],true);
  assert.equal(formulas['0:5:2'],true);assert.equal(formulas['0:8:1'],true);assert.equal(formulas['0:5:3'],undefined);
  assert.deepEqual(plain(app._rxFaixasFormulas_(livro.arquivos['xl/sheet1.xml'])),[{x1:2,x2:3,y1:5,y2:6}]);
  const paginaFormulas={};app._rxPreviaPagina_(livro,'Calculadora.xlsx',{busca:'achado',coluna:0},paginaFormulas);
  assert.deepEqual(Object.keys(paginaFormulas),['0:5:2']);
});
test('RPC RX lê livro uma vez e devolve metadados isolados de fórmulas',()=>{
  const {app,blob,leituras}=ambiente();app._rxArquivo_=()=>({id:'teste',nome:'Calculadora.xlsx',url:'url',atualizadoEm:1,calculadora:true,arquivo:{getBlob:()=>({...blob,setName(){return this;}})}});
  const r=app.obterPreviaReentradasExcel({id:'teste'});
  assert.equal(leituras(),1);assert.equal(r.formulas['0:1:3'],true);assert.equal(r.faixasFormulas[0].aba,0);
  const p=app.obterPreviaReentradasExcel({id:'teste',apenasPagina:true,busca:'achado'});
  assert.equal(leituras(),2);assert.equal(p.previaExcel,null);assert.equal(p.pagina.totalFiltradas,2);assert.equal(p.formulas['0:5:2'],true);
});
test('marca recálculo após edição sem perder parâmetros de iteração do arquivo',()=>{
  const {app}=ambiente();
  const xml='<workbook><calcPr calcId="123" calcMode="manual" fullCalcOnLoad="0" forceFullCalc="0" iterate="1" iterateCount="10"/><extLst/></workbook>';
  const marcado=app._rxMarcarRecalculo_(xml);
  assert.match(marcado,/calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/);
  assert.match(marcado,/calcId="123"/);assert.match(marcado,/iterate="1" iterateCount="10"/);
  assert.equal(app._rxMarcarRecalculo_(marcado),marcado);
  assert.match(app._rxMarcarRecalculo_('<workbook><extLst/></workbook>'),/<calcPr[^>]*\/><extLst/);
});
test('índice de salvamento acompanha somente worksheets da prévia e ignora chartsheets/external',()=>{
  const {app}=ambiente();
  const workbook='<workbook><sheets><sheet name="Gráfico" r:id="grafico"/><sheet name="Primeira" r:id="p1"/><sheet name="Outro gráfico" r:id="grafico2"/><sheet name="Segunda" r:id="p2"/><sheet name="Externa" r:id="externa"/></sheets></workbook>';
  const tipo='http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  const rels='<Relationships><Relationship Id="grafico" Target="chartsheets/sheet1.xml" Type="'+tipo+'chartsheet"/>'+ 
    '<Relationship Id="p1" Target="worksheets/sheet1.xml" Type="'+tipo+'worksheet"/>'+ 
    '<Relationship Id="grafico2" Target="chartsheets/sheet2.xml" Type="'+tipo+'chartsheet"/>'+ 
    '<Relationship Id="p2" Target="worksheets/sheet2.xml" Type="'+tipo+'worksheet"/>'+ 
    '<Relationship Id="externa" Target="https://example.test/sheet" Type="'+tipo+'worksheet" TargetMode="External"/></Relationships>';
  const abas=app._rxAbasEditaveis_(workbook,rels);
  assert.deepEqual(plain(abas),[{nome:'Primeira',caminho:'xl/worksheets/sheet1.xml'},{nome:'Segunda',caminho:'xl/worksheets/sheet2.xml'}]);
  assert.equal(abas[1].caminho,'xl/worksheets/sheet2.xml','índice 1 da prévia selecionou outra aba');
  const u=app.Utilities;
  u.zip=(pp,nome)=>u.newBlob(mx.escreverZip(pp.map(p=>({nome:p.getName(),dados:Buffer.from(p.getBytes())}))),app.RX_MIME,nome);
  u.unzip=b=>mx.lerZip(Buffer.from(b.getBytes())).map(p=>u.newBlob(p.dados,'application/xml',p.nome));
  const sheet=n=>'<worksheet><sheetData><row r="1"><c r="A1"><v>'+n+'</v></c></row></sheetData></worksheet>';
  const nome='CALCULADORA FORNECEDOR META - TESTE.xlsx',id='teste_chartsheet_20261029';
  let blob=u.zip([['xl/workbook.xml',workbook],['xl/_rels/workbook.xml.rels',rels],['xl/worksheets/sheet1.xml',sheet(10)],['xl/worksheets/sheet2.xml',sheet(20)],['xl/chartsheets/sheet1.xml','<chartsheet/>']].map(([n,s])=>u.newBlob(s,'application/xml',n)),nome);
  let mtime=1800000000000,updates=0;
  app._rxArquivo_=()=>({arquivo:{getBlob:()=>blob},id,nome,url:'url',atualizadoEm:mtime});
  app._bcDigest_=b=>Buffer.from(b.getBytes()).toString('base64');
  app.LockService={getScriptLock:()=>({tryLock:()=>true,releaseLock(){}})};
  app.Drive={Files:{update(meta,fileId,novo){assert.equal(fileId,id);blob=novo;mtime++;updates++;}}};
  const r=app.salvarEdicoesReentradasExcel(id,mtime,[{aba:1,linha:1,coluna:1,valor:'99'}]);
  assert.equal(r.sucesso,true,r.erro);assert.equal(updates,1);
  const pp=new Map(mx.lerZip(Buffer.from(blob.getBytes())).map(p=>[p.nome,p.dados.toString('utf8')]));
  assert.equal(pp.get('xl/worksheets/sheet1.xml'),sheet(10));
  assert.equal(pp.get('xl/worksheets/sheet2.xml'),sheet(99));
  assert.equal(pp.get('xl/chartsheets/sheet1.xml'),'<chartsheet/>');
});
