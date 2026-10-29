// Leitura sintética de XLSX e Google Sheets. Nenhuma API, Drive ou e-mail real.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function decodeXml(text) {
  return text.replace(/&(?:amp|quot|apos|lt|gt|#\d+|#x[0-9a-f]+);/gi, entity => {
    const names = {'&amp;':'&','&quot;':'"','&apos;':"'",'&lt;':'<','&gt;':'>'};
    if (names[entity]) return names[entity];
    return String.fromCodePoint(entity.startsWith('&#x') ? parseInt(entity.slice(3),16) : parseInt(entity.slice(2),10));
  });
}
function parseXml(xml) {
  const stack = [], roots = [];
  for (const match of xml.matchAll(/<[^>]+>|[^<]+/g)) {
    const part = match[0];
    if (part.startsWith('<?') || part.startsWith('<!')) continue;
    if (part.startsWith('</')) { stack.pop(); continue; }
    if (part.startsWith('<')) {
      const name = /^<([^\s/>]+)/.exec(part)[1];
      const attrs = {};
      for (const attr of part.matchAll(/([^\s=<>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) attrs[attr[1]] = decodeXml(attr[2] ?? attr[3]);
      const element = {
        name, attrs, children: [], texts: [],
        getName() { return this.name.split(':').pop(); },
        getText() { return this.texts.join(''); },
        getValue() { return this.texts.join('') + this.children.map(child => child.getValue()).join(''); },
        getChildren(childName) { return childName === undefined ? this.children : this.children.filter(child => child.getName() === childName); },
        getChild(childName) { return this.getChildren(childName)[0] || null; },
        getChildText(childName) { return this.getChild(childName)?.getText() ?? null; },
        getAttribute(attrName) { const found = Object.keys(this.attrs).find(key => key === attrName || key.split(':').pop() === attrName); return found === undefined ? null : {getValue: () => this.attrs[found]}; },
        getAttributes() { return Object.keys(this.attrs).map(key => ({getName:()=>key,getValue:()=>this.attrs[key]})); },
        getNamespace() { return {}; },
        getDescendants() { return this.children.flatMap(child => [{asElement:()=>child,getType:()=> 'ELEMENT'}, ...child.getDescendants()]); },
      };
      (stack.at(-1)?.children || roots).push(element);
      if (!part.endsWith('/>')) stack.push(element);
    } else if (stack.length) stack.at(-1).texts.push(decodeXml(part));
  }
  return {getRootElement: () => roots[0]};
}



const UUID = '75a9b7b1-0b22-4b37-8b78-9323c1024005';
const workbook = '<workbook><sheets><sheet name="Tiragem (legado)" r:id="rId1"/><sheet name="Árvore (legado)" r:id="rId2"/></sheets></workbook>';
const rels = '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/><Relationship Id="rId2" Target="worksheets/sheet2.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/></Relationships>';
const sheet1 = '<worksheet><dimension ref="A1:AB4"/><sheetData>' +
  '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="AB1" t="inlineStr"><is><t>Última coluna</t></is></c></row>' +
  '<row r="2"><c r="A2" t="inlineStr"><is><t>SKU-01</t></is></c><c r="B2"><v>25159</v></c></row>' +
  '<row r="4"><c r="A4" t="inlineStr"><is><t>SKU-02</t></is></c><c r="B4" t="b"><v>1</v></c></row>' +
  '</sheetData></worksheet>';
const sheet2 = '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>KIT-01</t></is></c></row></sheetData></worksheet>';
const zip = [
  ['xl/workbook.xml',workbook],
  ['xl/_rels/workbook.xml.rels',rels],
  ['xl/sharedStrings.xml','<sst><si><t>Código do SKU</t></si></sst>'],
  ['xl/worksheets/sheet1.xml',sheet1],
  ['xl/worksheets/sheet2.xml',sheet2]
].map(([name,text])=>({getName:()=>name,getDataAsString:()=>text}));
function context() {
  const calls={spg:0,email:0,drive:0,unzip:0};
  const sheetRows = [['SKU','Tiragem'],['SKU-11','500'],['SKU-22','250']];
  const sheet = {getName:()=> 'PPG',getLastRow:()=>3,getLastColumn:()=>2,
    getRange:(r,c,n,m)=>({getDisplayValues:()=>sheetRows.slice(r-1,r-1+n).map(row=>row.slice(c-1,c-1+m))})};
  const context={
    Utilities:{unzip:()=>{calls.unzip++;return zip;},base64Encode:bytes=>Buffer.from(bytes.map(b=>(b+256)%256)).toString('base64')},
    XmlService:{parse:parseXml,getNamespace:()=>({})},
    DriveApp:{getFileById:id=>{calls.drive++;return {getMimeType:()=> 'application/vnd.google-apps.spreadsheet'};}},
    SpreadsheetApp:{openById:()=>({getName:()=> 'Origem PPG',getSheets:()=>[sheet]})},
    _emailSpgId_:id=>String(id),
    _emailSpgPacote_:()=>{calls.spg++;return {hash:'a'.repeat(64),blob:{getName:()=> 'Chamado.xlsx',getBytes:()=>[80,75,0,0],copyBlob(){return this},setContentType(){return this}}};},
    _emailSpgXlsxRaiz_:xml=>parseXml(xml).getRootElement(),
    _emailSpgXlsxTexto_:elem=>elem.getChild('t')?.getText()||'',
    _emailSpgXlsxCaminho_:alvo=>'xl/'+alvo,
    GmailApp:{sendEmail(){calls.email++;}}
  };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(path.join(__dirname,'..','PreviaExcelEmail.js'),'utf8'),context);
  return {context,calls};
}
test('prévia lê abas reais, colunas além de Z e números das linhas do Excel',()=>{
  const {context:c}=context();
  const r=c.obterPreviaExcelEmail({origem:'spg',idChamado:'3100',anexoHash:'a'.repeat(64),aba:0});
  assert.deepEqual(Array.from(r.abas),['Tiragem (legado)','Árvore (legado)']);
  assert.equal(r.totalLinhas,3);
  assert.equal(r.totalColunas,28);
  assert.equal(r.colunas.at(-1),'AB');
  assert.equal(r.linhas[0].celulas[0],'Código do SKU');
  assert.equal(r.linhas[0].celulas[27],'Última coluna');
  assert.equal(r.linhas[1].numero,2);
  assert.equal(r.linhas[1].celulas[1],'25159');
  assert.equal(r.linhas[2].numero,4);
  assert.equal(r.linhas[2].celulas[1],'VERDADEIRO');
});
test('troca abas, busca todas as linhas e não altera o arquivo',()=>{
  const {context:c,calls}=context();
  const a=c.obterPreviaExcelEmail({origem:'spg',idChamado:'3100',anexoHash:'a'.repeat(64),aba:1});
  assert.equal(a.linhas[0].celulas[0],'KIT-01');
  const b=c.obterPreviaExcelEmail({origem:'spg',idChamado:'3100',anexoHash:'a'.repeat(64),busca:'sku-02'});
  assert.equal(b.totalFiltradas,1);
  assert.equal(b.linhas[0].numero,4);
  assert.equal(calls.email,0);
  assert.equal(calls.drive,0);
});
test('bloqueia prévia quando hash do Excel do SPG não coincide',()=>{
  const {context:c}=context();
  assert.throws(()=>c.obterPreviaExcelEmail({origem:'spg',idChamado:'3100',anexoHash:'b'.repeat(64)}),/mudaram/i);
});
test('fluxo antigo lê células da planilha do Drive sem enviar email',()=>{
  const {context:c,calls}=context();
  const r=c.obterPreviaExcelEmail({origem:'antigo',idArquivo:'abcDEF1234567890',aba:0,busca:'SKU-22'});
  assert.equal(r.totalLinhas,3);
  assert.equal(r.totalFiltradas,1);
  assert.equal(r.linhas[0].celulas[1],'250');
  assert.equal(calls.drive,1);
  assert.equal(calls.email,0);
});
test('valida entradas antes de ler anexo de origem desconhecida',()=>{
  const {context:c}=context();
  assert.throws(()=>c.obterPreviaExcelEmail({origem:'qualquer'}),/Origem/);
  assert.throws(()=>c.obterPreviaExcelEmail({origem:'antigo',idArquivo:'../../'}),/ID do anexo/);
});

test('consulta inicial inclui o conjunto das abas usando o mesmo blob já validado',()=>{
  const {context:c,calls}=context();
  vm.runInContext(fs.readFileSync(path.join(__dirname,'..','EmailSPG.js'),'utf8'),c);
  let recebidos=0;
  c._emailSpgPacote_=()=>{
    recebidos++;
    return {id:'3100',hash:'a'.repeat(64),dataDesejada:'15/11/2026',descricao:'Produção',
      blob:{getBytes:()=>[80,75,0,0],getName:()=> 'Chamado 3100.xlsx',copyBlob(){return this},setContentType(){return this}}};
  };
  const r=c.consultarPacoteEmailSPG('3100','ultima','reentrada');
  assert.equal(r.anexoValido,true);
  assert.equal(recebidos,1,'a prévia não consulta novamente o pacote/API');
  assert.equal(calls.unzip,1,'extrai o conjunto uma vez a partir do blob existente');
  assert.equal(r.previaExcel.arquivo,'Chamado 3100.xlsx');
  assert.equal(r.previaExcel.abas.length,2);
  assert.equal(r.previaExcel.abas[0].linhas[2].numero,4);
  assert.equal(r.previaExcel.abas[0].linhas[2].valores[0],'SKU-02');
  assert.equal(r.previaExcel.abas[0].totalColunas,28);
  assert.equal(calls.email,0);
  assert.equal(calls.drive,0);
});

test('conjunto completo preserva texto e dados fora da primeira página sem truncar busca',()=>{
  const {context:c}=context();
  const extra = '<worksheet><sheetData>' + Array.from({length:80},(_,i)=>
    '<row r="'+(i+1)+'"><c r="A'+(i+1)+'" t="inlineStr"><is><t>'+('SKU-'+i)+'</t></is></c></row>').join('') + '</sheetData></worksheet>';
  c.Utilities.unzip=()=>zip.map(p=>p.getName()==='xl/worksheets/sheet1.xml' ? {getName:p.getName,getDataAsString:()=>extra} : p);
  const conjunto=c.obterConjuntoPreviaExcelEmail({origem:'spg',idChamado:'3100',anexoHash:'a'.repeat(64)}).previaExcel;
  assert.equal(conjunto.abas[0].linhas.length,80);
  assert.equal(conjunto.abas[0].linhas[79].valores[0],'SKU-79');
});

test('conjunto grande retorna página inicial do mesmo blob sem repetir a API',()=>{
  const {context:c,calls}=context();
  c.PF_PREVIA_CONJUNTO_MAX_JSON=10;
  const r=c.obterConjuntoPreviaExcelEmail({origem:'spg',idChamado:'3100',anexoHash:'a'.repeat(64),busca:'SKU-02'});
  assert.equal(r.previaExcel,null);
  assert.equal(r.pagina.totalFiltradas,1);
  assert.equal(r.pagina.linhas[0].numero,4);
  assert.equal(calls.spg,1);
  assert.equal(calls.email,0);
});

test('consulta inicial não impede revisão se o conjunto ultrapassar o limite',()=>{
  const {context:c,calls}=context();
  vm.runInContext(fs.readFileSync(path.join(__dirname,'..','EmailSPG.js'),'utf8'),c);
  let consultas=0;
  c._emailSpgPacote_=()=>{
    consultas++;
    return {id:'3100',hash:'a'.repeat(64),dataDesejada:'15/11/2026',descricao:'Produção',blob:{getBytes:()=>[80,75,0,0],getName:()=> 'Chamado.xlsx',copyBlob(){return this},setContentType(){return this}}};
  };
  c.PF_PREVIA_CONJUNTO_MAX_JSON=10;
  const r=c.consultarPacoteEmailSPG('3100','ultima','reentrada');
  assert.equal(r.anexoValido,true);
  assert.equal(r.previaExcel,null);
  assert.equal(r.paginaPreviaExcel.aba,0);
  assert.equal(r.paginaPreviaExcel.linha,0);
  assert.equal(r.paginaPreviaExcel.coluna,0);
  assert.equal(r.paginaPreviaExcel.linhas[1].celulas[0],'SKU-01');
  assert.equal(consultas,1,'o fallback da primeira abertura reutiliza o pacote já consultado');
  assert.equal(calls.email,0);
  assert.equal(calls.drive,0);
});

test('download retorna bytes XLSX do pacote confirmado e não envia email',()=>{
  const {context:c,calls}=context();
  const r=c.baixarPreviaExcelEmail({origem:'spg',idChamado:'3100',anexoHash:'a'.repeat(64)});
  assert.equal(r.mimeType,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  assert.deepEqual(Array.from(Buffer.from(r.base64,'base64')),[80,75,0,0]);
  assert.equal(calls.spg,1);
  assert.equal(calls.email,0);
  assert.equal(calls.drive,0);
});

test('download bloqueia pacote alterado desde a consulta e entrada inválida',()=>{
  const {context:c,calls}=context();
  assert.throws(()=>c.baixarPreviaExcelEmail({origem:'spg',idChamado:'3100',anexoHash:'b'.repeat(64)}),/mudaram/i);
  assert.throws(()=>c.baixarPreviaExcelEmail({origem:'antigo',idArquivo:'../../',idChamado:'3100'}),/ID do anexo/);
  assert.equal(calls.email,0);
});

test('download de Sheets legado exporta XLSX sem criar nem alterar arquivo no Drive',()=>{
  const {context:c,calls}=context();
  let nome='',url='';
  c.ScriptApp={getOAuthToken:()=> 'autorizacao-teste'};
  c.UrlFetchApp={fetch:(destino,options)=>{
    url=destino;
    assert.equal(options.headers.Authorization,'Bearer autorizacao-teste');
    return {getResponseCode:()=>200,getBlob:()=>({setName(n){nome=n;return this;},getName:()=>nome,getBytes:()=>[80,75,3,4]})};
  }};
  const r=c.baixarPreviaExcelEmail({origem:'antigo',idArquivo:'abcDEF1234567890',idChamado:'3100'});
  assert.match(url,/\/files\/abcDEF1234567890\/export\?mimeType=/);
  assert.equal(r.nome,'Chamado 3100.xlsx');
  assert.equal(calls.email,0);
  assert.equal(calls.drive,1);
});

test('prévia ignora XML não utilizado e reutiliza o ZIP na paginação de fallback',()=>{
  const {context:c,calls}=context();
  c.PF_PREVIA_CONJUNTO_MAX_JSON=1;
  c.Utilities.unzip=()=>{
    calls.unzip++;
    return zip.concat([{getName:()=> 'xl/calcChain.xml',
      getDataAsString:()=>{throw Error('XML de cálculo não deve ser lido pela prévia');}}]);
  };
  const r=c.obterConjuntoPreviaExcelEmail({origem:'spg',idChamado:'3100',anexoHash:'a'.repeat(64),aba:0});
  assert.equal(r.previaExcel,null);
  assert.ok(r.pagina && r.pagina.linhas.length>0);
  assert.equal(r.pagina.linhas[0].celulas[0],'Código do SKU');
  assert.equal(calls.unzip,1);
});
test('primeira consulta de E-mail reaproveita a mesma leitura do Excel no fallback',()=>{
  const {context:c,calls}=context();
  vm.runInContext(fs.readFileSync(path.join(__dirname,'..','EmailSPG.js'),'utf8'),c);
  c.PF_PREVIA_CONJUNTO_MAX_JSON=1;
  c._emailSpgPacote_=()=>({id:'3100',hash:'a'.repeat(64),dataDesejada:'15/11/2026',
    descricao:'Produção',blob:{getBytes:()=>[80,75,0,0],getName:()=> 'Chamado 3100.xlsx',
      copyBlob(){return this},setContentType(){return this}}});
  const r=c.consultarPacoteEmailSPG('3100','ultima','reentrada');
  assert.equal(r.anexoValido,true);
  assert.equal(r.previaExcel,null);
  assert.ok(r.paginaPreviaExcel && r.paginaPreviaExcel.linhas.length>0);
  assert.equal(calls.unzip,1);
});
