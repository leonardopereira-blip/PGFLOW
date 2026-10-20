// Run: node --test tests/email-spg.test.cjs
// Nao consulta a API nem envia e-mails: teste isolado das regras de seguranca.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

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


function createContext(zip) {
  const context = {
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' },
      unzip: () => zip,
      computeDigest: (algo, data) => Array.from(require('node:crypto').createHash('sha256')
        .update(Array.isArray(data) ? Buffer.from(data.map(x => (x+256)%256)) : String(data))
        .digest()),
      newBlob: (bytes,type,name) => ({getBytes:()=>Array.isArray(bytes) ? bytes : Array.from(Buffer.from(String(bytes),'utf8')),
        getName:()=>name||'',getDataAsString:()=>Array.isArray(bytes)?Buffer.from(bytes).toString():String(bytes),
        setName(value){name=value;return this},setContentType(){return this},copyBlob(){return this}})
    },
    XmlService:{parse:parseXml,getNamespace:()=>({})},
    UrlFetchApp: {}, PropertiesService: {}, LockService: {},
    CacheService: {}, GmailApp: {}, SpreadsheetApp: {},
    _spgLogin_: ()=>'token',
    obterDadosGraficas: ()=>({}),
    ID_PLANILHA_LOG: 'test'
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.join(__dirname,'..','EmailSPG.js'),'utf8');
  vm.runInContext(source, context);
  return context;
}

const ctx = createContext([]);
test('normaliza data ISO e rejeita data inexistente', () => {
  assert.equal(ctx._emailSpgData_('2026-11-15T00:00:00Z'),'15/11/2026');
  assert.equal(ctx._emailSpgData_('31/02/2026'),'');
});
test('replica regra de coleta (nao sexta, sabado ou domingo)', () => {
  assert.equal(ctx._emailSpgColeta_('15/11/2026'),'12/11/2026');
});
test('busca campos aninhados sem usar equivalencias por aproximacao', () => {
  assert.equal(ctx._emailSpgCampo_({solicitacao:{dataDesejadaEntrega:'2026-11-15'}},
    ['dataDesejadaEntrega']),'2026-11-15');
  assert.equal(ctx._emailSpgCampo_({solicitacao:{dataPrevista:'2026-11-15'}},
    ['dataDesejadaEntrega']),'');
});
test('valida id e bloqueia entradas inesperadas', () => {
  assert.equal(ctx._emailSpgId_(' 3100 '),'3100');
  assert.throws(()=>ctx._emailSpgId_('../3100'));
});
test('deduplica dentro de cada grupo e mantem um email por grafica', () => {
  assert.equal(ctx._emailSpgDestinos_(['a@example.com, A@EXAMPLE.COM','b@example.com']).length,2);
  assert.equal(ctx._emailSpgDestinos_(['a@example.com, b@example.com','b@example.com, a@example.com']).length,2);
  assert.deepEqual(Array.from(ctx._emailSpgDestinos_(['a@example.com, comum@example.com','b@example.com, comum@example.com'])),['a@example.com, comum@example.com','b@example.com, comum@example.com']);
  assert.throws(()=>ctx._emailSpgDestinos_(['invalido']));
});

test('nunca usa descricao de um SKU como descricao do chamado', () => {
  assert.equal(ctx._emailSpgDescricao_({items:[{descricao:'SKU indevido'}]}),'');
  assert.equal(ctx._emailSpgDescricao_({solicitacao:{descricao:'Produção correta'}}),
    'Produção correta');
});

const UUID_CHAMADO = '1b06d3ed-0d44-4577-a638-8a29c2b33491';
const UUID_VERSAO = '75a9b7b1-0b22-4b37-8b78-9323c1024005';
function zip(terceiraValida) {
  const stringCell=(ref,value)=>'<c r="'+ref+'" t="inlineStr"><is><t>'+value+'</t></is></c>';
  const sheet=(headers,values)=>'<worksheet><sheetData><row r="1">'+headers.map((v,i)=>stringCell(String.fromCharCode(65+i)+'1',v)).join('')+'</row><row r="2">'+values.map((v,i)=>stringCell(String.fromCharCode(65+i)+'2',v)).join('')+'</row></sheetData></worksheet>';
  const sku='1. ID_Código SKU',kit='1. ID_Código Kit';
  return [
    ['xl/workbook.xml','<workbook><sheets><sheet name="Espec. e Tiragem de Itens" r:id="rId1"/>' +
      '<sheet name="Árvore e Tiragem de Kits" r:id="rId2"/><sheet name="Tiragem (legado)" r:id="rId3"/>'+
      '<sheet name="Esp. Técnica (legado)" r:id="rId4"/><sheet name="Árvore (legado)" r:id="rId5"/></sheets></workbook>'],
    ['xl/_rels/workbook.xml.rels','<Relationships>'+[1,2,3,4,5].map(i=>'<Relationship Id="rId'+i+'" Target="worksheets/sheet'+i+'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>').join('')+'</Relationships>'],
    ['xl/worksheets/_rels/sheet1.xml.rels','<Relationships><Relationship Id="link1" Target="https://spg.prod.arcocv.co/op/'+UUID_VERSAO+'" TargetMode="External"/></Relationships>'],
    ['xl/worksheets/sheet1.xml',sheet([sku],['SKU1'])],
    ['xl/worksheets/sheet2.xml',sheet([sku,kit],['SKU1','KIT1'])],
    ['xl/worksheets/sheet3.xml',sheet([sku,kit,'15. DES_CD envio','16. TIR_SKU'],['SKU1','KIT1','JGS','3'])],
    ['xl/worksheets/sheet4.xml',sheet([sku],[terceiraValida?'SKU1':''])],
    ['xl/worksheets/sheet5.xml',sheet([sku,kit],['SKU1','KIT1'])],
    ['xl/sharedStrings.xml','<sst><si><t>Texto original</t></si></sst>']
  ].map(([name,content])=>({getName:()=>name,getDataAsString:()=>content}));
}
function validBlob(){return {getBytes:()=>[80,75,...Array(110).fill(0)],copyBlob(){return this},setContentType(){return this}};}
test('excel completo com dados nas tres abas passa', () => {
  const c = createContext(zip(true));
  assert.equal(c._emailSpgChecarXlsx_(validBlob()),true);
});
test('excel com aba somente de cabecalho e bloqueado', () => {
  const c = createContext(zip(false));
  assert.throws(()=>c._emailSpgChecarXlsx_(validBlob()),
    /Esp\. Técnica/);
});
test('dados opcionais so podem ser omitidos por escolha explicita', () => {
  const p = {dataDesejada:'',descricao:''};
  assert.throws(()=>ctx._emailSpgSobrescrita_(p,{sobrescrita:{}}));
  const ok=ctx._emailSpgSobrescrita_(p,{sobrescrita:{omitirData:true,omitirDescricao:true}});
  assert.equal(ok.data,'');
  assert.equal(ok.descricao,'');
});




test('Cadastro valida e-mails, nomes e permite excluir entradas', () => {
  const original = ctx._emailSpgValidarCadastro_({
    Alpha:['A@example.com','a@example.com','b@example.com'],
    Beta:['c@example.com']
  });
  assert.equal(original.Alpha.length,2);
  assert.equal(original.Beta[0],'c@example.com');
  const excluido = ctx._emailSpgValidarCadastro_({Alpha:['a@example.com']});
  assert.equal(excluido.Alpha.length,1);
  assert.equal(Object.hasOwn(excluido,'Beta'),false);
  assert.throws(()=>ctx._emailSpgValidarCadastro_({Alpha:[]}),/1 a 80/);
  assert.throws(()=>ctx._emailSpgValidarCadastro_(JSON.parse('{"__proto__":["a@example.com"]}')),/grafica/);
  assert.throws(()=>ctx._emailSpgValidarCadastro_({Alpha:['e-mail invalido']}),/Email invalido/);
});
test('Lista inicial vem do portal sem Sheets e cadastro salvo preserva alteracoes e exclusoes', () => {
  const c = createContext([]), props = {};
  c.PropertiesService.getScriptProperties = () => ({getProperty:k=>props[k]||null});
  c.SpreadsheetApp.openById = () => {throw new Error('Nao consultar Sheets');};
  c.obterDadosGraficas = () => {throw new Error('Nao consultar cadastro antigo');};
  const initial = c.obterCadastroEmailsSPG();
  assert.equal(Object.keys(initial.graficas).length,10);
  assert.equal(initial.graficas.Logprint.length,11);
  assert.equal(initial.graficas.Teste[0],'leonardo.pereira@arcoeducacao.com.br');
  props[c.EMAIL_SPG_DESTINATARIOS_PROP]=JSON.stringify({Nova:['new@example.com']});
  assert.equal(Object.keys(c.obterCadastroEmailsSPG().graficas).join(','),'Nova');
  props[c.EMAIL_SPG_DESTINATARIOS_PROP]='{}';
  assert.equal(Object.keys(c.obterCadastroEmailsSPG().graficas).length,10);
  props[c.EMAIL_SPG_CADASTRO_EDITADO_PROP]='1';
  assert.equal(Object.keys(c.obterCadastroEmailsSPG().graficas).length,0);
});

test('ordem de graficas ou emails nao gera falso conflito de versao', () => {
  assert.equal(ctx._emailSpgVersaoCadastro_({B:['z@example.com','a@example.com'],A:['b@example.com']}),
    ctx._emailSpgVersaoCadastro_({A:['b@example.com'],B:['a@example.com','z@example.com']}));
});

test('pagina sem comentarios injeta versao salva e permite editar e salvar sem falso conflito', () => {
  const c=createContext([]),props={};
  c.PropertiesService.getScriptProperties=()=>({getProperty:k=>props[k]||null,setProperty:(k,v)=>{props[k]=v;}});
  c.LockService.getScriptLock=()=>({tryLock:()=>true,releaseLock(){}});
  const defaults=c.obterDadosGraficasEmailSPG();
  const old={};Object.keys(defaults).reverse().forEach(name=>{old[name]=defaults[name].slice().reverse();});
  props[c.EMAIL_SPG_DESTINATARIOS_PROP]=JSON.stringify(old);
  const html=fs.readFileSync(path.join(__dirname,'..','TelaEmail.html'),'utf8').replace('/*PF_EMAIL_CADASTRO_INICIAL*/','');
  c.HtmlService={createHtmlOutputFromFile:()=>({getContent:()=>html})};
  vm.runInContext(fs.readFileSync(path.join(__dirname,'..','WebApp.js'),'utf8'),c);
  const loaded=JSON.parse(c.include('TelaEmail').match(/var pfCadastroSalvo = (.*);/)[1]);
  assert.equal(loaded.versao,c.obterCadastroEmailsSPG().versao);
  loaded.graficas.Star7.push('novo@example.com');
  loaded.graficas.FormaCerta=loaded.graficas.FormaCerta.filter(e=>e!=='alessandra@formacerta.com.br');
  const saved=c.salvarCadastroEmailsSPG(loaded.graficas,loaded.versao);
  assert.ok(saved.graficas.Star7.includes('novo@example.com'));
  assert.ok(!saved.graficas.FormaCerta.includes('alessandra@formacerta.com.br'));
  assert.equal(c.obterCadastroEmailsSPG().versao,saved.versao);
});

test('lista fixa e cadastro vazio legado usam a mesma versao para o primeiro salvamento', () => {
  const c=createContext([]),props={EMAIL_SPG_DESTINATARIOS_GRAFICAS:'{}'};
  c.PropertiesService.getScriptProperties=()=>({getProperty:k=>props[k]||null,setProperty:(k,v)=>{props[k]=v;}});
  c.LockService.getScriptLock=()=>({tryLock:()=>true,releaseLock(){}});
  const html=fs.readFileSync(path.join(__dirname,'..','TelaEmail.html'),'utf8');
  const hash=html.match(/var pfCadastroInicial = \{graficas:pfGraficasPadrao,versao:"([a-f0-9]+)"\}/)[1];
  const loaded=c.obterCadastroEmailsSPG();
  assert.equal(loaded.versao,hash);
  loaded.graficas.Teste.push('outro@example.com');
  assert.equal(c.salvarCadastroEmailsSPG(loaded.graficas,hash).graficas.Teste.length,2);
});

test('Pagina inclui cadastro e versao no HTML sem RPC e protege texto de fechamento de script', () => {
  const c=createContext([]);
  const html=fs.readFileSync(path.join(__dirname,'..','TelaEmail.html'),'utf8');
  c.obterCadastroEmailsSPG=()=>({graficas:{'</script><img>':['test@example.com']},versao:'v1'});
  c.HtmlService={createHtmlOutputFromFile:()=>({getContent:()=>html})};
  vm.runInContext(fs.readFileSync(path.join(__dirname,'..','WebApp.js'),'utf8'),c);
  const rendered=c.include('TelaEmail');
  assert.doesNotMatch(rendered,/PF_EMAIL_CADASTRO_INICIAL/);
  assert.doesNotMatch(rendered,/\.obterCadastroEmailsSPG\(/);
  assert.doesNotMatch(rendered,/<\/script><img>/);
  const embedded=rendered.match(/var pfCadastroSalvo = (.*);/)[1];
  assert.equal(JSON.parse(embedded).versao,'v1');
  assert.equal(JSON.parse(embedded).graficas['</script><img>'][0],'test@example.com');
});

test('Cadastro compartilhado usa versao e impede sobrescrita concorrente', () => {
  const props = {};
  const locks = {tryLock:()=>true,releaseLock:()=>{}};
  ctx.PropertiesService.getScriptProperties = () => ({
    getProperty: key => props[key] || null,
    setProperty: (key,val) => {props[key]=val;}
  });
  ctx.LockService.getScriptLock = () => locks;
  const antes = ctx.obterCadastroEmailsSPG();
  assert.equal(typeof antes.versao,'string');
  const depois = ctx.salvarCadastroEmailsSPG({GraficaNova:['a@example.com']},antes.versao);
  assert.equal(depois.graficas.GraficaNova[0],'a@example.com');
  assert.throws(()=>ctx.salvarCadastroEmailsSPG({GraficaNova:['b@example.com']},antes.versao),
    /alterado por outra pessoa/);
  const final = ctx.salvarCadastroEmailsSPG({},depois.versao);
  assert.equal(Object.keys(final.graficas).length,0);
  assert.equal(ctx.obterCadastroEmailsSPG().vazioIntencional,true);
});


test('Interface contem popup de cadastro, exclusao, cancelamento e novo envio', () => {
  const html=fs.readFileSync(path.join(__dirname,'..','TelaEmail.html'),'utf8');
  assert.match(html,/id="pf-modal"/);
  assert.match(html,/Gerenciar e-mails/);
  assert.match(html,/Excluir gráfica/);
  assert.match(html,/Salvar cadastro/);
  assert.match(html,/Cancelar alterações/);
  assert.match(html,/Cancelar e limpar/g);
  assert.match(html,/Iniciar novo envio/);
  assert.doesNotMatch(html,/prompt\(/);
  assert.match(html,/pfRenderDestinatarios\(\)/);
  assert.match(html,/dispararEmailFinal\(pacote\)/);
  assert.match(html,/dispararEmailPacoteSPG\(pacote\)/);
  const source=html.match(/<script>([\s\S]*?)<\/script>/)[1];
  assert.doesNotThrow(()=>new Function(source));
});
test('Cancelar limpa identificacao, selecoes, origem, assunto, corpo e volta ao inicio', () => {
  const html=fs.readFileSync(path.join(__dirname,'..','TelaEmail.html'),'utf8');
  const source=html.match(/<script>([\s\S]*?)<\/script>/)[1];
  const elements={};
  const classe=()=>({
    active:{},
    toggle(nome,ligar){this.active[nome]=ligar;},
    contains(nome){return !!this.active[nome];}
  });
  const el=id=>elements[id]||(elements[id]={
    value:'sujo',textContent:'sujo',innerHTML:'',checked:false,disabled:false,
    style:{display:'block'},classList:classe(),hidden:false,removeAttribute(){},setAttribute(){},
    querySelectorAll(){return [];}
  });
  const checked=[{checked:true,value:'Grafica A'},{checked:false,value:'Grafica B'}];
  const radioTipo={checked:false,value:'inclusivo'};
  const radioVersao={checked:false,value:'manual'};
  const document={
    getElementById:el,
    querySelector(query){
      if(query==='input[name="tipoEmail"]:checked') return {value:radioTipo.checked?'reentrada':'inclusivo'};
      if(query==='input[name="tipoEmail"][value="reentrada"]')return radioTipo;
      if(query==='input[name="pfEmailVersao"][value="ultima"]')return radioVersao;
      if(query==='input[name="pfEmailVersao"]:checked')return radioVersao;
      throw Error('selector inesperado '+query);
    },
    querySelectorAll(query){return query.indexOf('pf-email-grafica')>=0?checked:[];}
  };
  const context={document,alert:()=>{throw Error('alerta inesperado');}};
  vm.createContext(context);
  vm.runInContext(source.replace(/if \(pfCadastroInicial\) \{\s*pfCarregarCadastro\(\[\]\);\s*\}/,''),context);
  context.pfEnviando=false;
  context.pfRascunhoCadastro={Grafica:['email@example.com']};
  context.dadosGlobais={idChamado:'3100',idArquivo:'antigo'};
  context.pfGruposSelecionados=['Grafica A'];
  el('inputChamado').value='3100';
  el('inputNumVersao').value='4';
  context.pfResetar();
  assert.equal(el('inputChamado').value,'');
  assert.equal(el('inputNumVersao').value,'');
  assert.equal(el('assuntoEmail').value,'');
  assert.equal(el('corpoEmail').value,'');
  assert.equal(checked[0].checked,false);
  assert.equal(context.dadosGlobais.idChamado,undefined);
  assert.equal(context.pfGruposSelecionados.length,0);
  assert.equal(el('passo1').classList.active.ativo,true);
  assert.equal(el('passo2').classList.active.ativo,false);
  assert.equal(el('pf-modal').hidden,true);

});
test('Inclusivos usam pacote API sem desvio automatico para legado', () => {
  const html=fs.readFileSync(path.join(__dirname,'..','TelaEmail.html'),'utf8');
  const source=html.match(/<script>([\s\S]*?)<\/script>/)[1];
  const vals={inputChamado:{value:'3100',style:{}},
    inputNumVersao:{value:'',style:{}},
    'mensagem-erro':{style:{},textContent:''},
    'pgflowOrigemSPG':{style:{},textContent:''}};
  let legado=0,api=0;
  const context={
    document:{
      getElementById:id=>vals[id]||(vals[id]={value:'',style:{},textContent:'',setAttribute(){}}),
      querySelector:q=>q.includes('tipoEmail')?{value:'inclusivo'}:{value:'ultima'},
      querySelectorAll:q=>q.includes('.pf-email-grafica:checked')?[{value:'Grafica A'}]:[]
    }
  };
  const runner={withSuccessHandler(){return this;},withFailureHandler(){return this;},consultarPacoteEmailSPG(id,versao,tipo){api++;assert.equal(tipo,'inclusivo');}};
  context.google={script:{run:runner}};
  vm.createContext(context);
  vm.runInContext(source.replace(/if \(pfCadastroInicial\) \{\s*pfCarregarCadastro\(\[\]\);\s*\}/,''),context);
  context.pfVersaoCadastro='versao';
  context.pfCarregarLegado=()=>{legado++;};
  context.avancarParaRevisao();
  assert.equal(legado,0);assert.equal(api,1);
  assert.equal(context.pfTipoAtual,'inclusivo');
});

// Contrato observado do SPG: numero no relatorio, identidade UUID nos detalhes.
function apiHarness(options={}) {
  const c=createContext(options.zip || zip(true));
  const state={network:[],emails:[],attempts:0,cache:new Map(),properties:new Map()};
  const detalhe={id:UUID_CHAMADO,numero:3100,status:'SOLICITADO',versaoId:UUID_VERSAO,
    dataDesejadaEntrega:'2026-11-15',versao:{id:UUID_VERSAO,opCode:'OP-26-520-762'}};
  const versao={id:UUID_VERSAO,status:'PUBLICADA',descricao:'Reentrada 3100 · SPOT · SPE',
    opCode:'OP-26-520-762',opParcial:true};
  const pages=[{page:1,totalPages:1,limit:100,total:1,rows:[
    {Chamado:3100,'ID Chamado':UUID_CHAMADO},
    {Chamado:3100,'ID Chamado':UUID_CHAMADO}]}];
  const data={detalhe,versao,pages};
  options.mutate?.(data);
  const response=(body,status=200,binary=false)=>({
    getResponseCode:()=>status,
    getContentText:()=>JSON.stringify(body),
    getContent:()=>binary?[80,75,...Array(110).fill(0)]:Array.from(Buffer.from(JSON.stringify(body)))
  });
  c.UrlFetchApp.fetch=(url,opts)=>{
    state.network.push({url,method:opts.method});
    const u=new URL(url);
    if(u.pathname==='/api/reentradas/sheets/relatorio')
      return response(data.pages[Number(u.searchParams.get('page'))-1]);
    if(u.pathname==='/api/reentradas/'+UUID_CHAMADO) return response({reentrada:data.detalhe});
    if(u.pathname==='/api/tiragem/'+UUID_VERSAO) return response({versao:data.versao});
    if(u.pathname==='/api/reentradas/'+UUID_CHAMADO+'/op') return response('',200,true);
    return response({erro:'rota inesperada'},400);
  };
  c.LockService.getScriptLock=()=>({tryLock:()=>true,waitLock(){},releaseLock(){}});
  c.CacheService.getUserCache=()=>({get:k=>state.cache.get(k)||null,put:(k,v)=>state.cache.set(k,v)});
  c.GmailApp.sendEmail=(...args)=>{
    state.attempts++;
    if(options.failAt===state.attempts) throw new Error('falha simulada');
    state.emails.push(args);
  };
  c.PropertiesService.getScriptProperties=()=>({getProperty:k=>state.properties.get(k)||null,
    setProperty:(k,v)=>state.properties.set(k,v)});
  const prepare=()=>c.consultarPacoteEmailSPG('3100','ultima','reentrada');
  const send=(prepared,extra={})=>c.dispararEmailPacoteSPG({idChamado:'3100',tipoEmail:'reentrada',
    assunto:'Assunto do teste',corpo:'Texto do teste',anexoHash:prepared.anexoHash,
    arrayDestinos:['a@example.com, interno@example.com','b@example.com, interno@example.com'],...extra});
  return {c,state,data,prepare,send};
}

test('identifica o SPOT de ampliados pelos metadados do chamado sem obter producao ou Excel',()=>{
  const h=apiHarness({mutate(d){
    d.detalhe.tipoSolicitacao='SPOT';d.detalhe.justificativa='Tiragem V1 Ampliados';
    d.detalhe.versao.status='PUBLICADA';
  }});
  let logins=0;
  h.c._spgLogin_=()=>{logins++;return 'token';};
  h.c.Utilities.unzip=()=>{throw new Error('nao deve ler Excel');};
  const result=h.c.identificarTipoEmailSPG(' 3100 ');
  assert.deepEqual(JSON.parse(JSON.stringify(result)),{
    encontrado:true,idChamado:'3100',tipoEmail:'inclusivo',tipoSolicitacao:'SPOT',
    statusChamado:'SOLICITADO',statusProducao:'PUBLICADA',cancelado:false
  });
  assert.equal(logins,1);assert.equal(h.state.network.length,2);
  assert.ok(h.state.network.every(r=>r.method==='get'));
  assert.ok(h.state.network[0].url.includes('/api/reentradas/sheets/relatorio?'));
  assert.ok(h.state.network[1].url.endsWith('/api/reentradas/'+UUID_CHAMADO));
  assert.equal(h.state.network.some(r=>r.url.includes('/api/tiragem/')||r.url.endsWith('/op')),false);
  assert.equal(h.state.attempts,0);assert.equal(h.state.emails.length,0);
  assert.equal(h.state.properties.size,0);assert.equal(h.state.cache.size,0);
});

test('selecao automatica distingue REENTRADA, INCLUSIVOS e palavra ampliado normalizada em SPOT',()=>{
  const cases=[
    ['REENTRADA','','reentrada'],[' reentrada ','Ampliados','reentrada'],
    ['INCLUSIVOS','','inclusivo'],['inclusívos','','inclusivo'],
    ...['Ampliados','AMPLIADO','ampliada','ampliadas','AMPLÍADOS','(ampliados).']
      .map(word=>[' sPoT ','Tiragem V1 '+word,'inclusivo'])
  ];
  for(const [tipo,justificativa,expected] of cases){
    const h=apiHarness({mutate(d){
      d.detalhe.tipoSolicitacao=tipo;d.detalhe.justificativa=justificativa;
    }}),result=h.c.identificarTipoEmailSPG('3100');
    assert.equal(result.encontrado,true);
    assert.equal(result.tipoEmail,expected,JSON.stringify([tipo,justificativa]));
    assert.equal(result.tipoSolicitacao,tipo.trim());
  }
});

test('SPOT comum e tipos desconhecidos nao assumem um modelo de email',()=>{
  const cases=[
    ['SPOT','Tiragem V1'],['SPOT','superampliados'],['SPOT','ampliados123'],
    ['SPOT','ampliadamente'],['SPOT','ampliadores'],['SPOT',null],
    ['OUTRO','Tiragem Ampliados'],['','Tiragem Ampliados'],[null,null]
  ];
  for(const [tipo,justificativa] of cases){
    const h=apiHarness({mutate(d){
      d.detalhe.tipoSolicitacao=tipo;d.detalhe.justificativa=justificativa;
    }}),result=h.c.identificarTipoEmailSPG('3100');
    assert.equal(result.encontrado,true);
    assert.equal(result.tipoEmail,'',JSON.stringify([tipo,justificativa]));
  }
});

test('identificacao nao busca tipo ou justificativa em SKU, producao ou campos parecidos',()=>{
  for(const tipo of [undefined,'SPOT']){
    const h=apiHarness({mutate(d){
      d.detalhe.tipoSolicitacao=tipo;
      d.detalhe.descricao='Ampliados';d.detalhe.justificativaSolicitacao='Ampliados';
      d.detalhe.items=[{tipoSolicitacao:'INCLUSIVOS',justificativa:'Ampliados'}];
      d.detalhe.versao.tipoSolicitacao='REENTRADA';
      d.detalhe.versao.justificativa='Ampliados';
      d.versao.tipoSolicitacao='INCLUSIVOS';d.versao.descricao='Ampliados';
    }}),result=h.c.identificarTipoEmailSPG('3100');
    assert.equal(result.encontrado,true);assert.equal(result.tipoEmail,'');
    assert.equal(h.state.network.length,2);
  }
});

test('identificacao so aceita identidade numerica e UUID confirmados no detalhe',()=>{
  for(const mutate of [d=>d.detalhe.numero=9999,d=>d.detalhe.id=UUID_VERSAO,
    d=>d.detalhe=null]){
    const h=apiHarness({mutate}),result=h.c.identificarTipoEmailSPG('3100');
    assert.equal(result.encontrado,false);assert.equal(result.idChamado,'3100');
    assert.match(result.erro,/confirmar a identificacao/);
    assert.equal(Object.hasOwn(result,'tipoEmail'),false);
    assert.equal(h.state.network.length,2);
  }
});

test('identificacao informa cancelamento do chamado e da versao com vinculo valido',()=>{
  for(const [statusChamado,statusProducao] of [
    ['CANCELADO','PUBLICADA'],[' cancelada ','PUBLICADA'],
    ['SOLICITADO','CANCELADA'],['SOLICITADO',' cancelado ']
  ]){
    const h=apiHarness({mutate(d){
      d.detalhe.tipoSolicitacao='REENTRADA';d.detalhe.status=statusChamado;
      d.detalhe.versao.status=statusProducao;
    }}),result=h.c.identificarTipoEmailSPG('3100');
    assert.equal(result.encontrado,true);assert.equal(result.cancelado,true);
    assert.equal(result.statusChamado,statusChamado.trim());
    assert.equal(result.statusProducao,statusProducao.trim());
    assert.equal(result.tipoEmail,'reentrada');assert.equal(h.state.network.length,2);
  }
});

test('identificacao ignora status de producao com vinculo de versao ausente ou invalido',()=>{
  for(const mutate of [d=>d.detalhe.versao.id=UUID_CHAMADO,
    d=>d.detalhe.versao=null,d=>delete d.detalhe.versaoId,
    d=>{d.detalhe.versaoId='invalido';d.detalhe.versao.id='invalido';}]){
    const h=apiHarness({mutate(d){d.detalhe.versao.status='CANCELADA';mutate(d);}});
    const result=h.c.identificarTipoEmailSPG('3100');
    assert.equal(result.encontrado,true);assert.equal(result.statusProducao,'');
    assert.equal(result.cancelado,false);assert.equal(h.state.network.length,2);
  }
  const h=apiHarness({mutate(d){d.detalhe.status='CANCELADO';d.detalhe.versao=null;}});
  assert.equal(h.c.identificarTipoEmailSPG('3100').cancelado,true);
});

test('falhas de identificacao retornam erro sem escolher tipo ou consultar anexos',()=>{
  const scenarios=[
    h=>{h.c._spgLogin_=()=>null;},
    h=>{h.c.UrlFetchApp.fetch=()=>{throw new Error('SPG indisponivel');};},
    h=>{h.data.pages[0].rows=[];},
    h=>{h.data.pages[0].rows[0]['ID Chamado']='invalido';},
    h=>{h.c.UrlFetchApp.fetch=()=>({getResponseCode:()=>200,getContentText:()=>'{invalido'});}
  ];
  for(const change of scenarios){
    const h=apiHarness();change(h);
    const result=h.c.identificarTipoEmailSPG('3100');
    assert.equal(result.encontrado,false);assert.equal(result.idChamado,'3100');
    assert.equal(typeof result.erro,'string');assert.ok(result.erro);
    assert.equal(Object.hasOwn(result,'tipoEmail'),false);
    assert.equal(h.state.network.some(r=>r.url.includes('/api/tiragem/')||r.url.endsWith('/op')),false);
    assert.equal(h.state.attempts,0);
  }
  const h=apiHarness(),result=h.c.identificarTipoEmailSPG('../3100');
  assert.equal(result.encontrado,false);assert.match(result.erro,/numero de chamado valido/);
  assert.equal(h.state.network.length,0);
});

test('consulta pacote completo por UUID e le dados da versao vinculada',()=>{
  const h=apiHarness(),p=h.prepare();
  assert.equal(p.status,'ok');assert.equal(p.anexoValido,true);
  assert.equal(p.statusChamado,'SOLICITADO');assert.equal(p.statusProducao,'PUBLICADA');
  assert.equal(p.cancelado,false);
  assert.equal(p.dataDesejada,'15/11/2026');assert.equal(p.dataColeta,'12/11/2026');
  assert.equal(p.descricao,'Reentrada 3100 · SPOT · SPE');
  assert.equal(p.op,'OP-26-520-762');assert.equal(p.versaoId,UUID_VERSAO);assert.equal(p.opParcial,true);
  assert.equal(h.state.network.length,4);
  assert.ok(h.state.network.every(r=>r.method==='get' && !/\/reentradas\/3100/.test(r.url)));
  assert.ok(h.state.network.some(r=>r.url.endsWith('/reentradas/'+UUID_CHAMADO+'/op')));
});

test('resolve numero em paginas seguintes e recusa numero ambiguo entre paginas',()=>{
  const h=apiHarness({mutate(d){
    d.pages[0]={page:1,totalPages:2,rows:[{Chamado:3099,'ID Chamado':UUID_CHAMADO}]};
    d.pages.push({page:2,totalPages:2,rows:[{Chamado:3100,'ID Chamado':UUID_CHAMADO}]});
  }});
  assert.equal(h.prepare().status,'ok');
  assert.equal(h.state.network.filter(r=>r.url.includes('/sheets/relatorio')).length,2);
  const duplicate=apiHarness({mutate(d){
    d.pages[0].totalPages=2;
    d.pages.push({page:2,totalPages:2,rows:[{Chamado:3100,'ID Chamado':'b06fe3e4-6233-4e52-907d-093e161ed69e'}]});
  }});
  const p=duplicate.prepare();assert.equal(p.anexoValido,false);assert.match(p.erro,/identificacoes diferentes/);
  assert.equal(duplicate.state.network.length,2);
});

test('recusa metadados de paginacao invalidos e chamado nao encontrado',()=>{
  for(const mutate of [d=>d.pages[0].totalPages=Infinity,d=>d.pages[0].page=2,d=>d.pages[0].rows=[]]){
    const h=apiHarness({mutate}),p=h.prepare();
    assert.equal(p.anexoValido,false);assert.equal(h.state.network.length,1);
  }
});

test('recusa detalhes de outro chamado e versao incoerente antes de obter anexo',()=>{
  for(const mutate of [d=>d.detalhe.numero=9999,d=>d.detalhe.id='outro',
    d=>d.detalhe.versaoId='invalido',d=>d.detalhe.versao.id='outra',d=>d.versao.id='outra']){
    const h=apiHarness({mutate}),p=h.prepare();
    assert.equal(p.status,'incompleto');assert.equal(p.anexoValido,false);
    assert.equal(h.state.network.some(r=>r.url.endsWith('/op')),false);
  }
});

test('chamado cancelado tem resposta propria antes de consultar producao ou Excel',()=>{
  for(const status of ['CANCELADO','CANCELADA']){
    const h=apiHarness({mutate(d){
      d.detalhe.status=status;d.detalhe.dataDesejadaEntrega=null;
      d.detalhe.versao.status='CANCELADA';
    }}),p=h.prepare();
    assert.equal(p.status,'cancelado');assert.equal(p.cancelado,true);
    assert.equal(p.statusChamado,status);assert.equal(p.statusProducao,'CANCELADA');
    assert.equal(p.anexoValido,false);assert.equal(p.faltantes.length,0);
    assert.match(p.erro,/Este chamado está cancelado/);
    assert.equal(h.state.network.length,2);
    assert.equal(h.state.network.some(r=>r.url.includes('/api/tiragem/')||r.url.endsWith('/op')),false);
    assert.equal(h.state.emails.length,0);
  }
});

test('producao cancelada tem alerta proprio sem classificar o chamado como cancelado',()=>{
  for(const status of ['CANCELADO','CANCELADA']){
    const h=apiHarness({mutate(d){d.versao.status=status;}}),p=h.prepare();
    assert.equal(p.status,'cancelado');assert.equal(p.cancelado,true);
    assert.equal(p.statusChamado,'SOLICITADO');assert.equal(p.statusProducao,status);
    assert.equal(p.anexoValido,false);assert.equal(p.faltantes.length,0);
    assert.match(p.erro,/versão de produção deste chamado está cancelada/);
    assert.doesNotMatch(p.erro,/Este chamado está cancelado/);
    assert.equal(h.state.network.length,3);
    assert.equal(h.state.network.some(r=>r.url.endsWith('/op')),false);
    assert.equal(h.state.emails.length,0);
  }
});

test('selecao de versao antiga consulta o status atual sem baixar o Excel atual',()=>{
  const h=apiHarness(),p=h.c.consultarPacoteEmailSPG('3100','manual','reentrada');
  assert.equal(p.status,'incompleto');assert.equal(p.cancelado,false);
  assert.equal(p.statusChamado,'SOLICITADO');assert.equal(p.statusProducao,'PUBLICADA');
  assert.equal(p.anexoValido,false);assert.match(p.erro,/versao antiga/);
  assert.equal(h.state.network.length,3);
  assert.equal(h.state.network.some(r=>r.url.endsWith('/op')),false);
});

test('selecao de versao antiga nao contorna cancelamento de chamado ou producao',()=>{
  for(const mutate of [d=>d.detalhe.status='CANCELADO',d=>d.versao.status='CANCELADA']){
    const h=apiHarness({mutate}),p=h.c.consultarPacoteEmailSPG('3100','manual','reentrada');
    assert.equal(p.status,'cancelado');assert.equal(p.cancelado,true);
    assert.equal(p.anexoValido,false);assert.match(p.erro,/cancelad[oa]/);
    assert.doesNotMatch(p.erro,/fluxo antigo/);
    assert.equal(h.state.network.some(r=>r.url.endsWith('/op')),false);
    assert.equal(h.state.emails.length,0);
  }
});

test('status confirmado permanece disponivel se outros metadados falharem',()=>{
  const h=apiHarness({mutate(d){d.detalhe.versaoId='invalido';}}),p=h.prepare();
  assert.equal(p.status,'incompleto');assert.equal(p.cancelado,false);
  assert.equal(p.statusChamado,'SOLICITADO');assert.equal(p.statusProducao,'');
  assert.equal(p.anexoValido,false);
  const wrong=apiHarness({mutate(d){d.detalhe.id='outro';d.detalhe.status='CANCELADO';}}).prepare();
  assert.equal(wrong.status,'incompleto');assert.equal(wrong.statusChamado,'');
  assert.equal(wrong.statusProducao,'');assert.equal(wrong.cancelado,false);
});

test('nao completa data ou descricao usando campos de produtos aninhados',()=>{
  const h=apiHarness({mutate(d){
    d.detalhe.dataDesejadaEntrega=null;
    d.detalhe.items=[{dataDesejadaEntrega:'2027-02-05',descricao:'SKU que nao e descricao do chamado'}];
    d.versao.descricao='';d.versao.items=[{descricao:'Outro SKU'}];
  }}),p=h.prepare();
  assert.equal(p.status,'incompleto');assert.equal(p.anexoValido,true);
  assert.equal(p.statusChamado,'SOLICITADO');assert.equal(p.statusProducao,'PUBLICADA');
  assert.equal(p.cancelado,false);
  assert.equal(p.dataDesejada,'');assert.equal(p.descricao,'');assert.equal(p.faltantes.length,2);
  const missing=h.send(p);assert.equal(missing.sucesso,false);assert.equal(h.state.emails.length,0);
  const explicit=h.send(p,{sobrescrita:{omitirData:true,omitirDescricao:true}});
  assert.equal(explicit.sucesso,true);assert.equal(h.state.emails.length,2);
});

test('anexo precisa identificar a versao vinculada do chamado',()=>{
  const other=zip(true).map(p=>({getName:p.getName,getDataAsString:()=>p.getDataAsString().replaceAll(UUID_VERSAO,'c1acf191-3f72-4bb5-8aa5-682c15e129e1')}));
  const h=apiHarness({zip:other}),p=h.prepare();
  assert.equal(p.anexoValido,false);assert.match(p.erro,/corresponde.*versão/);
  assert.equal(p.status,'incompleto');assert.equal(p.statusChamado,'SOLICITADO');
  assert.equal(p.statusProducao,'PUBLICADA');assert.equal(p.cancelado,false);
});

test('hash considera textos compartilhados e relacionamentos mas ignora docProps',()=>{
  const original=zip(true),c=createContext(original),blob=validBlob();
  const before=c._emailSpgHash_(blob);
  c.Utilities.unzip=()=>original.concat([{getName:()=> 'docProps/core.xml',getDataAsString:()=>'<modified>agora</modified>'}]);
  assert.equal(c._emailSpgHash_(blob),before);
  for(const name of ['xl/sharedStrings.xml','xl/_rels/workbook.xml.rels','xl/workbook.xml']){
    c.Utilities.unzip=()=>original.map(p=>({getName:p.getName,getDataAsString:()=>p.getDataAsString()+(p.getName()===name?'mudou':'')}));
    assert.notEqual(c._emailSpgHash_(blob),before);
  }
});

test('alteracao de dados ou status entre revisao e envio exige nova revisao',()=>{
  for(const mutate of [d=>d.detalhe.dataDesejadaEntrega='2026-11-16',d=>d.versao.descricao='Nova descricao',
    d=>d.versao.opParcial=false,d=>d.detalhe.status='VALIDADO',d=>d.versao.status='APROVADA']){
    const h=apiHarness(),p=h.prepare();mutate(h.data);
    const result=h.send(p);assert.equal(result.sucesso,false);
    assert.match(result.erro,/dados do chamado foram alterados/);assert.equal(h.state.emails.length,0);
  }
});

test('cancelamento entre revisao e envio bloqueia Gmail sem baixar outro Excel',()=>{
  for(const mutate of [d=>d.detalhe.status='CANCELADO',d=>d.versao.status='CANCELADA']){
    const h=apiHarness(),p=h.prepare();mutate(h.data);h.state.network.length=0;
    const result=h.send(p);assert.equal(result.sucesso,false);
    assert.match(result.erro,/cancelad[oa]/);assert.equal(h.state.emails.length,0);
    assert.equal(h.state.network.some(r=>r.url.endsWith('/op')),false);
  }
});

test('envia um email por grafica com contatos comuns presentes nos dois grupos',()=>{
  const h=apiHarness(),p=h.prepare(),result=h.send(p);
  assert.equal(result.sucesso,true);assert.equal(result.enviados,2);
  assert.deepEqual(h.state.emails.map(args=>args[0]),['a@example.com, interno@example.com','b@example.com, interno@example.com']);
  assert.equal(h.send(p).sucesso,false);assert.equal(h.state.emails.length,2);
});

test('valida todos os destinatarios antes do primeiro envio',()=>{
  const h=apiHarness(),p=h.prepare();
  const result=h.send(p,{arrayDestinos:['a@example.com','email-invalido']});
  assert.equal(result.sucesso,false);assert.equal(h.state.emails.length,0);
});

test('falha parcial conserva bloqueio e informa quantos grupos concluiram',()=>{
  const h=apiHarness({failAt:2}),p=h.prepare(),result=h.send(p);
  assert.equal(result.sucesso,false);assert.match(result.erro,/1 de 2 grupos/);
  assert.equal(h.state.emails.length,1);
  const retry=h.send(p);assert.equal(retry.sucesso,false);
  assert.match(retry.erro,/ja foi iniciado/);assert.equal(h.state.attempts,2);
});


test('duas graficas com mesmos contatos ainda geram dois emails',()=>{
  const h=apiHarness(),p=h.prepare();
  const result=h.send(p,{arrayDestinos:['a@example.com, A@example.com','a@example.com']});
  assert.equal(result.sucesso,true);assert.equal(result.enviados,2);
  assert.deepEqual(h.state.emails.map(args=>args[0]),['a@example.com','a@example.com']);
});


test('inclusivo SPOT ampliados consulta e envia com Excel vinculado da API',()=>{
  const h=apiHarness({mutate(d){d.detalhe.tipoSolicitacao='SPOT';d.detalhe.justificativa='Tiragem V1 Ampliados';}});
  const p=h.c.consultarPacoteEmailSPG('3100','ultima','inclusivo');
  assert.equal(p.status,'ok');assert.equal(p.anexoValido,true);
  assert.equal(h.send(p,{tipoEmail:'inclusivo'}).sucesso,true);assert.equal(h.state.emails.length,2);
});
test('modelo inclusivo recusa outro tipo antes de baixar ou enviar anexo',()=>{
  const h=apiHarness({mutate(d){d.detalhe.tipoSolicitacao='REENTRADA';}});
  const p=h.c.consultarPacoteEmailSPG('3100','ultima','inclusivo');
  assert.equal(p.anexoValido,false);assert.match(p.erro,/não está identificado como inclusivo/);
  assert.equal(h.send(p,{tipoEmail:'inclusivo'}).sucesso,false);
  assert.equal(h.state.network.some(r=>r.url.endsWith('/op')),false);assert.equal(h.state.emails.length,0);
});

// A trava de tipo e exclusiva da validacao de custos de Reentradas.
test('validacao de reentradas recusa chamado Principal antes de baixar OP',()=>{
  const h=apiHarness({mutate(d){d.detalhe.tipoSolicitacao='PRINCIPAL';}});
  assert.throws(()=>h.c._emailSpgPacote_('3100',true,'reentrada','reentrada'),/não é uma REENTRADA/);
  assert.equal(h.state.network.some(r=>r.url.endsWith('/op')),false);
});
