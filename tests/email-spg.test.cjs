// Run: node --test tests/email-spg.test.cjs
// Nao consulta a API nem envia e-mails: teste isolado das regras de seguranca.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function createContext(zip) {
  const context = {
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' },
      unzip: () => zip,
      computeDigest: (algo, data) => Array.from(require('node:crypto').createHash('sha256')
        .update(Array.isArray(data) ? Buffer.from(data.map(x => (x+256)%256)) : String(data))
        .digest()),
      newBlob: (bytes) => ({getBytes:()=>Array.isArray(bytes) ? bytes : Array.from(Buffer.from(String(bytes),'utf8')),
        setName(){return this}})
    },
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
test('destinatarios duplicados nao causam multiplos envios', () => {
  assert.equal(ctx._emailSpgDestinos_(['a@example.com, A@EXAMPLE.COM','b@example.com']).length,2);
  assert.throws(()=>ctx._emailSpgDestinos_(['invalido']));
});

test('nunca usa descricao de um SKU como descricao do chamado', () => {
  assert.equal(ctx._emailSpgDescricao_({items:[{descricao:'SKU indevido'}]}),'');
  assert.equal(ctx._emailSpgDescricao_({solicitacao:{descricao:'Produção correta'}}),
    'Produção correta');
});

function zip(terceiraValida) {
  const xml = valido => '<worksheet><sheetData><row r="1"><c><t>header</t></c></row>' +
    (valido?'<row r="2"><c><v>3</v></c></row>':'<row r="2"><c><v></v></c></row>') +
    '</sheetData></worksheet>';
  return [
    ['xl/workbook.xml','<workbook><sheets><sheet name="Tiragem" r:id="rId1"/>' +
      '<sheet name="Árvore de produto" r:id="rId2"/><sheet name="Especificação Técnica" r:id="rId3"/></sheets></workbook>'],
    ['xl/_rels/workbook.xml.rels','<Relationships>' +
      '<Relationship Id="rId1" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Target="worksheets/sheet2.xml"/>' +
      '<Relationship Id="rId3" Target="worksheets/sheet3.xml"/></Relationships>'],
    ['xl/worksheets/sheet1.xml',xml(true)],
    ['xl/worksheets/sheet2.xml',xml(true)],
    ['xl/worksheets/sheet3.xml',xml(terceiraValida)]
  ].map(([name,content])=>({getName:()=>name,getDataAsString:()=>content}));
}
test('excel completo com dados nas tres abas passa', () => {
  const c = createContext(zip(true));
  assert.equal(c._emailSpgChecarXlsx_({getBytes:()=>[80,75,...Array(110).fill(0)]}),true);
});
test('excel com aba somente de cabecalho e bloqueado', () => {
  const c = createContext(zip(false));
  assert.throws(()=>c._emailSpgChecarXlsx_({getBytes:()=>[80,75,...Array(110).fill(0)]}),
    /especificacao/);
});
test('dados opcionais so podem ser omitidos por escolha explicita', () => {
  const p = {dataDesejada:'',descricao:''};
  assert.throws(()=>ctx._emailSpgSobrescrita_(p,{sobrescrita:{}}));
  const ok=ctx._emailSpgSobrescrita_(p,{sobrescrita:{omitirData:true,omitirDescricao:true}});
  assert.equal(ok.data,'');
  assert.equal(ok.descricao,'');
});


test('Inclusivos nao sao consultados pela rota de reentradas', () => {
  const resultado = ctx.consultarPacoteEmailSPG('3100','ultima','inclusivo');
  assert.equal(resultado.anexoValido,false);
  assert.equal(resultado.status,'incompleto');
  assert.match(resultado.erro,/inclusivos/i);
});
test('Bloqueia disparo de inclusivos com o anexo de reentradas', () => {
  const resultado = ctx.dispararEmailPacoteSPG({
    idChamado:'3100',tipoEmail:'inclusivo',
    assunto:'Teste',corpo:'Teste',arrayDestinos:['a@example.com']
  });
  assert.equal(resultado.sucesso,false);
  assert.match(resultado.erro,/somente reentradas/i);
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
    style:{display:'block'},classList:classe(),hidden:false,removeAttribute(){},
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
      if(query==='input[name="opcaoVersao"][value="ultima"]')return radioVersao;
      if(query==='input[name="opcaoVersao"]:checked')return radioVersao;
      throw Error('selector inesperado '+query);
    },
    querySelectorAll(query){return query.indexOf('chk-grafica')>=0?checked:[];}
  };
  const context={document,alert:()=>{throw Error('alerta inesperado');}};
  vm.createContext(context);
  vm.runInContext(source,context);
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
  assert.equal(el('pf-dica-tipo').style.display,'none');
});
test('Inclusivos ficam no fluxo legado sem tentativa de usar rota de reentradas', () => {
  const html=fs.readFileSync(path.join(__dirname,'..','TelaEmail.html'),'utf8');
  const source=html.match(/<script>([\s\S]*?)<\/script>/)[1];
  const vals={inputChamado:{value:'3100',style:{}},
    inputNumVersao:{value:'',style:{}},
    'mensagem-erro':{style:{},textContent:''},
    'pgflowOrigemSPG':{style:{},textContent:''}};
  let legado=0;
  const context={
    document:{
      getElementById:id=>vals[id]||(vals[id]={value:'',style:{},textContent:''}),
      querySelector:q=>q.includes('tipoEmail')?{value:'inclusivo'}:{value:'ultima'},
      querySelectorAll:q=>q==='.chk-grafica:checked'?[{value:'Grafica A'}]:[]
    }
  };
  vm.createContext(context);
  vm.runInContext(source,context);
  context.pfVersaoCadastro='versao';
  context.pfCarregarLegado=()=>{legado++;};
  context.avancarParaRevisao();
  assert.equal(legado,1);
  assert.equal(context.pfTipoAtual,'inclusivo');
});
