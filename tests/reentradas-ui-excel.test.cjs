// Execute os controles reais da interface sem navegador ou serviços Google.
const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const html=fs.readFileSync(path.join(__dirname,'..','ValidaInterface.html'),'utf8');
const source=html.match(/<script>([\s\S]*?)<\/script>/)[1];
const plain=x=>JSON.parse(JSON.stringify(x));
function screen(){
  const elements=new Map(),requests=[],timers=[];
  function element(){
    const classes=new Set();let text='';
    return {value:'',hidden:false,disabled:false,style:{},children:[],innerHTML:'',
      get textContent(){return text},set textContent(v){text=String(v);this.children=[]},
      classList:{add:x=>classes.add(x),remove:x=>classes.delete(x),contains:x=>classes.has(x),
        toggle(x,on){if(on===undefined)on=!classes.has(x);on?classes.add(x):classes.delete(x)}},
      appendChild(x){this.children.push(x);return x},setAttribute(k,v){this[k]=v},
      querySelector(selector){return selector==='input'?this.children.find(x=>x.tag==='input')||null:null},
      focus(){},select(){}};
  }
  function el(id){if(!elements.has(id))elements.set(id,element());return elements.get(id)}
  const document={getElementById:el,querySelector:el,querySelectorAll:()=>[],
    createElement(tag){const e=element();e.tag=tag;return e}};
  const google={script:{}};
  Object.defineProperty(google.script,'run',{get(){
    const chain={withSuccessHandler(fn){this.success=fn;return this},withFailureHandler(fn){this.failure=fn;return this}};
    return new Proxy(chain,{get(target,key){return key in target?target[key]:
      (...args)=>requests.push({method:key,args:plain(args),success:target.success,failure:target.failure})}});
  }});
  const context={document,google,confirm:()=>true,setTimeout:(fn,delay)=>{timers.push({fn,delay});return timers.length},
    clearTimeout(){},setInterval:()=>1,clearInterval(){}};
  vm.createContext(context);vm.runInContext(source,context);
  el('rxOverlay').hidden=true;el('rxEditorPane').hidden=true;
  el('txtChamados').value='123';el('selectGraficas').value='LOGPRINT';
  context.verificarLiberacaoBotao();
  return {context,el,requests,timers,request:()=>requests.at(-1)};
}
const file={id:'ARQUIVO_1234567890',nome:'CALCULADORA FORNECEDOR LOGPRINT.xlsx',atualizadoEm:1000,
  url:'https://drive.google.com/file/d/ARQUIVO_1234567890',editavel:true,bloqueavel:true};
const conjunto={abas:[{nome:'CALC Geral',totalColunas:2,linhas:[{numero:5,valores:{0:'20',1:'texto'}}]}]};
function open(s){s.context.rxAbrirArquivo(file);s.request().success({arquivo:file,previaExcel:conjunto,formulas:{'0:5:1':true}})}

test('ajustes permitem voltar aos chamados e cancelar/iniciar nova antes da geração',()=>{
  const fase2=html.slice(html.indexOf('id="telaFase2"'),html.indexOf('id="telaFase3"'));
  assert.match(fase2,/onclick="rxVoltarChamados\(\)"/);
  assert.match(fase2,/onclick="resetarParaNovaCalculadora\(\)"/);
  const s=screen(),c=s.context;
  c.cacheChamadosValidos=['123'];s.el('txtChamados').disabled=true;s.el('selectGraficas').disabled=true;
  const seq=c._reentradasConsultaSeq;c.rxVoltarChamados();
  assert.equal(s.el('telaFase1').style.display,'block');assert.equal(s.el('telaFase2').style.display,'none');
  assert.equal(s.el('txtChamados').value,'123');assert.equal(s.el('selectGraficas').value,'LOGPRINT');
  assert.equal(s.el('txtChamados').disabled,false);assert.equal(s.el('selectGraficas').disabled,false);
  assert.equal(c.cacheChamadosValidos.length,0);assert(c._reentradasConsultaSeq>seq);assert.equal(s.requests.length,0);
  s.el('telaFase2').style.display='block';c.resetarParaNovaCalculadora();
  assert.equal(s.el('telaFase1').style.display,'block');assert.equal(s.el('telaFase2').style.display,'none');
  assert.equal(s.el('telaFase3').style.display,'none');assert.equal(s.el('txtChamados').value,'');
  assert.equal(s.el('selectGraficas').value,'LOGPRINT');assert.equal(s.requests.length,0);
});

test('resposta antiga de listagem não substitui atualização mais recente nem arquivo aberto',()=>{
  const s=screen(),c=s.context;c.rxAbrirLista();const old=s.request();c.rxCarregarLista();const fresh=s.request();
  fresh.success({arquivos:[file]});old.success({arquivos:[]});old.failure(new Error('antiga'));
  assert.equal(c.rxArquivos.length,1);assert.equal(s.el('rxListaStatus').textContent,'1 arquivo(s) encontrado(s).');
  c.rxCarregarLista();const stale=s.request();open(s);stale.success({arquivos:[]});
  assert.equal(c.rxArquivos.length,1);assert.equal(c.rxArquivo.id,file.id);
});

test('cancelar restauração não grava; confirmação impede duplicação e saída enquanto aguarda',()=>{
  const s=screen(),c=s.context;c.rxAbrirLista();s.request().success({arquivos:[file]});
  c.confirm=()=>false;c.rxBloquear(file);assert.equal(s.requests.length,1);
  c.confirm=()=>true;c.rxBloquear(file);const op=s.request();c.rxBloquear(file);c.rxFechar();
  assert.equal(s.requests.length,2);assert.equal(op.method,'bloquearArquivoReentradasExcel');
  assert.deepEqual(op.args,[file.id,file.atualizadoEm]);assert.equal(c.rxOcupado,true);
  assert.equal(s.el('rxOverlay').hidden,false);assert.match(s.el('rxListaStatus').textContent,/Aguarde/);
  op.failure(new Error('Falha controlada'));assert.equal(c.rxOcupado,false);
  assert.match(s.el('rxListaStatus').textContent,/Falha controlada/);c.rxFechar();assert.equal(s.el('rxOverlay').hidden,true);
});

test('fórmula com valor cacheado fica protegida pela coordenada sem impedir editar texto',()=>{
  const s=screen(),c=s.context;open(s);
  const row=s.el('rxCorpo').children[0],formula=row.children[1],texto=row.children[2];
  assert.equal(formula.textContent,'20');assert.match(formula.className,/rx-formula/);
  formula.ondblclick();assert.equal(formula.children.length,0);assert.match(s.el('rxMsg').textContent,/Fórmulas/);
  texto.ondblclick();const input=texto.children[0];input.value='alterado';input.onblur();
  assert.deepEqual(plain(c.rxPendente),{'0:5:2':{aba:0,linha:5,coluna:2,valor:'alterado'}});
});

test('células dependentes de fórmula matricial ou compartilhada são protegidas por intervalo',()=>{
  const s=screen(),c=s.context;c.rxAbrirArquivo(file);
  s.request().success({arquivo:file,previaExcel:conjunto,formulas:{},
    faixasFormulas:[{aba:0,x1:1,x2:2,y1:5,y2:6}]});
  assert.equal(c.rxEhFormula(0,5,2),true);assert.equal(c.rxEhFormula(0,6,2),true);
  assert.equal(c.rxEhFormula(1,5,2),false);assert.equal(c.rxEhFormula(0,7,2),false);
  const td=s.el('rxCorpo').children[0].children[2];td.ondblclick();
  assert.equal(td.children.length,0);assert.equal(c.rxTemEdicoes(),false);
  assert.match(s.el('rxMsg').textContent,/Fórmulas/);
});

test('mesclagem protege posições cobertas e orienta editar somente a célula âncora',()=>{
  const s=screen(),c=s.context;c.rxAbrirArquivo(file);
  s.request().success({arquivo:file,previaExcel:conjunto,formulas:{},
    mesclagens:[{aba:0,x1:1,x2:2,y1:5,y2:6}]});
  assert.equal(c.rxMesclagemCoberta(0,5,1),null);
  assert.equal(c.rxMesclagemCoberta(0,6,1).y1,5);
  assert.equal(c.rxMesclagemCoberta(1,5,2),null);assert.equal(c.rxMesclagemCoberta(0,7,2),null);
  const row=s.el('rxCorpo').children[0],anchor=row.children[1],covered=row.children[2];
  assert.match(covered.className,/rx-merged/);assert.match(covered.title,/Edite A5/);
  covered.ondblclick();assert.equal(covered.children.length,0);assert.equal(c.rxTemEdicoes(),false);
  assert.match(s.el('rxMsg').textContent,/Edite A5/);
  anchor.ondblclick();const input=anchor.children[0];input.value='nova âncora';input.onblur();
  assert.equal(c.rxPendente['0:5:1'].valor,'nova âncora');
  c.rxMoverColuna(1);assert.equal(c.rxMesclagemCoberta(0,5,2).x1,1);
  c.rxFechar();c.rxAbrirArquivo(file);assert.equal(c.rxMesclagens.length,0);
});

test('fechar/nova respeitam recusa ao descarte e fechar invalida prévia atrasada',()=>{
  const s=screen(),c=s.context;open(s);c.rxPendente={'0:5:2':{aba:0,linha:5,coluna:2,valor:'x'}};
  c.confirm=()=>false;c.rxFechar();c.resetarParaNovaCalculadora();
  assert.equal(c.rxTemEdicoes(),true);assert.equal(s.el('rxOverlay').hidden,false);assert.equal(s.el('txtChamados').value,'123');
  c.confirm=()=>true;c.rxFechar();assert.equal(c.rxTemEdicoes(),false);assert.equal(s.el('rxOverlay').hidden,true);
  c.rxAbrirArquivo(file);const old=s.request();c.rxFechar();old.success({arquivo:{...file,id:'antigo'},previaExcel:conjunto});
  assert.equal(c.rxArquivo.id,file.id);assert.equal(s.el('rxOverlay').hidden,true);
});

test('salvamento preserva pendências em erro ou confirmação incompleta e mantém mesmo ID',()=>{
  const s=screen(),c=s.context;open(s);
  const edit={aba:0,linha:5,coluna:2,valor:'alterado'};c.rxPendente={'0:5:2':edit};c.rxSalvarEdicoes();const first=s.request();
  assert.equal(first.method,'salvarEdicoesReentradasExcel');assert.deepEqual(first.args,[file.id,1000,[edit]]);
  c.rxFechar();assert.equal(s.el('rxOverlay').hidden,false);assert.equal(c.rxOcupado,true);
  first.success({});assert.equal(c.rxTemEdicoes(),true);assert.equal(c.rxOcupado,false);
  c.rxSalvarEdicoes();s.request().failure(new Error('rede'));assert.equal(c.rxTemEdicoes(),true);
  c.rxSalvarEdicoes();s.request().success({sucesso:true,id:'outro',atualizadoEm:2000});assert.equal(c.rxTemEdicoes(),true);
  c.rxSalvarEdicoes();s.request().success({sucesso:true,id:file.id,atualizadoEm:2000});
  assert.equal(c.rxTemEdicoes(),false);assert.equal(c.rxArquivo.atualizadoEm,2000);
  assert.equal(s.request().method,'obterPreviaReentradasExcel');assert.equal(s.request().args[0].id,file.id);
  assert.equal(s.request().args[0].atualizadoEm,2000);
});

test('conjunto leve muda página e busca localmente sem nova consulta ou perder edição',()=>{
  const s=screen(),c=s.context;open(s);const count=s.requests.length;
  c.rxPendente={'0:5:2':{aba:0,linha:5,coluna:2,valor:'alterado'}};
  c.rxEstado.busca='texto';c.rxPedirPagina();c.rxMoverLinha(1);c.rxMoverColuna(1);c.rxLimparBusca();
  assert.equal(s.requests.length,count);assert.equal(c.rxTemEdicoes(),true);assert.equal(c.rxFormulas['0:5:1'],true);
});

test('a edição 201 é recusada antes de entrar no lote, mantendo as 200 pendências salváveis',()=>{
  const s=screen(),c=s.context;open(s);
  for(let linha=10;linha<210;linha++)c.rxPendente['0:'+linha+':2']={aba:0,linha,coluna:2,valor:'x'};
  const td=s.el('rxCorpo').children[0].children[2];td.ondblclick();
  const input=td.children[0];input.value='excesso';input.onblur();
  assert.equal(Object.keys(c.rxPendente).length,200);assert.equal(c.rxPendente['0:5:2'],undefined);
  assert.match(s.el('rxMsg').textContent,/Salve as 200/);c.rxSalvarEdicoes();
  assert.equal(s.request().method,'salvarEdicoesReentradasExcel');assert.equal(s.request().args[2].length,200);
});
