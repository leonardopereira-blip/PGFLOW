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
      appendChild(x){this.children.push(x);return x},append(...xs){this.children.push(...xs)},replaceChildren(...xs){this.children=xs},insertBefore(x,other){const i=this.children.indexOf(other);this.children.splice(i<0?0:i,0,x);return x},setAttribute(k,v){this[k]=v},
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

test('Custo Validado requer comprovante escolhido, confirmacao e arquiva download',()=>{
  const s=screen(),c=s.context,calc={...file,calculadora:true,
    nome:'CALCULADORA FORNECEDOR LOGPRINT - CHAMADO 3127.xlsx'};
  c.rxAbrirArquivo(calc);
  s.request().success({arquivo:calc,previaExcel:conjunto,formulas:{}});
  assert.equal(s.el('rxCustoValidado').disabled,false);
  c.rxPendente={'0:5:2':{aba:0,linha:5,coluna:2,valor:'novo'}};c.rxRenderPagina();
  assert.equal(s.el('rxCustoValidado').disabled,true);
  c.rxFinalizarCusto();assert.equal(s.requests.length,1);
  c.rxPendente={};c.rxRenderPagina();c.rxAbrirConfirmacao();
  assert.equal(s.request().method,'obterConfirmacoesEmailCustoReentradas');
  assert.deepEqual(s.request().args,[calc.id,calc.atualizadoEm,'',0]);
  s.request().success({emails:[],sucesso:true,sugestaoJaTratada:false});
  s.el('rxEmailTrecho').value='Trecho selecionado livremente pelo usuário';c.rxAtualizarAceite();
  assert.equal(s.el('rxEmailValidar').disabled,false);
  c.rxFinalizarCusto();
  assert.equal(s.request().method,'confirmarCustoComEmailReentradas');
  assert.deepEqual(s.request().args.slice(0,2),[calc.id,calc.atualizadoEm]);
  assert.equal(s.request().args[2].confirmacaoExplicita,true);
  assert.equal(s.request().args[2].texto,'Trecho selecionado livremente pelo usuário');
  assert.equal(s.el('rxCustoValidado').disabled,true);
  const url='https://drive.google.com/uc?export=download&id='+calc.id;
  s.request().success({sucesso:true,id:calc.id,chamado:'3127',atualizadoEm:calc.atualizadoEm,
    validadoEm:1730000000000,downloadUrl:url});
  assert.equal(s.el('rxBaixarValidado').hidden,false);
  assert.equal(s.el('rxBaixarValidado').href,url);
  assert.match(s.el('rxValidacaoStatus').textContent,/Validado em/);
  c.rxPendente={'0:5:2':{aba:0,linha:5,coluna:2,valor:'corrigido'}};c.rxSalvarEdicoes();
  s.request().success({sucesso:true,id:calc.id,atualizadoEm:1730000001000});
  assert.equal(s.el('rxBaixarValidado').hidden,true);
});

test('a validação abre expandida e Recalcular atualiza o total imediatamente sem salvar nem navegar',()=>{
  const s=screen(),c=s.context,calc={...file,calculadora:true};
  c.rxAbrirLista();assert.equal(s.el('rxOverlay').classList.contains('rx-wide'),true);
  const resumo={abas:[{nome:'CALC Geral',totalColunas:10,linhas:[
    {numero:3,valores:{3:'CÓDIGO SKU',5:'AVULSO ACAB',6:'AVULSO GRAF',7:'KIT S/ MALETA',8:'KIT C/ MALETA',9:'TOTAL'}},
    {numero:4,valores:{9:'Exemplares'}},
    {numero:6,valores:{3:'SKU-3106',5:'800',6:'0',7:'0',8:'0',9:'30'}}]}]};
  c.rxAbrirArquivo(calc);s.request().success({arquivo:calc,previaExcel:resumo,formulas:{},
    faixasFormulas:[{aba:0,x1:10,x2:10,y1:6,y2:6}]});
  assert.equal(s.el('rxOverlay').classList.contains('rx-wide'),true);
  assert.equal(s.el('rxExpandir').textContent,'Reduzir');
  const chamadas=s.requests.length;
  c.rxPendente={'0:6:6':{aba:0,linha:6,coluna:6,valor:'1200'}};
  c.rxRecalcular();
  assert.equal(s.requests.length,chamadas,'recalcular nao deve chamar o servidor nem salvar o Excel');
  assert.equal(c.rxCalculados['0:6:10'],'1200');
  assert.equal(s.el('rxPrepararPane').hidden,true);
  assert.match(s.el('rxMsg').textContent,/1 totais/);
  c.rxRecalcular();assert.equal(s.requests.length,chamadas,'recalculo repetido nao pode gerar requisicoes');
  assert.equal(c.rxCalculados['0:6:10'],'1200');
});

test('sugestao de e-mail pode ser ignorada uma unica vez por calculadora',()=>{
  const s=screen(),c=s.context,calc={...file,calculadora:true};
  c.rxAbrirArquivo(calc);s.request().success({arquivo:calc,previaExcel:conjunto,formulas:{}});
  c.rxAbrirConfirmacao();const msg={id:'MENSAGEM123',assunto:'Custo validado 3127',remetente:'grafica@fornecedor.com',data:1730000000000,trecho:'OK, de acordo'};
  s.request().success({sucesso:true,emails:[{id:'THREAD123456789',assunto:'3127',mensagens:[msg]}],
    sugestaoJaTratada:false,sugestao:{mensagem:msg,threadId:'THREAD123456789'}});
  assert.equal(s.el('rxEmailAviso').hidden,false);
  c.rxIgnorarSugestao();assert.equal(s.request().method,'ignorarSugestaoEmailCustoReentradas');
  s.request().success({sucesso:true});assert.equal(s.el('rxEmailAviso').hidden,true);
  c.rxCarregarConfirmacoes();s.request().success({sucesso:true,emails:[{id:'THREAD123456789',mensagens:[msg]}],sugestaoJaTratada:true});
  assert.equal(s.el('rxEmailAviso').hidden,true);
});

test('fórmula com valor cacheado fica protegida pela coordenada sem impedir editar texto',()=>{
  const s=screen(),c=s.context;open(s);
  const row=s.el('rxCorpo').children[0],formula=row.children[1],texto=row.children[2];
  assert.equal(formula.textContent,'20');assert.match(formula.className,/rx-formula/);
  formula.ondblclick();assert.equal(formula.children.length,0);assert.match(s.el('rxMsg').textContent,/Fórmulas/);
  texto.ondblclick();const input=texto.children[0];input.value='alterado';input.onblur();
  assert.deepEqual(plain(c.rxPendente),{'0:5:2':{aba:0,linha:5,coluna:2,valor:'alterado'}});
});

test('Especificacoes apresenta titulo/subtitulo e permite separar Nao editorial em aba virtual persistida',()=>{
  const s=screen(),c=s.context,calc={...file,calculadora:true},
    nomes=['PPG - Especificações'],valores={0:'SKU-3106',1:'Não editorial',2:'205x275'},
    dados={abas:[{nome:nomes[0],totalColunas:3,linhas:[
      {numero:3,valores:{0:'CÓDIGO SKU',1:'CLASSIFICAÇÃO DO MATERIAL',2:'FORMATO'}},
      {numero:4,valores:{2:'mm'}},{numero:5,valores}]}]},
    quadro={indices:[0,1,2],cabecalhoAte:4,ultimaLinha:5,
      topo:[{numero:3,valores:{0:'CÓDIGO SKU',1:'CLASSIFICAÇÃO DO MATERIAL',2:'FORMATO'}},{numero:4,valores:{2:'mm'}}]};
  c.rxAbrirArquivo(calc);s.request().success({arquivo:calc,previaExcel:dados,quadro,naoEditoriais:[],formulas:{}});
  const headers=s.el('rxCab').children.at(-1).children;
  assert.equal(headers[1].textContent,'CÓDIGO SKU');
  assert.equal(headers[3].children[0].textContent,'mm');
  const sku=s.el('rxCorpo').children[0].children[1];
  const botao=sku.children[0];assert.equal(botao.textContent,'Separar');botao.onclick({stopPropagation(){}});
  assert.equal(s.request().method,'separarNaoEditorialReentradas');
  s.request().success({sucesso:true,naoEditoriais:[{aba:0,linha:5}]});
  assert.equal(c.rxNomesAbas.at(-1),'Não editorial');
  assert.equal(c.rxPagina.totalFiltradas,1);
  assert.equal(c.rxPagina.linhas[0].numero,5);
  assert.equal(c.rxPagina.aba,1);
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
  assert.equal(s.el('rxPrepararPane').hidden,true);
  assert.equal(s.request().method,'salvarEdicoesReentradasExcel','salvar preserva a revisão sem recarregar tudo');
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

const nomesAbas=['Primeira aba','Segunda aba'];
function sheetResponse(indice,options={}){
  return {arquivo:options.arquivo||file,nomesAbas:nomesAbas,
    abaCompleta:{indice,nome:nomesAbas[indice],totalColunas:options.totalColunas||2,
      linhas:options.linhas||[{numero:5,valores:{0:'20',1:'texto da aba '+indice}}]},
    previaExcel:null,pagina:null,formulas:options.formulas||{},
    faixasFormulas:options.faixasFormulas||[],mesclagens:options.mesclagens||[]};
}
function pageResponse(aba,linha=0){
  const colunas=Array.from({length:65},(_,i)=>{
    let n=i+1,s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s;
  });
  return {arquivo:file,nomesAbas:nomesAbas,abaCompleta:null,previaExcel:null,formulas:{},
    faixasFormulas:[],mesclagens:[],pagina:{arquivo:file.nome,abas:nomesAbas,aba,linha,coluna:0,
      totalLinhas:120,totalFiltradas:120,totalColunas:70,colunas,
      linhas:Array.from({length:50},(_,i)=>({numero:linha+i+1,
        celulas:Array.from({length:65},(_,c)=>'L'+(linha+i+1)+'C'+(c+1))}))}};
}

test('abas carregadas voltam sem consulta e preservam edição, fórmulas e mesclagens por coordenada',()=>{
  const s=screen(),c=s.context;c.rxAbrirArquivo(file);
  s.request().success(sheetResponse(0,{formulas:{'0:5:1':true},
    mesclagens:[{aba:0,x1:1,x2:2,y1:8,y2:8}]}));
  const texto=s.el('rxCorpo').children[0].children[2];texto.ondblclick();
  texto.children[0].value='edição mantida';texto.children[0].onblur();
  c.rxSelecionarAba(1);const second=s.request();
  assert.equal(second.args[0].aba,1);second.success(sheetResponse(1,{
    faixasFormulas:[{aba:1,x1:2,x2:2,y1:5,y2:8}]}));
  const requests=s.requests.length;c.rxSelecionarAba(0);
  assert.equal(s.requests.length,requests);assert.equal(c.rxPagina.aba,0);
  assert.equal(s.el('rxCorpo').children[0].children[2].textContent,'edição mantida');
  assert.equal(c.rxEhFormula(0,5,1),true);assert.equal(c.rxEhFormula(0,5,2),false);
  assert.equal(c.rxEhFormula(1,5,2),true);
  assert.equal(c.rxMesclagemCoberta(0,8,2).x1,1);assert.equal(c.rxMesclagemCoberta(1,8,2),null);
  c.rxSelecionarAba(1);assert.equal(s.requests.length,requests);
});

test('voltar a uma aba restaura sua busca e posição sem depender de consulta remota',()=>{
  const s=screen(),c=s.context;c.rxAbrirArquivo(file);
  const linhas=Array.from({length:120},(_,i)=>({numero:i+1,valores:{0:'linha '+(i+1),12:'campo '+(i+1)}}));
  s.request().success(sheetResponse(0,{linhas,totalColunas:24}));
  c.rxEstado.linha=50;c.rxEstado.coluna=12;c.rxEstado.busca='linha';c.rxPedirPagina();
  c.rxSelecionarAba(1);s.request().success(sheetResponse(1));
  const requests=s.requests.length;c.rxSelecionarAba(0);
  assert.equal(s.requests.length,requests);
  assert.deepEqual(plain(c.rxEstado),{aba:0,linha:50,coluna:12,busca:'linha'});
  assert.equal(s.el('rxBuscarCelulas').value,'linha');assert.equal(c.rxPagina.linha,50);
  assert.equal(c.rxPagina.coluna,0,'a tabela usa rolamento horizontal contínuo');
});

test('arquivo grande usa rolagem horizontal contínua e reutiliza páginas de linhas',()=>{
  const s=screen(),c=s.context;c.rxAbrirArquivo(file);s.request().success(pageResponse(0));
  const firstCount=s.requests.length;
  c.rxMoverColuna(1);assert.equal(c.rxPagina.coluna,0);
  assert.equal(s.requests.length,firstCount);assert.equal(c.rxPagina.colunas.length,65);
  assert(s.el('rxGrid').scrollLeft>0);
  c.rxMoverColuna(-1);assert.equal(s.el('rxGrid').scrollLeft,0);assert.equal(s.requests.length,firstCount);
  c.rxMoverLinha(1);assert.equal(s.requests.length,firstCount+1);
  assert.equal(s.request().args[0].linha,50);s.request().success(pageResponse(0,50));
  c.rxMoverLinha(-1);assert.equal(s.requests.length,firstCount+1);
  assert.equal(c.rxPagina.linha,0);assert.equal(c.rxPagina.linhas[0].numero,1);
});

test('trocas rápidas reutilizam leitura em andamento e resposta atrasada não troca a aba atual',()=>{
  const s=screen(),c=s.context;c.rxAbrirArquivo(file);s.request().success(sheetResponse(0));
  c.rxSelecionarAba(1);const delayed=s.request(),count=s.requests.length;
  c.rxPedirPagina();assert.equal(s.requests.length,count,'a mesma aba não deve disparar leitura duplicada');
  c.rxSelecionarAba(0);delayed.success(sheetResponse(1));
  assert.equal(c.rxEstado.aba,0);assert.equal(c.rxPagina.aba,0);
  c.rxSelecionarAba(1);assert.equal(s.requests.length,count,'resposta atrasada válida fica disponível no cache');
  assert.equal(c.rxPagina.aba,1);
});

test('salvamento preserva a aba atual, mas invalida respostas antigas de outras abas',()=>{
  const s=screen(),c=s.context;c.rxAbrirArquivo(file);s.request().success(sheetResponse(0));
  c.rxSelecionarAba(1);const oldSheet=s.request();c.rxSelecionarAba(0);
  c.rxPendente={'0:5:2':{aba:0,linha:5,coluna:2,valor:'novo'}};
  c.rxSalvarEdicoes();s.request().success({sucesso:true,id:file.id,atualizadoEm:2000});
  assert.equal(c.rxEstado.aba,0);assert.equal(s.el('rxPrepararPane').hidden,true);
  oldSheet.success(sheetResponse(1));
  const count=s.requests.length;c.rxSelecionarAba(1);
  assert.equal(s.requests.length,count+1,'a aba nao carregada precisa ser lida na versao nova');
  assert.equal(s.request().args[0].atualizadoEm,2000);
  const previousFileRequest=s.request(),other={...file,id:'OUTRO_ARQUIVO_1234567890',atualizadoEm:3000};
  c.rxAbrirArquivo(other);const otherRequest=s.request();
  previousFileRequest.success(sheetResponse(1));
  assert.equal(c.rxArquivo.id,other.id);assert.equal(c.rxPagina,null);
  otherRequest.success(sheetResponse(0,{arquivo:other}));
  assert.equal(c.rxArquivo.id,other.id);assert.equal(c.rxPagina.aba,0);
});

for(const codigo of ['REVISAO_AUSENTE','REVISAO_DESATUALIZADA']){
  test(codigo+': preparar revisão exige clique e retoma leitura sem salvar o XLSX',()=>{
    const s=screen(),c=s.context;c.rxAbrirArquivo(file);
    s.request().success({erro:'Prepare a revisão deste arquivo.',codigo});
    assert.equal(s.el('rxPrepararPane').hidden,false);
    assert.equal(s.requests.length,1,'a falta de revisão não deve abrir ou preparar o Excel automaticamente');
    assert.equal(c.rxArquivo.id,file.id);assert.equal(c.rxTemEdicoes(),false);
    c.rxPrepararRevisao();const prepare=s.request();
    assert.equal(prepare.method,'prepararRevisaoReentradasExcel');
    assert.deepEqual(prepare.args,[file.id,file.atualizadoEm]);assert.equal(c.rxOcupado,true);
    assert.equal(s.el('rxPrepararBtn').disabled,true);
    c.rxPrepararRevisao();assert.equal(s.requests.length,2,'clique repetido não duplica preparação');
    prepare.success({sucesso:true,id:file.id,atualizadoEm:file.atualizadoEm});
    assert.equal(c.rxOcupado,false);assert.equal(s.el('rxPrepararBtn').disabled,false);
    assert.equal(s.el('rxPrepararPane').hidden,true);
    assert.equal(s.request().method,'obterPreviaReentradasExcel');
    assert.equal(s.request().args[0].id,file.id);
    assert.equal(s.request().args[0].atualizadoEm,file.atualizadoEm);
    s.request().success(sheetResponse(0));assert.equal(c.rxPagina.aba,0);
    assert.equal(s.requests.some(r=>r.method==='salvarEdicoesReentradasExcel'),false);
  });
}

test('clique simples e Enter respeitam fórmulas e mesclagens e permitem editar campo livre pelo teclado',()=>{
  const s=screen(),c=s.context;c.rxAbrirArquivo(file);
  s.request().success(sheetResponse(0,{totalColunas:3,
    linhas:[{numero:5,valores:{0:'20',1:'mesclada',2:'campo livre'}}],formulas:{'0:5:1':true},
    mesclagens:[{aba:0,x1:1,x2:2,y1:5,y2:5}]}));
  const row=s.el('rxCorpo').children[0],formula=row.children[1],covered=row.children[2],free=row.children[3];
  let prevented=0;const enter={key:'Enter',preventDefault(){prevented++}};
  formula.onclick();assert.equal(formula.children.length,0);assert.match(s.el('rxMsg').textContent,/Fórmulas/);
  formula.onkeydown(enter);assert.equal(formula.children.length,0);
  covered.onclick();assert.equal(covered.children.length,0);assert.match(s.el('rxMsg').textContent,/Edite A5/);
  covered.onkeydown(enter);assert.equal(covered.children.length,0);assert.equal(c.rxTemEdicoes(),false);
  free.onkeydown(enter);assert.equal(free.children[0].tag,'input');
  free.children[0].value='valor novo';free.children[0].onkeydown(enter);
  assert.equal(c.rxPendente['0:5:3'].valor,'valor novo');assert.equal(prevented,4);
});

test('Excel salvo com revisão pendente conserva os dados atuais sem abrir tela vazia',()=>{
  const s=screen(),c=s.context;c.rxAbrirArquivo(file);s.request().success(sheetResponse(0));
  c.rxPendente={'0:5:2':{aba:0,linha:5,coluna:2,valor:'salvo'}};
  c.rxSalvarEdicoes();const save=s.request(),count=s.requests.length;
  save.success({sucesso:true,id:file.id,atualizadoEm:2000,revisaoPendente:true,
    aviso:'O Excel foi salvo e conferido. Prepare a revisão novamente para continuar.'});
  assert.equal(s.requests.length,count,'revisão pendente deve aguardar preparação explícita');
  assert.equal(c.rxTemEdicoes(),false);assert.equal(c.rxArquivo.atualizadoEm,2000);
  assert.equal(c.rxOcupado,false);assert.equal(s.el('rxSalvar').disabled,true);
  assert.equal(s.el('rxPrepararPane').hidden,true);
  assert.match(s.el('rxMsg').textContent,/Alterações salvas/);
  assert.equal(c.rxPagina.aba,0);
});

test('busca cancelada ao fechar não reinicia a posição de uma nova sessão do arquivo',()=>{
  const s=screen(),c=s.context;c.rxAbrirArquivo(file);s.request().success(sheetResponse(0));
  s.el('rxBuscarCelulas').value='busca cancelada';c.rxAgendarBusca();c.rxFechar();
  c.rxAbrirArquivo(file);
  const linhas=Array.from({length:120},(_,i)=>({numero:i+1,valores:{0:'linha '+(i+1)}}));
  s.request().success(sheetResponse(0,{linhas}));c.rxMoverLinha(1);
  assert.equal(c.rxEstado.linha,50);c.rxSelecionarAba(1);s.request().success(sheetResponse(1));
  c.rxSelecionarAba(0);assert.equal(c.rxEstado.linha,50);
  assert.equal(c.rxEstado.busca,'');assert.equal(s.el('rxBuscarCelulas').value,'');
});

test('a revisão apresenta validação de custos e não oferece navegação paginada de colunas',()=>{
  assert.match(html,/Validar custos/);
  assert.doesNotMatch(html,/Conferir e editar Excel/);
  assert.match(html,/rx-column-controls[^>]*hidden/);
  const s=screen(),c=s.context;open(s);
  assert.equal(c.rxPagina.colunas.length,2);
  assert.equal(s.el('rxColAnterior').disabled,true);
  assert.equal(s.el('rxColProxima').disabled,true);
});
test('largura completa conserva coordenadas originais e destaca referências discretas',()=>{
  const s=screen(),c=s.context;c.rxAbrirArquivo(file);
  s.request().success(sheetResponse(0,{totalColunas:65,
    linhas:[{numero:5,valores:{2:'SKU-001',3:'83.85',64:'FINAL'}}],
    formulas:{'0:5:4':true}}));
  assert.equal(c.rxPagina.colunas.length,65);
  const linha=s.el('rxCorpo').children[0];
  assert.equal(linha.children[3].textContent,'SKU-001');
  assert.equal(linha.children[4].textContent,'83,85');
  assert.equal(linha.children[65].textContent,'FINAL');
  linha.children[4].ondblclick();
  assert.match(s.el('rxMsg').textContent,/Fórmulas/);
  linha.children[3].ondblclick();
  const input=linha.children[3].children[0];input.value='SKU-002';input.onblur();
  assert.equal(c.rxPendente['0:5:3'].valor,'SKU-002');
});

test('lista de gráficas exibe rótulos no menu nativo inclusive com tema escuro',()=>{
  assert.match(html, /#selectGraficas[{]color-scheme:light/);
  assert.match(html, /#selectGraficas option[{][^}]*color:#1f2a44/);
  const s=screen(),c=s.context;
  c.preencherPastasGraficasApp([{id:'CASA',nome:'Casa da Gráfica'},{id:'LOGPRINT',nome:'Logprint'}]);
  const options=s.el('selectGraficas').children;
  assert.equal(options.length,2);
  assert.equal(options[0].textContent,'Casa da Gráfica');
  assert.equal(options[0].label,'Casa da Gráfica');
  assert.equal(options[1].textContent,'Logprint');
  assert.equal(s.el('selectGraficas').disabled,false);
});

test('seletor próprio mostra os nomes e mantém o ID esperado pela validação',()=>{
  const s=screen(),c=s.context;
  c.preencherPastasGraficasApp([{id:'CASA',nome:'Casa da Gráfica'},
    {id:'LOGPRINT',nome:'Logprint'},{id:'FELK',nome:'Felk'}]);
  const menu=s.el('grPickerMenu'),botao=s.el('grPickerButton');
  assert.equal(menu.children.length,3);
  assert.deepEqual(menu.children.map(o=>o.textContent),['Casa da Gráfica','Logprint','Felk']);
  assert.equal(botao.disabled,false);
  c.alternarPickerGraficas();assert.equal(menu.hidden,false);
  assert.equal(botao['aria-expanded'],'true');
  menu.children[2].onclick();
  assert.equal(s.el('selectGraficas').value,'FELK');
  assert.equal(botao.textContent,'Felk');
  assert.equal(menu.hidden,true);assert.equal(botao['aria-expanded'],'false');
  assert.equal(s.el('btnAvancar').disabled,false);
});

test('seletor próprio respeita bloqueio e oferece navegação por teclado',()=>{
  const s=screen(),c=s.context;
  c.preencherPastasGraficasApp([{id:'CASA',nome:'Casa da Gráfica'},{id:'LOGPRINT',nome:'Logprint'}]);
  const menu=s.el('grPickerMenu'),botao=s.el('grPickerButton');
  c.alternarPickerGraficas();assert.equal(menu.hidden,false);
  let prevented=0;menu.children[0].onkeydown({key:'ArrowDown',preventDefault(){prevented++;}});
  menu.children[1].onkeydown({key:'Escape',preventDefault(){prevented++;}});
  assert.equal(prevented,2);assert.equal(menu.hidden,true);
  c.definirBloqueioGraficas(true);
  assert.equal(botao.disabled,true);assert.equal(s.el('selectGraficas').disabled,true);
  c.alternarPickerGraficas();assert.equal(menu.hidden,true);
  c.definirBloqueioGraficas(false);assert.equal(botao.disabled,false);
  c.preencherPastasGraficasApp([]);
  assert.equal(botao.disabled,true);assert.match(botao.textContent,/Nenhuma gráfica/);
});
