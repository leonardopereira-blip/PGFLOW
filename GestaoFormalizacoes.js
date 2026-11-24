// Painel somente leitura. Nao participa do disparo nem da validacao de custos.
function _gfEmailEndereco_(s) {
  var m=String(s||'').toLowerCase().match(/[a-z0-9.!#$%&'*+/=?^_\x60{|}~-]+@[a-z0-9.-]+\.[a-z]{2,}/g);
  return m || [];
}
function _gfChave_(s) {
  return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
}
function _gfGraficaTeste_(s) {return /^teste(?:[\s_-]*\d+)?$/.test(_gfChave_(s));}
function _gfMapaDestinos_(catalogo) {
  var mapa=Object.create(null);
  Object.keys(catalogo||{}).forEach(function(nome) {
    if (_gfGraficaTeste_(nome)) return;
    (catalogo[nome]||[]).forEach(function(raw) {
      _gfEmailEndereco_(raw).forEach(function(email) {
        if (/@arcoeducacao\.com\.br$/.test(email)) return;
        if (!mapa[email]) mapa[email]=[];
        if (mapa[email].indexOf(nome)===-1) mapa[email].push(nome);
      });
    });
  });
  return mapa;
}
function _gfNumerosSPG_() {
  var cache=CacheService.getScriptCache(), chave='gf_chamados_v1';
  try {var salvo=cache.get(chave);if(salvo)return JSON.parse(salvo);}catch(e){}
  var token=_spgLogin_();
  if(!token)throw Error('Autenticacao do SPG indisponivel.');
  var chamados={}, paginas=1;
  for(var pagina=1;pagina<=paginas;pagina++){
    var dados=_emailSpgLerJson_('/api/reentradas/sheets/relatorio?page='+pagina+'&limit=100',token);
    paginas=_reRecentesPagina_(dados,pagina,paginas);
    _reRecentesAgrupar_(dados.rows,chamados);
  }
  // A lista inclui SPOT para nao excluir ampliados com justificativa ausente do relatorio.
  // Nao interpreta a existencia no SPG como evidência de formalização.
  var lista=Object.keys(chamados).filter(function(n){
    var c=chamados[n];return !c.tipo || /^(reentrada|inclusivos?|spot)$/.test(c.tipo);
  }).sort(function(a,b){return Number(b)-Number(a);});
  try{cache.put(chave,JSON.stringify(lista),180);}catch(e){}
  return lista;
}
function _gfLerMensagem_(msg, ids, remetentes, mapa) {
  if(typeof msg.isDraft==='function' && msg.isDraft())return null;
  var from=_gfEmailEndereco_(msg.getFrom());
  if(from.length!==1 || remetentes.indexOf(from[0])===-1)return null;
  var assunto=_gfChave_(msg.getSubject());
  // A evidencia precisa do assunto padrao e do anexo do mesmo chamado.
  // Evita confundir respostas, reencaminhamentos ou anexos da validacao de custo.
  var match=/^solicitacao de prazo:\s*chamado\s+(\d{1,12})(?:\s*-\s*.*)?$/.exec(assunto);
  if(!match || !ids[match[1]])return null;
  var id=match[1];
  var anexos=msg.getAttachments({includeInlineImages:false,includeAttachments:true});
  if(!anexos.some(function(a){return _gfChave_(a.getName())==='chamado '+id+'.xlsx';}))return null;
  var destinatarios=_gfEmailEndereco_(msg.getTo()).filter(function(e){return !/@arcoeducacao\.com\.br$/.test(e);});
  // Envio interno/teste nao conta como formalizacao para grafica.
  if(!destinatarios.length)return null;
  var nomes=[], ambiguo=false;
  destinatarios.forEach(function(email){
    var correspondencias=mapa[email]||[];
    if(correspondencias.length===1 && nomes.indexOf(correspondencias[0])===-1)nomes.push(correspondencias[0]);
    if(correspondencias.length!==1)ambiguo=true;
  });
  return {id:id,data:msg.getDate().getTime(),graficas:nomes,ambiguo:ambiguo};
}
function _gfConsultarEnviados_(numeros,catalogo) {
  var ids=Object.create(null),historico=Object.create(null),mapa=_gfMapaDestinos_(catalogo);
  numeros.forEach(function(n){ids[n]=true;});
  var remetentes=_formalizacaoRemetentes_().map(function(r){return r.email;});
  var executor=_formalizacaoNormalizarEmail_(Session.getEffectiveUser().getEmail());
  if(executor && remetentes.indexOf(executor)===-1)remetentes.push(executor);
  var vistos=Object.create(null),completo=true;
  for(var i=0;i<numeros.length;i+=10){
    var lote=numeros.slice(i,i+10);
    var query='in:sent {'+lote.map(function(n){return 'filename:"Chamado '+n+'.xlsx"';}).join(' ')+'}';
    // Paginacao limitada para manter o painel leve; nao declara ausencia em leitura truncada.
    for(var start=0;start<500;start+=100){
      var threads=GmailApp.search(query,start,100);
      var mensagens=GmailApp.getMessagesForThreads(threads);
      mensagens.forEach(function(grupo){grupo.forEach(function(msg){
        var idMsg=msg.getId();if(vistos[idMsg])return;vistos[idMsg]=true;
        var e=_gfLerMensagem_(msg,ids,remetentes,mapa);
        if(!e)return;
        var item=historico[e.id]||(historico[e.id]={data:0,graficas:Object.create(null),ambiguo:false});
        item.data=Math.max(item.data,e.data);item.ambiguo=item.ambiguo||e.ambiguo;
        e.graficas.forEach(function(g){item.graficas[g]=Math.max(item.graficas[g]||0,e.data);});
      });});
      if(threads.length<100)break;
      if(start===400)completo=false;
    }
  }
  return {historico:historico,completo:completo};
}
function obterGestaoFormalizacoes(pagina,busca) {
  try {
    var ativo=_formalizacaoNormalizarEmail_(Session.getActiveUser().getEmail());
    if(!/@arcoeducacao\.com\.br$/.test(ativo))throw Error('Conta corporativa nao identificada.');
    pagina=Number(pagina)||1;
    if(pagina<1 || pagina%1!==0)throw Error('Pagina invalida.');
    busca=String(busca||'').trim();
    if(busca && !/^\d{1,12}$/.test(busca))throw Error('Informe apenas o numero do chamado.');
    var lista=_gfNumerosSPG_();
    if(busca)lista=lista.filter(function(n){return n===busca;});
    var porPagina=30,total=lista.length,paginas=Math.max(1,Math.ceil(total/porPagina));
    pagina=Math.min(pagina,paginas);
    var numeros=lista.slice((pagina-1)*porPagina,pagina*porPagina);
    var leitura;
    try {leitura=_gfConsultarEnviados_(numeros,obterDadosGraficasEmailSPG());}
    catch(e){leitura={historico:{},completo:false,erro:'Nao foi possivel consultar os e-mails enviados.'};}
    function fmt(ts){return ts?Utilities.formatDate(new Date(ts),'America/Sao_Paulo','dd/MM/yyyy HH:mm'):'';}
    return {pagina:pagina,paginas:paginas,total:total,completo:leitura.completo,erroHistorico:leitura.erro||'',
      chamados:numeros.map(function(n){
        var h=leitura.historico[n],nomes=h?Object.keys(h.graficas).sort():[];
        return {chamado:n,status:nomes.length?'formalizado':h?'grafica_nao_identificada':leitura.completo?'nao_localizado':'indeterminado',
          dataEnvio:h?fmt(h.data):'',graficas:nomes.map(function(nome){return {nome:nome,data:fmt(h.graficas[nome])};}),
          parcial:!!(h&&h.ambiguo)};
      })};
  }catch(e){return {erro:String(e.message||e),chamados:[]};}
}
