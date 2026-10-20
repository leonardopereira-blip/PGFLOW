// Origem temporaria e somente leitura para chamados da solicitacao antiga.
// Invocada apenas apos confirmacao explicita no fluxo de reentradas.
// Ao encerrar o legado, remover este modulo e a selecao de origem no fluxo.
var RE_LEGADO_ALOCACAO_CONFIG = {
  idPlanilha: '1pR9a5-oogD7ZYHxImsE9ABNpHeVyCJaTcIab7_U_bI0',
  abas: ['Especificação Técnica', 'Árvore de produto', 'Tiragem'],
  colunaChamado: 'ID (Coluna A)',
  colunaLink: 'PPG utilizada (Último Link)',
  colunaVersao: 'Versão da PPG'
};

function _reLegIndice_(headers,nome,aba) {
  var indices=[];
  headers.forEach(function(h,i){if(_reSpNormal_(h)===_reSpNormal_(nome))indices.push(i);});
  if(indices.length!==1)throw Error('Alocação antiga: coluna '+nome+' ausente ou duplicada em '+aba+'.');
  return indices[0];
}

function _reLegValor_(valor,fuso) {
  if(Object.prototype.toString.call(valor)!=='[object Date]')return valor;
  if(!isFinite(valor.getTime()))throw Error('Alocação antiga: data inválida.');
  // getValues devolve Date para qualquer celula com formato de data, inclusive
  // numeros como 20 paginas. Recuperar o serial original pelo horario local da
  // planilha, sem o deslocamento UTC e sem transformar numeros/textos comuns.
  var partes=Utilities.formatDate(valor,fuso,"yyyy-MM-dd'T'HH:mm:ss.SSS")
    .match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})\.(\d{3})$/);
  if(!partes)throw Error('Alocação antiga: não foi possível converter a data.');
  return (Date.UTC(Number(partes[1]),Number(partes[2])-1,Number(partes[3]),
    Number(partes[4]),Number(partes[5]),Number(partes[6]),Number(partes[7]))-
    Date.UTC(1899,11,30))/86400000;
}

function _reLegLerChamado_(ss,nome,numero) {
  var aba=ss.getSheetByName(nome);
  if(!aba)throw Error('Alocação antiga: aba '+nome+' não encontrada.');
  var ultimaLinha=aba.getLastRow(),ultimaColuna=aba.getLastColumn();
  if(!ultimaLinha||!ultimaColuna)throw Error('Alocação antiga: aba '+nome+' vazia.');
  var headers=aba.getRange(1,1,1,ultimaColuna).getValues()[0].map(function(h){return String(h==null?'':h).trim();});
  var config=RE_LEGADO_ALOCACAO_CONFIG;
  var id=_reLegIndice_(headers,config.colunaChamado,nome);
  var link=_reLegIndice_(headers,config.colunaLink,nome);
  var versao=_reLegIndice_(headers,config.colunaVersao,nome);
  var encontrados=ultimaLinha<2?[]:aba.getRange(2,id+1,ultimaLinha-1,1)
    .createTextFinder(numero).matchEntireCell(true).useRegularExpression(false).findAll();
  // A pesquisa e restrita a coluna ID. Ler apenas os blocos de linhas deste
  // chamado, evitando carregar dezenas de milhares de linhas das tres abas.
  var fuso=ss.getSpreadsheetTimeZone();
  var linhas=[],numeros=encontrados.map(function(r){return r.getRow();})
    .sort(function(a,b){return a-b;}).filter(function(n,i,a){return i===0||n!==a[i-1];});
  for(var p=0;p<numeros.length;){
    var inicio=numeros[p],fim=inicio;
    while(p+1<numeros.length&&numeros[p+1]===fim+1){p++;fim++;}
    var bloco=aba.getRange(inicio,1,fim-inicio+1,ultimaColuna).getValues();
    bloco.forEach(function(r){
      r=r.map(function(valor){return _reLegValor_(valor,fuso);});
      if(String(r[id]==null?'':r[id]).trim()!==numero)throw Error('Alocação antiga: chamado inconsistente em '+nome+'.');
      linhas.push(r);
    });
    p++;
  }
  return {nome:nome,headers:headers,linhas:linhas,indiceLink:link,indiceVersao:versao};
}

function _reLegVersao_(valor,nome) {
  var s=String(valor==null?'':valor).trim(),n=Number(s);
  if(!/^\d+$/.test(s)||!isFinite(n)||n<0||Math.floor(n)!==n) {
    throw Error('Alocação antiga: versão da PPG inválida em '+nome+'.');
  }
  return n;
}

function _reLegUltimaVersao_(matrizes,numero) {
  var maxima=-1,linkUnico='';
  matrizes.forEach(function(m){
    m.linhas.forEach(function(r){maxima=Math.max(maxima,_reLegVersao_(r[m.indiceVersao],m.nome));});
  });
  if(maxima<0)throw Error('Chamado '+numero+' não encontrado na alocação antiga.');
  var filtradas=matrizes.map(function(m){
    var linhas=m.linhas.filter(function(r){return _reLegVersao_(r[m.indiceVersao],m.nome)===maxima;});
    if(!linhas.length)throw Error('Alocação antiga: versão '+maxima+' incompleta para o chamado '+numero+' (sem '+m.nome+').');
    linhas.forEach(function(r){
      var link=String(r[m.indiceLink]==null?'':r[m.indiceLink]).trim();
      if(!link)throw Error('Alocação antiga: versão '+maxima+' sem link da PPG em '+m.nome+'.');
      if(linkUnico&&link!==linkUnico)throw Error('Alocação antiga: links da PPG divergentes na versão '+maxima+' do chamado '+numero+'.');
      linkUnico=link;
    });
    return {headers:m.headers,linhas:linhas};
  });
  // Nunca recuar para uma versao anterior se a mais recente estiver parcial.
  return {versao:maxima,link:linkUnico,matrizes:filtradas};
}

function _reLegSemCodigo_(valor) {
  var s=String(valor==null?'':valor).trim();
  return s===''||s==='-';
}

function _reSpPrepararDadosLegado_(numero,grafica) {
  numero=String(numero==null?'':numero).trim();
  if(!/^\d{2,8}$/.test(numero))throw Error('Número do chamado inválido.');
  var config=RE_LEGADO_ALOCACAO_CONFIG;
  var ss=SpreadsheetApp.openById(config.idPlanilha);
  var lidas=config.abas.map(function(nome){return _reLegLerChamado_(ss,nome,numero);});
  var origem=_reLegUltimaVersao_(lidas,numero),e=origem.matrizes[0],a=origem.matrizes[1],t=origem.matrizes[2];
  [
    [e,['1. ID_Código SKU']],
    [a,['1. ID_Código Kit','1. ID_Código SKU']],
    [t,['1. ID_Código Kit','1. ID_Código SKU','15. DES_CD envio','16. TIR_SKU']]
  ].forEach(function(item,i){
    item[1].forEach(function(nome){_reLegIndice_(item[0].headers,nome,config.abas[i]);});
  });
  // A grafica seleciona apenas o molde de custos, como no caminho da API.
  // A OP inteira e preservada, inclusive antes da alocacao de uma grafica.
  var skus=Object.create(null),kits=Object.create(null),pares=Object.create(null),destinos=Object.create(null);
  var linhasSku=t.linhas.filter(function(r){return !_reLegSemCodigo_(_reSpBuscar_(r,t.headers,'1. ID_Código SKU'));});
  if(!linhasSku.length)throw Error('Tiragem da alocação antiga sem linhas de SKU.');
  linhasSku.forEach(function(r){
    var sku=_reSpBuscar_(r,t.headers,'1. ID_Código SKU'),kit=_reSpBuscar_(r,t.headers,'1. ID_Código Kit');
    var cd=_reSpBuscar_(r,t.headers,'15. DES_CD envio');
    if(_reLegSemCodigo_(kit)||_reLegSemCodigo_(cd))throw Error('Tiragem da alocação antiga sem SKU, kit ou CD de destino.');
    _reSpQuant_(_reSpBuscar_(r,t.headers,'16. TIR_SKU'));
    skus[sku]=true;kits[kit]=true;pares[JSON.stringify([kit,sku])]=true;
    destinos[JSON.stringify([kit,cd])]=true;
  });
  // O legado repete a tiragem no pai do kit (SKU '-' ou vazio) e nos filhos.
  // O pai e estrutural: nao e outro SKU e nao pode duplicar os exemplares.
  t.linhas.forEach(function(r){
    if(!_reLegSemCodigo_(_reSpBuscar_(r,t.headers,'1. ID_Código SKU')))return;
    var kit=_reSpBuscar_(r,t.headers,'1. ID_Código Kit'),cd=_reSpBuscar_(r,t.headers,'15. DES_CD envio');
    if(!destinos[JSON.stringify([kit,cd])])throw Error('Tiragem da alocação antiga: kit sem SKU no CD de destino: '+kit+' / '+cd);
  });
  var espec=e.linhas.filter(function(r){return !!skus[_reSpBuscar_(r,e.headers,'1. ID_Código SKU')];});
  var especSkus=Object.create(null),arvorePares=Object.create(null);
  espec.forEach(function(r){especSkus[_reSpBuscar_(r,e.headers,'1. ID_Código SKU')]=true;});
  Object.keys(skus).forEach(function(sku){
    if(!especSkus[sku])throw Error('SKU sem especificação na alocação antiga: '+sku);
  });
  var arvore=a.linhas.filter(function(r){
    var kit=_reSpBuscar_(r,a.headers,'1. ID_Código Kit'),sku=_reSpBuscar_(r,a.headers,'1. ID_Código SKU');
    var par=JSON.stringify([kit,sku]),pai=_reLegSemCodigo_(sku);
    if(!pai&&pares[par])arvorePares[par]=true;
    return !!kits[kit]&&(pai||!!pares[par]);
  });
  Object.keys(pares).forEach(function(par){
    if(!arvorePares[par])throw Error('Kit/SKU sem árvore na alocação antiga: '+JSON.parse(par).join(' / '));
  });
  return {
    pacote:{id:numero,versaoId:'legado:'+origem.versao,origem:'legado',
      versaoPpg:origem.versao,linkPpg:origem.link,fontePlanilha:config.idPlanilha},
    espec:{headers:e.headers,linhas:espec},
    arvore:{headers:a.headers,linhas:arvore},
    tiragem:{headers:t.headers,linhas:linhasSku}
  };
}
