// Gerador nativo XLSX - somente VALIDACAO DE CUSTOS REENTRADAS.
// Sem fila de e-mail, sem alteracao na Validacao Principal.
// Calculadora e De-Para sao criados do zero a partir dos moldes mapeados
// (MoldeXlsx.js + Molde*.js): formulas, tabelas, estilos e abas do modelo,
// sem ler planilhas no Drive.
var RE_SP_CONFIG = {
  pastaDestino: '1DKie1AHE7EFHa6roJY2XxE2ZlFhC_lCy',
  senhaProtecao: 'PGFLOW2026',
  limiteSkus: 300
};
// Moldes por grafica. Nova grafica: mapear o Excel com tools/mapear-molde.cjs e registrar aqui.
var RE_SP_MOLDES = {
  LOGPRINT: {
    nome: 'Logprint',
    calc: function(){return _moldeLogprintCalc_();},
    depara: function(){return _moldeLogprintDepara_();}
  },
  CASADAGRAFICA: {
    nome: 'Casa da Gráfica',
    calc: function(){return _moldeCasaDaGraficaCalc_();},
    depara: function(){return _moldeLogprintDepara_();}
  },
  FORMACERTA: {
    nome: 'Forma Certa',
    calc: function(){return _moldeFormaCertaCalc_();},
    depara: function(){return _moldeLogprintDepara_();}
  },
  META: {
    nome: 'Meta',
    calc: function(){return _moldeMetaCalc_();},
    depara: function(){return _moldeLogprintDepara_();}
  },
  REPROSET: {
    nome: 'Reproset',
    calc: function(){return _moldeReprosetCalc_();},
    depara: function(){return _moldeLogprintDepara_();}
  }
};
var _reSpMoldesCache_ = {};
function _reSpMolde_(grafica,tipo) {
  var g=RE_SP_MOLDES[_reSpNormal_(grafica)];
  if(!g||!g[tipo])throw Error('Sem molde '+tipo+' para a grafica '+grafica);
  var chave=_reSpNormal_(grafica)+'|'+tipo;
  if(!_reSpMoldesCache_[chave])_reSpMoldesCache_[chave]=g[tipo]();
  return _mxPartes_(_reSpMoldesCache_[chave]);
}

function _reSpEscape_(v) {
  return String(v == null ? '' : v).replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
}
function _reSpUnescape_(v) {
  return String(v || '').replace(/&#x([0-9a-f]+);/gi,function(_,n){return String.fromCharCode(parseInt(n,16));})
    .replace(/&#([0-9]+);/g,function(_,n){return String.fromCharCode(+n);})
    .replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&gt;/g,'>').replace(/&lt;/g,'<').replace(/&amp;/g,'&');
}
function _reSpCol_(n) {
  var out='';
  for (;n>0;n=Math.floor((n-1)/26)) out=String.fromCharCode(65+(n-1)%26)+out;
  return out;
}
function _reSpIdx_(col) {
  var n=0; for(var i=0;i<col.length;i++)n=n*26+col.charCodeAt(i)-64;
  return n-1;
}
function _reSpPartes_(blob) {
  var partes=Utilities.unzip(blob.copyBlob().setContentType('application/zip'));
  var d={};partes.forEach(function(p){d[p.getName()]=p;});
  return {partes:partes,mapa:d};
}
function _reSpStrings_(xml) {
  var out=[],si=/<si(?:\s[^>]*)?>([\s\S]*?)<\/si>/g,m;
  while ((m=si.exec(xml||''))!==null) {
    var frag=m[1],t=/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g,a='',mt;
    while ((mt=t.exec(frag))!==null)a+=_reSpUnescape_(mt[1]);
    out.push(a);
  }
  return out;
}
function _reSpLerAba_(mapa,numero,strings) {
  var p=mapa['xl/worksheets/sheet'+numero+'.xml'];
  if(!p)throw Error('XLSX sem aba '+numero);
  var xml=p.getDataAsString('UTF-8'),rows=[],re=/<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g,m;
  while((m=re.exec(xml))!==null){
    var obj={},cr=/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g,c;
    while((c=cr.exec(m[2]))!==null){
      var adr=/\br="([A-Z]+)\d+"/.exec(c[1]);if(!adr)continue;
      var idx=_reSpIdx_(adr[1]),tipo=/\bt="([^"]+)"/.exec(c[1]);
      var v=/<v>([\s\S]*?)<\/v>/.exec(c[2]||''),inline=/<is(?:\s[^>]*)?>([\s\S]*?)<\/is>/.exec(c[2]||'');
      var val=v?_reSpUnescape_(v[1]):'';
      if(tipo&&tipo[1]==='s')val=strings[+val]||'';
      else if(inline){var inner='',tr=/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g,tm;while((tm=tr.exec(inline[1]))!==null)inner+=_reSpUnescape_(tm[1]);val=inner;}
      else if(v&&(!tipo||tipo[1]==='n')&&/^[+-]?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(val))val=Number(val);
      if(val!=='')obj[idx]=val;
    }
    rows.push({numero:+m[1],cells:obj});
  }
  return rows;
}
function _reSpAbaPorNome_(partes,nome) {
  var w=partes.mapa['xl/workbook.xml'];if(!w)throw Error('XLSX sem workbook.xml');
  var xml=w.getDataAsString('UTF-8'),re=/<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"[^>]*\/>/g,m,id='';
  while((m=re.exec(xml))!==null)if(_reSpUnescape_(m[1])===nome)id=m[2];
  if(!id)throw Error('Aba obrigatoria ausente: '+nome);
  var rel=partes.mapa['xl/_rels/workbook.xml.rels'];
  var x=rel.getDataAsString('UTF-8'),rs=/<Relationship\b[^>]*\/>/g,y;
  while((y=rs.exec(x))!==null){
    var a=y[0],rid=/\bId="([^"]+)"/.exec(a),target=/\bTarget="([^"]+)"/.exec(a);
    if(rid&&target&&rid[1]===id){
      var caminho=target[1];
      return (caminho[0]==='/'?caminho.substr(1):'xl/'+caminho).replace(/\/[^/]+\/\.\.\//g,'/');
    }
  }
  throw Error('Relacionamento da aba ausente: '+nome);
}
function _reSpMatriz_(partes,nome) {
  var caminho=_reSpAbaPorNome_(partes,nome);
  var match=/sheet(\d+)\.xml$/.exec(caminho);
  if(!match)throw Error('Endereco de aba invalido: '+nome);
  var strings=_reSpStrings_(partes.mapa['xl/sharedStrings.xml']?
    partes.mapa['xl/sharedStrings.xml'].getDataAsString('UTF-8'):'');
  var rows=_reSpLerAba_(partes.mapa,+match[1],strings);
  var cab=rows.filter(function(x){return x.numero===1;})[0];
  if(!cab)throw Error(nome+': cabecalho vazio');
  var max=Math.max.apply(null,Object.keys(cab.cells).map(Number))+1;
  var headers=[];for(var i=0;i<max;i++)headers.push(String(cab.cells[i]||'').trim());
  var linhas=rows.filter(function(x){return x.numero>1;}).map(function(x){
    var a=[];for(var i=0;i<max;i++)a[i]=x.cells[i]===undefined?'':x.cells[i];
    return a;
  }).filter(function(a){return a.some(function(v){return v!=='';});});
  return {headers:headers,linhas:linhas};
}
function _reSpNormal_(s){
  return String(s||'').trim().normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/[^A-Za-z0-9]/g,'').toUpperCase();
}
function _reSpNumeroFonte_(valor) {
  if(typeof valor!=='string')return valor;
  var texto=valor.trim().replace(/[\s\u00a0]/g,'');
  if(!texto)return valor;
  if(/^[+-]?\d{1,3}(?:\.\d{3})+(?:,\d+)?$/.test(texto))texto=texto.replace(/\./g,'').replace(',','.');
  else if(/^[+-]?\d+(?:,\d+)$/.test(texto))texto=texto.replace(',','.');
  else if(!/^[+-]?\d+(?:\.\d+)?$/.test(texto))return valor;
  var numero=Number(texto);
  return isFinite(numero)?numero:valor;
}
function _reSpNumeroPorCabecalho_(cabecalho,valor) {
  var c=_reSpNormal_(cabecalho);
  var numerico=/^16TIR(?:SKU|ITEM|MIOLO|CAPAAVULSA|KIT)$/.test(c) ||
    /^3DIM(?:B(?:ABERTO|FECHADO)MM|A(?:ABERTO|FECHADO)MM|EMM|PESOKG|QUANTIDADETOTALDEPAGINAS)$/.test(c) ||
    /^[45678]MIO(?:PRIN|PROF|ENCA|ADES)QUANTIDADEDEPAGINAS$/.test(c) ||
    c==='4CAPORELHA' || c==='11REFQUANTIDADEDEPERSONALIZADOS' ||
    c==='13MANTOTALDEITENSCOLECIONADOS';
  return numerico?_reSpNumeroFonte_(valor):valor;
}
function _reSpMapearCab_(dest,cab,linhas) {
  var mapa={};cab.forEach(function(c,i){if(c)mapa[_reSpNormal_(c)]=i;});
  var idx=dest.map(function(h){var n=_reSpNormal_(h);return mapa[n]===undefined?-1:mapa[n];});
  var rows=linhas.map(function(a){return idx.map(function(i,j){
    return i===-1?'':_reSpNumeroPorCabecalho_(dest[j],a[i]);
  });});
  return rows.filter(function(a){return a[0]||a[1]||a[3];});
}
function _reSpBuscar_(row,headers,nome) {
  var idx=headers.map(_reSpNormal_).indexOf(_reSpNormal_(nome));
  return idx<0?'':(row[idx]===undefined?'':row[idx]);
}
function _reSpQuant_(v) {
  if(v===null||v===undefined||String(v).trim()==='')throw Error('Tiragem vazia');
  var str=String(v).trim().replace(/\s/g,'');
  if(str.indexOf(',')>=0)str=str.replace(/\./g,'').replace(',','.');
  else if(/^\d{1,3}(\.\d{3})+$/.test(str))str=str.replace(/\./g,'');
  var n=Number(str);
  if(!isFinite(n)||n<0)throw Error('Tiragem invalida: '+v);
  return n;
}
function _reSpBool_(v) {
  if(v===true||/^(TRUE|SIM)$/i.test(String(v).trim()))return 'Sim';
  if(v===false||/^(FALSE|NÃO|NAO)$/i.test(String(v).trim()))return 'Não';
  return v===null||v===undefined?'':v;
}
function _reSpCelula_(ref,val,style) {
  if(val===undefined||val===null||val==='')return '';
  var s=style?' s="'+style+'"':'';
  if(typeof val==='number'&&isFinite(val))return '<c r="'+ref+'"'+s+'><v>'+val+'</v></c>';
  return '<c r="'+ref+'"'+s+' t="inlineStr"><is><t>'+_reSpEscape_(val)+'</t></is></c>';
}
function _reSpLinha_(numero,cells,estilos) {
  var cs='';Object.keys(cells).map(Number).sort(function(a,b){return a-b;}).forEach(function(i){
    cs+=_reSpCelula_(_reSpCol_(i+1)+numero,cells[i],estilos&&estilos[i]);
  });
  return '<row r="'+numero+'">'+cs+'</row>';
}
function _reSpEstilos_(xml,linha) {
  var a=new RegExp('<row\\b[^>]*\\br="'+linha+'"[^>]*>([\\s\\S]*?)<\\/row>').exec(xml);
  var mapa={};if(!a)return mapa;var re=/<c\b([^>]*?)(?:\/>|>)/g,m;
  while((m=re.exec(a[1]))!==null){var r=/\br="([A-Z]+)\d+"/.exec(m[1]),s=/\bs="(\d+)"/.exec(m[1]);if(r&&s)mapa[_reSpIdx_(r[1])]=s[1];}
  return mapa;
}
function _reSpTrocarLinhas_(xml,minimo,linhas) {
  var tag=/<sheetData(?:\s[^>]*)?>[\s\S]*?<\/sheetData>/.exec(xml);
  if(!tag)throw Error('sheetData ausente');
  var orig=tag[0],inicio=orig.indexOf('>')+1,fim=orig.lastIndexOf('</sheetData>');
  var antes=orig.substring(inicio,fim),re=/<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>|<row\b[^>]*\br="(\d+)"[^>]*\/>/g,m,fixos='';
  while((m=re.exec(antes))!==null){if(+(m[1]||m[2])<minimo)fixos+=m[0];}
  return xml.replace(tag[0],orig.substring(0,inicio)+fixos+linhas.join('')+'</sheetData>');
}
function _reSpSenhaHash_(password) {
  var hash=0;
  for(var i=password.length-1;i>=0;i--){hash=((hash<<1)|(hash>>14))&32767;hash^=password.charCodeAt(i);}
  hash=((hash<<1)|(hash>>14))&32767;hash^=password.length;hash^=0xCE4B;
  return ('0000'+hash.toString(16).toUpperCase()).slice(-4);
}
function _reSpProteger_(xml) {
  // Etapa FINAL, nunca antes do preenchimento das abas.
  var tag='<sheetProtection password="'+_reSpSenhaHash_(RE_SP_CONFIG.senhaProtecao)+'" sheet="1" objects="1" scenarios="1"/>';
  xml=xml.replace(/<sheetProtection\b[^>]*\/>/g,'').replace(/<sheetProtection\b[^>]*>[\s\S]*?<\/sheetProtection>/g,'');
  var i=xml.indexOf('</sheetData>');if(i<0)throw Error('XML sem sheetData');
  i+='</sheetData>'.length;
  return xml.substring(0,i)+tag+xml.substring(i);
}
function _reSpFullCalc_(xml) {
  var pr='<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>';
  // Ordem XLSX: calcPr antes de pivotCaches e extLst. Nunca anexar apos extLst.
  if(/<calcPr\b[^>]*\/>/.test(xml))return xml.replace(/<calcPr\b[^>]*\/>/,pr);
  return xml.replace(/(?=<pivotCaches\b|<extLst\b|<\/workbook>)/,pr);
}
function _reSpSalvarZip_(arquivo,partes,alteradas,nome) {
  var saida=partes.partes.map(function(p){
    var val=alteradas[p.getName()];
    return val===undefined?p:Utilities.newBlob(val,'text/xml',p.getName());
  });
  var xlsx=Utilities.zip(saida,nome).setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  return DriveApp.getFolderById(RE_SP_CONFIG.pastaDestino).createFile(xlsx).setName(nome);
}
function _reSpPrepararDados_(numero,grafica) {
  if(!RE_SP_MOLDES[_reSpNormal_(grafica)])throw Error('Modelo ainda nao disponivel para esta grafica.');
  var pacote=_emailSpgPacote_(numero,false,'reentrada','reentrada');
  if(pacote.id!==numero||!pacote.versaoId||!pacote.blob)throw Error('Chamado ou versao inconsistentes na API');
  if(/CANCELAD/i.test(pacote.statusChamado+' '+pacote.statusProducao))throw Error('Chamado cancelado');
  var raw=_reSpPartes_(pacote.blob);
  var e=_reSpMatriz_(raw,'Esp. Técnica (legado)');
  var a=_reSpMatriz_(raw,'Árvore (legado)');
  var t=_reSpMatriz_(raw,'Tiragem (legado)');
  if(!e.linhas.length||!a.linhas.length||!t.linhas.length)throw Error('OP do SPG sem as tres matrizes de producao');
  // A grafica escolhida define o modelo de custos para a OP inteira do chamado.
  // O calculo pode anteceder a alocacao; nao consulta nem altera alocacoes no SPG.
  var tGraf=t.linhas,skus=Object.create(null),kits=Object.create(null),pares=Object.create(null);
  tGraf.forEach(function(l){
    var sku=_reSpBuscar_(l,t.headers,'1. ID_Código SKU'),kit=_reSpBuscar_(l,t.headers,'1. ID_Código Kit'),
        cd=_reSpBuscar_(l,t.headers,'15. DES_CD envio');
    if(!sku||!kit||!cd)throw Error('Tiragem do chamado sem SKU, kit ou CD de destino.');
    _reSpQuant_(_reSpBuscar_(l,t.headers,'16. TIR_SKU'));
    skus[sku]=true;kits[kit]=true;pares[JSON.stringify([kit,sku])]=true;
  });
  var espec=e.linhas.filter(function(l){
    return !!skus[_reSpBuscar_(l,e.headers,'1. ID_Código SKU')];
  }),especSkus=Object.create(null),arvorePares=Object.create(null);
  espec.forEach(function(l){especSkus[_reSpBuscar_(l,e.headers,'1. ID_Código SKU')]=true;});
  Object.keys(skus).forEach(function(sku){
    if(!especSkus[sku])throw Error('SKU sem especificacao no SPG: '+sku);
  });
  var arvore=a.linhas.filter(function(l){
    var kit=_reSpBuscar_(l,a.headers,'1. ID_Código Kit'),sku=_reSpBuscar_(l,a.headers,'1. ID_Código SKU'),
        par=JSON.stringify([kit,sku]);
    if(sku&&pares[par])arvorePares[par]=true;
    return !!kits[kit]&&(!sku||!!pares[par]);
  });
  Object.keys(pares).forEach(function(par){
    if(!arvorePares[par])throw Error('Kit/SKU sem arvore no SPG: '+JSON.parse(par).join(' / '));
  });
  return {pacote:pacote,espec:{headers:e.headers,linhas:espec},
    arvore:{headers:a.headers,linhas:arvore},tiragem:{headers:t.headers,linhas:tGraf}};
}
function _reSpMontarBase_(source,headersModelo) {
  return _reSpMapearCab_(headersModelo,source.headers,source.linhas);
}
function _reSpConstrucaoNC_(espec,arvore,tiragem,respostas) {
  var cab=espec.headers,idx={},dados=espec.linhas,ret=[];
  cab.forEach(function(c,i){idx[_reSpNormal_(c)]=i;});
  var man={};(respostas||[]).forEach(function(r){man[Number(r.index)]=r.valor;});
  var arvCab=arvore.headers;
  var arv={};arvore.linhas.forEach(function(r){var kit=_reSpBuscar_(r,arvCab,'1. ID_Código Kit');if(kit&&!arv[kit])arv[kit]=r;});
  var aux={};tiragem.linhas.forEach(function(r){
    var h=tiragem.headers,sku=_reSpBuscar_(r,h,'1. ID_Código SKU');
    var kit=_reSpBuscar_(r,h,'1. ID_Código Kit');
    if(!sku||!kit)return;
    var key=kit+'|'+sku,quant=_reSpQuant_(_reSpBuscar_(r,h,'16. TIR_SKU'));
    if(!aux[key])aux[key]={kit:kit,sku:sku,quant:0};
    aux[key].quant+=quant;
  });
  var matCalc=Object.keys(aux).map(function(k){var a=aux[k],pai=arv[a.kit]||[],elementos=_reSpBuscar_(pai,arvCab,'13. MAN_Total de itens colecionados'),emb=_reSpBuscar_(pai,arvCab,'13. MAN_Embalagem'),obs=_reSpBuscar_(pai,arvCab,'13. MAN_Observação de embalagem');
    var shrink=/shrink|shirink/i.test(obs),maleta=/maleta/i.test(emb);
    return {kit:a.kit,sku:a.sku,quant:a.quant,elementos:Number(elementos)||0,emb:emb,obs:obs,
      shrinkIndividual:shrink&&Number(elementos)===1?'Sim':'Não',
      shrinkKit:shrink&&Number(elementos)>1?'Sim':'Não',maleta:maleta?'Sim':'Não'};
  });
  var porSku={};matCalc.forEach(function(x){
    var p=porSku[x.sku]||(porSku[x.sku]={avulso:0,kit:0,maleta:0,total:0});
    if(x.kit===x.sku)p.avulso+=x.quant;
    else if(x.maleta==='Sim')p.maleta+=x.quant;
    else p.kit+=x.quant;
    p.total+=x.quant;
  });
  var normPapel=function(s){s=String(s==null?'':s).trim().toUpperCase();if(!s||s==='-'||s==='NULL')return 0;return s.replace(/^CARTÃO (\d+)G$/,'CARTÃO TRIPLEX $1G').replace(/^OFFSET (\d+)G$/,'OFFSET $1G').replace(/^LAMINAÇÃO BOPP BRILHO$/,'LAMINAÇÃO BRILHO').replace(/^LAMINAÇÃO BOPP FOSCO$/,'LAMINAÇÃO FOSCA');};
  var capa=function(s){s=String(s||'').trim().toUpperCase();return s==='DURA'?'CAPA DURA':s;};
  var enc=function(s){s=String(s||'').toUpperCase();return s==='CANOA'?'GRAMPO CANOA':s==='LOMBADA_QUADRADA_PUR'?'PUR':s;};
  var manual=function(col,def){return man[col]===undefined?def:man[col];};
  dados.forEach(function(src){
    var x=function(col){var i=idx[_reSpNormal_(col)];return i===undefined?'':src[i];};
    var sku=x('1. ID_Código SKU');if(!sku)return;
    var r=[],v=function(n,value){r[n-3]=value;},g=function(i){return src[_reSpIdx_(i)]||'';};
    // Colunas C:CC. Posicoes rigorosamente do NC - PPG Especificacoes do De-Para 2740.
    ['B','C','D','F','H','I','M','J','L','K','O','P','X','T','W','Y','Z','AA'].forEach(function(c,j){v(j+3,g(c));});
    var cap=capa(g('AB'));v(21,cap);v(22,cap==='SIMPLES'||cap==='CAPA DURA'?4:cap==='EMPASTADA'?8:0);
    v(23,cap==='CAPA DURA'?String(g('AE')).charAt(0)+'x0':g('AE'));
    v(24,normPapel(cap==='CAPA DURA'?g('AF'):g('AD')));
    v(25,/personalizada/i.test(g('BU'))?'Sim':'Não');v(26,g('BW'));
    var orelha=_reSpBool_(g('AC')),larguraOrelha=Number(String(g('AC')).trim().replace(',','.'));
    v(27,orelha==='Sim'||(isFinite(larguraOrelha)&&larguraOrelha>0)?'Sim':'Não');
    v(28,_reSpNormal_(g('F'))==='MARALTO'&&isFinite(larguraOrelha)?larguraOrelha:0);
    v(29,_reSpBool_(g('AK')));v(30,normPapel(g('AH')));v(31,normPapel(g('AI')));
    for(var col=32;col<=37;col++)v(col,manual(col,0));
    v(38,g('BR'));v(39,cap==='CAPA DURA'&&enc(g('BM'))!=='ESPIRAL'?8:cap==='CAPA DURA'?4:0);
    v(40,cap==='CAPA DURA'?String(g('AE')).slice(-1)+'x0':0);
    v(41,normPapel(g('AG')));v(42,g('AL'));v(43,g('AN'));v(44,normPapel(g('AM')));
    v(45,g('BW'));v(46,/^(CYAN|MAGENTA)$/i.test(g('AO'))?'Sim':'Não');
    v(47,g('BT'));v(48,g('AO'));v(49,_reSpBool_(g('AQ')));v(50,_reSpBool_(g('AP')));
    v(51,manual(51,0));v(52,g('AR'));v(53,g('AS'));v(54,g('AU'));v(55,normPapel(g('AT')));
    v(56,_reSpBool_(g('AX')));v(57,_reSpBool_(g('AW')));v(58,g('AY'));
    v(59,g('AZ'));v(60,g('BB'));v(61,normPapel(g('BA')));v(62,_reSpBool_(g('BE')));
    v(63,_reSpBool_(g('BD')));v(64,g('BF'));v(65,g('BG'));v(66,g('BI'));
    v(67,normPapel(g('BH')));v(68,manual(68,0));v(69,_reSpBool_(g('BK')));
    v(70,g('BL'));v(71,enc(g('BM')));v(72,manual(72,0));v(73,manual(73,0));
    v(74,g('BN'));v(75,manual(75,0));v(76,g('BQ'));
    var entries=matCalc.filter(function(y){return y.sku===sku;});
    v(77,entries.some(function(y){return y.shrinkIndividual==='Sim';})?'Sim':'Não');
    v(78,manual(78,'Não'));
    v(79,entries.some(function(y){return y.shrinkKit==='Sim';})?'Sim':'Não');
    v(80,entries.some(function(y){return y.maleta==='Sim';})?'Sim':'Não');
    v(81,0);
    if(r.length!==79)throw Error('NC com numero incorreto de colunas');
    // Nunca gravar literais booleanos em campos Sim/Nao do contrato de custos.
    [25,27,29,46,49,50,51,56,57,62,63,68,69,75,77,78,79,80].forEach(function(col){
      r[col-3]=_reSpBool_(r[col-3]);
    });
    // Areas, dimensoes e campos manuais quantitativos precisam ser numericos no XLSX.
    [22,28,33,35,37,39,72,73,81].forEach(function(col){
      r[col-3]=_reSpNumeroFonte_(r[col-3]);
    });
    ret.push(r);
  });
  if(!ret.length)throw Error('Nenhuma especificacao util apos o filtro por grafica');
  return {especificacoes:ret,calc:matCalc,porSku:porSku};
}
function listarPastasGraficasReentradasSPG(){
  return Object.keys(RE_SP_MOLDES).map(function(id){return {nome:RE_SP_MOLDES[id].nome,id:id};});
}
// Unico ponto de troca de fonte. Remover junto com LegadoAlocacaoReentradas.js
// quando a transicao dos chamados antigos terminar.
function _reSpDadosChamado_(numero,grafica,origensLegado,especificacoesGerais){
  var legado=Array.isArray(origensLegado)&&origensLegado.some(function(id){return String(id).trim()===numero;});
  if(!legado)return _reSpPrepararDados_(numero,grafica);
  // A confirmacao da tela nao pode trocar um chamado existente/cancelado por Sheets.
  try{
    _emailSpgPacote_(numero,true,'reentrada','reentrada');
  }catch(err){
    if(err.codigo==='CHAMADO_NAO_ENCONTRADO')return _reSpPrepararDadosLegado_(numero,grafica,especificacoesGerais);
    throw err;
  }
  throw Error('Este chamado existe no SPG. Consulte novamente os dados da API.');
}
function _reSpValoresResumo_(dados,campos) {
  var unicos={},valores=[];
  dados.linhas.forEach(function(linha){
    campos.forEach(function(campo){
      var valor=_reSpBuscar_(linha,dados.headers,campo);
      if(valor===null||valor===undefined)return;
      valor=String(valor).trim();
      if(!valor||valor==='-'||/^(N\/A|NÃO INFORMADO|NAO INFORMADO)$/i.test(valor))return;
      var chave=_reSpNormal_(valor);
      if(!chave||unicos[chave])return;
      unicos[chave]=true;
      valores.push(valor);
    });
  });
  return valores.join(', ');
}
function buscarResumoReentradasSPG(chamados,origensLegado,especificacoesGerais){
  return String(chamados||'').split(',').map(function(s){return s.trim();}).filter(Boolean).map(function(id){
    try{
      var p=_reSpDadosChamado_(id,'LOGPRINT',origensLegado,especificacoesGerais);
      var skus={};
      p.espec.linhas.forEach(function(r){
        var sku=_reSpBuscar_(r,p.espec.headers,'1. ID_Código SKU');
        if(sku)skus[sku]=true;
      });
      var total=0;
      p.tiragem.linhas.forEach(function(r){
        total+=_reSpQuant_(_reSpBuscar_(r,p.tiragem.headers,'16. TIR_SKU'));
      });
      return {
        chamado:id,
        origem:p.pacote&&p.pacote.origem==='legado'?'legado':'spg',
        legadoTemporario:!!(p.pacote&&p.pacote.origem==='legado'),
        skusEspecificacaoGeral:p.skusEspecificacaoGeral||[],
        skusIgnorados:p.skusIgnorados||[],
        tiragemIgnorada:p.tiragemIgnorada||0,
        semSkusParaGerar:!!p.semSkusParaGerar,
        avisos:p.avisos||[],
        totalSkus:Object.keys(skus).length,
        tiragemTotal:total,
        marcas:_reSpValoresResumo_(p.espec,['2. INF_Marca']),
        series:_reSpValoresResumo_(p.espec,['2. INF_Série']),
        papeisCapa:_reSpValoresResumo_(p.espec,['4. CAP_Papel da capa']),
        papeisMiolo:_reSpValoresResumo_(p.espec,['5. MIO_PRIN_Papel']),
        cdsDestino:_reSpValoresResumo_(p.tiragem,['15. DES_CD envio'])
      };
    }catch(err){
      return {chamado:id,erro:err.message,naoEncontrado:err.codigo==='CHAMADO_NAO_ENCONTRADO',
        origem:Array.isArray(origensLegado)&&origensLegado.some(function(n){return String(n).trim()===id;})?'legado':'',
        especificacoesAusentes:err.codigo==='ESPECIFICACAO_LEGADO_AUSENTE'?err.especificacoesAusentes:[],totalSkus:0,tiragemTotal:0,
        marcas:'',series:'',papeisCapa:'',papeisMiolo:'',cdsDestino:''};
    }
  });
}
function validarReentradasSPG(grafica,chamados,origensLegado,especificacoesGerais) {
  if(!RE_SP_MOLDES[_reSpNormal_(grafica)])return {erro:'Calculadora indisponivel para esta grafica.'};
  var ids=String(chamados||'').split(',').map(function(s){return s.trim();}).filter(Boolean);
  if(!ids.length)return {erro:'Informe os chamados.'};
  var validos=[],invalidos=[],avisos=[];
  ids.forEach(function(id){try{
    var p=_reSpDadosChamado_(id,grafica,origensLegado,especificacoesGerais);
    (p.avisos||[]).forEach(function(aviso){avisos.push('Chamado '+id+': '+aviso);});
    if(p.semSkusParaGerar){invalidos.push(id+': nenhum SKU com especificação disponível; chamado não incluído.');return;}
    validos.push(id);
  }catch(err){invalidos.push(id+': '+err.message);}});
  if(!validos.length)return {erro:'Nenhum chamado com SKUs disponíveis para gerar: '+invalidos.join(' | '),avisos:avisos};
  var manuais=[
    [32,'RELEVO','Sim / Não'],[33,'RELEVO - ÁREA','cm²'],[34,'UV LOCALIZADO','Sim / Não'],
    [35,'UV LOCALIZADO - ÁREA','cm²'],[36,'HOT STAMPING','Sim / Não'],[37,'HOT STAMPING - ÁREA','cm²'],
    [51,'LAM. DADOS VARIÁVEIS','Sim / Não'],[68,'SERRILHA / FITALITO (ADESIVOS)','Sim / Não'],
    [72,'BITOLA','mm'],[73,'FILAMENTO','mm'],[75,'ESCANTEAMENTO','Sim / Não'],
    [78,'SHRINK COLETIVO','Sim / Não']
  ].map(function(r){return {index:String(r[0]),titulo:r[1],instrucao:r[2],
    opcoes:/Sim/.test(r[2])?['Não','Sim']:null};});
  return {bundle:true,sucesso:true,validos:validos,invalidos:invalidos,avisos:avisos,colunasManuais:manuais};
}
function gerarValidacaoReentradasSPG(grafica,chamados,opcaoVersao,versoesManuais,respostas,fornecimento,origensLegado,especificacoesGerais) {
  if(!RE_SP_MOLDES[_reSpNormal_(grafica)])return {erro:'Modelo ainda nao disponivel para esta grafica.'};
  if ((opcaoVersao && opcaoVersao !== 'spg' && opcaoVersao !== 'ultima') || String(versoesManuais || '').trim()) {
    return {sucesso:false,erro:'A validação usa somente a versão vinculada ao chamado no SPG. Atualize a página.'};
  }
  try{
    var result=[],failed=[],ignorados=[];
    (chamados||[]).forEach(function(chamado){
      try{
        var item=_reSpGerarUm_(String(chamado),grafica,respostas,fornecimento||{},origensLegado,especificacoesGerais);
        if(item.semSkusParaGerar)ignorados.push(item);else result.push(item);
      }
      catch(e){failed.push(chamado+': '+e.message);}
    });
    var avisos=failed.slice();
    result.concat(ignorados).forEach(function(item){(item.avisos||[]).forEach(function(aviso){avisos.push(aviso);});});
    return {sucesso:!!result.length,erro:result.length?undefined:(failed.join(' | ')||'Nenhum SKU com especificação disponível. Nenhum arquivo foi gerado.'),
      resumo:result,arquivos:result,ignorados:ignorados,avisos:avisos,jobs:[]};
  }catch(e){return {erro:e.message};}
}
// Edita SOMENTE as entradas da CALC Geral mantendo intactas as formulas em
// todas as linhas e colunas, inclusive as mais de 100 colunas de precificacao.
function _reSpPreencherCalcGeral_(xml,skus,porSku,chamado,fornecimento,grafica){
  var parte=/<sheetData(?:\s[^>]*)?>[\s\S]*?<\/sheetData>/.exec(xml);
  if(!parte)throw Error('CALC Geral sem sheetData');
  var original=parte[0],linhas=/<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/g;
  var f=fornecimento||{},elements=f.elementos||{},tec=_reSpNormal_(f.tecnologia||'DIGITAL');
  var forn=f.fornecedor||'GRÁFICA';
  if(['DIGITAL','PLANA','ROTATIVA'].indexOf(tec)<0)throw Error('Tecnologia invalida: '+tec);
  var formaCerta=_reSpNormal_(grafica)==='FORMACERTA';
  var inputs=formaCerta?['C','D','F','G','H','X','AA',
    'CA','CE','CI','CM','CQ','CU','CY','DC','DG','DK','DO','DS','DW','EA','EE','EI','EM','EQ']:
    ['C','D','F','G','H','I','Y','AB','AE',
    'CF','CK','CP','CU','CZ','DE','DI',
    'DN','DS','DX','EC','EH','EM','EQ',
    'EV','FA','FF','FK','FP','FU','FY',
    'CE','CJ','CO','CT','CY','DD','DH',
    'DM','DR','DW','EB','EG','EL','EP',
    'EU','EZ','FE','FJ','FO','FT','FX'];
  var index={};inputs.forEach(function(c){index[_reSpIdx_(c)]=true;});
  var seen=0,regex=/<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/g;
  var rewritten=original.replace(regex,function(full,n){
    var rn=+n;if(rn<6)return full;
    var k=rn-6,sku=skus[k],v={};
    if(sku!==undefined){
      seen++;
      var x=porSku[sku];
      // Campos de entrada extraidos diretamente do contrato da calculadora:
      // C=chamado, D=SKU, F/G/H/I=tiragens, Y=tecnologia, AB=acabamento.
      v[_reSpIdx_('C')]='Chamado '+chamado;
      v[_reSpIdx_('D')]=sku;
      v[_reSpIdx_('F')]=x.avulso;
      v[_reSpIdx_('G')]=formaCerta?x.kit:0;
      v[_reSpIdx_('H')]=formaCerta?x.maleta:x.kit;
      if(!formaCerta)v[_reSpIdx_('I')]=x.maleta;
      v[_reSpIdx_(formaCerta?'X':'Y')]=tec;
      v[_reSpIdx_(formaCerta?'AA':'AB')]=forn;
      var blocos=formaCerta?{
        DIGITAL:['CA','CE','CI','CM','CQ','CU'],
        PLANA:['CY','DC','DG','DK','DO','DS'],
        ROTATIVA:['DW','EA','EE','EI','EM','EQ']
      }:{
        DIGITAL:['CF','CK','CP','CU','CZ','DE','DI'],
        PLANA:['DN','DS','DX','EC','EH','EM','EQ'],
        ROTATIVA:['EV','FA','FF','FK','FP','FU','FY']
      };
      var keys=['capa','forro','miolo_principal','miolo_professor','encartes','adesivos','ajuste_papel'];
      Object.keys(blocos).forEach(function(t){
        blocos[t].forEach(function(col,i){
          var user=elements[t]||{},val=user[keys[i]]||'';
          if(val!=='ARCO'&&val!=='GRÁFICA'&&val!=='ARCO TRIANG') {
            if(val!=='')throw Error('Fornecimento invalido: '+val);
            val=keys[i]==='capa'||keys[i]==='forro'||keys[i]==='miolo_principal'?'GRÁFICA':'';
          }
          if(val)v[_reSpIdx_(col)]=val;
        });
      });
    }
    var head=/^<row\b[^>]*>/.exec(full)[0],end='</row>';
    var cells=/<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g,found={},match;
    while((match=cells.exec(full))!==null){
      var raw=match[0],m=/\br="([A-Z]+)\d+"/.exec(raw);
      if(!m)continue;
      var i=_reSpIdx_(m[1]);
      if(index[i]&& !/<f(?:\s|>)/.test(raw)){
        // descarta todas as entradas 2910 do modelo, inclusive linhas sobrando
        var sty=/\bs="(\d+)"/.exec(raw);
        if(sty)found[i]={style:sty[1]};
      }else found[i]={xml:raw};
    }
    Object.keys(v).forEach(function(k){
      var i=Number(k),prior=found[i]||{};
      if(prior.xml&&/<f(?:\s|>)/.test(prior.xml))throw Error('Coluna de entrada tem formula: '+_reSpCol_(i+1));
      found[i]={xml:_reSpCelula_(_reSpCol_(i+1)+rn,v[i],prior.style)};
    });
    var result=Object.keys(found).map(Number).sort(function(a,b){return a-b;})
      .map(function(i){return found[i].xml||'';}).join('');
    return head+result+end;
  });
  if(seen!==skus.length)throw Error('CALC Geral sem capacidade para '+skus.length+' SKUs (linhas: '+seen+')');
  return xml.replace(original,rewritten);
}

function _reSpGerarUm_(chamado,grafica,respostas,fornecimento,origensLegado,especificacoesGerais) {
  if(!/^\d{2,8}$/.test(chamado))throw Error('Numero do chamado invalido');
  var api=_reSpDadosChamado_(chamado,grafica,origensLegado,especificacoesGerais);
  if(api.semSkusParaGerar)return {chamado:chamado,origem:'legado',semSkusParaGerar:true,
    skusIgnorados:api.skusIgnorados||[],avisos:api.avisos||[],skus:0,tiragem:0};
  var deparaOrig=_reSpMolde_(grafica,'depara');
  var calcOrig=_reSpMolde_(grafica,'calc');
  var dnomes=['PPG - Especificações','PPG - Arvore de Produto','PPG - Tiragem',
    'NC - PPG Especificações','NC - CALC Geral','TAB - Ajustes de Nomes'];
  var cnomes=['PPG Especificações','CALC Geral','CALC Digital','CALC Plana',
    'CALC Rotativa','CALC Acabamento','TAB Regras','TAB Formatos',
    'TAB Papel','TAB Impressão','TAB Acabamento','TAB Versões'];
  if(_reSpNormal_(grafica)==='FORMACERTA')cnomes.push('IMPUT Forma Certa');
  dnomes.forEach(function(n){_reSpAbaPorNome_(deparaOrig,n);});
  cnomes.forEach(function(n){_reSpAbaPorNome_(calcOrig,n);});
  var listas=[api.espec,api.arvore,api.tiragem],bases=[],alterDepara={};
  for(var s=0;s<3;s++){
    var caminho=_reSpAbaPorNome_(deparaOrig,dnomes[s]),xml=deparaOrig.mapa[caminho].getDataAsString('UTF-8');
    var model=_reSpMatriz_(deparaOrig,dnomes[s]);
    var data=_reSpMontarBase_(listas[s],model.headers);
    var minimo=s===0?1:3;
    if(!data.length)throw Error('Aba de origem '+dnomes[s]+' ficou vazia');
    bases.push({headers:model.headers,linhas:data});
    var styles=_reSpEstilos_(xml,2),rows=data.map(function(v,i){
      var o={};v.forEach(function(x,j){if(x!==''&&x!==null&&x!==undefined)o[j]=_reSpBool_(x);});
      if(s===2)o[19]=v[0]===v[3]?'Avulso':'Kit';
      return _reSpLinha_(i+2,o,styles);
    });
    // Preservar o cabecalho original, limpar quaisquer dados antigos do 2740.
    alterDepara[caminho]=_reSpTrocarLinhas_(xml,2,rows);
  }
  var nc=_reSpConstrucaoNC_(bases[0],bases[1],bases[2],respostas);
  if(nc.especificacoes.length>RE_SP_CONFIG.limiteSkus)throw Error('Excesso de SKUs para calculadora; maximo '+RE_SP_CONFIG.limiteSkus);
  var dc=_reSpAbaPorNome_(deparaOrig,'NC - PPG Especificações'),
    dxml=deparaOrig.mapa[dc].getDataAsString('UTF-8'),dstyle=_reSpEstilos_(dxml,6);
  var ncrows=nc.especificacoes.map(function(r,i){var o={};r.forEach(function(x,j){o[j+2]=x;});return _reSpLinha_(i+6,o,dstyle);});
  alterDepara[dc]=_reSpTrocarLinhas_(dxml,6,ncrows);
  var da=_reSpAbaPorNome_(deparaOrig,'NC - CALC Geral'),axml=deparaOrig.mapa[da].getDataAsString('UTF-8'),astyle=_reSpEstilos_(axml,4);
  // A:J (kits/componentes) e L:U (totais por SKU) sao blocos paralelos,
  // NAO sequenciais: o consumidor legado le L:S a partir da linha 4.
  var bySku=Object.keys(nc.porSku).sort(),cgrows=[];
  for(var ic=0;ic<Math.max(nc.calc.length,bySku.length);ic++){
    var x=nc.calc[ic],sku=bySku[ic],g={};
    if(x){g={0:x.kit,1:x.sku,2:x.quant,3:x.kit===x.sku?'TRUE':'FALSE',4:x.elementos,
      5:x.emb,6:x.shrinkIndividual,7:x.shrinkKit,8:x.maleta,9:x.obs};}
    if(sku){
      var n=nc.porSku[sku];
      g[11]=sku;g[12]=n.avulso;g[13]=n.total;
      g[15]=n.avulso;g[16]=n.kit;g[17]=n.maleta;g[18]=n.total;g[19]='TRUE';
    }
    cgrows.push(_reSpLinha_(ic+4,g,astyle));
  }
  alterDepara[da]=_reSpTrocarLinhas_(axml,4,cgrows);
  // Insere especificacoes apenas na aba de entrada; as 11 demais mantem formula original.
  var cp=_reSpAbaPorNome_(calcOrig,'PPG Especificações'),cpx=calcOrig.mapa[cp].getDataAsString('UTF-8'),cpstyle=_reSpEstilos_(cpx,5);
  var cr=nc.especificacoes.map(function(r,i){var o={};r.forEach(function(x,j){o[j+2]=x;});return _reSpLinha_(i+5,o,cpstyle);});
  var alterCalc={};alterCalc[cp]=_reSpTrocarLinhas_(cpx,5,cr);
  var geral=_reSpAbaPorNome_(calcOrig,'CALC Geral'),gx=calcOrig.mapa[geral].getDataAsString('UTF-8');
  alterCalc[geral]=_reSpPreencherCalcGeral_(gx,bySku,nc.porSku,chamado,fornecimento,grafica);
  if(_reSpNormal_(grafica)==='FORMACERTA'){
    var imput=_reSpAbaPorNome_(calcOrig,'IMPUT Forma Certa');
    alterCalc[imput]=_reFcExpandirImput_(calcOrig.mapa[imput].getDataAsString('UTF-8'),bySku.length);
  }
  // Marcar todas as formulas para recalculo integral, sem substituir tarifas da Logprint.
  alterDepara['xl/workbook.xml']=_reSpFullCalc_(deparaOrig.mapa['xl/workbook.xml'].getDataAsString('UTF-8'));
  alterCalc['xl/workbook.xml']=_reSpFullCalc_(calcOrig.mapa['xl/workbook.xml'].getDataAsString('UTF-8'));
  [deparaOrig,calcOrig].forEach(function(w,idx){
    var alter=idx===0?alterDepara:alterCalc,names=idx===0?dnomes:cnomes;
    // A senha definida neste gerador e aplicada somente apos todos os dados.
    names.forEach(function(n){var p=_reSpAbaPorNome_(w,n);alter[p]=_reSpProteger_(alter[p]===undefined?w.mapa[p].getDataAsString('UTF-8'):alter[p]);});
  });
  var nomeGrafica=RE_SP_MOLDES[_reSpNormal_(grafica)].nome;
  var nomeD='PLANILHA DE PARA PPG '+nomeGrafica.toUpperCase()+' - CHAMADO '+chamado+'.xlsx';
  var nomeC='CALCULADORA FORNECEDOR '+nomeGrafica.toUpperCase()+' - LIVROS REENTRADA 2026 - NOVA VERSÃO - CHAMADO '+chamado+'.xlsx';
  // Somente aqui os arquivos sao criados, depois de validar todas as matrizes.
  var fd=_reSpSalvarZip_(null,deparaOrig,alterDepara,nomeD);
  var fc=_reSpSalvarZip_(null,calcOrig,alterCalc,nomeC);
  return {chamado:chamado,grafica:nomeGrafica,origem:api.pacote.origem==='legado'?'legado':'spg',skus:bySku.length,
    tiragem:Object.keys(nc.porSku).reduce(function(s,k){return s+nc.porSku[k].total;},0),
    depara:fd.getUrl(),calculadora:fc.getUrl(),arquivo:nomeC,
    skusEspecificacaoGeral:api.skusEspecificacaoGeral||[],avisos:api.avisos||[],
    skusIgnorados:api.skusIgnorados||[],tiragemIgnorada:api.tiragemIgnorada||0,
    aviso:'O Excel recalcula custos ao abrir; precos nao homologados sem abrir e validar.'};
}
