// Gerador nativo XLSX - somente VALIDACAO DE CUSTOS REENTRADAS.
// Sem SpreadsheetApp, sem fila de e-mail, sem alteracao na Validacao Principal.
// Excel original e mantido: formulas, tabelas, estilos e 12 abas.
var RE_SP_CONFIG = {
  deparaModelo: '1TU5fAYIAznMlva10vvElmD-icFuygbws',
  calculadoraLogprint: '1cb3GthAbz9yPtILBMMfk_du8gGNyI9i6',
  pastaDestino: '1DKie1AHE7EFHa6roJY2XxE2ZlFhC_lCy',
  senhaProtecao: 'PGFLOW2026',
  limiteSkus: 300
};

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
function _reSpNormal_(s){return String(s||'').trim().replace(/\s+/g,' ').toUpperCase();}
function _reSpMapearCab_(dest,cab,linhas) {
  var mapa={};cab.forEach(function(c,i){if(c)mapa[_reSpNormal_(c)]=i;});
  var idx=dest.map(function(h){var n=_reSpNormal_(h);return mapa[n]===undefined?-1:mapa[n];});
  var rows=linhas.map(function(a){return idx.map(function(i){return i===-1?'':a[i];});});
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
  return v===null||v===undefined?'':String(v);
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
  xml=xml.replace(/<calcPr\b[^>]*\/>/g,'');
  return xml.replace('</workbook>','<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>');
}
function _reSpSalvarZip_(arquivo,partes,alteradas,nome) {
  var saida=partes.partes.map(function(p){
    var val=alteradas[p.getName()];
    return val===undefined?p:Utilities.newBlob(val,'text/xml',p.getName());
  });
  var xlsx=Utilities.zip(saida,nome).setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  return DriveApp.getFolderById(RE_SP_CONFIG.pastaDestino).createFile(xlsx).setName(nome);
}
function _reSpPrepararDados_(numero,grafica,versaoManual) {
  var pacote=_emailSpgPacote_(numero,false,'reentrada');
  if(pacote.id!==numero||!pacote.versaoId||!pacote.blob)throw Error('Chamado ou versao inconsistentes na API');
  if(/CANCELAD/i.test(pacote.statusChamado+' '+pacote.statusProducao))throw Error('Chamado cancelado');
  var raw=_reSpPartes_(pacote.blob);
  var e=_reSpMatriz_(raw,'Esp. Técnica (legado)');
  var a=_reSpMatriz_(raw,'Árvore (legado)');
  var t=_reSpMatriz_(raw,'Tiragem (legado)');
  if(!e.linhas.length||!a.linhas.length||!t.linhas.length)throw Error('OP do SPG sem as tres matrizes de producao');
  // A OP pode conter alocacoes de varias graficas. Sem atribuicao demonstravel, NAO atribuir a Logprint por suposicao.
  var colunaGrafica=t.headers.find(function(h){return /^(GRAFICA|GRÁFICA|GRAFICA DESTINO|GRÁFICA DESTINO)$/i.test(h);});
  var tGraf=colunaGrafica?t.linhas.filter(function(l){return _reSpNormal_(_reSpBuscar_(l,t.headers,colunaGrafica))===_reSpNormal_(grafica);}):[];
  if(!tGraf.length){
    // Exclusivamente o contrato de alocacao de versao, nunca o consolidado.
    // Se nao fornecer grafica, kit, sku, CD e quantidade inequivocos, bloquear.
    var dadosAlocacao;
    try{dadosAlocacao=_emailSpgLerJson_('/api/tiragem/'+pacote.versaoId+'/alocacao',_spgLogin_());}
    catch(exc){throw Error('Nao foi possivel segmentar o chamado por grafica na API: '+exc.message);}
    var aloc=Array.isArray(dadosAlocacao)?dadosAlocacao:
      (dadosAlocacao.alocacoes||dadosAlocacao.rows||dadosAlocacao.data||[]);
    if(!Array.isArray(aloc))throw Error('Resposta de alocacoes desconhecida; processamento bloqueado');
    var chave=function(a,b,c){return [_reSpNormal_(a),_reSpNormal_(b),_reSpNormal_(c)].join('|');};
    var xAloc={};
    aloc.forEach(function(z){
      var nome=z.nomeGrafica||(z.grafica&&(z.grafica.nome||z.grafica.name))||z.graficaNome||'';
      if(_reSpNormal_(nome)!==_reSpNormal_(grafica))return;
      var sku=z.skuAtual||z.sku||z.codigoSku||'',kit=z.kitCode||z.kit||z.codigoKit||'';
      var cd=z.cdDestino||z.cd||z.destinoCd||'';
      var q=z.quantidade===undefined?z.tiragem:z.quantidade;
      if(!sku||!kit||!cd||q===undefined)throw Error('Alocacao sem chave/quantidade completa: abortado');
      var k=chave(kit,sku,cd);
      if(xAloc[k])throw Error('Alocacoes duplicadas precisam de conciliacao: '+k);
      xAloc[k]=_reSpQuant_(q);
    });
    if(!Object.keys(xAloc).length)throw Error('A API nao confirmou tiragem da '+grafica+' nesta versao');
    tGraf=[];
    t.linhas.forEach(function(ln){
      var kit=_reSpBuscar_(ln,t.headers,'1. ID_Código Kit'),sku=_reSpBuscar_(ln,t.headers,'1. ID_Código SKU'),
          cd=_reSpBuscar_(ln,t.headers,'15. DES_CD envio'),k=chave(kit,sku,cd);
      if(xAloc[k]===undefined)return;
      var item=ln.slice(),cabe=t.headers;
      ['16. TIR_SKU','16. TIR_Item','16. TIR_Miolo','16. TIR_Capa avulsa','16. TIR_Kit'].forEach(function(h){
        var i=cabe.indexOf(h);if(i>=0)item[i]=xAloc[k];
      });
      tGraf.push(item);delete xAloc[k];
    });
    if(Object.keys(xAloc).length)throw Error('Alocacoes da grafica nao conciliadas com a OP: '+Object.keys(xAloc).join(', '));
  }
  if(!tGraf.length)throw Error('Nenhuma tiragem vinculada a grafica '+grafica+' na versao '+pacote.versaoId);
  var skus={},kits={};
  tGraf.forEach(function(l){
    var sku=_reSpBuscar_(l,t.headers,'1. ID_Código SKU'),kit=_reSpBuscar_(l,t.headers,'1. ID_Código Kit');
    if(sku)skus[sku]=true;if(kit)kits[kit]=true;
  });
  return {pacote:pacote,espec:{headers:e.headers,linhas:e.linhas.filter(function(l){
    return !!skus[_reSpBuscar_(l,e.headers,'1. ID_Código SKU')];
  })},arvore:{headers:a.headers,linhas:a.linhas.filter(function(l){
    return !!kits[_reSpBuscar_(l,a.headers,'1. ID_Código Kit')];
  })},tiragem:{headers:t.headers,linhas:tGraf}};
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
  var normPapel=function(s){s=String(s||'').trim().toUpperCase();return s.replace(/^CARTÃO (\d+)G$/,'CARTÃO TRIPLEX $1G').replace(/^OFFSET (\d+)G$/,'OFFSET $1G').replace(/^LAMINAÇÃO BOPP BRILHO$/,'LAMINAÇÃO BRILHO').replace(/^LAMINAÇÃO BOPP FOSCO$/,'LAMINAÇÃO FOSCA');};
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
    v(27,g('AC')&&Number(g('AC'))!==0?'Sim':'Não');v(28,_reSpNormal_(g('F'))==='MARALTO'?g('AC'):0);
    v(29,_reSpBool_(g('AK')));v(30,normPapel(g('AH')));v(31,normPapel(g('AI')));
    for(var col=32;col<=37;col++)v(col,manual(col,col%2===0?'Não':0));
    v(38,g('BR'));v(39,cap==='CAPA DURA'&&enc(g('BM'))!=='ESPIRAL'?8:cap==='CAPA DURA'?4:0);
    v(40,cap==='CAPA DURA'?String(g('AE')).slice(-1)+'x0':g('AE'));
    v(41,normPapel(g('AG')));v(42,g('AL'));v(43,g('AN'));v(44,normPapel(g('AM')));
    v(45,g('BW'));v(46,/^(CYAN|MAGENTA)$/i.test(g('AO'))?'Sim':'Não');
    v(47,g('BT'));v(48,g('AO'));v(49,_reSpBool_(g('AQ')));v(50,_reSpBool_(g('AP')));
    v(51,manual(51,'Não'));v(52,g('AR'));v(53,g('AS'));v(54,g('AU'));v(55,normPapel(g('AT')));
    v(56,_reSpBool_(g('AX')));v(57,_reSpBool_(g('AW')));v(58,g('AY'));
    v(59,g('AZ'));v(60,g('BB'));v(61,normPapel(g('BA')));v(62,_reSpBool_(g('BE')));
    v(63,_reSpBool_(g('BD')));v(64,g('BF'));v(65,g('BG'));v(66,g('BI'));
    v(67,normPapel(g('BH')));v(68,manual(68,'Não'));v(69,_reSpBool_(g('BK')));
    v(70,g('BL'));v(71,enc(g('BM')));v(72,manual(72,0));v(73,manual(73,0));
    v(74,g('BN'));v(75,manual(75,'Não'));v(76,g('BQ'));
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
    ret.push(r);
  });
  if(!ret.length)throw Error('Nenhuma especificacao util apos o filtro por grafica');
  return {especificacoes:ret,calc:matCalc,porSku:porSku};
}
function listarPastasGraficasReentradasSPG(){
  // Disponibilidade de modelos: por ora somente Logprint homologada.
  return [{nome:'Logprint',id:'LOGPRINT'}];
}
function buscarResumoReentradasSPG(chamados){
  return String(chamados||'').split(',').map(function(s){return s.trim();}).filter(Boolean).map(function(id){
    try{var p=_reSpPrepararDados_(id,'LOGPRINT');
      var skus={};p.espec.linhas.forEach(function(r){var s=_reSpBuscar_(r,p.espec.headers,'1. ID_Código SKU');if(s)skus[s]=1;});
      var total=0,cds={};p.tiragem.linhas.forEach(function(r){total+=_reSpQuant_(_reSpBuscar_(r,p.tiragem.headers,'16. TIR_SKU'));var cd=_reSpBuscar_(r,p.tiragem.headers,'15. DES_CD envio');if(cd)cds[cd]=1;});
      return {chamado:id,totalSkus:Object.keys(skus).length,tiragemTotal:total,cdsDestino:Object.keys(cds).join(', '),marcas:'',series:'',papeisCapa:'',papeisMiolo:''};
    }catch(err){return {chamado:id,erro:err.message,totalSkus:0,tiragemTotal:0,marcas:'',series:'',papeisCapa:'',papeisMiolo:'',cdsDestino:''};}
  });
}
function validarReentradasSPG(grafica,chamados) {
  if(_reSpNormal_(grafica)!=='LOGPRINT')return {erro:'Apenas calculadora Logprint esta disponivel.'};
  var ids=String(chamados||'').split(',').map(function(s){return s.trim();}).filter(Boolean);
  if(!ids.length)return {erro:'Informe os chamados.'};
  var validos=[],invalidos=[];
  ids.forEach(function(id){try{_reSpPrepararDados_(id,grafica);validos.push(id);}catch(err){invalidos.push(id+': '+err.message);}});
  if(!validos.length)return {erro:'Nenhum chamado validado: '+invalidos.join(' | ')};
  var manuais=[
    [32,'RELEVO','Sim / Não'],[33,'RELEVO - ÁREA','cm²'],[34,'UV LOCALIZADO','Sim / Não'],
    [35,'UV LOCALIZADO - ÁREA','cm²'],[36,'HOT STAMPING','Sim / Não'],[37,'HOT STAMPING - ÁREA','cm²'],
    [51,'LAM. DADOS VARIÁVEIS','Sim / Não'],[68,'SERRILHA / FITALITO (ADESIVOS)','Sim / Não'],
    [72,'BITOLA','mm'],[73,'FILAMENTO','mm'],[75,'ESCANTEAMENTO','Sim / Não'],
    [78,'SHRINK COLETIVO','Sim / Não']
  ].map(function(r){return {index:String(r[0]),titulo:r[1],instrucao:r[2],
    opcoes:/Sim/.test(r[2])?['Não','Sim']:null};});
  return {bundle:true,sucesso:true,validos:validos,invalidos:invalidos,colunasManuais:manuais};
}
function gerarValidacaoReentradasSPG(grafica,chamados,opcaoVersao,versoesManuais,respostas,fornecimento) {
  if(_reSpNormal_(grafica)!=='LOGPRINT')return {erro:'Modelo ainda nao homologado para esta grafica.'};
  try{
    var result=[],failed=[];
    (chamados||[]).forEach(function(chamado){
      try{result.push(_reSpGerarUm_(String(chamado),grafica,respostas,fornecimento||{}));}
      catch(e){failed.push(chamado+': '+e.message);}
    });
    return {sucesso:!!result.length,erro:result.length?undefined:failed.join(' | '),resumo:result,arquivos:result,avisos:failed,jobs:[]};
  }catch(e){return {erro:e.message};}
}
function _reSpGerarUm_(chamado,grafica,respostas,fornecimento) {
  if(!/^\d{2,8}$/.test(chamado))throw Error('Numero do chamado invalido');
  var api=_reSpPrepararDados_(chamado,grafica);
  var deparaOrig=_reSpPartes_(DriveApp.getFileById(RE_SP_CONFIG.deparaModelo).getBlob());
  var calcOrig=_reSpPartes_(DriveApp.getFileById(RE_SP_CONFIG.calculadoraLogprint).getBlob());
  var dnomes=['PPG - Especificações','PPG - Arvore de Produto','PPG - Tiragem',
    'NC - PPG Especificações','NC - CALC Geral','TAB - Ajustes de Nomes'];
  var cnomes=['PPG Especificações','CALC Geral','CALC Digital','CALC Plana',
    'CALC Rotativa','CALC Acabamento','TAB Regras','TAB Formatos',
    'TAB Papel','TAB Impressão','TAB Acabamento','TAB Versões'];
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
  var cgrows=nc.calc.map(function(x,i){var g={0:x.kit,1:x.sku,2:x.quant,3:x.kit===x.sku?'TRUE':'FALSE',4:x.elementos,
    5:x.emb,6:x.shrinkIndividual,7:x.shrinkKit,8:x.maleta,9:x.obs};return _reSpLinha_(i+4,g,astyle);});
  var idx=nc.calc.length+4,bySku=Object.keys(nc.porSku).sort();
  bySku.forEach(function(sku,i){var n=nc.porSku[sku],j={11:sku,12:n.avulso,13:n.total,15:n.avulso,16:n.kit,17:n.maleta,18:n.total,19:'TRUE'};cgrows.push(_reSpLinha_(idx+i,j,astyle));});
  alterDepara[da]=_reSpTrocarLinhas_(axml,4,cgrows);
  // Insere especificacoes apenas na aba de entrada; as 11 demais mantem formula original.
  var cp=_reSpAbaPorNome_(calcOrig,'PPG Especificações'),cpx=calcOrig.mapa[cp].getDataAsString('UTF-8'),cpstyle=_reSpEstilos_(cpx,5);
  var cr=nc.especificacoes.map(function(r,i){var o={};r.forEach(function(x,j){o[j+2]=x;});return _reSpLinha_(i+5,o,cpstyle);});
  var alterCalc={};alterCalc[cp]=_reSpTrocarLinhas_(cpx,5,cr);
  var geral=_reSpAbaPorNome_(calcOrig,'CALC Geral'),gx=calcOrig.mapa[geral].getDataAsString('UTF-8'),gstyle=_reSpEstilos_(gx,6);
  // Esta planilha contem formulas ate as ultimas colunas. Preserva formulas da linha modelo.
  var padrao=/<row\b[^>]*\br="6"[^>]*>([\s\S]*?)<\/row>/.exec(gx);
  if(!padrao)throw Error('Modelo calculadora sem linha de formulas 6');
  var opF=fornecimento||{},tecnologia=opF.tecnologia||'DIGITAL',forn=opF.fornecedor||'GRÁFICA';
  var gr=[];
  bySku.forEach(function(sku,i){
    var rowno=i+6,tot=nc.porSku[sku],vals={2:'Chamado '+chamado,3:sku,5:tot.avulso,6:0,7:tot.kit,8:tot.maleta,
      24:tecnologia,27:forn,30:0};
    var espec=nc.especificacoes.filter(function(r){return r[0]===sku;})[0];
    if(!espec)throw Error('SKU sem especificacao: '+sku);
    var base=padrao[1].replace(/(<c\b[^>]*\br="[A-Z]+)6(")/g,'$1'+rowno+'$2')
      .replace(/(\$?[A-Z]{1,3}\$?)6\b/g,'$1'+rowno);
    // Retirar valores/constantes da linha original de entrada (inclusive SKU 2740).
    Object.keys(vals).forEach(function(k){
      var col=_reSpCol_(+k+1),ref=col+rowno,cellRe=new RegExp('<c\\b[^>]*\\br="'+ref+'"[^>]*(?:\\/\\>|>[\\s\\S]*?<\\/c>)');
      base=base.replace(cellRe,'');
    });
    var add='';Object.keys(vals).forEach(function(k){add+=_reSpCelula_(_reSpCol_(+k+1)+rowno,vals[k],gstyle[k]);});
    gr.push('<row r="'+rowno+'">'+add+base+'</row>');
  });
  alterCalc[geral]=_reSpTrocarLinhas_(gx,6,gr);
  // Marcar todas as formulas para recalculo integral, sem substituir tarifas da Logprint.
  alterDepara['xl/workbook.xml']=_reSpFullCalc_(deparaOrig.mapa['xl/workbook.xml'].getDataAsString('UTF-8'));
  alterCalc['xl/workbook.xml']=_reSpFullCalc_(calcOrig.mapa['xl/workbook.xml'].getDataAsString('UTF-8'));
  [deparaOrig,calcOrig].forEach(function(w,idx){
    var alter=idx===0?alterDepara:alterCalc,names=idx===0?dnomes:cnomes;
    // A senha definida neste gerador e aplicada somente apos todos os dados.
    names.forEach(function(n){var p=_reSpAbaPorNome_(w,n);alter[p]=_reSpProteger_(alter[p]===undefined?w.mapa[p].getDataAsString('UTF-8'):alter[p]);});
  });
  var nomeD='PLANILHA DE PARA PPG LOGPRINT - CHAMADO '+chamado+'.xlsx';
  var nomeC='CALCULADORA FORNECEDOR LOGPRINT - LIVROS REENTRADA 2026 - NOVA VERSÃO - CHAMADO '+chamado+'.xlsx';
  // Somente aqui os arquivos sao criados, depois de validar todas as matrizes.
  var fd=_reSpSalvarZip_(null,deparaOrig,alterDepara,nomeD);
  var fc=_reSpSalvarZip_(null,calcOrig,alterCalc,nomeC);
  return {chamado:chamado,grafica:'LOGPRINT',skus:bySku.length,
    tiragem:Object.keys(nc.porSku).reduce(function(s,k){return s+nc.porSku[k].total;},0),
    depara:fd.getUrl(),calculadora:fc.getUrl(),arquivo:nomeC,
    aviso:'O Excel recalcula custos ao abrir; precos nao homologados sem abrir e validar.'};
}
