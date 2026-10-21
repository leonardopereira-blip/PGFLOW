// A Logprint foi entregue sem protecao e sem estilos desbloqueados. Reaplica
// somente a mascara de protecao do mesmo layout (Casa da Grafica), mantendo
// os valores, formulas e a formatacao propria da Logprint.
function _reLpAttr_(tag,nome,valor){
  var re=new RegExp('\\s'+nome+'="[^"]*"');
  if(re.test(tag))return tag.replace(re,' '+nome+'="'+valor+'"');
  return tag.replace(/\/?>(?=$)/,function(fim){return ' '+nome+'="'+valor+'"'+fim;});
}
function _reLpIgual_(a,b){return a.locked===b.locked&&a.hidden===b.hidden;}
function _reLpAbas_(workbook,rels){
  var porId={},out=[],m,re=/<Relationship\b[^>]*\/>/g;
  while((m=re.exec(rels))!==null){var r=_reMpAtributos_(m[0]);porId[r.Id]=r.Target;}
  re=/<sheet\b[^>]*>/g;
  while((m=re.exec(workbook))!==null){
    var s=_reMpAtributos_(m[0]);
    if(!porId[s['r:id']])throw Error('Aba sem relacionamento na protecao Logprint: '+s.name);
    out.push({nome:s.name,caminho:_reMpCaminho_(porId[s['r:id']])});
  }
  return out;
}
function _reLpColunas_(xml){
  var out=[],re=/<col\b[^>]*\/>/g,m;
  while((m=re.exec(xml))!==null){var a=_reMpAtributos_(m[0]);out.push({
    inicio:Number(a.min),fim:Number(a.max),estilo:a.style===undefined?0:Number(a.style),xml:m[0]
  });}
  return out;
}
function _reLpColuna_(colunas,n){
  for(var i=0;i<colunas.length;i++)if(colunas[i].inicio<=n&&colunas[i].fim>=n)return colunas[i];
  return null;
}
function _reLpMascara_(xml,estilos){
  var padrao=estilos[0],colunas=_reLpColunas_(xml),linhas={},celulas={},m;
  var re=/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g;
  while((m=re.exec(xml))!==null){
    var row=m[0],a=_reMpAtributos_((/^<row\b[^>]*>/.exec(row)||[])[0]),rn=Number(a.r),pRow=null;
    if(_reMpBool_(a.customFormat,0))pRow=linhas[rn]=estilos[Number(a.s||0)];
    var cr=/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g,c;
    while((c=cr.exec(row))!==null){
      var ca=_reMpAtributos_((/^<c\b[^>]*>/.exec(c[0])||[])[0]);
      var letra=/^[A-Z]+/.exec(ca.r||'');if(!letra)throw Error('Celula sem referencia na mascara Logprint.');
      // Uma celula materializada sem s usa XF0 no Excel. A heranca de linha
      // ou coluna vale para coordenadas sem elemento <c>, nao para <c sem s>.
      var col=_reLpColuna_(colunas,_reSpIdx_(letra[0])+1),base=pRow||(col?estilos[col.estilo]:padrao),p=estilos[Number(ca.s||0)];
      if(!p)throw Error('Estilo ausente na mascara Logprint: '+ca.s);
      if(!_reLpIgual_(p,base))celulas[ca.r]=p;
    }
  }
  return {padrao:padrao,colunas:colunas,linhas:linhas,celulas:celulas,estilos:estilos};
}
function _reLpTemExcecao_(aba,estilos){
  // O molde compacto permite descartar as nove abas inteiramente bloqueadas
  // sem materializar novamente suas centenas de milhares de formulas.
  var textos=[aba.inicio].concat(aba.celulas,aba.linhas.map(function(l){return l.h;}));
  for(var i=0;i<textos.length;i++){
    var re=/\b(?:s|style)="(\d+)"/g,m;
    while((m=re.exec(textos[i]))!==null)if(!_reLpIgual_(estilos[Number(m[1])],estilos[0]))return true;
  }
  return false;
}
function _reLpProtegerAba_(xml,tag){
  xml=xml.replace(/<sheetProtection\b[^>]*?(?:\/>|>[\s\S]*?<\/sheetProtection>)/g,'');
  if(xml.indexOf('</sheetData>')<0)throw Error('Aba Logprint sem sheetData.');
  return xml.replace('</sheetData>','</sheetData>'+tag);
}
function _reLpAplicarMascara_(xml,mascara,clonar){
  var origCols=_reLpColunas_(xml),limites={1:true,16385:true},novasCols=[];
  origCols.concat(mascara.colunas).forEach(function(c){limites[c.inicio]=true;limites[c.fim+1]=true;});
  var cortes=Object.keys(limites).map(Number).sort(function(a,b){return a-b;});
  for(var i=0;i<cortes.length-1;i++){
    var inicio=cortes[i],fim=cortes[i+1]-1,orig=_reLpColuna_(origCols,inicio),ref=_reLpColuna_(mascara.colunas,inicio);
    var p=ref?mascara.estilos[ref.estilo]:mascara.padrao;
    if(!orig&&_reLpIgual_(p,mascara.padrao))continue;
    var si=clonar(orig?orig.estilo:0,p),tag=orig?orig.xml:'<col/>';
    tag=_reLpAttr_(_reLpAttr_(_reLpAttr_(tag,'min',inicio),'max',fim),'style',si);
    novasCols.push({inicio:inicio,fim:fim,estilo:si,xml:tag});
  }
  var cols='<cols>'+novasCols.map(function(c){return c.xml;}).join('')+'</cols>';
  if(/<cols\b/.test(xml))xml=xml.replace(/<cols\b[^>]*>[\s\S]*?<\/cols>/,cols);
  else xml=xml.replace(/<sheetData\b/,cols+'<sheetData');
  return xml.replace(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g,function(row){
    var head=(/^<row\b[^>]*>/.exec(row)||[])[0],a=_reMpAtributos_(head),rn=Number(a.r),pRow=mascara.linhas[rn];
    if(_reMpBool_(a.customFormat,0)&&!pRow)throw Error('Formato de linha Logprint inesperado: '+rn);
    var novoHead=head,novoRowStyle=null;
    if(pRow){
      novoRowStyle=clonar(a.s===undefined?0:Number(a.s),pRow);
      novoHead=_reLpAttr_(_reLpAttr_(head,'s',novoRowStyle),'customFormat','1');
    }
    row=row.replace(head,novoHead);
    return row.replace(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g,function(cell){
      var ch=(/^<c\b[^>]*>/.exec(cell)||[])[0],ca=_reMpAtributos_(ch),letra=/^[A-Z]+/.exec(ca.r||'');
      if(!letra)throw Error('Celula Logprint sem referencia.');
      var cn=_reSpIdx_(letra[0])+1,ref=_reLpColuna_(mascara.colunas,cn);
      var p=mascara.celulas[ca.r]||pRow||(ref?mascara.estilos[ref.estilo]:mascara.padrao);
      if(!p.locked&&/<f(?:\s|>)/.test(cell))throw Error('Mascara de referencia desbloquearia formula Logprint: '+ca.r);
      var si=Number(ca.s||0),destino=clonar(si,p);
      if(ca.s===undefined&&destino===0)return cell;
      return cell.replace(ch,_reLpAttr_(ch,'s',destino));
    });
  });
}
function _reSpAplicarProtecaoLogprint_(partes,alteradas,referencia,restaurarTudo){
  var ref=referencia||_moldeCasaDaGraficaCalc_(),estilosRef=_reMpEstilos_(ref.textos['xl/styles.xml']);
  var original=alteradas['xl/styles.xml']===undefined?partes.mapa['xl/styles.xml'].getDataAsString('UTF-8'):alteradas['xl/styles.xml'];
  var tags=_reMpXfs_(original,'cellXfs'),estilos=_reMpEstilos_(original),cache={};
  var clonar=function(id,p){
    if(!tags[id]||!estilos[id])throw Error('Estilo Logprint ausente: '+id);
    if(_reLpIgual_(estilos[id],p))return id;
    var chave=id+'|'+p.locked+'|'+p.hidden;if(cache[chave]!==undefined)return cache[chave];
    var tag=tags[id].replace(/\sapplyProtection="[^"]*"/,'').replace(/<protection\b[^>]*?(?:\/>|>[\s\S]*?<\/protection>)/g,'');
    if(/\/>$/.test(tag))tag=tag.replace(/\/>$/,'></xf>');
    var head=/^<xf\b[^>]*>/.exec(tag)[0];tag=tag.replace(head,_reLpAttr_(head,'applyProtection','1'));
    tag=tag.replace(/(?=<extLst\b|<\/xf>)/,'<protection locked="'+p.locked+'" hidden="'+p.hidden+'"/>');
    cache[chave]=tags.length;tags.push(tag);estilos.push({locked:p.locked,hidden:p.hidden});return cache[chave];
  };
  if(!_reLpIgual_(estilos[0],estilosRef[0]))throw Error('O estilo padrao de protecao do arquivo foi alterado. Nenhum arquivo foi salvo; restaure esse estilo antes de tentar novamente.');
  var refs={},refAbas=_reLpAbas_(ref.textos['xl/workbook.xml'],ref.textos['xl/_rels/workbook.xml.rels']);
  refAbas.forEach(function(a){refs[a.nome]=a.caminho;});
  var abas=_reLpAbas_(partes.mapa['xl/workbook.xml'].getDataAsString('UTF-8'),partes.mapa['xl/_rels/workbook.xml.rels'].getDataAsString('UTF-8'));
  abas.forEach(function(a){
    var caminho=refs[a.nome],molde=caminho&&ref.abas[caminho];if(!molde)throw Error('Aba Logprint sem referencia compativel: '+a.nome);
    var protegido=/<sheetProtection\b[^>]*>/.exec(molde.fim);
    var xml=alteradas[a.caminho]===undefined?partes.mapa[a.caminho].getDataAsString('UTF-8'):alteradas[a.caminho];
    if(!protegido||!_reMpBool_(_reMpAtributos_(protegido[0]).sheet,0)){
      if(!restaurarTudo)throw Error('Aba de referencia sem protecao: '+a.nome);
      // Uma aba livre no perfil continua livre, mesmo que o arquivo escolhido
      // tenha recebido protecao manualmente depois de ter sido gerado.
      alteradas[a.caminho]=xml.replace(/<sheetProtection\b[^>]*?(?:\/>|>[\s\S]*?<\/sheetProtection>)/g,'');
      if(protegido)alteradas[a.caminho]=_reLpProtegerAba_(alteradas[a.caminho],protegido[0]);
      return;
    }
    var op=_reMpAtributos_(protegido[0]),tag='<sheetProtection password="'+_reSpSenhaHash_('PGFLOW2026')+'"';
    Object.keys(RE_MP_OPCOES).forEach(function(k){tag+=' '+k+'="'+_reMpBool_(op[k],RE_MP_OPCOES[k])+'"';});tag+='/>';
    if(restaurarTudo||_reLpTemExcecao_(molde,estilosRef))xml=_reLpAplicarMascara_(xml,_reLpMascara_(_mxGerarAba_(molde),estilosRef),clonar);
    alteradas[a.caminho]=_reLpProtegerAba_(xml,tag);
  });
  alteradas['xl/styles.xml']=original.replace(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/,function(x){
    return _reLpAttr_(/^<cellXfs\b[^>]*>/.exec(x)[0],'count',tags.length)+tags.join('')+'</cellXfs>';
  });
}
