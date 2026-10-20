// Metadados para restaurar a protecao original, sem alterar valores, formulas
// ou estilos do XLSX. A macro aplica os ranges na ordem em que sao emitidos.
var RE_MP_NAMESPACE='urn:pgflow:protecao:v1';
var RE_MP_OPCOES={sheet:0,objects:0,scenarios:0,formatCells:1,
  formatColumns:1,formatRows:1,insertColumns:1,insertRows:1,
  insertHyperlinks:1,deleteColumns:1,deleteRows:1,selectLockedCells:0,
  sort:1,autoFilter:1,pivotTables:1,selectUnlockedCells:0};

function _reMpAtributos_(tag){
  var out={},re=/([A-Za-z_][A-Za-z0-9_.:-]*)\s*=\s*(["'])([\s\S]*?)\2/g,m;
  while((m=re.exec(tag||''))!==null)out[m[1]]=_reSpUnescape_(m[3]);
  return out;
}
function _reMpBool_(valor,padrao){
  if(valor===undefined)return padrao;
  if(/^(1|true)$/i.test(valor))return 1;
  if(/^(0|false)$/i.test(valor))return 0;
  throw Error('Flag de protecao invalida: '+valor);
}
function _reMpIgual_(a,b){return a.locked===b.locked&&a.hidden===b.hidden;}
function _reMpXfs_(xml,nome){
  var parte=new RegExp('<'+nome+'\\b[^>]*>([\\s\\S]*?)<\\/'+nome+'>').exec(xml);
  if(!parte)return [];
  var out=[],re=/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g,m;
  while((m=re.exec(parte[1]))!==null)out.push(m[0]);
  return out;
}
function _reMpEstilos_(xml){
  var padrao={locked:1,hidden:0};
  var ler=function(xf,base){
    var a=_reMpAtributos_((/^<xf\b[^>]*>/.exec(xf)||[])[0]);
    if(a.applyProtection!==undefined&&!_reMpBool_(a.applyProtection,0))return base;
    var p=/<protection\b[^>]*\/>/.exec(xf);
    if(!p)return _reMpBool_(a.applyProtection,0)?padrao:base;
    var b=_reMpAtributos_(p[0]);
    return {locked:_reMpBool_(b.locked,1),hidden:_reMpBool_(b.hidden,0)};
  };
  var bases=_reMpXfs_(xml,'cellStyleXfs').map(function(x){return ler(x,padrao);});
  var estilos=_reMpXfs_(xml,'cellXfs').map(function(x){
    var a=_reMpAtributos_((/^<xf\b[^>]*>/.exec(x)||[])[0]);
    return ler(x,bases[Number(a.xfId||0)]||padrao);
  });
  if(!estilos.length)throw Error('Molde sem cellXfs para restaurar protecao');
  return estilos;
}
function _reMpEstilo_(estilos,id){
  var n=Number(id);
  if(!/^\d+$/.test(String(id))||!estilos[n])throw Error('Estilo de protecao ausente: '+id);
  return estilos[n];
}
function _reMpCaminho_(target){
  var out=[],base=target.charAt(0)==='/'?target.substring(1):'xl/'+target;
  base.split('/').forEach(function(p){if(p==='..')out.pop();else if(p&&p!=='.')out.push(p);});
  return out.join('/');
}
function _reMpIntervalos_(lista){
  // As faixas de colunas/linhas de uma worksheet sao ordenadas e nao se sobrepoem.
  var out=[];
  lista.forEach(function(x){
    var prev=out[out.length-1];
    if(prev&&prev.fim+1===x.inicio&&_reMpIgual_(prev,x))prev.fim=x.fim;
    else out.push({inicio:x.inicio,fim:x.fim,locked:x.locked,hidden:x.hidden});
  });
  return out;
}
function _reMpRetangulo_(address){
  var m=/^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(address),r;
  if(m)r={x1:_reSpIdx_(m[1])+1,x2:_reSpIdx_(m[2])+1,y1:1,y2:1048576};
  else if((m=/^([1-9]\d*):([1-9]\d*)$/.exec(address)))r={x1:1,x2:16384,y1:Number(m[1]),y2:Number(m[2])};
  else if((m=/^([A-Z]{1,3})([1-9]\d*)(?::([A-Z]{1,3})([1-9]\d*))?$/.exec(address)))
    r={x1:_reSpIdx_(m[1])+1,x2:_reSpIdx_(m[3]||m[1])+1,y1:Number(m[2]),y2:Number(m[4]||m[2])};
  if(!r||r.x1>r.x2||r.y1>r.y2||r.x2>16384||r.y2>1048576)throw Error('Range de protecao invalido: '+address);
  return r;
}
function _reMpSubtrair_(r,m){
  var x1=Math.max(r.x1,m.x1),x2=Math.min(r.x2,m.x2),y1=Math.max(r.y1,m.y1),y2=Math.min(r.y2,m.y2);
  if(x1>x2||y1>y2)return [r];
  var out=[];
  if(r.y1<y1)out.push({x1:r.x1,x2:r.x2,y1:r.y1,y2:y1-1});
  if(y2<r.y2)out.push({x1:r.x1,x2:r.x2,y1:y2+1,y2:r.y2});
  if(r.x1<x1)out.push({x1:r.x1,x2:x1-1,y1:y1,y2:y2});
  if(x2<r.x2)out.push({x1:x2+1,x2:r.x2,y1:y1,y2:y2});
  return out;
}
function _reMpAddress_(r){
  if(r.y1===1&&r.y2===1048576)return _reSpCol_(r.x1)+':'+_reSpCol_(r.x2);
  if(r.x1===1&&r.x2===16384)return r.y1+':'+r.y2;
  var a=_reSpCol_(r.x1)+r.y1,b=_reSpCol_(r.x2)+r.y2;
  return a===b?a:a+':'+b;
}
function _reMpRange_(address,p,mesclagens){
  var retangulos=[_reMpRetangulo_(address)];
  (mesclagens||[]).forEach(function(merge){
    var proximos=[];
    retangulos.forEach(function(r){proximos=proximos.concat(_reMpSubtrair_(r,merge.retangulo));});
    retangulos=proximos;
  });
  return retangulos.map(function(r){
    return '<range address="'+_reMpAddress_(r)+'" locked="'+p.locked+'" hidden="'+p.hidden+'"/>';
  }).join('');
}
function _reMpRangesAba_(xml,estilos){
  var padrao=estilos[0],colunas=[],linhas=[],celulas=[],mesclagens=[],porAncora={},porLinha={},m;
  var cols=/<cols\b[^>]*>([\s\S]*?)<\/cols>/.exec(xml),reCol=/<col\b[^>]*\/>/g;
  while(cols&&(m=reCol.exec(cols[1]))!==null){
    var a=_reMpAtributos_(m[0]);
    if(a.style===undefined)continue;
    var p=_reMpEstilo_(estilos,a.style);
    colunas.push({inicio:Number(a.min),fim:Number(a.max),locked:p.locked,hidden:p.hidden});
  }
  var mergeRe=/<mergeCell\b[^>]*\/>/g;
  while((m=mergeRe.exec(xml))!==null){
    var ma=_reMpAtributos_(m[0]),mr=/^([A-Z]{1,3})([1-9]\d*)(?::([A-Z]{1,3})([1-9]\d*))?$/.exec(ma.ref||'');
    if(!mr)throw Error('Mesclagem de protecao invalida: '+ma.ref);
    var ancora=mr[1]+mr[2],col=_reSpIdx_(mr[1])+1,base=padrao;
    for(var j=0;j<colunas.length;j++)if(colunas[j].inicio<=col&&colunas[j].fim>=col){base=colunas[j];break;}
    var merge={address:ma.ref,locked:base.locked,hidden:base.hidden,retangulo:_reMpRetangulo_(ma.ref)};
    mesclagens.push(merge);porAncora[ancora]=merge;
    if(!porLinha[Number(mr[2])])porLinha[Number(mr[2])]=[];
    porLinha[Number(mr[2])].push(merge);
  }
  var reRow=/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g;
  while((m=reRow.exec(xml))!==null){
    var row=m[0],h=_reMpAtributos_((/^<row\b[^>]*>/.exec(row)||[])[0]);
    var rn=Number(h.r),pRow=null;
    if(_reMpBool_(h.customFormat,0)){
      pRow=_reMpEstilo_(estilos,h.s===undefined?'0':h.s);
      linhas.push({inicio:rn,fim:rn,locked:pRow.locked,hidden:pRow.hidden});
      (porLinha[rn]||[]).forEach(function(merge){merge.locked=pRow.locked;merge.hidden=pRow.hidden;});
    }
    var reCell=/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g,c,anterior=null;
    while((c=reCell.exec(row))!==null){
      var ca=_reMpAtributos_((/^<c\b[^>]*>/.exec(c[0])||[])[0]);
      var ref=/^([A-Z]{1,3})([1-9]\d*)$/.exec(ca.r||'');
      if(!ref||Number(ref[2])!==rn)throw Error('Referencia de protecao invalida: '+ca.r);
      // Celula explicita sem s usa XF0. Heranca de linha/coluna vale somente
      // para coordenadas sem elemento <c>, como confirmado no Excel nativo.
      var cn=_reSpIdx_(ref[1])+1,pCell=_reMpEstilo_(estilos,ca.s===undefined?'0':ca.s),pBase=pRow;
      if(porAncora[ca.r]){porAncora[ca.r].locked=pCell.locked;porAncora[ca.r].hidden=pCell.hidden;}
      if(!pBase){
        pBase=padrao;
        for(var k=0;k<colunas.length;k++)if(colunas[k].inicio<=cn&&colunas[k].fim>=cn){pBase=colunas[k];break;}
      }
      if(_reMpIgual_(pCell,pBase))continue;
      if(anterior&&anterior.fim+1===cn&&_reMpIgual_(anterior,pCell))anterior.fim=cn;
      else{
        anterior={inicio:cn,fim:cn,linha:rn,ultimaLinha:rn,locked:pCell.locked,hidden:pCell.hidden};
        celulas.push(anterior);
      }
    }
  }
  var out=[];
  _reMpIntervalos_(colunas).forEach(function(c){
    if(!_reMpIgual_(c,padrao))out.push(_reMpRange_(_reSpCol_(c.inicio)+':'+_reSpCol_(c.fim),c,mesclagens));
  });
  _reMpIntervalos_(linhas).forEach(function(l){out.push(_reMpRange_(l.inicio+':'+l.fim,l,mesclagens));});
  // Retangulos de mesmas colunas e protecao em linhas consecutivas viram um range.
  var porFaixa={},compactadas=[];
  celulas.forEach(function(c){
    var chave=[c.inicio,c.fim,c.locked,c.hidden].join('|'),prev=porFaixa[chave];
    if(prev&&prev.ultimaLinha+1===c.linha)prev.ultimaLinha=c.linha;
    else{compactadas.push(c);porFaixa[chave]=c;}
  });
  compactadas.forEach(function(c){
    var inicio=_reSpCol_(c.inicio)+c.linha,fim=_reSpCol_(c.fim)+c.ultimaLinha;
    out.push(_reMpRange_(inicio===fim?inicio:inicio+':'+fim,c,mesclagens));
  });
  mesclagens.forEach(function(merge){
    out.push('<merge address="'+merge.address+'" locked="'+merge.locked+'" hidden="'+merge.hidden+'"/>');
  });
  return out.join('');
}

function _reSpMapaProtecao_(partes,alteradas){
  var ler=function(nome,final){
    if(final&&alteradas&&alteradas[nome]!==undefined)return alteradas[nome];
    var blob=partes.mapa[nome];
    if(!blob)throw Error('Parte ausente para mapa de protecao: '+nome);
    return blob.getDataAsString('UTF-8');
  };
  var workbook=ler('xl/workbook.xml'),rels=ler('xl/_rels/workbook.xml.rels'),porId={},m;
  var relRe=/<Relationship\b[^>]*\/>/g;
  while((m=relRe.exec(rels))!==null){var r=_reMpAtributos_(m[0]);porId[r.Id]=r.Target;}
  var estilos=null,out=['<protection xmlns="'+RE_MP_NAMESPACE+'" version="1">'];
  var sheetRe=/<sheet\b[^>]*>/g;
  while((m=sheetRe.exec(workbook))!==null){
    var s=_reMpAtributos_(m[0]),target=porId[s['r:id']];
    if(!target)throw Error('Aba sem relacionamento para mapa: '+s.name);
    var caminho=_reMpCaminho_(target),original=ler(caminho),tag=/<sheetProtection\b[^>]*>/.exec(original);
    if(!tag)continue;
    var opcoes=_reMpAtributos_(tag[0]);
    if(!_reMpBool_(opcoes.sheet,0))continue;
    if(!estilos)estilos=_reMpEstilos_(ler('xl/styles.xml'));
    out.push('<sheet name="'+_reSpEscape_(s.name)+'" defaultLocked="'+estilos[0].locked+'" defaultHidden="'+estilos[0].hidden+'"');
    Object.keys(RE_MP_OPCOES).forEach(function(k){out.push(' '+k+'="'+_reMpBool_(opcoes[k],RE_MP_OPCOES[k])+'"');});
    out.push('>',_reMpRangesAba_(ler(caminho,true),estilos),'</sheet>');
  }
  out.push('</protection>');
  return out.join('');
}
