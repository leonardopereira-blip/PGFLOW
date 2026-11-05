// Grafo de dependencias da calculadora. Funcoes puras, sem Drive/Gmail.
// Os padroes/intervalos sao produzidos por _rxRevisaoMontar_ uma unica vez.
// O grafo indica o que precisa ser recalculado; ele NAO inventa resultados de
// formulas que ainda nao possuam avaliador numerico validado.

function _rxDepsReferencias_(expressao,abaAtual){
  // O modelo guarda offsets relativos. Usa uma linha artificial para resolver
  // as referencias, conservando a diferenca entre $A$6 e A6.
  var base=100000,texto=_mxAbs_(expressao,base);
  texto=texto.replace(/"(?:[^"]|"")*"/g,'""');
  var refs=[],re=/(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?(\$?[A-Z]{1,3})(\$?\d+)?(?::(\$?[A-Z]{1,3})(\$?\d+)?)?/g,m;
  while((m=re.exec(texto))!==null){
    if(!m[4]&&!m[5]&&!m[1]&&!m[2])continue;
    var anterior=texto[m.index-1]||'',seguinte=texto[re.lastIndex]||'';
    if(/[A-Za-z0-9_.]/.test(anterior)||/[A-Za-z0-9_.]/.test(seguinte))continue;
    if(!m[4]&&!m[5]&&!m[6])continue;
    var r1=m[4]?(Number(m[4].replace('$',''))):1;
    var r2=m[6]?(Number(m[6].replace('$',''))):m[5]?1048576:r1;
    var c1=_reSpIdx_(m[3].replace('$',''));
    var c2=m[5]?_reSpIdx_(m[5].replace('$','')):c1;
    if(c1<0||c2<0||c1>16383||c2>16383)continue;
    var rel1=!!m[4]&&m[4][0]!=='$',rel2=!!m[6]&&m[6][0]!=='$';
    refs.push({aba:(m[1]||m[2]||abaAtual).replace(/''/g,"'"),
      c1:Math.min(c1,c2),c2:Math.max(c1,c2),
      r1:r1,r2:r2,rel1:rel1,rel2:rel2,base:base});
  }
  return refs;
}

function _rxDepsAbrange_(ref,linhaFonte,linhaAlvo){
  var inicio=ref.rel1?ref.r1-ref.base+linhaAlvo:ref.r1;
  var fim=ref.rel2?ref.r2-ref.base+linhaAlvo:ref.r2;
  return linhaFonte>=Math.min(inicio,fim)&&linhaFonte<=Math.max(inicio,fim);
}

function _rxDepsFaixaAlvos_(ref,linhaFonte,inicio,fim){
  // A dependência de uma coluna inteira, ou de uma faixa absoluta que contém
  // a origem, precisa ser aplicada a todas as linhas do bloco de fórmulas.
  if(!ref.rel1&&!ref.rel2){
    return linhaFonte>=ref.r1&&linhaFonte<=ref.r2?[inicio,fim]:null;
  }
  if(ref.rel1&&ref.rel2){
    var a=linhaFonte-(ref.r1-ref.base),b=linhaFonte-(ref.r2-ref.base);
    return [Math.max(inicio,Math.min(a,b)),Math.min(fim,Math.max(a,b))];
  }
  if(ref.rel1){
    if(linhaFonte>ref.r2)return null;
    return [inicio,Math.min(fim,linhaFonte-(ref.r1-ref.base))];
  }
  if(linhaFonte<ref.r1)return null;
  return [Math.max(inicio,linhaFonte-(ref.r2-ref.base)),fim];
}

function _rxDepsIndice_(modelo){
  var nomes=(modelo.abas||[]).map(function(a){return a.nome;}),porNome={};
  nomes.forEach(function(n,i){porNome[n]=i;});
  var indice=Object.create(null),refsCache={};
  (modelo.posicoesCalculo||[]).forEach(function(colunas,aba){
    Object.keys(colunas||{}).forEach(function(col){
      (colunas[col]||[]).forEach(function(faixa){
        var codigo=faixa[2],chave=aba+':'+codigo;
        if(!refsCache[chave])refsCache[chave]=_rxDepsReferencias_(modelo.padroesCalculo[codigo],nomes[aba]);
        refsCache[chave].forEach(function(ref){
          var origem=porNome[ref.aba];if(origem===undefined)return;
          var registro={aba:aba,col:Number(col),inicio:faixa[0],fim:faixa[1],ref:ref};
          for(var c=ref.c1;c<=ref.c2;c++){
            var k=origem+':'+c;(indice[k]||(indice[k]=[])).push(registro);
          }
        });
      });
    });
  });
  return indice;
}

function _rxDepsCandidatos_(indice,origem,visitar){
  var candidatos=indice[origem.aba+':'+(origem.coluna-1)]||[];
  for(var i=0;i<candidatos.length;i++){
    var c=candidatos[i],faixa=_rxDepsFaixaAlvos_(c.ref,origem.linha,c.inicio,c.fim);
    if(!faixa)continue;
    for(var rn=faixa[0];rn<=faixa[1];rn++){
      if(!_rxDepsAbrange_(c.ref,origem.linha,rn))continue;
      if(visitar({aba:c.aba,linha:rn,coluna:c.col+1})===false)return false;
    }
  }
  return true;
}

function _rxCalculoAfetados_(modelo,edicoes,limite){
  limite=limite||30000;
  var entradas=Object.create(null),indice=_rxDepsIndice_(modelo);
  var fila=[];
  (edicoes||[]).forEach(function(e){
    var k=e.aba+':'+e.linha+':'+e.coluna;
    if(!entradas[k]){entradas[k]=true;fila.push({aba:e.aba,linha:e.linha,coluna:e.coluna});}
  });
  var afetados=Object.create(null),avisos=[],total=0,esgotado=false;
  for(var pos=0;pos<fila.length&&!esgotado;pos++){
    _rxDepsCandidatos_(indice,fila[pos],function(c){
        var key=c.aba+':'+c.linha+':'+c.coluna;
        if(entradas[key]||afetados[key])return;
        if(total>=limite){esgotado=true;
          avisos.push('O limite de dependencias foi atingido; e necessario continuar a analise.');return false;}
        afetados[key]=true;total++;fila.push(c);
    });
  }
  return {celulas:Object.keys(afetados),total:total,avisos:avisos};
}
