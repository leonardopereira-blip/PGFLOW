// Avaliacao pura das expressoes da calculadora a partir de seus valores
// atuais. O resultado so e aceito quando todas as dependencias sao resolvidas.
// Qualquer construcao nao suportada devolve erro explicito, nunca um preco
// inferido a partir do cache antigo do arquivo.

function _rxAvTokens_(origem){
  var s=String(origem||'').replace(/^=/,'').replace(/(?:_xlfn\.|_xludf\.)/gi,''),out=[],p=0;
  while(p<s.length){
    var c=s[p],tail=s.substring(p),m;
    if(/\s/.test(c)){p++;continue;}
    if(c==='"'){
      var texto='';p++;
      while(p<s.length){if(s[p]==='"'){if(s[p+1]==='"'){texto+='"';p+=2;continue;}p++;break;}
        texto+=s[p++];}
      out.push({tipo:'texto',valor:texto});continue;
    }
    if(c==="'"){
      var aba='';p++;
      while(p<s.length){if(s[p]==="'"){if(s[p+1]==="'"){aba+="'";p+=2;continue;}p++;break;}
        aba+=s[p++];}
      out.push({tipo:'aba',valor:aba});continue;
    }
    m=/^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/.exec(tail);
    if(m){out.push({tipo:'numero',valor:Number(m[0])});p+=m[0].length;continue;}
    m=/^[A-Za-z_$][A-Za-z0-9_.$]*/.exec(tail);
    if(m){out.push({tipo:'ident',valor:m[0]});p+=m[0].length;continue;}
    m=/^(?:<>|<=|>=)/.exec(tail);
    if(m){out.push({tipo:m[0]});p+=2;continue;}
    if('+-*/^&%=<>(),:!'.indexOf(c)>=0){out.push({tipo:c});p++;continue;}
    throw Error('Expressao nao suportada no caractere '+c+'.');
  }
  out.push({tipo:'fim'});return out;
}

function _rxAvCompilar_(texto){
  var toks=_rxAvTokens_(texto),pos=0;
  function atual(){return toks[pos];}
  function pegar(tipo){if(atual().tipo!==tipo)throw Error('Expressao invalida: esperado '+tipo+'.');return toks[pos++];}
  var precedencia={'=':1,'<>':1,'<':1,'>':1,'<=':1,'>=':1,'&':2,'+':3,'-':3,'*':4,'/':4,'^':5};
  function primaria(){
    var t=atual();
    if(t.tipo==='+'||t.tipo==='-'){pos++;return {tipo:'unario',op:t.tipo,valor:expressao(6)};}
    if(t.tipo==='('){pos++;var e=expressao(0);pegar(')');return e;}
    if(t.tipo==='numero'||t.tipo==='texto'){pos++;return {tipo:'literal',valor:t.valor};}
    if(t.tipo==='ident'||t.tipo==='aba'){
      pos++;var nome=t.valor;
      if(t.tipo==='ident'&&atual().tipo==='('){
        pos++;var args=[];
        while(atual().tipo!==')'){
          if(atual().tipo===','){args.push({tipo:'literal',valor:''});pos++;continue;}
          args.push(expressao(0));
          if(atual().tipo!==',')break;pos++;
        }
        pegar(')');return {tipo:'funcao',nome:nome.toUpperCase(),args:args};
      }
      if(atual().tipo==='!'){pos++;var cell=pegar('ident');
        return referencia(nome,cell.valor);}
      if(/^TRUE$/i.test(nome))return {tipo:'literal',valor:true};
      if(/^FALSE$/i.test(nome))return {tipo:'literal',valor:false};
      return referencia(null,nome);
    }
    throw Error('Expressao incompleta ou nao suportada.');
  }
  function referencia(aba,celula){
    if(!/^\$?[A-Z]{1,3}(?:\$?\d+)?$/i.test(celula))throw Error('Referencia desconhecida '+celula+'.');
    var fim=celula;
    if(atual().tipo===':'){
      pos++;
      if(atual().tipo==='aba'||(atual().tipo==='ident'&&toks[pos+1].tipo==='!')){
        var outra=toks[pos++].valor;pegar('!');if(aba!==null&&outra!==aba)throw Error('Intervalo entre abas nao suportado.');aba=outra;
      }
      fim=pegar('ident').valor;
    }
    return {tipo:'ref',aba:aba,inicio:celula,fim:fim};
  }
  function expressao(min){
    var e=primaria();
    while(atual().tipo==='%'&&min<=6){pos++;e={tipo:'unario',op:'%',valor:e};}
    while(precedencia[atual().tipo]!==undefined&&precedencia[atual().tipo]>=min){
      var op=atual().tipo,p=precedencia[op];pos++;
      e={tipo:'binario',op:op,esq:e,dir:expressao(op==='^'?p:p+1)};
    }
    return e;
  }
  var raiz=expressao(0);if(atual().tipo!=='fim')throw Error('Expressao incompleta em '+atual().tipo+'.');
  return raiz;
}

function _rxAvReferencia_(ref,abaAtual){
  function ponto(s){var m=/^\$?([A-Z]{1,3})(?:\$?(\d+))?$/i.exec(s);
    if(!m)throw Error('Endereco de celula invalido.');
    return {c:_reSpIdx_(m[1].toUpperCase()),r:m[2]?Number(m[2]):null};}
  var a=ponto(ref.inicio),b=ponto(ref.fim);
  return {aba:ref.aba||abaAtual,c1:Math.min(a.c,b.c),c2:Math.max(a.c,b.c),r1:a.r||1,r2:b.r||1048576};
}

function _rxAvNumero_(v){
  if(v===null||v===undefined||v==='')return 0;
  if(typeof v==='boolean')return v?1:0;
  if(typeof v==='number')return v;
  var s=String(v).trim();if(!s)return 0;
  if(!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/.test(s))throw Error('Valor nao numerico: '+s.slice(0,30));
  return Number(s);
}
function _rxAvIgual_(a,b){
  if((a===null||a===undefined||a==='')&&(b===null||b===undefined||b===''))return true;
  if(typeof a==='string'&&typeof b==='string')return a.toLowerCase()===b.toLowerCase();
  var na=Number(a),nb=Number(b);if(isFinite(na)&&isFinite(nb))return na===nb;
  return String(a).toLowerCase()===String(b).toLowerCase();
}
function _rxAvArredondar_(n,d,cima){
  var escala=Math.pow(10,_rxAvNumero_(d)),x=n*escala;
  return (cima?Math.sign(x)*Math.ceil(Math.abs(x)-1e-12):Math.sign(x)*Math.floor(Math.abs(x)+.5+1e-12))/escala;
}
function _rxAvCrit_(valor,criterio){
  var m=/^(<=|>=|<>|=|<|>)(.*)$/.exec(String(criterio)),op=m?m[1]:'=',alvo=m?m[2]:criterio;
  if(op==='=')return _rxAvIgual_(valor,alvo);
  if(op==='<>')return !_rxAvIgual_(valor,alvo);
  var va=Number(valor),vb=Number(alvo),a=isFinite(va)&&isFinite(vb)?va:String(valor).toLowerCase();
  var b=isFinite(va)&&isFinite(vb)?vb:String(alvo).toLowerCase();
  return op==='<'?a<b:op==='>'?a>b:op==='<='?a<=b:a>=b;
}

// O provedor recebe (nomeAba, linha, coluna) e devolve o valor atual.
// Linhas e colunas sao 1-based; os intervalos sao lidos apenas ate o maximo
// ocupado informado por maxLinha, sem materializar 1.048.576 posicoes.
function _rxAvExecutar_(arvore,contexto){
  var atual=contexto.aba,ler=contexto.ler,maxLinha=contexto.maxLinha;
  function resolve(no){
    if(no.tipo==='literal')return no.valor;
    if(no.tipo==='ref'){
      var ref=_rxAvReferencia_(no,atual);
      if(ref.r1===ref.r2&&ref.c1===ref.c2)return ler(ref.aba,ref.r1,ref.c1+1);
      return {range:ref};
    }
    if(no.tipo==='unario'){
      var n=_rxAvNumero_(resolve(no.valor));return no.op==='%'?n/100:no.op==='-'?-n:n;
    }
    if(no.tipo==='binario'){
      var a=resolve(no.esq),b=resolve(no.dir);
      if(no.op==='&')return String(a??'')+String(b??'');
      if(['=','<>','<','>','<=','>='].indexOf(no.op)>=0)return _rxAvCrit_(a,no.op+String(b??''));
      a=_rxAvNumero_(a);b=_rxAvNumero_(b);
      if(no.op==='+')return a+b;if(no.op==='-')return a-b;if(no.op==='*')return a*b;
      if(no.op==='/'){if(!b)throw Error('#DIV/0!');return a/b;}
      if(no.op==='^')return Math.pow(a,b);
    }
    if(no.tipo!=='funcao')throw Error('Operador desconhecido.');
    var fn=no.nome,arg=no.args||[],av=i=>i<arg.length?resolve(arg[i]):'',num=i=>_rxAvNumero_(av(i));
    function faixa(i){var v=av(i);if(!v||!v.range)throw Error(fn+' exige intervalo de celulas.');return v.range;}
    function rows(ref){return Math.min(ref.r2,Math.max(ref.r1,maxLinha(ref.aba)));}
    function somar(v){
      if(!v||!v.range)return _rxAvNumero_(v);
      var ref=v.range,s=0;for(var r=ref.r1;r<=rows(ref);r++)for(var c=ref.c1;c<=ref.c2;c++){
        var raw=ler(ref.aba,r,c+1);try{s+=_rxAvNumero_(raw);}catch(e){} }
      return s;
    }
    if(fn==='IF')return av(0)?av(1):av(2);
    if(fn==='IFNA'||fn==='IFERROR'){try{return av(0);}catch(e){
      if(fn==='IFERROR'||String(e.message)==='#N/A')return av(1);throw e;}}
    if(fn==='AND')return arg.every((_,i)=>!!av(i));
    if(fn==='OR')return arg.some((_,i)=>!!av(i));
    if(fn==='SUM'||fn==='SOMA')return arg.reduce((s,_,i)=>s+somar(av(i)),0);
    if(fn==='ROUND')return _rxAvArredondar_(num(0),num(1),false);
    if(fn==='ROUNDUP')return _rxAvArredondar_(num(0),num(1),true);
    if(fn==='CONCATENATE'||fn==='CONCAT')return arg.map((_,i)=>String(av(i)??'')).join('');
    if(fn==='LEFT')return String(av(0)).slice(0,arg.length>1?Math.max(0,num(1)):1);
    if(fn==='RIGHT'){var s=String(av(0)),n=arg.length>1?Math.max(0,num(1)):1;return s.slice(-n);}
    if(fn==='COUNTIF'||fn==='SUMIF'){
      var ran=faixa(0),crit=av(1),sum=fn==='SUMIF'&&arg.length>2?faixa(2):ran,cont=0;
      for(var rr=ran.r1;rr<=rows(ran);rr++)for(var cc=ran.c1;cc<=ran.c2;cc++){
        if(_rxAvCrit_(ler(ran.aba,rr,cc+1),crit)){
          if(fn==='COUNTIF')cont++;
          else cont+=_rxAvNumero_(ler(sum.aba,sum.r1+(rr-ran.r1),sum.c1+(cc-ran.c1)+1));
        }
      }return cont;
    }
    if(fn==='MATCH'||fn==='XLOOKUP'){
      var valor=av(0),procurar=faixa(1),retornar=fn==='XLOOKUP'?faixa(2):null;
      var modo=fn==='MATCH'?(arg.length>2?num(2):1):(arg.length>4?num(4):0);
      var horizontal=procurar.r1===procurar.r2,limite=horizontal?procurar.c2:rows(procurar);
      var achou=-1,indice=0;
      for(var j=horizontal?procurar.c1:procurar.r1;j<=limite;j++){
        var x=horizontal?ler(procurar.aba,procurar.r1,j+1):ler(procurar.aba,j,procurar.c1+1);
        if(_rxAvIgual_(x,valor)){achou=j;indice=j-(horizontal?procurar.c1:procurar.r1)+1;break;}
        if(modo!==0&&Number(x)<=Number(valor)){achou=j;indice=j-(horizontal?procurar.c1:procurar.r1)+1;}
      }
      if(achou<0){if(fn==='XLOOKUP'&&arg.length>3)return av(3);throw Error('#N/A');}
      if(fn==='MATCH')return indice;
      var bruto=horizontal?ler(retornar.aba,retornar.r1,retornar.c1+achou-procurar.c1+1):
        ler(retornar.aba,retornar.r1+achou-procurar.r1,retornar.c1+1);
      return bruto===''?0:bruto;
    }
    if(fn==='INDEX'){
      var rg=faixa(0),r=num(1),c=arg.length>2?num(2):1;
      var resultado=ler(rg.aba,rg.r1+r-1,rg.c1+c);
      return resultado===''?0:resultado;
    }
    if(fn==='SMALL'){
      var f=faixa(0),v=[];for(var k=f.r1;k<=rows(f);k++){
        var z=ler(f.aba,k,f.c1+1);if(z!==''&&z!=null&&isFinite(Number(z)))v.push(Number(z));}
      v.sort((a,b)=>a-b);return v[num(1)-1]??'';
    }
    throw Error('Funcao ainda nao validada: '+fn);
  }
  return resolve(arvore);
}

// Recebe edicoes em coordenadas originais (1-based) e produz somente deltas.
// Nao grava arquivos e nao modifica o modelo: o usuario pode recalcular
// quantas vezes desejar, conferindo a tela antes de salvar.
function _rxCalculoAplicar_(modelo,edicoes,limite){
  if(!modelo||!modelo.padroesCalculo||!modelo.posicoesCalculo)
    throw Error('O modelo precisa ser preparado com dependencias para recalcular.');
  var nomes=modelo.abas.map(function(a){return a.nome;}),porNome={},linhas=[],maximos=[];
  nomes.forEach(function(n,i){
    porNome[n]=i;var cache=Object.create(null),max=0;
    (modelo.abas[i].linhas||[]).forEach(function(l){cache[l.numero]=l.valores||{};max=Math.max(max,l.numero);});
    linhas.push(cache);maximos.push(max);
  });
  var entradas=Object.create(null);
  edicoes.forEach(function(e){
    if(!Number.isInteger(e.aba)||e.aba<0||e.aba>=nomes.length||
        !Number.isInteger(e.linha)||e.linha<1||!Number.isInteger(e.coluna)||e.coluna<1)
      throw Error('Coordenada de edicao invalida.');
    entradas[e.aba+':'+e.linha+':'+e.coluna]=e.valor;
  });
  var indice=_rxDepsIndice_(modelo),cache=Object.create(null),pilha=Object.create(null),falhas=Object.create(null);
  var recalculando=Object.create(null),fila=edicoes.map(function(e){return {aba:e.aba,linha:e.linha,coluna:e.coluna};});
  var atingidos=Object.create(null),contador=0,avisos=[],maximo=limite||1000;
  function expressao(aba,r,c){
    var grupos=modelo.posicoesCalculo[aba]||{},lista=grupos[c-1]||[];
    for(var i=0;i<lista.length;i++){var f=lista[i];
      if(r>=f[0]&&r<=f[1])return _mxAbs_(modelo.padroesCalculo[f[2]],r);}
    return null;
  }
  function ler(nome,r,c){
    var aba=porNome[nome];if(aba===undefined)throw Error('Aba de referencia ausente: '+nome);
    var key=aba+':'+r+':'+c;
    if(Object.prototype.hasOwnProperty.call(entradas,key))return entradas[key];
    if(Object.prototype.hasOwnProperty.call(cache,key))return cache[key];
    if(recalculando[key]){
      if(pilha[key])throw Error('Dependencia circular em '+nome+'!'+c+':'+r);
      var f=expressao(aba,r,c);if(!f)throw Error('Expressao nao encontrada para '+key);
      pilha[key]=true;
      try{var ast=_rxAvCompilar_(f),res=_rxAvExecutar_(ast,{aba:nome,ler:ler,
          maxLinha:function(outra){return maximos[porNome[outra]]||0;}});
        if(typeof res==='number'&&!isFinite(res))throw Error('Resultado numerico invalido.');
        cache[key]=res;return res;
      }finally{delete pilha[key];}
    }
    var valores=linhas[aba][r]||{},v=valores[c-1];
    if(typeof v==='string'&&v[0]==='='){
      var outra=expressao(aba,r,c);
      if(!outra)throw Error('Formula sem resultado disponivel em '+key);
      var calc=_rxAvExecutar_(_rxAvCompilar_(outra),{aba:nome,ler:ler,
        maxLinha:function(n){return maximos[porNome[n]]||0;}});
      cache[key]=calc;return calc;
    }
    return v===undefined?'':v;
  }
  function diferente(antigo,novo){
    if(typeof antigo==='number'||typeof novo==='number'){
      var n=Number(antigo),m=Number(novo);
      if(isFinite(n)&&isFinite(m))return Math.abs(n-m)>1e-9*Math.max(1,Math.abs(n),Math.abs(m));
    }
    return !_rxAvIgual_(antigo,novo);
  }
  for(var pos=0;pos<fila.length&&contador<maximo;pos++){
    _rxDepsCandidatos_(indice,fila[pos],function(c){
      var key=c.aba+':'+c.linha+':'+c.coluna;
      if(Object.prototype.hasOwnProperty.call(entradas,key))return;
      if(contador>=maximo)return false;
      var antigo=Object.prototype.hasOwnProperty.call(cache,key)?cache[key]:
        (linhas[c.aba][c.linha]||{})[c.coluna-1]??'';
      delete cache[key];recalculando[key]=true;
      try{
        var novo=ler(nomes[c.aba],c.linha,c.coluna);
        delete falhas[key];
        if(diferente(antigo,novo)){
          atingidos[key]=c;contador++;fila.push(c);
        }
      }catch(e){
        var erro=String(e.message||e).slice(0,150);
        if(!falhas[key]||falhas[key].erro!==erro){
          falhas[key]={aba:c.aba,linha:c.linha,coluna:c.coluna,erro:erro};
          contador++;fila.push(c);
        }
        delete cache[key];
      }finally{delete recalculando[key];}
    });
  }
  if(contador>=maximo)avisos.push('O limite de processamento foi atingido. A revisao ainda tem dependencias pendentes.');
  var resultados=Object.keys(atingidos).filter(function(k){return !falhas[k];}).map(function(k){
    var c=atingidos[k];return {aba:c.aba,linha:c.linha,coluna:c.coluna,valor:cache[k]};
  });
  return {resultados:resultados,pendentes:Object.keys(falhas).map(function(k){return falhas[k];}),
    avisos:avisos,totalDependencias:contador};
}
