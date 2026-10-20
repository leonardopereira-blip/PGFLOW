// Molde XLSX: mapeia um Excel modelo (layout, estilos, cores, formulas, valores,
// mesclagens, filtros, validacoes, imagens) em dados guardados no script e
// recria o .xlsx do zero, sem ler nenhuma planilha no Drive.
// Cada parte recriada e identica byte a byte a do modelo mapeado; a unica parte
// omitida e xl/calcChain.xml (ordem interna de calculo que o Excel reconstroi).
// Funcoes puras (sem servicos Google) para rodar igual no Apps Script e nos testes.
var MX_VERSAO = 1;
var MX_CALCCHAIN = 'xl/calcChain.xml';

// Linhas de formulas repetidas ficam iguais quando o numero da linha vira
// deslocamento relativo (A7 na linha 7 => A{0}). Transformacao reversivel:
// \u0001..\u0003 sao proibidos em XML 1.0, entao nunca colidem com o modelo.
function _mxRel_(texto,linha) {
  return texto.replace(/(^|[^A-Za-z0-9_.\u0001])(\$?[A-Z]{1,3})([1-9]\d{0,6})(?![\d(])/g,
    function(m,p,c,r){return p+c+'\u0001'+(Number(r)-linha)+'\u0002';});
}
function _mxAbs_(texto,linha) {
  return texto.replace(/\u0001(-?\d+)\u0002/g,function(m,d){return String(linha+Number(d));});
}

function _mxPedacos_(re,texto,nome) {
  var out=[],pos=0,m;re.lastIndex=0;
  while((m=re.exec(texto))!==null){
    if(m.index!==pos)throw Error('Molde: trecho nao mapeado em '+nome+' na posicao '+pos);
    out.push(m[0]);pos=re.lastIndex;
  }
  if(pos!==texto.length)throw Error('Molde: trecho final nao mapeado em '+nome);
  return out;
}

function _mxMapearAba_(xml,nome) {
  if(/[\u0001-\u0003]/.test(xml))throw Error('Molde: caractere reservado em '+nome);
  var i=xml.indexOf('<sheetData>'),j=xml.lastIndexOf('</sheetData>');
  var aba={inicio:xml.substring(0,i+11),fim:xml.substring(j),celulas:[],linhas:[],seq:[],valores:[]};
  var cIdx={},lIdx={};
  _mxPedacos_(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g,xml.substring(i+11,j),nome).forEach(function(row){
    var num=Number(/^<row\b[^>]*?\sr="(\d+)"/.exec(row)[1]);
    var cab=/^<row\b[^>]*>/.exec(row)[0],fechada=/\/>$/.test(cab);
    var ids=[];
    if(!fechada){
      var corpo=row.substring(cab.length,row.length-6);
      _mxPedacos_(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g,corpo,nome).forEach(function(c){
        var v=/<v>([^<]*)<\/v>/.exec(c);
        if(v){aba.valores.push(v[1]);c=c.replace(v[0],'<v>\u0003</v>');}
        var chave=_mxRel_(c,num);
        if(cIdx[chave]===undefined){cIdx[chave]=aba.celulas.length;aba.celulas.push(chave);}
        ids.push(cIdx[chave]);
      });
    }
    var h=_mxRel_(cab.replace(/(\sr=")\d+"/,'$1\u00010\u0002"'),num);
    var lchave=h+'\u0003'+ids.join(',');
    if(lIdx[lchave]===undefined){lIdx[lchave]=aba.linhas.length;aba.linhas.push({h:h,c:ids});}
    aba.seq.push(num,lIdx[lchave]);
  });
  return aba;
}

// Pre-compila "texto{d}texto{valor}" em tokens: string fixa, numero (deslocamento) ou null (valor).
function _mxCompilar_(tpl) {
  var out=[],re=/\u0001(-?\d+)\u0002|\u0003/g,pos=0,m;
  while((m=re.exec(tpl))!==null){
    if(m.index>pos)out.push(tpl.substring(pos,m.index));
    out.push(m[1]===undefined?null:Number(m[1]));pos=re.lastIndex;
  }
  if(pos<tpl.length)out.push(tpl.substring(pos));
  return out;
}

function _mxGerarAba_(aba) {
  var cel=aba.celulas.map(_mxCompilar_),lin=aba.linhas.map(function(l){return {h:_mxCompilar_(l.h),c:l.c,fechada:/\/>$/.test(l.h)};});
  var out=[aba.inicio],vi=0;
  var emitir=function(toks,num){
    for(var t=0;t<toks.length;t++){
      var k=toks[t];
      out.push(typeof k==='string'?k:k===null?aba.valores[vi++]:String(num+k));
    }
  };
  for(var s=0;s<aba.seq.length;s+=2){
    var num=aba.seq[s],l=lin[aba.seq[s+1]];
    emitir(l.h,num);
    for(var c=0;c<l.c.length;c++)emitir(cel[l.c[c]],num);
    if(!l.fechada)out.push('</row>');
  }
  if(vi!==aba.valores.length)throw Error('Molde: valores fora de sincronia');
  out.push(aba.fim);
  return out.join('');
}

function _mxSemCalcChain_(arquivos) {
  if(!arquivos.some(function(a){return a.nome===MX_CALCCHAIN;}))return arquivos;
  return arquivos.filter(function(a){return a.nome!==MX_CALCCHAIN;}).map(function(a){
    if(a.nome==='[Content_Types].xml')
      return {nome:a.nome,texto:a.texto.replace(/<Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/,'')};
    if(a.nome==='xl/_rels/workbook.xml.rels')
      return {nome:a.nome,texto:a.texto.replace(/<Relationship\b[^>]*Target="[^"]*calcChain\.xml"[^>]*\/>/,'')};
    return a;
  });
}

// arquivos: [{nome,texto}|{nome,base64}] na ordem do zip original.
// Mapeia e confere: so retorna o molde se a recriacao for identica.
function _mxMapear_(arquivos,origem) {
  var fonte=_mxSemCalcChain_(arquivos);
  var molde={versao:MX_VERSAO,origem:origem||{},ordem:[],textos:{},binarios:{},abas:{}};
  fonte.forEach(function(a){
    molde.ordem.push(a.nome);
    if(a.base64!==undefined)molde.binarios[a.nome]=a.base64;
    else if(/^xl\/worksheets\/[^/]+\.xml$/.test(a.nome)&&a.texto.indexOf('<sheetData>')>=0)
      molde.abas[a.nome]=_mxMapearAba_(a.texto,a.nome);
    else molde.textos[a.nome]=a.texto;
  });
  var gerado=_mxGerarArquivos_(molde);
  fonte.forEach(function(a,k){
    var g=gerado[k];
    if(g.nome!==a.nome||g.texto!==a.texto||g.base64!==a.base64)
      throw Error('Molde: recriacao difere do modelo em '+a.nome);
  });
  return molde;
}

function _mxGerarArquivos_(molde) {
  if(!molde||molde.versao!==MX_VERSAO)throw Error('Molde em versao desconhecida');
  return molde.ordem.map(function(n){
    if(molde.binarios[n]!==undefined)return {nome:n,base64:molde.binarios[n]};
    if(molde.abas[n])return {nome:n,texto:_mxGerarAba_(molde.abas[n])};
    if(molde.textos[n]===undefined)throw Error('Molde sem a parte '+n);
    return {nome:n,texto:molde.textos[n]};
  });
}

// ---- Apps Script ----
function _mxDecodificar_(base64) {
  var gz=Utilities.newBlob(Utilities.base64Decode(base64),'application/x-gzip');
  return JSON.parse(Utilities.ungzip(gz).getDataAsString('UTF-8'));
}
function _mxCodificar_(molde) {
  var blob=Utilities.newBlob(JSON.stringify(molde),'application/json','molde.json');
  return Utilities.base64Encode(Utilities.gzip(blob).getBytes());
}
// Mesmo formato de _reSpPartes_: {partes:[Blob], mapa:{nome:Blob}}.
function _mxPartes_(molde) {
  var partes=_mxGerarArquivos_(molde).map(function(a){
    return a.base64!==undefined
      ? Utilities.newBlob(Utilities.base64Decode(a.base64),'application/octet-stream',a.nome)
      : Utilities.newBlob(a.texto,'text/xml',a.nome);
  });
  var mapa={};partes.forEach(function(p){mapa[p.getName()]=p;});
  return {partes:partes,mapa:mapa};
}
// Para quando o Excel com novos valores for enviado pela tela (etapa futura).
function _mxMapearBlob_(blob,origem) {
  var arquivos=Utilities.unzip(blob.copyBlob().setContentType('application/zip')).map(function(p){
    var n=p.getName();
    return /\.(xml|rels|vml)$/i.test(n)
      ? {nome:n,texto:p.getDataAsString('UTF-8')}
      : {nome:n,base64:Utilities.base64Encode(p.getBytes())};
  });
  return _mxMapear_(arquivos,origem);
}
