// Modelo de revisao independente do XLSX. Preparado junto da geracao/gravacao.
// A subpasta tecnica nao participa da lista de calculadoras de Reentradas.
var RX_REVISAO_PASTA='.pgflow-revisoes';
var RX_REVISAO_SCHEMA=1;

function _rxRevisaoErro_(codigo,mensagem){var e=Error(mensagem);e.codigo=codigo;return e;}
function _rxRevisaoPasta_(criar){
  var raiz=DriveApp.getFolderById(RX_PASTA),pastas=raiz.getFoldersByName(RX_REVISAO_PASTA);
  if(pastas.hasNext())return pastas.next();
  return criar?raiz.createFolder(RX_REVISAO_PASTA):null;
}
function _rxRevisaoArquivo_(pasta,id){
  var arquivos=pasta.getFilesByName(id+'.json.gz');
  return arquivos.hasNext()?arquivos.next():null;
}
function _rxRevisaoTexto_(texto){
  return String(texto||'').replace(/&#x([0-9a-f]+);/gi,function(_,n){return String.fromCodePoint(parseInt(n,16));})
    .replace(/&#([0-9]+);/g,function(_,n){return String.fromCodePoint(Number(n));})
    .replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&gt;/g,'>').replace(/&lt;/g,'<').replace(/&amp;/g,'&');
}
function _rxRevisaoValor_(atributos,corpo,strings){
  var tipo=(/\bt="([^"]+)"/.exec(atributos)||[])[1]||'';
  var v=/<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/.exec(corpo),valor=v?_rxRevisaoTexto_(v[1]):'';
  if(tipo==='s')return /^\d+$/.test(valor)?String(strings[Number(valor)]||''):'';
  if(tipo==='inlineStr'){
    var texto='',re=/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g,m;
    while((m=re.exec(corpo))!==null)texto+=_rxRevisaoTexto_(m[1]);
    return texto;
  }
  if(tipo==='b')return valor==='1'?'VERDADEIRO':valor==='0'?'FALSO':valor;
  if(!valor&&/<f\b/.test(corpo)){
    var formula=/<f\b[^>]*>([\s\S]*?)<\/f>/.exec(corpo);
    return '='+_rxRevisaoTexto_(formula?formula[1]:'');
  }
  return valor;
}
function _rxRevisaoMontar_(partes,alteradas){
  var mapa={};(partes.partes||partes).forEach(function(p){mapa[p.getName()]=p;});
  alteradas=alteradas||{};
  function ler(nome){
    if(alteradas[nome]!==undefined)return alteradas[nome];
    if(!mapa[nome])throw Error('Parte necessaria para revisao ausente: '+nome);
    return mapa[nome].getDataAsString('UTF-8');
  }
  var nomes=_rxAbasEditaveis_(ler('xl/workbook.xml'),ler('xl/_rels/workbook.xml.rels'));
  var strings=mapa['xl/sharedStrings.xml']?_reSpStrings_(ler('xl/sharedStrings.xml')):[];
  var modelo={schema:RX_REVISAO_SCHEMA,abas:[],formulas:{},faixasFormulas:[],mesclagens:[]};
  nomes.forEach(function(n,indice){
    var xml=ler(n.caminho),aba={nome:n.nome,totalColunas:1,linhas:[]};
    var dimensao=/<dimension\b[^>]*\bref="([^"]+)"/.exec(xml);
    if(dimensao)aba.totalColunas=Math.max(1,_pfPreviaColunaNumero_(dimensao[1].split(':').pop())+1);
    var linhas=/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g,r,retangulos={};
    while((r=linhas.exec(xml))!==null){
      var numero=Number((/\br="(\d+)"/.exec(r[1])||[])[1]);
      if(!Number.isInteger(numero)||numero<1||numero>1048576)throw Error('Linha invalida no modelo de revisao.');
      var item={numero:numero,valores:{}},celulas=/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g,c,cols=[];
      while((c=celulas.exec(r[2]||''))!==null){
        var ref=(/\br="([A-Z]+)\d+"/.exec(c[1])||[])[1];if(!ref)continue;
        var coluna=_reSpIdx_(ref);if(coluna<0||coluna>=16384)throw Error('Coluna invalida no modelo de revisao.');
        aba.totalColunas=Math.max(aba.totalColunas,coluna+1);
        var corpo=c[2]||'',valor=_rxRevisaoValor_(c[1],corpo,strings);
        if(valor!=='')item.valores[coluna]=valor;
        if(/<f\b/.test(corpo))cols.push(coluna+1);
      }
      aba.linhas.push(item);
      // Colunas de formula contiguas e repetidas em linhas viram retangulos.
      cols.sort(function(a,b){return a-b;});
      for(var i=0;i<cols.length;i++){
        var inicio=cols[i],fim=inicio;
        while(i+1<cols.length&&cols[i+1]===fim+1)fim=cols[++i];
        var chave=inicio+':'+fim,anterior=retangulos[chave];
        if(anterior&&anterior.y2===numero-1)anterior.y2=numero;
        else{anterior={aba:indice,x1:inicio,x2:fim,y1:numero,y2:numero};retangulos[chave]=anterior;modelo.faixasFormulas.push(anterior);}
      }
    }
    _rxFaixasFormulas_(xml).forEach(function(f){f.aba=indice;modelo.faixasFormulas.push(f);});
    _rxFaixasMesclagens_(xml).forEach(function(f){f.aba=indice;modelo.mesclagens.push(f);});
    modelo.abas.push(aba);
  });
  return modelo;
}
function _rxRevisaoSalvar_(arquivo,modelo,jaTravado){
  var meta=arquivo.arquivo?arquivo:{id:arquivo.getId(),nome:arquivo.getName(),url:arquivo.getUrl(),atualizadoEm:arquivo.getLastUpdated().getTime()};
  modelo.arquivo={id:meta.id,nome:meta.nome,url:meta.url,atualizadoEm:meta.atualizadoEm,calculadora:/^CALCULADORA FORNECEDOR /i.test(meta.nome)};
  var nome=meta.id+'.json.gz',json=Utilities.newBlob(JSON.stringify(modelo),'application/json',meta.id+'.json');
  var blob=Utilities.gzip(json).setName(nome),lock=null,pegou=false;
  try{
    // Geração pode ocorrer em paralelo; serializa a criação da pasta/modelo.
    // Salvar/importar já detêm este mesmo lock e informam isso explicitamente.
    if(!jaTravado){
      lock=LockService.getScriptLock();pegou=lock.tryLock(3000);
      if(!pegou)throw Error('Outra revisao esta sendo preparada. Tente novamente.');
    }
    var pasta=_rxRevisaoPasta_(true),existente=_rxRevisaoArquivo_(pasta,meta.id);
    if(existente)Drive.Files.update({title:nome,mimeType:'application/gzip'},existente.getId(),blob,{newRevision:false,supportsAllDrives:true});
    else pasta.createFile(blob).setName(nome);
    return modelo;
  }finally{if(pegou)lock.releaseLock();}
}
function _rxRevisaoLer_(arq){
  var pasta=_rxRevisaoPasta_(false),arquivo=pasta&&_rxRevisaoArquivo_(pasta,arq.id);
  if(!arquivo)throw _rxRevisaoErro_('REVISAO_AUSENTE','Este arquivo ainda nao tem uma revisao preparada. Prepare a revisao uma vez para navegar sem reler o Excel.');
  var modelo;
  try{modelo=JSON.parse(Utilities.ungzip(arquivo.getBlob()).getDataAsString('UTF-8'));}
  catch(e){throw _rxRevisaoErro_('REVISAO_AUSENTE','Nao foi possivel ler a revisao preparada. Prepare-a novamente.');}
  if(!modelo||modelo.schema!==RX_REVISAO_SCHEMA||!modelo.arquivo||modelo.arquivo.id!==arq.id||!Array.isArray(modelo.abas))
    throw _rxRevisaoErro_('REVISAO_AUSENTE','O modelo de revisao precisa ser preparado novamente.');
  if(modelo.arquivo.atualizadoEm!==arq.atualizadoEm)
    throw _rxRevisaoErro_('REVISAO_DESATUALIZADA','O arquivo mudou depois da revisao preparada. Prepare a revisao novamente para conferir a versao atual.');
  return modelo;
}
function _rxRevisaoResposta_(arq,modelo,pedido){
  var consulta=_pfPreviaConsulta_(pedido),aba=modelo.abas[consulta.aba];
  if(!aba)throw Error('Aba da revisao nao encontrada.');
  var nomes=modelo.abas.map(function(a){return a.nome;}),formulas={};
  Object.keys(modelo.formulas||{}).forEach(function(k){if(k.split(':')[0]===String(consulta.aba))formulas[k]=modelo.formulas[k];});
  var faixas=(modelo.faixasFormulas||[]).filter(function(f){return f.aba===consulta.aba;}),mesclas=(modelo.mesclagens||[]).filter(function(f){return f.aba===consulta.aba;});
  var completa=null,pagina=null,limite=typeof PF_PREVIA_CONJUNTO_MAX_JSON==='number'?PF_PREVIA_CONJUNTO_MAX_JSON:1000000;
  if(pedido.porAba&&JSON.stringify({aba:aba,formulas:formulas,faixas:faixas,mesclas:mesclas}).length<=limite)
    completa={indice:consulta.aba,nome:aba.nome,totalColunas:aba.totalColunas,linhas:aba.linhas};
  if(!completa){
    var termo=consulta.buscaNormalizada,filtradas=termo?aba.linhas.filter(function(r){
      return Object.keys(r.valores).some(function(k){return String(r.valores[k]).toLocaleLowerCase('pt-BR').indexOf(termo)>=0;});
    }):aba.linhas;
    pagina=_pfPreviaResultado_(arq.nome,nomes,consulta.aba,filtradas,aba.totalColunas,aba.linhas.length,filtradas,consulta);
  }
  return {arquivo:{id:arq.id,nome:arq.nome,url:arq.url,atualizadoEm:arq.atualizadoEm,calculadora:arq.calculadora},
    previaExcel:null,nomesAbas:nomes,abaCompleta:completa,pagina:pagina,formulas:formulas,faixasFormulas:faixas,mesclagens:mesclas,revisaoPreparada:true};
}
function prepararRevisaoReentradasExcel(id,atualizadoEm){
  var lock=LockService.getScriptLock(),pegou=false;
  try{
    if(!Number.isSafeInteger(Number(atualizadoEm))||!/^\d{10,16}$/.test(String(atualizadoEm||'')))throw Error('Atualize a lista antes de preparar a revisao.');
    pegou=lock.tryLock(3000);if(!pegou)throw Error('Outra operacao esta em andamento. Tente novamente.');
    var arq=_rxArquivo_(id,atualizadoEm),partes=_reSpPartes_(arq.arquivo.getBlob());
    var modelo=_rxRevisaoMontar_(partes);_rxArquivo_(id,atualizadoEm);_rxRevisaoSalvar_(arq,modelo,true);
    return {sucesso:true,id:arq.id,atualizadoEm:arq.atualizadoEm};
  }catch(e){return {erro:e.message,codigo:e.codigo||'REVISAO_FALHOU'};}
  finally{if(pegou)lock.releaseLock();}
}
