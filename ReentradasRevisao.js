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
    // O MIME existente e preservado; especifica-lo no metadata de um update
    // pode fazer a API recusar a revisao com "mime type change is forbidden".
    if(existente)Drive.Files.update({title:nome},existente.getId(),blob,{newRevision:false,supportsAllDrives:true});
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
  // Os resultados <v> de formulas do molde podem pertencer ao chamado antigo.
  // So projeta formulas de tiragem conhecidas; nao simula os custos do Excel.
  _rxRevisaoProjetarTotais_(modelo);
  return modelo;
}
function _rxRevisaoProjetarTotais_(modelo){
  var calculados=0,corrigidos=0;
  (modelo.abas||[]).forEach(function(aba,indice){
    if(!/^CALC[\s-]*Geral$/i.test(String(aba.nome||'')))return;
    var linhas=aba.linhas||[],titulo=(linhas.find(function(r){return r.numero===3;})||{}).valores||{};
    var sub=(linhas.find(function(r){return r.numero===4;})||{}).valores||{};
    var colSku=Number(Object.keys(titulo).find(function(k){return /C[ÓO]DIGO\s*SKU/i.test(String(titulo[k]));}));
    var colTotal=Number(Object.keys(titulo).find(function(k){return /^TOTAL$/i.test(String(titulo[k]).trim())&&
      /Exemplares/i.test(String(sub[k]||''));}));
    if(!Number.isInteger(colSku)||!Number.isInteger(colTotal)||colTotal<=colSku+2||colTotal>colSku+8)return;
    var formula=(modelo.faixasFormulas||[]).filter(function(f){return f.aba===indice&&colTotal+1>=f.x1&&colTotal+1<=f.x2;});
    linhas.forEach(function(r){
      if(r.numero<=5||!formula.some(function(f){return r.numero>=f.y1&&r.numero<=f.y2;}))return;
      var v=r.valores||{},sku=String(v[colSku]||'').trim();if(!sku||/^=/.test(sku))return;
      var soma=0;
      for(var c=colSku+2;c<colTotal;c++){
        var valor=String(v[c]===undefined?'0':v[c]).trim();
        if(valor!==''&&!/^-?\d+(?:\.\d+)?$/.test(valor))return;
        soma+=Number(valor||0);
      }
      if(!isFinite(soma)||soma<0)return;
      calculados++;
      if(String(v[colTotal]||'')!==String(soma)){v[colTotal]=String(soma);corrigidos++;}
    });
  });
  return {calculados:calculados,corrigidos:corrigidos};
}
function recalcularTotaisReentradasExcel(id,atualizadoEm){
  try{
    var arq=_rxArquivo_(id,atualizadoEm);
    if(!arq.calculadora)throw Error('Selecione a calculadora para recalcular os totais.');
    var revisao=_rxRevisaoLer_(arq),info=_rxRevisaoProjetarTotais_(revisao);
    // A leitura ja projetou os totais. Conta apenas linhas cobertas pelo contrato.
    if(!info.calculados)throw Error('Nao foi encontrada formula de total de exemplares compativel.');
    return {sucesso:true,id:arq.id,atualizadoEm:arq.atualizadoEm,
      totaisRecalculados:info.calculados,
      aviso:'As tiragens foram recalculadas pelo script. Precos e outras formulas do Excel continuam sujeitos a recalculo no arquivo.'};
  }catch(e){return {erro:e.message||String(e)};}
}
// Organiza apenas a leitura visual. Coordenadas e conteudo do XLSX permanecem intactos.
function _rxRevisaoLayout_(aba,modelo,indice){
  var linhas=aba.linhas||[],nome=String(aba.nome||''),mapa={};
  linhas.forEach(function(r){mapa[r.numero]=r.valores||{};});
  var calculadora=/^CALC(?:[\s-]|$)/i.test(nome);
  var especificacoes=/^(?:NC - )?PPG[ -]+Especifica/i.test(nome);
  var merges=(modelo.mesclagens||[]).filter(function(m){return m.aba===indice&&m.y1<=4&&m.y2<=5;});
  // Nas especificacoes e no resumo, linhas 3/4 sao TITULO/SUBTITULO.
  // Nunca devem aparecer como registros editaveis da calculadora.
  var cabecalhoAte=calculadora?5:especificacoes?4:(merges.length&&linhas.some(function(r){return r.numero>2;})?2:0);
  var campos=(calculadora||especificacoes)?(mapa[3]||{}):(mapa[5]||mapa[4]||mapa[3]||mapa[2]||mapa[1]||{});
  var skuCol=-1;
  Object.keys(campos).some(function(k){
    var valor=String(campos[k]||'').toUpperCase();
    if(/\b(SKU|COD(?:IGO)?(?: DO)? (?:ITEM|PRODUTO)|ITEM|PRODUTO)\b/.test(valor)){skuCol=Number(k);return true;}
    return false;
  });
  var ultima=0,minCol=Infinity,maxCol=-1,semFormulas=/^=/,ativas=[];
  linhas.forEach(function(r){
    if(r.numero<=cabecalhoAte)return;
    var vals=r.valores||{},chaves=Object.keys(vals);
    var colSku=skuCol>=0?String(vals[skuCol]??'').trim():'';
    var contemDado=skuCol>=0?!!colSku&&!semFormulas.test(colSku):
      chaves.some(function(k){var v=String(vals[k]??'').trim();return v!==''&&!semFormulas.test(v)&&!/^0(?:[.,]0+)?$/.test(v);});
    if(!contemDado)return;
    ativas.push(r);
    ultima=Math.max(ultima,r.numero);
    chaves.forEach(function(k){
      var v=String(vals[k]??'').trim();
      if(v!==''&&!semFormulas.test(v)&&!/^0(?:[.,]0+)?$/.test(v)){
        minCol=Math.min(minCol,Number(k));maxCol=Math.max(maxCol,Number(k));
      }
    });
  });
  if(!ultima){
    linhas.forEach(function(r){if(Object.keys(r.valores||{}).length)ultima=Math.max(ultima,r.numero);});
  }
  if(maxCol<0){
    linhas.forEach(function(r){if(r.numero>ultima)return;
      Object.keys(r.valores||{}).forEach(function(k){if(String(r.valores[k]||'').trim()){minCol=Math.min(minCol,Number(k));maxCol=Math.max(maxCol,Number(k));}});
    });
  }
  if(maxCol<0){minCol=0;maxCol=0;}
  // Campos com titulo real e formula em um ITEM ativo precisam continuar
  // visiveis, mesmo quando o resultado calculado em cache for zero/vazio.
  // Nao amplia ate as colunas tecnicas reservadas sem titulo (ex.: GA/GB/GC).
  var titulo=mapa[3]||{},subtitulo=mapa[4]||{};
  if(calculadora||especificacoes){
    ativas.forEach(function(r){
      Object.keys(r.valores||{}).forEach(function(k){
        var c=Number(k),v=String(r.valores[k]??'').trim();
        if(c<minCol||!semFormulas.test(v))return;
        var rotulo=String(titulo[k]??'').trim(),sub=String(subtitulo[k]??'').trim();
        if(rotulo||sub)maxCol=Math.max(maxCol,c);
      });
    });
  }
  // Mesclagens e cabecalhos nao ampliam as margens vazias do quadro.
  var indices=[];for(var c=minCol;c<=maxCol;c++)indices.push(c);
  return {indices:indices,cabecalhoAte:cabecalhoAte,ultimaLinha:ultima,
    topo:linhas.filter(function(r){return r.numero<=cabecalhoAte;}).map(function(r){return {numero:r.numero,valores:r.valores};}),
    mesclagens:merges};
}
function _rxRevisaoResposta_(arq,modelo,pedido){
  pedido=pedido||{};
  var consulta=_pfPreviaConsulta_(pedido),aba=modelo.abas[consulta.aba];
  if(!aba)throw Error('Aba da revisao nao encontrada.');
  var nomes=modelo.abas.map(function(a){return a.nome;}),formulas={};
  Object.keys(modelo.formulas||{}).forEach(function(k){if(k.split(':')[0]===String(consulta.aba))formulas[k]=modelo.formulas[k];});
  var faixas=(modelo.faixasFormulas||[]).filter(function(f){return f.aba===consulta.aba;}),mesclas=(modelo.mesclagens||[]).filter(function(f){return f.aba===consulta.aba;});
  var quadro=_rxRevisaoLayout_(aba,modelo,consulta.aba),completa=null,pagina=null;
  var limite=typeof PF_PREVIA_CONJUNTO_MAX_JSON==='number'?PF_PREVIA_CONJUNTO_MAX_JSON:1000000;
  if(pedido.porAba&&!pedido.somenteNaoEditorial&&JSON.stringify({aba:aba,formulas:formulas,faixas:faixas,mesclas:mesclas}).length<=limite)
    completa={indice:consulta.aba,nome:aba.nome,totalColunas:aba.totalColunas,linhas:aba.linhas};
  if(!completa){
    var termo=consulta.buscaNormalizada,filtradas=aba.linhas.filter(function(r){
      if(r.numero<=quadro.cabecalhoAte||r.numero>quadro.ultimaLinha)return false;
      if(pedido.somenteNaoEditorial&&!(modelo.naoEditoriais||[]).some(function(x){return x.aba===consulta.aba&&x.linha===r.numero;}))return false;
      if(!pedido.somenteNaoEditorial&&/^(?:NC - )?PPG[ -]+Especifica/i.test(aba.nome)&&
        (modelo.naoEditoriais||[]).some(function(x){return x.aba===consulta.aba&&x.linha===r.numero;}))return false;
      return !termo||Object.keys(r.valores).some(function(k){
        return String(r.valores[k]).toLocaleLowerCase('pt-BR').indexOf(termo)>=0;});
    });
    var inicio=Math.min(Math.max(0,consulta.linha),filtradas.length);
    pagina={arquivo:arq.nome,abas:nomes,aba:consulta.aba,linha:inicio,coluna:0,
      totalLinhas:aba.linhas.length,totalFiltradas:filtradas.length,totalColunas:aba.totalColunas,
      indices:quadro.indices,colunas:quadro.indices.map(function(c){return _rxRevisaoLetra_(c+1);}),
      linhas:filtradas.slice(inicio,inicio+50).map(function(r){return {numero:r.numero,
        celulas:quadro.indices.map(function(c){return String(r.valores[c]??'');})};})};
  }
  return {arquivo:{id:arq.id,nome:arq.nome,url:arq.url,atualizadoEm:arq.atualizadoEm,calculadora:arq.calculadora},
    previaExcel:null,nomesAbas:nomes,abaCompleta:completa,pagina:pagina,quadro:quadro,
    formulas:formulas,faixasFormulas:faixas,mesclagens:mesclas,
    naoEditoriais:modelo.naoEditoriais||[],recalculoPendente:!!modelo.recalculoPendente,
    revisaoPreparada:true};
}
function _rxRevisaoLetra_(n){var s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;}
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

// Separacao de itens na interface; nao cria nem move abas no XLSX original.
function separarNaoEditorialReentradas(id,atualizadoEm,indice,numero){
  var lock=LockService.getScriptLock(),pegou=false;
  try{
    pegou=lock.tryLock(3000);if(!pegou)throw Error('Outra operacao de revisao esta em andamento.');
    var arq=_rxArquivo_(id,atualizadoEm),modelo=_rxRevisaoLer_(arq),aba=modelo.abas[indice];
    if(!aba||!/^(?:NC - )?PPG[ -]+Especifica/i.test(aba.nome))throw Error('Selecione um item da etapa Especificacoes tecnicas.');
    var q=_rxRevisaoLayout_(aba,modelo,indice),linha=aba.linhas.find(function(r){return r.numero===numero;});
    if(!linha||numero<=q.cabecalhoAte||numero>q.ultimaLinha)throw Error('Item da revisao nao encontrado.');
    var titulo=(aba.linhas.find(function(r){return r.numero===3;})||{}).valores||{};
    var coluna=Object.keys(titulo).find(function(k){return /CLASSIFICA/i.test(String(titulo[k]).normalize('NFD').replace(/[\u0300-\u036f]/g,''));});
    if(coluna===undefined)throw Error('A classificacao editorial nao foi encontrada nesta calculadora.');
    var classificacao=String(linha.valores[coluna]||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().trim();
    if(!/^(NAO EDITORIAL|NAO-EDITORIAL|NAO_EDITORIAL)$/.test(classificacao))
      throw Error('Somente itens classificados como Nao editorial podem ser separados.');
    var separadas=modelo.naoEditoriais||[];
    if(!separadas.some(function(x){return x.aba===indice&&x.linha===numero;})){
      separadas.push({aba:indice,linha:numero});modelo.naoEditoriais=separadas;
      _rxRevisaoSalvar_(arq,modelo,true);
    }
    return {sucesso:true,naoEditoriais:separadas};
  }catch(e){return {erro:e.message||String(e)};}
  finally{if(pegou)lock.releaseLock();}
}
