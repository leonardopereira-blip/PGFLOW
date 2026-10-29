// Prévia e edição de XLSX gerado: somente arquivos da pasta de Reentradas.
// Alterações substituem o conteúdo do MESMO arquivo (nova revisão no Drive).
var RX_PASTA = '1DKie1AHE7EFHa6roJY2XxE2ZlFhC_lCy';
var RX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

function _rxArquivo_(id, atualizadoEm) {
  if (!/^[A-Za-z0-9_-]{10,200}$/.test(String(id||''))) throw Error('Arquivo inválido.');
  var arq=DriveApp.getFileById(id), nome=arq.getName();
  if(arq.isTrashed()||arq.getMimeType()!==RX_MIME||!/\.xlsx$/i.test(nome))
    throw Error('Somente arquivos XLSX ativos são aceitos.');
  var pais=arq.getParents(),valido=false;
  while(pais.hasNext())if(pais.next().getId()===RX_PASTA)valido=true;
  if(!valido||!/^CALCULADORA FORNECEDOR |^PLANILHA DE PARA PPG /i.test(nome))
    throw Error('Arquivo não pertence aos documentos gerados de Reentradas.');
  var data=arq.getLastUpdated().getTime();
  if(atualizadoEm!==undefined&&atualizadoEm!==null&&Number(atualizadoEm)!==data)
    throw Error('O arquivo foi alterado no Drive. Atualize a prévia antes de salvar.');
  return {arquivo:arq,id:arq.getId(),nome:nome,url:arq.getUrl(),atualizadoEm:data,
    calculadora:/^CALCULADORA FORNECEDOR /i.test(nome)};
}

function listarArquivosReentradasExcel(){
  var arquivos=[],iterator=DriveApp.getFolderById(RX_PASTA).getFiles();
  while(iterator.hasNext()){
    var a=iterator.next(),nome=a.getName(),tipo=a.getMimeType();
    if(a.isTrashed())continue;
    var calculadora=/^CALCULADORA FORNECEDOR /i.test(nome);
    var depara=/^PLANILHA DE PARA PPG /i.test(nome);
    var item={id:a.getId(),nome:nome,url:a.getUrl(),atualizadoEm:a.getLastUpdated().getTime(),
      calculadora:calculadora,editavel:tipo===RX_MIME&&(calculadora||depara),bloqueavel:false};
    if(calculadora){
      try{_bcMetadados_(a);_bcGrafica_(nome);item.bloqueavel=true;}
      catch(e){item.motivo=e.message;}
    }else item.motivo=depara?'O De-Para permanece com a proteção do modelo original.':'Tipo de arquivo sem restauração de proteção configurada.';
    arquivos.push(item);
  }
  arquivos.sort(function(a,b){return b.atualizadoEm-a.atualizadoEm||a.nome.localeCompare(b.nome);});
  return {arquivos:arquivos};
}

function obterPreviaReentradasExcel(pedido){
  var arq=_rxArquivo_(pedido&&pedido.id,pedido&&pedido.atualizadoEm);
  var blob=arq.arquivo.getBlob().setName(arq.nome);
  var consulta={aba:pedido.aba||0,linha:pedido.linha||0,coluna:pedido.coluna||0,busca:pedido.busca||''};
  var conjunto=pedido.apenasPagina?null:_pfPreviaConjuntoXlsx_(blob);
  return {arquivo:{id:arq.id,nome:arq.nome,url:arq.url,atualizadoEm:arq.atualizadoEm,
      calculadora:arq.calculadora},
    previaExcel:conjunto,pagina:conjunto?null:_pfPreviaXlsx_(blob,consulta)};
}

function _rxValores_(valor,celula){
  if(valor===undefined||valor===null)throw Error('Valor de célula ausente.');
  var texto=String(valor);
  if(texto.length>2000)throw Error('Valor de célula grande demais.');
  if(!texto)return {texto:''};
  var atr=(/^<c\b[^>]*>/.exec(celula||'')||[])[0]||'';
  var t=(/\bt="([^"]+)"/.exec(atr)||[])[1]||'';
  var numerico=/^[+-]?(?:0|[1-9]\d{0,14})(?:[.,]\d+)?$/.test(texto) &&
    !/^[+-]?0\d/.test(texto);
  if(numerico&&(!t||t==='n'||!celula)){
    var n=Number(texto.replace(',','.'));
    if(isFinite(n))return {numero:n};
  }
  return {texto:texto};
}

function _rxNovaCelula_(ref,original,valor){
  if(original&&/<f\b/.test(original))
    throw Error('A célula '+ref+' contém fórmula e não pode ser substituída.');
  if(original&&/<(?:extLst|is\b[^>]*><r\b)/.test(original))
    throw Error('A célula '+ref+' possui estrutura especial. Edite diretamente no Excel.');
  var head=(/^<c\b[^>]*\/?>/.exec(original||'')||[])[0]||'';
  var extras=head.replace(/^<c\b/,'').replace(/\/?>$/,'').match(/\s([A-Za-z_:][\w:.-]*)="[^"]*"/g)||[];
  extras.forEach(function(atrib){
    var nome=/^\s([A-Za-z_:][\w:.-]*)=/.exec(atrib)[1];
    if(['r','s','t'].indexOf(nome)<0)throw Error('Célula '+ref+' com metadados não suportados.');
  });
  var sty=(/\bs="(\d+)"/.exec(head)||[])[1],atr=' r="'+ref+'"'+(sty?' s="'+sty+'"':'');
  var v=_rxValores_(valor,original);
  if(v.texto==='')return '<c'+atr+'/>';
  if(v.numero!==undefined)return '<c'+atr+'><v>'+v.numero+'</v></c>';
  return '<c'+atr+' t="inlineStr"><is><t'+(/(^\s|\s$)/.test(v.texto)?' xml:space="preserve"':'')+'>'+_reSpEscape_(v.texto)+'</t></is></c>';
}

function _rxEditarAba_(xml,edicoes){
  var dados=/<sheetData(?:\s[^>]*)?>[\s\S]*?<\/sheetData>/.exec(xml);
  if(!dados)throw Error('Aba Excel sem dados.');
  var porLinha={};edicoes.forEach(function(e){
    (porLinha[e.linha]||(porLinha[e.linha]=[])).push(e);
  });
  var encontradas={},troca=dados[0].replace(/<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/g,function(row,numero){
    var edits=porLinha[Number(numero)];if(!edits)return row;
    var atual={},cells=/<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g,m;
    while((m=cells.exec(row))!==null){
      var ref=(/\br="([A-Z]+\d+)"/.exec(m[0])||[])[1];
      if(ref)atual[ref]=m[0];
    }
    edits.forEach(function(e){
      var ref=_reSpCol_(e.coluna)+e.linha;
      atual[ref]=_rxNovaCelula_(ref,atual[ref],e.valor);
      encontradas[e.linha+':'+e.coluna]=true;
    });
    var head=(/^<row\b[^>]*>/.exec(row)||[])[0];
    if(!head)throw Error('Linha Excel inválida.');
    var fim=row.replace(head,'').replace(/<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g,'').replace(/<\/row>$/,'').trim();
    if(fim)throw Error('Linha Excel com estrutura não suportada.');
    return head+Object.keys(atual).sort(function(a,b){
      return _reSpIdx_((/^[A-Z]+/.exec(a)||[''])[0])-_reSpIdx_((/^[A-Z]+/.exec(b)||[''])[0]);
    }).map(function(ref){return atual[ref];}).join('')+'</row>';
  });
  edicoes.forEach(function(e){
    if(!encontradas[e.linha+':'+e.coluna])throw Error('Linha '+e.linha+' não encontrada no Excel.');
  });
  return xml.replace(dados[0],troca);
}

function salvarEdicoesReentradasExcel(id,atualizadoEm,edicoes){
  var lock=LockService.getScriptLock(),pegou=false,gravou=false;
  try{
    if(!Array.isArray(edicoes)||!edicoes.length||edicoes.length>200)throw Error('Selecione de 1 a 200 células por salvamento.');
    pegou=lock.tryLock(3000);if(!pegou)throw Error('Outra operação está em andamento. Tente novamente.');
    var arq=_rxArquivo_(id,atualizadoEm),p=_reSpPartes_(arq.arquivo.getBlob());
    if(p.partes.length!==Object.keys(p.mapa).length)throw Error('Partes duplicadas no Excel.');
    var names=_reLpAbas_(p.mapa['xl/workbook.xml'].getDataAsString('UTF-8'),
      p.mapa['xl/_rels/workbook.xml.rels'].getDataAsString('UTF-8'));
    var porAba={},seen={};
    edicoes.forEach(function(e){
      if(!e||!Number.isInteger(e.aba)||e.aba<0||e.aba>=names.length||
        !Number.isInteger(e.linha)||e.linha<1||e.linha>1000000||
        !Number.isInteger(e.coluna)||e.coluna<1||e.coluna>16384)throw Error('Endereço de célula inválido.');
      var key=e.aba+':'+e.linha+':'+e.coluna;
      if(seen[key])throw Error('Célula repetida no salvamento.');
      seen[key]=true;(porAba[e.aba]||(porAba[e.aba]=[])).push(e);
    });
    var alter={};
    Object.keys(porAba).forEach(function(indice){
      var caminho=names[Number(indice)].caminho,parte=p.mapa[caminho];
      if(!parte)throw Error('Aba não encontrada.');
      alter[caminho]=_rxEditarAba_(parte.getDataAsString('UTF-8'),porAba[indice]);
    });
    var zipped=Utilities.zip(p.partes.map(function(part){
      return alter[part.getName()]===undefined?part:
        Utilities.newBlob(alter[part.getName()],'application/xml',part.getName());
    }),arq.nome).setContentType(RX_MIME);
    _rxArquivo_(id,atualizadoEm);
    var hash=_bcDigest_(zipped);
    Drive.Files.update({title:arq.nome,mimeType:RX_MIME},id,zipped,
      {newRevision:true,supportsAllDrives:true});
    gravou=true;
    var final=_rxArquivo_(id);
    if(_bcDigest_(final.arquivo.getBlob())!==hash)throw Error('Conferência do XLSX salvo não corresponde.');
    return {sucesso:true,id:final.id,nome:final.nome,url:final.url,
      atualizadoEm:final.atualizadoEm,celulas:edicoes.length};
  }catch(e){return {erro:(gravou?'Arquivo salvo, mas a releitura não confirmou: ':'')+e.message};}
  finally{if(pegou)lock.releaseLock();}
}

function bloquearArquivoReentradasExcel(id,atualizadoEm){
  try{
    // O serviço de restauração existente também valida XLSM e preserva o mesmo ID.
    var arq=DriveApp.getFileById(id),meta=_bcMetadados_(arq);
    if(meta.atualizadoEm!==Number(atualizadoEm))throw Error('Arquivo alterado. Atualize a lista.');
    _bcGrafica_(meta.nome);
    return bloquearCalculadoraSelecionada(id,String(atualizadoEm));
  }catch(e){return {erro:e.message};}
}
