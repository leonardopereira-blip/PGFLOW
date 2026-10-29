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
  var livro=_pfPreviaLivroXlsx_(blob),formulas={},faixas=[],mesclagens=[];
  var conjunto=pedido.apenasPagina?null:_rxPreviaConjunto_(livro,arq.nome,formulas);
  if(!conjunto)formulas={};
  var pagina=conjunto?null:_rxPreviaPagina_(livro,arq.nome,consulta,formulas);
  (conjunto?livro.caminhos:[livro.caminhos[pagina.aba]]).forEach(function(caminho,i){
    _rxFaixasFormulas_(livro.arquivos[caminho]).forEach(function(r){
      r.aba=conjunto?i:pagina.aba;faixas.push(r);
    });
    _rxFaixasMesclagens_(livro.arquivos[caminho]).forEach(function(r){
      r.aba=conjunto?i:pagina.aba;mesclagens.push(r);
    });
  });
  return {arquivo:{id:arq.id,nome:arq.nome,url:arq.url,atualizadoEm:arq.atualizadoEm,
      calculadora:arq.calculadora},
    previaExcel:conjunto,pagina:pagina,formulas:formulas,faixasFormulas:faixas,mesclagens:mesclagens};
}

// O formato e os limites sao os mesmos da previa existente. O livro e lido
// uma unica vez e a informacao de formula fica isolada em Reentradas.
function _rxPreviaPagina_(livro,nome,pedido,formulas){
  var consulta=_pfPreviaConsulta_(pedido),nomes=livro.nomes;
  if(consulta.aba>=nomes.length)throw Error('Aba da prévia não encontrada.');
  var xml=livro.arquivos[livro.caminhos[consulta.aba]];
  if(!xml)throw Error('Não foi possível ler a aba selecionada do Excel.');
  var raiz=_emailSpgXlsxRaiz_(xml),sheetData=raiz.getChild('sheetData',raiz.getNamespace());
  if(!sheetData)throw Error('Esta aba não possui área de células legível.');
  var linhasXml=sheetData.getChildren('row',sheetData.getNamespace()),colunas=0;
  var dimensao=raiz.getChild('dimension',raiz.getNamespace()),ref=dimensao&&dimensao.getAttribute('ref');
  if(ref)colunas=Math.max(0,_pfPreviaColunaNumero_(ref.getValue().split(':').pop())+1);
  var selecionadas=[],termo=consulta.buscaNormalizada,total=linhasXml.length;
  linhasXml.forEach(function(linha,i){
    var valores={},formulaCols=[],numero=linha.getAttribute('r'),rn=numero?Number(numero.getValue()):i+1;
    if(!termo&&(i<consulta.linha||i>=consulta.linha+PF_PREVIA_LINHAS_PAGINA)){
      if(!ref)linha.getChildren('c',linha.getNamespace()).forEach(function(c){
        var p=c.getAttribute('r');if(p)colunas=Math.max(colunas,_pfPreviaColunaNumero_(p.getValue())+1);
      });
      selecionadas.push({numero:rn,valores:valores});return;
    }
    linha.getChildren('c',linha.getNamespace()).forEach(function(celula){
      var posicao=celula.getAttribute('r'),indice=posicao?_pfPreviaColunaNumero_(posicao.getValue()):-1;
      if(indice<0)return;
      if(indice+1>colunas)colunas=indice+1;
      var valor=_pfPreviaValorCelula_(celula,livro.compartilhadas);
      if(valor!=='')valores[indice]=valor;
      if(celula.getChild('f',celula.getNamespace()))formulaCols.push(indice+1);
    });
    if(!termo||Object.keys(valores).some(function(c){return String(valores[c]).toLocaleLowerCase('pt-BR').indexOf(termo)>=0;})){
      if(selecionadas.length>=consulta.linha&&selecionadas.length<consulta.linha+PF_PREVIA_LINHAS_PAGINA)
        formulaCols.forEach(function(c){formulas[consulta.aba+':'+rn+':'+c]=true;});
      selecionadas.push({numero:rn,valores:valores});
    }
  });
  var pagina=_pfPreviaResultado_(nome,nomes,consulta.aba,selecionadas,Math.max(1,colunas),total,selecionadas,consulta);
  var visiveis={};pagina.linhas.forEach(function(linha){visiveis[linha.numero]=true;});
  Object.keys(formulas).forEach(function(k){
    var c=k.split(':');
    if(!visiveis[Number(c[1])]||Number(c[2])<=pagina.coluna||Number(c[2])>pagina.coluna+pagina.colunas.length)delete formulas[k];
  });
  return pagina;
}
function _rxPreviaConjunto_(livro,nome,formulas){
  var tamanhoXml=livro.caminhos.reduce(function(total,caminho){return total+String(livro.arquivos[caminho]||'').length;},0);
  if(tamanhoXml>PF_PREVIA_CONJUNTO_MAX_XML)return null;
  var conjunto={arquivo:nome,abas:[]},celulas=0;
  for(var a=0;a<livro.nomes.length;a++){
    var xml=livro.arquivos[livro.caminhos[a]];
    if(!xml)throw Error('Não foi possível ler uma aba do Excel.');
    var raiz=_emailSpgXlsxRaiz_(xml),ns=raiz.getNamespace(),dados=raiz.getChild('sheetData',ns);
    if(!dados)throw Error('Esta aba não possui área de células legível.');
    var aba={nome:livro.nomes[a],totalColunas:1,linhas:[]};
    var dimensao=raiz.getChild('dimension',ns),ref=dimensao&&dimensao.getAttribute('ref');
    if(ref)aba.totalColunas=Math.max(1,_pfPreviaColunaNumero_(ref.getValue().split(':').pop())+1);
    var linhasXml=dados.getChildren('row',dados.getNamespace());
    for(var i=0;i<linhasXml.length;i++){
      var linha=linhasXml[i],numero=linha.getAttribute('r'),rn=numero?Number(numero.getValue()):i+1;
      var item={numero:rn,valores:{}},valoresXml=linha.getChildren('c',linha.getNamespace());
      for(var c=0;c<valoresXml.length;c++){
        if(++celulas>PF_PREVIA_CONJUNTO_MAX_CELULAS)return null;
        var referencia=valoresXml[c].getAttribute('r'),coluna=referencia?_pfPreviaColunaNumero_(referencia.getValue()):-1;
        if(coluna<0)continue;
        aba.totalColunas=Math.max(aba.totalColunas,coluna+1);
        var valor=_pfPreviaValorCelula_(valoresXml[c],livro.compartilhadas);
        if(valor!=='')item.valores[coluna]=valor;
        if(valoresXml[c].getChild('f',valoresXml[c].getNamespace()))formulas[a+':'+rn+':'+(coluna+1)]=true;
      }
      aba.linhas.push(item);
    }
    conjunto.abas.push(aba);
    if(JSON.stringify(conjunto).length>PF_PREVIA_CONJUNTO_MAX_JSON)return null;
  }
  return conjunto;
}

function _rxValores_(valor,celula){
  if(valor===undefined||valor===null)throw Error('Valor de célula ausente.');
  var texto=String(valor);
  if(texto.length>2000)throw Error('Valor de célula grande demais.');
  for(var i=0;i<texto.length;i++){
    var codigo=texto.charCodeAt(i);
    if(codigo>=0xD800&&codigo<=0xDBFF){
      var proximo=texto.charCodeAt(++i);
      if(!(proximo>=0xDC00&&proximo<=0xDFFF))throw Error('Texto contém caractere inválido para o Excel.');
    }else if((codigo<32&&codigo!==9&&codigo!==10&&codigo!==13)||
      (codigo>=0xDC00&&codigo<=0xDFFF)||codigo===0xFFFE||codigo===0xFFFF)
      throw Error('Texto contém caractere inválido para o Excel.');
  }
  if(!texto)return {texto:''};
  var atr=(/^<c\b[^>]*>/.exec(celula||'')||[])[0]||'';
  var t=(/\bt="([^"]+)"/.exec(atr)||[])[1]||'';
  if(t==='b'){
    if(/^(?:1|true|verdadeiro)$/i.test(texto))return {booleano:1};
    if(/^(?:0|false|falso)$/i.test(texto))return {booleano:0};
    throw Error('Esta célula aceita somente VERDADEIRO ou FALSO.');
  }
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
  if(v.booleano!==undefined)return '<c'+atr+' t="b"><v>'+v.booleano+'</v></c>';
  if(v.numero!==undefined)return '<c'+atr+'><v>'+v.numero+'</v></c>';
  return '<c'+atr+' t="inlineStr"><is><t'+(/(^\s|\s$)/.test(v.texto)?' xml:space="preserve"':'')+'>'+_reSpEscape_(v.texto)+'</t></is></c>';
}

function _rxFaixasFormulas_(xml){
  var faixas=[],re=/<f\b[^>]*>/g,m;
  while((m=re.exec(xml))!==null){
    var ref=(/\bref="([^"]+)"/.exec(m[0])||[])[1];
    if(!ref)continue;
    var r=/^\$?([A-Z]{1,3})\$?([1-9]\d*)(?::\$?([A-Z]{1,3})\$?([1-9]\d*))?$/.exec(ref);
    if(!r)throw Error('Intervalo de fórmula não suportado: '+ref);
    var x1=_reSpIdx_(r[1])+1,x2=_reSpIdx_(r[3]||r[1])+1,y1=Number(r[2]),y2=Number(r[4]||r[2]);
    if(x1>x2||y1>y2||x2>16384||y2>1048576)throw Error('Intervalo de fórmula inválido: '+ref);
    if(x1!==x2||y1!==y2)faixas.push({x1:x1,x2:x2,y1:y1,y2:y2});
  }
  return faixas;
}
function _rxEmFaixaFormula_(faixas,linha,coluna){
  return faixas.some(function(r){return coluna>=r.x1&&coluna<=r.x2&&linha>=r.y1&&linha<=r.y2;});
}
function _rxFaixasMesclagens_(xml){
  var out=[],re=/<mergeCell\b[^>]*\/>/g,m;
  while((m=re.exec(xml))!==null)out.push(_reMpRetangulo_(_reMpAtributos_(m[0]).ref));
  return out;
}
function _rxMarcarRecalculo_(xml){
  var tags=xml.match(/<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/g)||[];
  if(tags.length>1)throw Error('Configuração de cálculo Excel duplicada.');
  if(!tags.length)return _reSpFullCalc_(xml);
  var tag=tags[0],head=(/^<calcPr\b[^>]*>/.exec(tag)||[])[0];
  var ativo=function(nome,valor){return new RegExp('\\b'+nome+'\\s*=\\s*(["\'])'+valor+'\\1').test(head);};
  if(ativo('calcMode','auto')&&ativo('fullCalcOnLoad','(?:1|true)')&&ativo('forceFullCalc','(?:1|true)'))return xml;
  var final=/\/>$/.test(head)?'/>':'>';
  var novo=head.replace(/\s(?:calcMode|fullCalcOnLoad|forceFullCalc)\s*=\s*(?:"[^"]*"|'[^']*')/g,'')
    .replace(/\s*\/?>$/,' calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"'+final);
  return xml.replace(tag,tag.replace(head,novo));
}
function _rxAbasEditaveis_(workbook,rels){
  var destinos={},abas=[],re=/<Relationship\b[^>]*\/>/g,m;
  while((m=re.exec(rels))!==null){
    var r=_reMpAtributos_(m[0]);
    if(r.Id&&r.Target&&/\/worksheet$/.test(r.Type||'')&&r.TargetMode!=='External')
      destinos[r.Id]=_reMpCaminho_(r.Target);
  }
  re=/<sheet\b[^>]*>/g;
  while((m=re.exec(workbook))!==null){
    var s=_reMpAtributos_(m[0]);
    if(s.name&&destinos[s['r:id']])abas.push({nome:s.name,caminho:destinos[s['r:id']]});
  }
  if(!abas.length)throw Error('O Excel não contém abas editáveis.');
  return abas;
}

function _rxEditarAba_(xml,edicoes){
  var dados=/<sheetData(?:\s[^>]*)?>[\s\S]*?<\/sheetData>/.exec(xml);
  if(!dados)throw Error('Aba Excel sem dados.');
  var formulas=_rxFaixasFormulas_(dados[0]),mesclagens=_rxFaixasMesclagens_(xml);
  var porLinha={};edicoes.forEach(function(e){
    if(_rxEmFaixaFormula_(formulas,e.linha,e.coluna))
      throw Error('A célula '+_reSpCol_(e.coluna)+e.linha+' pertence a um intervalo de fórmula e não pode ser substituída.');
    if(mesclagens.some(function(r){return e.coluna>=r.x1&&e.coluna<=r.x2&&e.linha>=r.y1&&e.linha<=r.y2&&
      (e.coluna!==r.x1||e.linha!==r.y1);}))
      throw Error('Edite somente a primeira célula da área mesclada.');
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
    if(!/^\d{10,16}$/.test(String(atualizadoEm||''))||!Number.isSafeInteger(Number(atualizadoEm)))
      throw Error('Atualize a prévia antes de salvar o arquivo.');
    if(!Array.isArray(edicoes)||!edicoes.length||edicoes.length>200)throw Error('Selecione de 1 a 200 células por salvamento.');
    pegou=lock.tryLock(3000);if(!pegou)throw Error('Outra operação está em andamento. Tente novamente.');
    var arq=_rxArquivo_(id,atualizadoEm),p=_reSpPartes_(arq.arquivo.getBlob());
    if(p.partes.length!==Object.keys(p.mapa).length)throw Error('Partes duplicadas no Excel.');
    var names=_rxAbasEditaveis_(p.mapa['xl/workbook.xml'].getDataAsString('UTF-8'),
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
    alter['xl/workbook.xml']=_rxMarcarRecalculo_(p.mapa['xl/workbook.xml'].getDataAsString('UTF-8'));
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
