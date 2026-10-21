// Restauracao manual de calculadoras Excel existentes. Trabalha no ZIP/XLSX;
// nao converte arquivos para Google Sheets nem recria valores ou formulas.
var BC_PASTA='1DKie1AHE7EFHa6roJY2XxE2ZlFhC_lCy';
var BC_XLSX='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
var BC_XLSM='application/vnd.ms-excel.sheet.macroenabled.12';

function abrirBloqueioCalculadoras(){
  var html=HtmlService.createHtmlOutputFromFile('BloquearCalculadoras').setWidth(920).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html,'Restaurar protecao das calculadoras');
}
function _bcGrafica_(nome){
  var texto=String(nome||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/[^A-Z0-9]+/g,' ');
  var opcoes={LOGPRINT:/\bLOGPRINT\b/,CASADAGRAFICA:/\bCASA\s*DA\s*GRAFICA\b/,
    FORMACERTA:/\bFORMA\s*CERTA\b/,META:/\bMETA\b/,REPROSET:/\bREPROSET\b/};
  var achadas=Object.keys(opcoes).filter(function(k){return opcoes[k].test(texto);});
  if(achadas.length!==1)throw Error('A grafica do arquivo nao corresponde a um unico modelo suportado.');
  return achadas[0];
}
function _bcMetadados_(arquivo){
  var nome=arquivo.getName(),tipo=String(arquivo.getMimeType()).toLowerCase(),pais=arquivo.getParents(),naPasta=false;
  while(pais.hasNext())if(pais.next().getId()===BC_PASTA)naPasta=true;
  if(arquivo.isTrashed()||!naPasta)throw Error('O arquivo nao esta na pasta de calculadoras ou foi excluido.');
  if(!/calculadora/i.test(nome)||! /\.xls[xm]$/i.test(nome))throw Error('Selecione um arquivo Calculadora em formato XLSX ou XLSM.');
  if(tipo!==BC_XLSX&&tipo!==BC_XLSM)throw Error('O tipo do arquivo nao e uma pasta de trabalho Excel suportada.');
  if((/\.xlsx$/i.test(nome)&&tipo!==BC_XLSX)||(/\.xlsm$/i.test(nome)&&tipo!==BC_XLSM))throw Error('O nome e o tipo Excel do arquivo nao correspondem.');
  return {id:arquivo.getId(),nome:nome,tipo:tipo,atualizadoEm:arquivo.getLastUpdated().getTime(),url:arquivo.getUrl()};
}
function listarCalculadorasParaBloqueio(){
  try{
    var arquivos=DriveApp.getFolderById(BC_PASTA).getFiles(),lista=[];
    while(arquivos.hasNext()){
      var arquivo=arquivos.next(),item;
      try{item=_bcMetadados_(arquivo);}catch(ignorar){continue;}
      try{item.grafica=_bcGrafica_(item.nome);item.disponivel=true;}
      catch(e){item.disponivel=false;item.motivo=e.message;}
      lista.push(item);
    }
    lista.sort(function(a,b){return b.atualizadoEm-a.atualizadoEm||a.nome.localeCompare(b.nome);});
    return {arquivos:lista};
  }catch(e){return {erro:'Nao foi possivel listar as calculadoras: '+e.message};}
}
function _bcPerfil_(grafica){
  var g=RE_SP_MOLDES[grafica];if(!g)throw Error('Modelo de calculadora nao suportado.');
  var molde=g.calc();
  return {molde:molde,protecao:grafica==='LOGPRINT'?_moldeCasaDaGraficaCalc_():molde};
}
function _bcMergeRefs_(xml){
  var out=[],re=/<mergeCell\b[^>]*\/>/g,m;
  while((m=re.exec(xml))!==null)out.push(_reMpAtributos_(m[0]).ref);
  return out.sort().join('|');
}
function _bcValidarLayout_(partes,molde){
  if(!partes.mapa['xl/workbook.xml']||!partes.mapa['xl/_rels/workbook.xml.rels']||!partes.mapa['xl/styles.xml'])throw Error('Estrutura de pasta de trabalho Excel incompleta.');
  var abas=_reLpAbas_(partes.mapa['xl/workbook.xml'].getDataAsString('UTF-8'),partes.mapa['xl/_rels/workbook.xml.rels'].getDataAsString('UTF-8'));
  var esperadas=_reLpAbas_(molde.textos['xl/workbook.xml'],molde.textos['xl/_rels/workbook.xml.rels']),porNome={};
  esperadas.forEach(function(a){porNome[a.nome]=a.caminho;});
  if(abas.length!==esperadas.length)throw Error('O conjunto de abas mudou. O arquivo nao corresponde ao modelo da grafica.');
  abas.forEach(function(a){
    var caminho=porNome[a.nome],aba=caminho&&molde.abas[caminho];
    if(!aba||!partes.mapa[a.caminho])throw Error('Aba sem perfil de protecao compativel: '+a.nome);
    var xml=partes.mapa[a.caminho].getDataAsString('UTF-8');
    if(_bcMergeRefs_(xml)!==_bcMergeRefs_(aba.fim))throw Error('O layout da aba '+a.nome+' mudou. Nenhum arquivo foi alterado.');
  });
  return abas;
}
function _bcSemMudancaDeConteudo_(xml){
  return xml.replace(/<sheetProtection\b[^>]*?(?:\/>|>[\s\S]*?<\/sheetProtection>)/g,'')
    .replace(/<cols\b[^>]*>[\s\S]*?<\/cols>/g,'')
    .replace(/<(?:c|row)\b[^>]*>/g,function(tag){return tag.replace(/\s(?:s|customFormat)="[^"]*"/g,'');});
}
function _bcLimparVba_(partes,alteradas){
  var removidas={},tipos='[Content_Types].xml';
  if(!partes.mapa[tipos])throw Error('Tipos de conteudo ausentes no Excel.');
  Object.keys(partes.mapa).forEach(function(n){
    if(/(?:^|\/)vbaProject(?:Signature(?:Agile|V3)?)?\.bin(?:\.rels)?$/i.test(n)||
      /^customXml\/(?:pgflowProtection(?:Props)?\.xml|_rels\/pgflowProtection\.xml\.rels)$/.test(n))removidas[n]=true;
  });
  Object.keys(partes.mapa).filter(function(n){return /\.rels$/i.test(n)&&!removidas[n];}).forEach(function(n){
    var xml=partes.mapa[n].getDataAsString('UTF-8'),final=xml.replace(/<Relationship\b[^>]*\/>/g,function(tag){
      var a=_reMpAtributos_(tag);
      return /\/vbaProject(?:Signature(?:Agile|V3)?)?$/i.test(a.Type||'')||/pgflowProtection(?:Props)?\.xml(?:$|[?#])/i.test(a.Target||'')?'':tag;
    });
    if(final!==xml)alteradas[n]=final;
  });
  var ct=partes.mapa[tipos].getDataAsString('UTF-8');
  if(/macroSheet|dialogsheet/i.test(ct))throw Error('O arquivo contem um tipo de macro que nao pode ser removido por esta funcao.');
  ct=ct.replace(/application\/vnd\.ms-excel\.sheet\.macroEnabled\.main\+xml/g,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml')
    .replace(/<(?:Override|Default)\b[^>]*\/>/g,function(tag){
      var a=_reMpAtributos_(tag),nome=(a.PartName||'').replace(/^\//,'');
      return removidas[nome]||/vbaProject/i.test(a.ContentType||'')?'':tag;
    });
  alteradas[tipos]=ct;
  var w=partes.mapa['xl/workbook.xml'].getDataAsString('UTF-8');
  alteradas['xl/workbook.xml']=w.replace(/<workbookPr\b[^>]*>/,function(tag){return tag.replace(/\scodeName="[^"]*"/,'');});
  // Nao ha conversao de Sheets nem partes VBA novas. Partes nao relacionadas
  // a protecao/VBA continuam sendo os mesmos blobs binarios recebidos.
  return removidas;
}
function _bcMontarXlsx_(blob,grafica,nome){
  var partes=_reSpPartes_(blob);
  if(partes.partes.length!==Object.keys(partes.mapa).length)throw Error('O arquivo Excel contem partes duplicadas.');
  var perfil=_bcPerfil_(grafica),abas=_bcValidarLayout_(partes,perfil.molde),alteradas={};
  _reSpAplicarProtecaoLogprint_(partes,alteradas,perfil.protecao,true);
  abas.forEach(function(a){
    var antes=partes.mapa[a.caminho].getDataAsString('UTF-8'),depois=alteradas[a.caminho];
    if(_bcSemMudancaDeConteudo_(antes)!==_bcSemMudancaDeConteudo_(depois))throw Error('A restauracao alteraria conteudo da aba '+a.nome+'. Nenhum arquivo foi alterado.');
  });
  var removidas=_bcLimparVba_(partes,alteradas),nomeFinal=nome.replace(/\.xlsm$/i,'.xlsx');
  var saida=partes.partes.filter(function(p){return !removidas[p.getName()];}).map(function(p){
    return alteradas[p.getName()]===undefined?p:Utilities.newBlob(alteradas[p.getName()],'application/xml',p.getName());
  });
  var arquivo=Utilities.zip(saida,nomeFinal).setContentType(BC_XLSX);
  return {blob:arquivo,nome:nomeFinal,abas:abas.length};
}
function _bcDigest_(blob){
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,blob.getBytes()));
}
function bloquearCalculadoraSelecionada(id,atualizadoEm){
  var lock=LockService.getScriptLock(),adquirido=false,gravou=false;
  try{
    if(!/^[A-Za-z0-9_-]{10,200}$/.test(String(id||''))||!/^\d{10,16}$/.test(String(atualizadoEm||'')))throw Error('Selecione novamente um arquivo da lista.');
    adquirido=lock.tryLock(2000);if(!adquirido)throw Error('Outra restauracao esta em andamento. Aguarde e atualize a lista.');
    var arquivo=DriveApp.getFileById(id),antes=_bcMetadados_(arquivo);
    if(antes.atualizadoEm!==Number(atualizadoEm))throw Error('O arquivo mudou desde que a lista foi carregada. Atualize a lista antes de continuar.');
    var grafica=_bcGrafica_(antes.nome),pronto=_bcMontarXlsx_(arquivo.getBlob(),grafica,antes.nome);
    var atual=_bcMetadados_(DriveApp.getFileById(id));
    if(atual.atualizadoEm!==antes.atualizadoEm||atual.nome!==antes.nome||atual.tipo!==antes.tipo)throw Error('O arquivo mudou durante a preparacao. Nenhum arquivo foi sobrescrito; atualize a lista.');
    var esperado=_bcDigest_(pronto.blob);
    Drive.Files.update({title:pronto.nome,mimeType:BC_XLSX},id,pronto.blob,{newRevision:true,supportsAllDrives:true});
    gravou=true;
    var finalArquivo=DriveApp.getFileById(id),final=_bcMetadados_(finalArquivo);
    if(final.id!==id||final.nome!==pronto.nome||final.tipo!==BC_XLSX||_bcDigest_(finalArquivo.getBlob())!==esperado)throw Error('A releitura do Drive divergiu do arquivo preparado.');
    return {sucesso:true,id:final.id,nome:final.nome,url:final.url,atualizadoEm:final.atualizadoEm,abas:pronto.abas};
  }catch(e){return {erro:(gravou?'O Drive recebeu o arquivo, mas a verificacao final nao foi confirmada: ':'')+e.message};}
  finally{if(adquirido)lock.releaseLock();}
}
