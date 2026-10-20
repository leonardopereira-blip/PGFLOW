// Macro manual de protecao da calculadora. Nao cria botoes nem eventos.
// O mapa de Locked/Hidden e extraido dos estilos finais do proprio molde.
function _reSpAdicionarMacro_(partes,alteradas,protecaoLogprintPreparada) {
  var baseProtecao=partes;
  // A Logprint veio livre para edicao; o usuario autorizou expressamente o
  // perfil das calculadoras de mesmo layout. A macro deve guardar esse perfil.
  if(protecaoLogprintPreparada){
    baseProtecao={mapa:{}};
    Object.keys(partes.mapa).forEach(function(n){baseProtecao.mapa[n]=alteradas[n]===undefined?partes.mapa[n]:Utilities.newBlob(alteradas[n],'application/xml',n);});
  }
  var mapa=_reSpMapaProtecao_(baseProtecao,alteradas);
  var rels='xl/_rels/workbook.xml.rels',tipos='[Content_Types].xml';
  var ler=function(nome){return alteradas[nome]===undefined?partes.mapa[nome].getDataAsString('UTF-8'):alteradas[nome];};
  var r=ler(rels),t=ler(tipos);
  if(/vbaProject|pgflowProtection/.test(r))throw Error('Calculadora ja tem macro ou mapa de protecao.');
  alteradas[rels]=r.replace('</Relationships>',
    '<Relationship Id="rIdPGFLOWVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>'+
    '<Relationship Id="rIdPGFLOWProtection" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/pgflowProtection.xml"/></Relationships>');
  if(t.indexOf('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml')<0)throw Error('Tipo de calculadora inesperado.');
  alteradas[tipos]=t.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml','application/vnd.ms-excel.sheet.macroEnabled.main+xml')
    .replace('</Types>',
      '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>'+
      '<Override PartName="/customXml/pgflowProtectionProps.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>');
  var w=ler('xl/workbook.xml');
  alteradas['xl/workbook.xml']=w.replace(/<workbookPr\b[^>]*\/>/,function(tag){
    if(/\bcodeName=/.test(tag))return tag;
    return tag.replace('/>',' codeName="ThisWorkbook"/>');
  });
  var novos=[
    Utilities.newBlob(Utilities.base64Decode(_reSpMacroVba_()),'application/vnd.ms-office.vbaProject','xl/vbaProject.bin'),
    Utilities.newBlob(mapa,'application/xml','customXml/pgflowProtection.xml'),
    Utilities.newBlob('<?xml version="1.0" encoding="UTF-8"?>'+
      '<ds:datastoreItem ds:itemID="{16DE2932-7B14-44AD-8C6B-6B253F1DCB78}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">'+
      '<ds:schemaRefs><ds:schemaRef ds:uri="urn:pgflow:protecao:v1"/></ds:schemaRefs></ds:datastoreItem>',
      'application/xml','customXml/pgflowProtectionProps.xml'),
    Utilities.newBlob('<?xml version="1.0" encoding="UTF-8"?>'+
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
      '<Relationship Id="rIdProps" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="pgflowProtectionProps.xml"/></Relationships>',
      'application/xml','customXml/_rels/pgflowProtection.xml.rels')
  ];
  novos.forEach(function(p){
    if(partes.mapa[p.getName()])throw Error('Parte duplicada na calculadora: '+p.getName());
    partes.partes.push(p);partes.mapa[p.getName()]=p;
  });
}
