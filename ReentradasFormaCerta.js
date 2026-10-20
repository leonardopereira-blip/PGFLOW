// A IMPUT do modelo Forma Certa tem 28 linhas completas e uma coluna
// preenchida ate o fim do Excel. O molde guarda somente a capacidade util;
// estas funcoes completam as formulas para a quantidade de SKUs do chamado.
function _reFcDeslocarFormula_(formula,delta) {
  // Textos literais e referencias com linha absoluta permanecem intactos.
  return formula.split(/("(?:[^"]|"")*")/).map(function(parte,i){
    if(i%2)return parte;
    return parte.replace(/(^|[^A-Za-z0-9_.])(\$?[A-Z]{1,3})(\$?)([1-9]\d{0,6})(?![\d(])/g,
      function(t,p,col,fixa,linha){return p+col+fixa+(fixa?linha:Number(linha)+delta);});
  }).join('');
}

function _reFcExpandirImput_(xml,quantidade) {
  if(quantidade<=28)return xml;
  if(quantidade>300)throw Error('Forma Certa: limite de 300 SKUs.');
  var modelos={},re=/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g,c;
  while((c=re.exec(xml))!==null){
    var f=/<f\b([^>]*)>([\s\S]*?)<\/f>/.exec(c[2]||'');
    var si=f&&/\bsi="(\d+)"/.exec(f[1]),ref=/\br="([A-Z]+)(\d+)"/.exec(c[1]);
    if(f&&si&&ref&&/\bt="shared"/.test(f[1]))modelos[si[1]]={col:ref[1],linha:+ref[2],formula:f[2]};
  }
  var modelo=/<row\b[^>]*\br="30"[^>]*>[\s\S]*?<\/row>/.exec(xml);
  if(!modelo)throw Error('Forma Certa: linha de formulas de referencia ausente.');
  var ultima=quantidade+2,alteradas=0;
  var criar=function(numero){
    var linha=modelo[0].replace(/^(<row\b[^>]*\br=")30"/,'$1'+numero+'"');
    return linha.replace(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g,function(celula,attrs,corpo){
      var endereco=/\br="([A-Z]+)30"/.exec(attrs);
      if(!endereco)throw Error('Forma Certa: endereco de celula inconsistente.');
      attrs=attrs.replace(/\br="[A-Z]+30"/,'r="'+endereco[1]+numero+'"');
      var formula=/<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(corpo||'');
      if(!formula)return '<c'+attrs+'/>';
      var texto=formula[2],si=/\bsi="(\d+)"/.exec(formula[1]);
      if(texto===undefined&&si){
        var base=modelos[si[1]];
        if(!base||base.col!==endereco[1])throw Error('Forma Certa: formula compartilhada sem origem vertical.');
        texto=_reFcDeslocarFormula_(base.formula,30-base.linha);
      }
      if(texto===undefined)throw Error('Forma Certa: formula vazia sem origem.');
      // O Excel recalcula as formulas; nao copiar resultados do chamado modelo.
      return '<c'+attrs+'><f>'+_reFcDeslocarFormula_(texto,numero-30)+'</f></c>';
    });
  };
  xml=xml.replace(/<row\b[^>]*\br="(\d+)"[^>]*(?:\/>|>[\s\S]*?<\/row>)/g,function(linha,n){
    if(+n<=30||+n>ultima)return linha;
    alteradas++;return criar(+n);
  });
  if(alteradas!==ultima-30)throw Error('Forma Certa: molde sem a capacidade de linhas esperada.');
  if(ultima>103)xml=xml.replace(/<dataValidation\b[\s\S]*?(?:\/>|<\/dataValidation>)/g,function(tag){
    return tag.replace(/\bsqref="([^"]*)"/,function(t,refs){return 'sqref="'+refs.replace(/(\$?[A-Z]+\$?)103\b/g,'$1'+ultima)+'"';});
  });
  return xml;
}
