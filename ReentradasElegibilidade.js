// A elegibilidade afeta somente a calculadora. As matrizes tecnicas de origem
// continuam completas; seus campos Grafico/Editorial sao lidos antes do molde.
function _reSpElegCabecalho_(valor) {
  return String(valor == null ? '' : valor).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function _reSpElegNegativo_(valor, campo) {
  if (valor === false) return true;
  if (typeof valor !== 'string') return false;
  var rotulo = valor.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase()
    .replace(/[\s\u00a0\-\u2010-\u2015]+/g, '');
  // So interpreta negativos explicitos na coluna dedicada. Nao infere por
  // classificacao do produto, descricao, numeros, ausencia ou texto livre.
  return rotulo === 'FALSE' || rotulo === 'NAO' || rotulo === 'NAO' + campo;
}

function _reSpSkusNaoCalculaveis_(especificacao, tiragem) {
  var ordem = [], porSku = Object.create(null);
  var fontes = [especificacao, tiragem];
  function indices(matriz) {
    var headers = matriz && Array.isArray(matriz.headers) ? matriz.headers : [];
    var mapa = { sku: -1, GRAFICO: [], EDITORIAL: [] };
    headers.forEach(function(header, i) {
      var nome = _reSpElegCabecalho_(header);
      if (nome === '1IDCODIGOSKU' && mapa.sku < 0) mapa.sku = i;
      if (nome === 'GRAFICO' || nome === 'EDITORIAL') mapa[nome].push(i);
    });
    return mapa;
  }
  function codigo(linha, indice) {
    return indice < 0 || !linha ? '' : String(linha[indice] == null ? '' : linha[indice]).trim();
  }
  var indiceTiragem = indices(tiragem);
  ((tiragem && tiragem.linhas) || []).forEach(function(linha) {
    var sku = codigo(linha, indiceTiragem.sku);
    if (sku && !porSku[sku]) {
      porSku[sku] = {};
      ordem.push(sku);
    }
  });
  fontes.forEach(function(matriz) {
    var mapa = indices(matriz);
    ((matriz && matriz.linhas) || []).forEach(function(linha) {
      var sku = codigo(linha, mapa.sku), registro = porSku[sku];
      if (!registro) return;
      ['GRAFICO', 'EDITORIAL'].forEach(function(campo) {
        if (mapa[campo].some(function(i) { return _reSpElegNegativo_(linha[i], campo); })) {
          registro[campo] = true;
        }
      });
    });
  });
  return ordem.filter(function(sku) {
    return porSku[sku].GRAFICO || porSku[sku].EDITORIAL;
  }).map(function(sku) {
    var motivos = [];
    if (porSku[sku].GRAFICO) motivos.push('Não gráfico');
    if (porSku[sku].EDITORIAL) motivos.push('Não editorial');
    return { sku: sku, motivos: motivos };
  });
}
