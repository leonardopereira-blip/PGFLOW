// Prévia de leitura do anexo de formalização. Não grava arquivos nem envia e-mails.
var PF_PREVIA_LINHAS_PAGINA = 50;
var PF_PREVIA_COLUNAS_PAGINA = 65;

function _pfPreviaNumero_(valor, limite) {
  var n = Number(valor);
  return isFinite(n) && n >= 0 ? Math.min(Math.floor(n), limite) : 0;
}
function _pfPreviaColunaNumero_(ref) {
  var m = /^([A-Z]+)\d*$/i.exec(String(ref || ''));
  if (!m) return -1;
  var n = 0, letras = m[1].toUpperCase();
  for (var i = 0; i < letras.length; i++) n = n * 26 + letras.charCodeAt(i) - 64;
  return n - 1;
}
function _pfPreviaColunaLetra_(indice) {
  var s = '', n = indice + 1;
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}
function _pfPreviaValorCelula_(celula, strings) {
  var tipo = celula.getAttribute('t');
  tipo = tipo ? tipo.getValue() : '';
  var v = celula.getChild('v', celula.getNamespace());
  var valor = v ? v.getText() : '';
  if (tipo === 's') return /^\d+$/.test(valor) ? String(strings[Number(valor)] || '') : '';
  if (tipo === 'inlineStr') {
    var inline = celula.getChild('is', celula.getNamespace());
    return inline ? _emailSpgXlsxTexto_(inline) : '';
  }
  if (tipo === 'b') return valor === '1' ? 'VERDADEIRO' : valor === '0' ? 'FALSO' : valor;
  if (!valor) {
    var formula = celula.getChild('f', celula.getNamespace());
    if (formula) return '=' + formula.getText();
  }
  return valor;
}
function _pfPreviaConsulta_(pedido) {
  if (!pedido || typeof pedido !== 'object') throw new Error('Solicitação de prévia inválida.');
  var busca = String(pedido.busca || '').trim().substring(0, 100);
  return {
    aba: _pfPreviaNumero_(pedido.aba, 500),
    linha: _pfPreviaNumero_(pedido.linha, 1000000),
    coluna: _pfPreviaNumero_(pedido.coluna, 100000),
    busca: busca,
    buscaNormalizada: busca.toLocaleLowerCase('pt-BR')
  };
}
function _pfPreviaResultado_(nome, nomes, selecionada, linhas, colunas, total, filtro, consulta) {
  var primeira = Math.min(consulta.linha, filtro.length);
  var coluna = Math.min(consulta.coluna, Math.max(0, colunas - 1));
  var ultima = Math.min(primeira + PF_PREVIA_LINHAS_PAGINA, filtro.length);
  var largura = Math.min(PF_PREVIA_COLUNAS_PAGINA, Math.max(0, colunas - coluna));
  var exibidas = [];
  for (var i = primeira; i < ultima; i++) {
    var item = filtro[i];
    var celulas = Array(largura).fill('');
    Object.keys(item.valores).forEach(function(indice) {
      var j = Number(indice) - coluna;
      if (j >= 0 && j < largura) celulas[j] = String(item.valores[indice] || '').substring(0, 2400);
    });
    exibidas.push({numero: item.numero, celulas: celulas});
  }
  var colunasVisiveis = [];
  for (var j = 0; j < largura; j++) colunasVisiveis.push(_pfPreviaColunaLetra_(coluna + j));
  return {
    arquivo: nome, abas: nomes, aba: selecionada, linha: primeira, coluna: coluna,
    totalLinhas: total, totalFiltradas: filtro.length, totalColunas: colunas,
    colunas: colunasVisiveis, linhas: exibidas, busca: consulta.busca
  };
}
function _pfPreviaXlsx_(blob, pedido) {
  var consulta = _pfPreviaConsulta_(pedido);
  if (blob.getBytes().length > 18 * 1024 * 1024) throw new Error('Excel grande demais para a prévia do portal.');
  var partes = Utilities.unzip(blob.copyBlob().setContentType('application/zip'));
  var arquivos = {};
  partes.forEach(function(parte) {
    if (/^(?:xl\/|_rels\/)/i.test(parte.getName()) && /\.(?:xml|rels)$/i.test(parte.getName())) {
      arquivos[parte.getName()] = parte.getDataAsString('UTF-8');
    }
  });
  if (!arquivos['xl/workbook.xml'] || !arquivos['xl/_rels/workbook.xml.rels']) {
    throw new Error('O Excel não contém a estrutura de abas necessária para a prévia.');
  }
  var livro = _emailSpgXlsxRaiz_(arquivos['xl/workbook.xml']);
  var ns = livro.getNamespace();
  var folhas = livro.getChild('sheets', ns);
  var relacoes = _emailSpgXlsxRaiz_(arquivos['xl/_rels/workbook.xml.rels']);
  var destinos = {};
  relacoes.getChildren('Relationship', relacoes.getNamespace()).forEach(function(r) {
    var id = r.getAttribute('Id'), tipo = r.getAttribute('Type'), alvo = r.getAttribute('Target');
    if (id && tipo && alvo && /\/worksheet$/.test(tipo.getValue()) &&
        !(r.getAttribute('TargetMode') && r.getAttribute('TargetMode').getValue() === 'External')) {
      destinos[id.getValue()] = _emailSpgXlsxCaminho_(alvo.getValue());
    }
  });
  var nomes = [], caminhos = [];
  var nsRel = XmlService.getNamespace('http://schemas.openxmlformats.org/officeDocument/2006/relationships');
  (folhas ? folhas.getChildren('sheet', ns) : []).forEach(function(f) {
    var nome = f.getAttribute('name'), id = f.getAttribute('id', nsRel);
    if (nome && id && destinos[id.getValue()]) {
      nomes.push(nome.getValue()); caminhos.push(destinos[id.getValue()]);
    }
  });
  if (!nomes.length) throw new Error('O Excel não contém abas acessíveis.');
  if (consulta.aba >= nomes.length) throw new Error('Aba da prévia não encontrada.');
  var compartilhadas = [];
  if (arquivos['xl/sharedStrings.xml']) {
    var tabela = _emailSpgXlsxRaiz_(arquivos['xl/sharedStrings.xml']);
    compartilhadas = tabela.getChildren('si', tabela.getNamespace()).map(_emailSpgXlsxTexto_);
  }
  var xml = arquivos[caminhos[consulta.aba]];
  if (!xml) throw new Error('Não foi possível ler a aba selecionada do Excel.');
  var raiz = _emailSpgXlsxRaiz_(xml);
  var sheetData = raiz.getChild('sheetData', raiz.getNamespace());
  if (!sheetData) throw new Error('Esta aba não possui área de células legível.');
  var linhasXml = sheetData.getChildren('row', sheetData.getNamespace());
  var colunas = 0;
  var dimensao = raiz.getChild('dimension', raiz.getNamespace());
  var ref = dimensao && dimensao.getAttribute('ref');
  if (ref) {
    var finalRef = ref.getValue().split(':').pop();
    colunas = Math.max(0, _pfPreviaColunaNumero_(finalRef) + 1);
  }
  var selecionadas = [], termo = consulta.buscaNormalizada, total = linhasXml.length;
  linhasXml.forEach(function(linha, i) {
    var valores = {}, numero = linha.getAttribute('r');
    if (!termo && (i < consulta.linha || i >= consulta.linha + PF_PREVIA_LINHAS_PAGINA)) {
      // Fora da página: não decodifica milhares de células sem necessidade.
      if (!ref) linha.getChildren('c', linha.getNamespace()).forEach(function(c) {
        var p = c.getAttribute('r');
        if (p) colunas = Math.max(colunas, _pfPreviaColunaNumero_(p.getValue()) + 1);
      });
      selecionadas.push({numero: numero ? Number(numero.getValue()) : i + 1, valores: valores});
      return;
    }
    linha.getChildren('c', linha.getNamespace()).forEach(function(celula) {
      var posicao = celula.getAttribute('r'), indice = posicao ? _pfPreviaColunaNumero_(posicao.getValue()) : -1;
      if (indice < 0) return;
      if (indice + 1 > colunas) colunas = indice + 1;
      var valor = _pfPreviaValorCelula_(celula, compartilhadas);
      if (valor !== '') valores[indice] = valor;
    });
    if (!termo || Object.keys(valores).some(function(c) {
      return String(valores[c]).toLocaleLowerCase('pt-BR').indexOf(termo) >= 0;
    })) {
      selecionadas.push({numero: numero ? Number(numero.getValue()) : i + 1, valores: valores});
    }
  });
  return _pfPreviaResultado_(blob.getName() || 'Chamado.xlsx', nomes, consulta.aba,
    selecionadas, Math.max(1, colunas), total, selecionadas, consulta);
}
function _pfPreviaSheets_(idArquivo, pedido) {
  var consulta = _pfPreviaConsulta_(pedido);
  var livro = SpreadsheetApp.openById(idArquivo);
  var folhas = livro.getSheets(), nomes = folhas.map(function(f) { return f.getName(); });
  if (consulta.aba >= folhas.length) throw new Error('Aba da prévia não encontrada.');
  var folha = folhas[consulta.aba];
  var total = folha.getLastRow(), colunas = Math.max(1, folha.getLastColumn());
  var encontrados = [];
  if (consulta.busca) {
    // Pesquisa por valor visível, limitada à aba escolhida; nenhuma célula é alterada.
    var valores = total ? folha.getRange(1, 1, total, colunas).getDisplayValues() : [];
    valores.forEach(function(linha, i) {
      if (linha.some(function(v) { return String(v).toLocaleLowerCase('pt-BR').indexOf(consulta.buscaNormalizada) >= 0; })) {
        encontrados.push(i + 1);
      }
    });
  } else {
    for (var i = 1; i <= total; i++) encontrados.push(i);
  }
  var primeiro = Math.min(consulta.linha, encontrados.length);
  var ultima = Math.min(primeiro + PF_PREVIA_LINHAS_PAGINA, encontrados.length);
  var coluna = Math.min(consulta.coluna, colunas - 1);
  var largura = Math.min(PF_PREVIA_COLUNAS_PAGINA, colunas - coluna);
  var linhas = [];
  if (!consulta.busca && ultima > primeiro) {
    var bloco = folha.getRange(primeiro + 1, coluna + 1, ultima - primeiro, largura).getDisplayValues();
    bloco.forEach(function(v, i) { linhas.push({numero: primeiro + i + 1, celulas: v}); });
  } else {
    for (var j = primeiro; j < ultima; j++) {
      var numero = encontrados[j];
      linhas.push({numero: numero, celulas: folha.getRange(numero, coluna + 1, 1, largura).getDisplayValues()[0]});
    }
  }
  var letras = [];
  for (var k = 0; k < largura; k++) letras.push(_pfPreviaColunaLetra_(coluna + k));
  return {arquivo: livro.getName(), abas: nomes, aba: consulta.aba, linha: primeiro,
    coluna: coluna, totalLinhas: total, totalFiltradas: encontrados.length, totalColunas: colunas,
    colunas: letras, linhas: linhas, busca: consulta.busca};
}
function obterPreviaExcelEmail(pedido) {
  if (!pedido || (pedido.origem !== 'spg' && pedido.origem !== 'antigo')) {
    throw new Error('Origem da prévia inválida.');
  }
  if (pedido.origem === 'spg') {
    var id = _emailSpgId_(pedido.idChamado);
    if (!/^[a-f0-9]{64}$/i.test(String(pedido.anexoHash || ''))) {
      throw new Error('Consulte o chamado novamente antes de visualizar o Excel.');
    }
    var pacote = _emailSpgPacote_(id, false, pedido.tipoEmail);
    if (pacote.hash !== pedido.anexoHash) {
      throw new Error('O Excel ou os dados do SPG mudaram. Volte e consulte o chamado novamente.');
    }
    return _pfPreviaXlsx_(pacote.blob, pedido);
  }
  var arquivoId = String(pedido.idArquivo || '').trim();
  if (!/^[A-Za-z0-9_-]{15,}$/.test(arquivoId)) throw new Error('ID do anexo inválido.');
  var arquivo = DriveApp.getFileById(arquivoId);
  if (arquivo.getMimeType() === 'application/vnd.google-apps.spreadsheet') {
    return _pfPreviaSheets_(arquivoId, pedido);
  }
  if (!/spreadsheetml\.sheet|application\/zip|application\/vnd\.ms-excel/i.test(arquivo.getMimeType())) {
    throw new Error('O anexo antigo não está em um formato de planilha compatível com a prévia.');
  }
  return _pfPreviaXlsx_(arquivo.getBlob().setName(arquivo.getName()), pedido);
}
