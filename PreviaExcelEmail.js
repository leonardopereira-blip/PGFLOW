// Prévia de leitura do anexo de formalização. Não grava arquivos nem envia e-mails.
var PF_PREVIA_LINHAS_PAGINA = 50;
var PF_PREVIA_COLUNAS_PAGINA = 65;
// Conjuntos leves ficam no navegador; arquivos maiores usam a leitura paginada.
var PF_PREVIA_CONJUNTO_MAX_JSON = 1000000;
var PF_PREVIA_CONJUNTO_MAX_XML = 4000000;
var PF_PREVIA_CONJUNTO_MAX_CELULAS = 100000;

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
function _pfPreviaLivroXlsx_(blob) {
  if (blob.getBytes().length > 18 * 1024 * 1024) throw new Error('Excel grande demais para a prévia do portal.');
  var partes = Utilities.unzip(blob.copyBlob().setContentType('application/zip'));
  // As partes não utilizadas pela prévia não precisam virar texto.
  var arquivos = {}, porNome = {};
  partes.forEach(function(parte) {
    if (/^(?:xl\/|_rels\/)/i.test(parte.getName()) && /\.(?:xml|rels)$/i.test(parte.getName())) {
      porNome[parte.getName()] = parte;
    }
  });
  function ler(nome) {
    if (porNome[nome] && arquivos[nome] === undefined)
      arquivos[nome] = porNome[nome].getDataAsString('UTF-8');
    return arquivos[nome];
  }
  ler('xl/workbook.xml');
  ler('xl/_rels/workbook.xml.rels');
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
  var compartilhadas = [];
  if (porNome['xl/sharedStrings.xml']) {
    var tabela = _emailSpgXlsxRaiz_(ler('xl/sharedStrings.xml'));
    compartilhadas = tabela.getChildren('si', tabela.getNamespace()).map(_emailSpgXlsxTexto_);
  }
  caminhos.forEach(ler);
  return {nomes: nomes, caminhos: caminhos, arquivos: arquivos, compartilhadas: compartilhadas};
}
function _pfPreviaXlsx_(blob, pedido, livroPronto) {
  var consulta = _pfPreviaConsulta_(pedido);
  var livro = livroPronto || _pfPreviaLivroXlsx_(blob);
  var nomes = livro.nomes, caminhos = livro.caminhos;
  var arquivos = livro.arquivos, compartilhadas = livro.compartilhadas;
  if (consulta.aba >= nomes.length) throw new Error('Aba da prévia não encontrada.');
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

function _pfPreviaConjuntoXlsx_(blob, livroPronto) {
  var livro = livroPronto || _pfPreviaLivroXlsx_(blob);
  var tamanhoXml = livro.caminhos.reduce(function(total, caminho) {
    return total + String(livro.arquivos[caminho] || '').length;
  }, 0);
  if (tamanhoXml > PF_PREVIA_CONJUNTO_MAX_XML) return null;
  var conjunto = {arquivo: blob.getName() || 'Chamado.xlsx', abas: []};
  var celulas = 0;
  for (var a = 0; a < livro.nomes.length; a++) {
    var xml = livro.arquivos[livro.caminhos[a]];
    if (!xml) throw new Error('Não foi possível ler uma aba do Excel.');
    var raiz = _emailSpgXlsxRaiz_(xml), ns = raiz.getNamespace();
    var dados = raiz.getChild('sheetData', ns);
    if (!dados) throw new Error('Esta aba não possui área de células legível.');
    var aba = {nome: livro.nomes[a], totalColunas: 1, linhas: []};
    var dimensao = raiz.getChild('dimension', ns);
    var ref = dimensao && dimensao.getAttribute('ref');
    if (ref) aba.totalColunas = Math.max(1, _pfPreviaColunaNumero_(ref.getValue().split(':').pop()) + 1);
    var linhasXml = dados.getChildren('row', dados.getNamespace());
    for (var i = 0; i < linhasXml.length; i++) {
      var linha = linhasXml[i], numero = linha.getAttribute('r');
      var item = {numero: numero ? Number(numero.getValue()) : i + 1, valores: {}};
      var valoresXml = linha.getChildren('c', linha.getNamespace());
      for (var c = 0; c < valoresXml.length; c++) {
        if (++celulas > PF_PREVIA_CONJUNTO_MAX_CELULAS) return null;
        var referencia = valoresXml[c].getAttribute('r');
        var coluna = referencia ? _pfPreviaColunaNumero_(referencia.getValue()) : -1;
        if (coluna < 0) continue;
        aba.totalColunas = Math.max(aba.totalColunas, coluna + 1);
        var valor = _pfPreviaValorCelula_(valoresXml[c], livro.compartilhadas);
        if (valor !== '') item.valores[coluna] = valor;
      }
      aba.linhas.push(item);
    }
    conjunto.abas.push(aba);
    // Nunca entrega um conjunto parcial: busca e paginação devem alcançar tudo.
    if (JSON.stringify(conjunto).length > PF_PREVIA_CONJUNTO_MAX_JSON) return null;
  }
  return conjunto;
}

function _pfPreviaConjuntoSheets_(idArquivo) {
  var livro = SpreadsheetApp.openById(idArquivo), folhas = livro.getSheets();
  var totalCelulas = folhas.reduce(function(total, f) {
    return total + f.getLastRow() * Math.max(1, f.getLastColumn());
  }, 0);
  if (totalCelulas > PF_PREVIA_CONJUNTO_MAX_CELULAS) return null;
  var conjunto = {arquivo: livro.getName(), abas: []};
  for (var a = 0; a < folhas.length; a++) {
    var folha = folhas[a], total = folha.getLastRow(), colunas = Math.max(1, folha.getLastColumn());
    var valores = total ? folha.getRange(1, 1, total, colunas).getDisplayValues() : [];
    var aba = {nome: folha.getName(), totalColunas: colunas, linhas: []};
    valores.forEach(function(linha, i) {
      var item = {numero: i + 1, valores: {}};
      linha.forEach(function(v, c) { if (v !== '') item.valores[c] = String(v); });
      aba.linhas.push(item);
    });
    conjunto.abas.push(aba);
    if (JSON.stringify(conjunto).length > PF_PREVIA_CONJUNTO_MAX_JSON) return null;
  }
  return conjunto;
}

function _pfPreviaPacoteSpg_(pedido) {
  var id = _emailSpgId_(pedido.idChamado);
  if (!/^[a-f0-9]{64}$/i.test(String(pedido.anexoHash || ''))) {
    throw new Error('Consulte o chamado novamente antes de visualizar o Excel.');
  }
  var pacote = _emailSpgPacote_(id, false, pedido.tipoEmail);
  if (pacote.hash !== pedido.anexoHash) {
    throw new Error('O Excel ou os dados do SPG mudaram. Volte e consulte o chamado novamente.');
  }
  return pacote;
}

function _pfPreviaArquivoAntigo_(pedido) {
  var arquivoId = String(pedido.idArquivo || '').trim();
  if (!/^[A-Za-z0-9_-]{15,}$/.test(arquivoId)) throw new Error('ID do anexo inválido.');
  var arquivo = DriveApp.getFileById(arquivoId);
  if (arquivo.getMimeType() !== 'application/vnd.google-apps.spreadsheet' &&
      !/spreadsheetml\.sheet|application\/zip|application\/vnd\.ms-excel/i.test(arquivo.getMimeType())) {
    throw new Error('O anexo antigo não está em um formato de planilha compatível com a prévia.');
  }
  return {id: arquivoId, arquivo: arquivo};
}

function obterConjuntoPreviaExcelEmail(pedido) {
  if (!pedido || (pedido.origem !== 'spg' && pedido.origem !== 'antigo')) throw new Error('Origem da prévia inválida.');
  if (pedido.origem === 'spg') {
    var blobSpg = _pfPreviaPacoteSpg_(pedido).blob;
    var livroSpg = _pfPreviaLivroXlsx_(blobSpg);
    var conjuntoSpg = _pfPreviaConjuntoXlsx_(blobSpg, livroSpg);
    return {previaExcel: conjuntoSpg, pagina: conjuntoSpg ? null : _pfPreviaXlsx_(blobSpg, pedido, livroSpg)};
  }
  var fonte = _pfPreviaArquivoAntigo_(pedido);
  if (fonte.arquivo.getMimeType() === 'application/vnd.google-apps.spreadsheet') {
    var conjuntoSheets = _pfPreviaConjuntoSheets_(fonte.id);
    return {previaExcel: conjuntoSheets, pagina: conjuntoSheets ? null : _pfPreviaSheets_(fonte.id, pedido)};
  }
  var blobAntigo = fonte.arquivo.getBlob().setName(fonte.arquivo.getName());
  var livroAntigo = _pfPreviaLivroXlsx_(blobAntigo);
  var conjuntoAntigo = _pfPreviaConjuntoXlsx_(blobAntigo, livroAntigo);
  return {previaExcel: conjuntoAntigo, pagina: conjuntoAntigo ? null : _pfPreviaXlsx_(blobAntigo, pedido, livroAntigo)};
}

function baixarPreviaExcelEmail(pedido) {
  if (!pedido || (pedido.origem !== 'spg' && pedido.origem !== 'antigo')) throw new Error('Origem da prévia inválida.');
  var numero = _emailSpgId_(pedido.idChamado);
  var mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  var blob;
  if (pedido.origem === 'spg') blob = _pfPreviaPacoteSpg_(pedido).blob;
  else {
    var fonte = _pfPreviaArquivoAntigo_(pedido);
    if (fonte.arquivo.getMimeType() === 'application/vnd.google-apps.spreadsheet') {
      var resposta = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fonte.id +
        '/export?mimeType=' + encodeURIComponent(mime), {
        headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}, muteHttpExceptions: true});
      if (resposta.getResponseCode() !== 200) throw new Error('Não foi possível baixar o Excel do Drive.');
      blob = resposta.getBlob();
    } else blob = fonte.arquivo.getBlob();
    blob.setName('Chamado ' + numero + '.xlsx');
  }
  var bytes = blob.getBytes();
  if (bytes.length > 18 * 1024 * 1024) throw new Error('O Excel é grande demais para baixar pelo portal.');
  if (bytes.length < 4 || (bytes[0] & 255) !== 80 || (bytes[1] & 255) !== 75) throw new Error('Não foi possível obter um arquivo XLSX válido.');
  return {nome: blob.getName(), mimeType: mime, base64: Utilities.base64Encode(bytes)};
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
    var pacote = _pfPreviaPacoteSpg_(pedido);
    return _pfPreviaXlsx_(pacote.blob, pedido);
  }
  var fonte = _pfPreviaArquivoAntigo_(pedido), arquivoId = fonte.id, arquivo = fonte.arquivo;
  if (arquivo.getMimeType() === 'application/vnd.google-apps.spreadsheet') {
    return _pfPreviaSheets_(arquivoId, pedido);
  }
  if (!/spreadsheetml\.sheet|application\/zip|application\/vnd\.ms-excel/i.test(arquivo.getMimeType())) {
    throw new Error('O anexo antigo não está em um formato de planilha compatível com a prévia.');
  }
  return _pfPreviaXlsx_(arquivo.getBlob().setName(arquivo.getName()), pedido);
}
