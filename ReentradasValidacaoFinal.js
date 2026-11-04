// A validacao final e um aceite explicito do usuario, independente da geracao.
// O XLSX original permanece no Drive com o mesmo ID, nome, formulas e revisoes.
var RX_CUSTO_VALIDADO_SCHEMA = 1;
var RX_CUSTO_VALIDADO_NOME = 'custos-validados.json';

function _rxCustoChamado_(nome) {
  var match = /\bCHAMADO\s+(\d+)\.xlsx$/i.exec(String(nome || ''));
  return match ? String(Number(match[1])) : '';
}

function _rxCustoRegistros_() {
  var pasta = _rxRevisaoPasta_(false);
  var arquivos = pasta && pasta.getFilesByName(RX_CUSTO_VALIDADO_NOME);
  if (!arquivos || !arquivos.hasNext()) return {schema: RX_CUSTO_VALIDADO_SCHEMA, arquivos: {}};
  var dados = JSON.parse(arquivos.next().getBlob().getDataAsString('UTF-8'));
  if (!dados || dados.schema !== RX_CUSTO_VALIDADO_SCHEMA || !dados.arquivos ||
      typeof dados.arquivos !== 'object' || Array.isArray(dados.arquivos)) {
    throw Error('O registro de validacoes precisa de verificacao. Nenhum aceite foi gravado.');
  }
  return dados;
}

function _rxCustoSalvarRegistros_(dados) {
  var pasta = _rxRevisaoPasta_(true);
  var arquivos = pasta.getFilesByName(RX_CUSTO_VALIDADO_NOME);
  var blob = Utilities.newBlob(JSON.stringify(dados), 'application/json', RX_CUSTO_VALIDADO_NOME);
  if (arquivos.hasNext()) {
    var existente = arquivos.next();
    Drive.Files.update({title: RX_CUSTO_VALIDADO_NOME, mimeType: 'application/json'},
      existente.getId(), blob, {newRevision: false, supportsAllDrives: true});
  } else {
    pasta.createFile(blob).setName(RX_CUSTO_VALIDADO_NOME);
  }
}

function _rxCustoUrlDownload_(id) {
  return 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(id);
}

function _rxCustoRegistroVigente_(registro, arquivo) {
  return !!(registro && registro.id === arquivo.id &&
    registro.atualizadoEm === arquivo.atualizadoEm &&
    registro.nome === arquivo.nome && arquivo.calculadora);
}

// Nunca e chamado durante a geracao nem ao simplesmente abrir o arquivo.
function marcarCustoValidadoReentradas(id, atualizadoEm) {
  var lock = LockService.getScriptLock(), pegou = false;
  try {
    if (!/^\d{10,16}$/.test(String(atualizadoEm || '')) ||
        !Number.isSafeInteger(Number(atualizadoEm))) throw Error('Atualize a revisao antes de validar.');
    pegou = lock.tryLock(3000);
    if (!pegou) throw Error('Outra operacao esta em andamento. Tente novamente.');
    var arq = _rxArquivo_(id, atualizadoEm);
    if (!arq.calculadora) throw Error('O aceite de custo e exclusivo da calculadora revisada.');
    var chamado = _rxCustoChamado_(arq.nome);
    if (!chamado) throw Error('Nao foi possivel identificar o chamado desta calculadora.');
    _rxRevisaoLer_(arq); // exige revisao preparada para a versao exata do XLSX
    var dados = _rxCustoRegistros_(), confirmado = _rxArquivo_(id, atualizadoEm);
    var registro = {id: confirmado.id, nome: confirmado.nome, chamado: chamado,
      atualizadoEm: confirmado.atualizadoEm, validadoEm: Date.now()};
    dados.arquivos[id] = registro;
    _rxCustoSalvarRegistros_(dados);
    return {sucesso: true, id: id, chamado: chamado, validadoEm: registro.validadoEm,
      atualizadoEm: registro.atualizadoEm, downloadUrl: _rxCustoUrlDownload_(id)};
  } catch (e) { return {erro: e.message || String(e)}; }
  finally { if (pegou) lock.releaseLock(); }
}

function _rxCustoListaDeArquivos_(arquivos) {
  var registros = _rxCustoRegistros_().arquivos;
  arquivos.forEach(function(arquivo) {
    var r = registros[arquivo.id];
    if (_rxCustoRegistroVigente_(r, arquivo)) {
      arquivo.validadoEm = r.validadoEm;
      arquivo.downloadUrl = _rxCustoUrlDownload_(arquivo.id);
    }
  });
}

// Na lista de chamados mostra a ultima validacao, mesmo que o arquivo tenha
// mudado depois. Download so aparece se a versao ainda for exatamente a validada.
function _rxCustoResumoChamados_() {
  var registros = _rxCustoRegistros_().arquivos, lista = Object.keys(registros);
  if (!lista.length) return {};
  var atuais = {}, iterator = DriveApp.getFolderById(RX_PASTA).getFiles();
  while (iterator.hasNext()) {
    var a = iterator.next(), id = a.getId();
    if (registros[id] && !a.isTrashed() && a.getMimeType() === RX_MIME) {
      atuais[id] = {id: id, nome: a.getName(), atualizadoEm: a.getLastUpdated().getTime(),
        calculadora: /^CALCULADORA FORNECEDOR /i.test(a.getName())};
    }
  }
  var porChamado = {};
  lista.forEach(function(id) {
    var r = registros[id];
    if (!r || !/^\d+$/.test(String(r.chamado)) || !Number.isSafeInteger(r.validadoEm)) return;
    var anterior = porChamado[r.chamado];
    if (!anterior || anterior.ultimaValidacaoEm < r.validadoEm) {
      var atual = atuais[id];
      porChamado[r.chamado] = {ultimaValidacaoEm: r.validadoEm,
        downloadUrl: atual && _rxCustoRegistroVigente_(r, atual) ? _rxCustoUrlDownload_(id) : ''};
    }
  });
  return porChamado;
}
