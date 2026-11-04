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

function _rxCustoValorNumero_(valor,descricao){
  var texto=String(valor===undefined?'':valor).trim();
  if(!/^-?\d+(?:\.\d+)?$/.test(texto))throw Error(descricao+' ainda nao possui resultado numerico calculado.');
  var n=Number(texto);if(!isFinite(n)||n<0)throw Error(descricao+' possui valor invalido.');
  return n;
}
// Le exclusivamente os CAMPOS FINAIS do CALC Geral, jamais valores do SPG
// anteriores a revisao. Confere o TOTAL de exemplares com as quatro tiragens.
function _rxCustoTotaisRevisados_(modelo){
  var abas=modelo&&modelo.abas||[],aba=abas.find(function(a){return /^CALC[\s-]*Geral$/i.test(a.nome);});
  if(!aba)throw Error('A etapa Resumo de custos nao foi encontrada para apurar o chamado.');
  var linhas=aba.linhas||[],cabecalho=(linhas.find(function(r){return r.numero===3;})||{}).valores||{};
  var subtitulo=(linhas.find(function(r){return r.numero===4;})||{}).valores||{};
  var colSku=Number(Object.keys(cabecalho).find(function(k){return /C[ÓO]DIGO\s*SKU/i.test(cabecalho[k]);}));
  var colTiragem=Number(Object.keys(cabecalho).find(function(k){return /^TOTAL$/i.test(String(cabecalho[k]).trim())&&
    /Exemplares/i.test(String(subtitulo[k]||''));}));
  var colPreco=Number(Object.keys(cabecalho).find(function(k){return /SERVI[CÇ]O\s*\+\s*PAPEL\s+TOTAL/i.test(String(cabecalho[k]));}));
  if(!Number.isInteger(colSku)||!Number.isInteger(colTiragem)||!Number.isInteger(colPreco)||
    colTiragem<=colSku+2||colTiragem>colSku+8||
    !Object.keys(cabecalho).length||!(''+cabecalho[colSku]).match(/SKU/i)||
    !/BRL\s*\/\s*Total/i.test(String(subtitulo[colPreco+1]||'')))
    throw Error('As colunas de tiragem e preco final deste molde precisam ser conferidas antes do aceite.');
  var tiragem=0,valor=0,quantidade=0;
  linhas.forEach(function(r){
    if(r.numero<=5)return;
    var v=r.valores||{},sku=String(v[colSku]||'').trim();if(!sku||/^=/.test(sku))return;
    var total=_rxCustoValorNumero_(v[colTiragem],'Total de exemplares do SKU '+sku);
    var partes=0;
    for(var c=colSku+2;c<colTiragem;c++)partes+=_rxCustoValorNumero_(v[c]===undefined?'0':v[c],'Tiragem do SKU '+sku);
    if(Math.abs(partes-total)>0.001)throw Error('O total de exemplares do SKU '+sku+
      ' diverge da tiragem informada. Recalcule a calculadora antes de validar o custo.');
    tiragem+=total;
    valor+=_rxCustoValorNumero_(v[colPreco+1],'Preco final do SKU '+sku);
    quantidade++;
  });
  if(!quantidade)throw Error('Nenhum SKU calculado foi encontrado para validar o custo.');
  return {tiragemValidada:tiragem,valorChamado:Math.round((valor+Number.EPSILON)*100)/100};
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
    var revisao=_rxRevisaoLer_(arq); // exige revisao preparada para a versao exata do XLSX
    if(revisao&&revisao.recalculoPendente)
      throw Error('Este custo foi alterado e ainda precisa recalcular os resultados das formulas antes do aceite final.');
    var totais=_rxCustoTotaisRevisados_(revisao);
    var dados = _rxCustoRegistros_(), confirmado = _rxArquivo_(id, atualizadoEm);
    var registro = {id: confirmado.id, nome: confirmado.nome, chamado: chamado,
      atualizadoEm: confirmado.atualizadoEm, validadoEm: Date.now(),
      tiragemValidada:totais.tiragemValidada,valorChamado:totais.valorChamado};
    dados.arquivos[id] = registro;
    _rxCustoSalvarRegistros_(dados);
    return {sucesso: true, id: id, chamado: chamado, validadoEm: registro.validadoEm,
      tiragemValidada:registro.tiragemValidada,valorChamado:registro.valorChamado,
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
  var atuais = {}, porChamado = {}, iterator = DriveApp.getFolderById(RX_PASTA).getFiles();
  while (iterator.hasNext()) {
    var a = iterator.next(), id = a.getId();
    if(a.isTrashed()||a.getMimeType()!==RX_MIME)continue;
    var nome=a.getName(), chamadoArquivo=/^CALCULADORA FORNECEDOR .*\bCHAMADO\s+(\d+)\.xlsx$/i.exec(nome);
    if(chamadoArquivo){
      var chamadoGerado=String(Number(chamadoArquivo[1]));
      var dataCriacao=a.getDateCreated().getTime(),anteriorGerado=porChamado[chamadoGerado];
      if(!anteriorGerado||!anteriorGerado.ultimoGeradoEm||anteriorGerado.ultimoGeradoEm<dataCriacao){
        porChamado[chamadoGerado]=Object.assign({},anteriorGerado||{},
          {ultimoGeradoEm:dataCriacao,ultimoArquivoId:id});
      }
    }
    if (registros[id]) {
      atuais[id] = {id: id, nome: a.getName(), atualizadoEm: a.getLastUpdated().getTime(),
        calculadora: /^CALCULADORA FORNECEDOR /i.test(a.getName())};
    }
  }
  lista.forEach(function(id) {
    var r = registros[id];
    if (!r || !/^\d+$/.test(String(r.chamado)) || !Number.isSafeInteger(r.validadoEm)) return;
    var anterior = porChamado[r.chamado];
    if (!anterior || !anterior.ultimaValidacaoEm || anterior.ultimaValidacaoEm < r.validadoEm) {
      var atual = atuais[id];
      porChamado[r.chamado] = Object.assign({},anterior||{},{ultimaValidacaoEm: r.validadoEm,
        tiragemValidada:r.tiragemValidada,valorChamado:r.valorChamado,
        downloadUrl: atual && _rxCustoRegistroVigente_(r, atual) ? _rxCustoUrlDownload_(id) : ''});
    }
  });
  return porChamado;
}
