// Consulta apenas a API do SPG. A opcao de alocacao legada nao participa desta lista.
function _reRecentesTexto_(valor) {
  return valor == null || typeof valor === 'object' ? '' : String(valor).trim();
}

function _reRecentesChave_(valor) {
  return _reRecentesTexto_(valor).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/[^a-z0-9]/g, '');
}

function _reRecentesCampo_(objeto, nomes) {
  if (!objeto || typeof objeto !== 'object' || Array.isArray(objeto)) return '';
  var chaves = Object.keys(objeto);
  for (var i = 0; i < nomes.length; i++) {
    for (var j = 0; j < chaves.length; j++) {
      if (_reRecentesChave_(chaves[j]) !== _reRecentesChave_(nomes[i])) continue;
      var valor = _reRecentesTexto_(objeto[chaves[j]]);
      if (valor) return valor;
    }
  }
  return '';
}

function _reRecentesTemCampo_(objeto, nome) {
  if (!objeto || typeof objeto !== 'object' || Array.isArray(objeto)) return false;
  return Object.keys(objeto).some(function(chave) {
    return _reRecentesChave_(chave) === _reRecentesChave_(nome);
  });
}

function _reRecentesNumero_(valor) {
  var numero = _reRecentesTexto_(valor);
  if (!/^\d{1,12}$/.test(numero) || Number(numero) < 1) {
    throw new Error('O SPG retornou um numero de chamado invalido.');
  }
  return String(Number(numero));
}

function _reRecentesAdicionar_(lista, valor) {
  valor = _reRecentesTexto_(valor);
  if (valor && lista.indexOf(valor) === -1) lista.push(valor);
}

function _reRecentesPagina_(dados, pagina, paginas) {
  var total = Number(dados && dados.totalPages);
  if (!dados || !Array.isArray(dados.rows) || !/^\d+$/.test(_reRecentesTexto_(dados.totalPages)) || !isFinite(total) || total < 0 ||
      total % 1 !== 0 || total > 1000 || Number(dados.page) !== pagina ||
      (pagina > 1 && total !== paginas) ||
      (total === 0 && (pagina !== 1 || dados.rows.length))) {
    throw new Error('O SPG retornou um relatorio ou paginacao invalidos.');
  }
  return total;
}

function _reRecentesAgrupar_(linhas, chamados) {
  linhas.forEach(function(linha) {
    var numero = _reRecentesNumero_(_reRecentesCampo_(linha, ['Chamado']));
    var atual = chamados[numero];
    if (!atual) atual = chamados[numero] = { chamado: numero, id: '', tipo: '',
      status: [], marca: [], envio: [], alocadoEm: [], campos: {} };
    var id = _reRecentesCampo_(linha, ['ID Chamado']);
    if (id && atual.id && id.toLowerCase() !== atual.id.toLowerCase()) {
      throw new Error('O SPG retornou identificacoes diferentes para o mesmo chamado.');
    }
    if (id) atual.id = id;
    var tipo = _reRecentesChave_(_reRecentesCampo_(linha, ['Tipo']));
    if (tipo && atual.tipo && tipo !== atual.tipo) {
      throw new Error('O SPG retornou tipos diferentes para o mesmo chamado.');
    }
    if (tipo) atual.tipo = tipo;
    if (_reRecentesTemCampo_(linha, 'Tipo')) atual.campos.tipo = true;
    var campos = { status: 'Status', marca: 'Marca', envio: 'Envio', alocadoEm: 'Alocado Em' };
    Object.keys(campos).forEach(function(campo) {
      if (_reRecentesTemCampo_(linha, campos[campo])) atual.campos[campo] = true;
      _reRecentesAdicionar_(atual[campo], _reRecentesCampo_(linha, [campos[campo]]));
    });
  });
}

function _reRecentesCompletar_(chamado, token) {
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(chamado.id)) {
    throw new Error('O SPG nao retornou uma identificacao valida para consultar o chamado.');
  }
  var resposta = _emailSpgLerJson_('/api/reentradas/' + chamado.id, token);
  var detalhe = resposta && resposta.reentrada;
  if (!detalhe || _reRecentesTexto_(detalhe.id).toLowerCase() !== chamado.id.toLowerCase() ||
      _reRecentesNumero_(detalhe.numero) !== chamado.chamado) {
    throw new Error('Nao foi possivel confirmar a identificacao do chamado no SPG.');
  }
  var tipo = _reRecentesChave_(detalhe.tipoSolicitacao);
  if (tipo) chamado.tipo = tipo;
  if (!chamado.tipo) throw new Error('O SPG nao informou o tipo de solicitacao do chamado.');
  if (chamado.tipo !== 'reentrada') return;
  if (!chamado.campos.status) _reRecentesAdicionar_(chamado.status, detalhe.status);
  if (!chamado.campos.marca) _reRecentesAdicionar_(chamado.marca, detalhe.brand && detalhe.brand.name);
  if (!chamado.campos.envio) _reRecentesAdicionar_(chamado.envio, detalhe.versao && detalhe.versao.envio);
  if (!chamado.campos.alocadoEm) _reRecentesAdicionar_(chamado.alocadoEm, detalhe.alocadoEm);
}

function listarUltimasReentradasSPG() {
  try {
    var token = _spgLogin_();
    if (!token) throw new Error('Autenticacao do SPG indisponivel.');
    var chamados = {}, paginas = 1;
    // A API pode repetir um chamado em varios SKUs e paginas, sem ordenar por numero.
    for (var pagina = 1; pagina <= paginas; pagina++) {
      var dados = _emailSpgLerJson_('/api/reentradas/sheets/relatorio?page=' + pagina + '&limit=100', token);
      paginas = _reRecentesPagina_(dados, pagina, paginas);
      _reRecentesAgrupar_(dados.rows, chamados);
    }
    var numeros = Object.keys(chamados).sort(function(a, b) { return Number(b) - Number(a); });
    var recentes = [];
    for (var i = 0; i < numeros.length && recentes.length < 30; i++) {
      var chamado = chamados[numeros[i]];
      if (chamado.campos.tipo && chamado.tipo !== 'reentrada') continue;
      // Campos vazios sao valores legitimos (por exemplo, chamado ainda nao alocado).
      // O detalhe so e necessario se a API omitir uma coluna, nunca para preencher vazios.
      if (!chamado.campos.tipo || !chamado.campos.status || !chamado.campos.marca ||
          !chamado.campos.envio || !chamado.campos.alocadoEm) {
        _reRecentesCompletar_(chamado, token);
      }
      if (chamado.tipo !== 'reentrada') continue;
      recentes.push({ chamado: chamado.chamado, status: chamado.status.join(', '),
        marca: chamado.marca.join(', '), envio: chamado.envio.join(', '), alocadoEm: chamado.alocadoEm.join(', ') });
    }
    return { chamados: recentes };
  } catch (e) {
    return { chamados: [], erro: String(e.message || e) };
  }
}
