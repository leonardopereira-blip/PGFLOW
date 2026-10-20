// Excecao temporaria do legado: completa somente SKUs autorizados pelo usuario.
// Os dados gerais e a formatacao sao os mesmos do relatorio SPG, sem planilhas auxiliares.
// Globais do Apps Script duram somente a execucao atual; nenhum cache e persistido.
var _reEspecGeralCacheExecucao_ = null;

function _reEspecGeralTexto_(valor) {
  return valor == null ? '' : String(valor).trim();
}

function _reEspecGeralChave_(valor) {
  return _reEspecGeralTexto_(valor).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/gi, '').toUpperCase();
}

function _reEspecGeralIndice_(headers, nome) {
  var chave = _reEspecGeralChave_(nome), encontrados = [];
  headers.forEach(function(h, i) { if (_reEspecGeralChave_(h) === chave) encontrados.push(i); });
  if (encontrados.length !== 1) throw new Error('Coluna obrigatoria ausente ou repetida: ' + nome + '.');
  return encontrados[0];
}

function _reEspecGeralErroAusentes_(skus) {
  var erro = new Error('SKU sem especificação na alocação antiga: ' + skus.join(', ') +
    '. Você pode escolher usar a especificação técnica geral do SPG para cada SKU.');
  erro.codigo = 'ESPECIFICACAO_LEGADO_AUSENTE';
  erro.especificacoesAusentes = skus.slice();
  return erro;
}

function _reEspecGeralInteiro_(valor, minimo) {
  return typeof valor === 'number' && isFinite(valor) && Math.floor(valor) === valor && valor >= minimo;
}

function _reEspecGeralPagina_(json, pagina, base) {
  var p = json && json.pagination;
  if (!json || !Array.isArray(json.products) || !p ||
      !_reEspecGeralInteiro_(p.total, 0) || !_reEspecGeralInteiro_(p.page, 1) ||
      !_reEspecGeralInteiro_(p.limit, 1) || !_reEspecGeralInteiro_(p.totalPages, 0) ||
      p.page !== pagina || p.totalPages > 1000 || p.limit > 1000 ||
      p.totalPages !== Math.ceil(p.total / p.limit) ||
      (base && (p.total !== base.total || p.limit !== base.limit || p.totalPages !== base.totalPages)) ||
      json.products.length !== Math.min(p.limit, Math.max(0, p.total - (pagina - 1) * p.limit)) ||
      json.products.some(function(r) { return !r || typeof r !== 'object' || Array.isArray(r); })) {
    throw new Error('O SPG retornou uma especificacao geral ou paginacao invalida. Tente novamente.');
  }
  return p;
}

function _reEspecGeralLerResposta_(resposta) {
  if (resposta.getResponseCode() !== 200) {
    throw new Error('Falha na consulta da especificacao geral no SPG (HTTP ' + resposta.getResponseCode() + '). Tente novamente.');
  }
  try { return JSON.parse(resposta.getContentText()); }
  catch (e) { throw new Error('O SPG retornou uma especificacao geral invalida. Tente novamente.'); }
}

function _reEspecGeralRequisicao_(pagina, token) {
  return {
    url: SPG_CONFIG.urlEspec + (SPG_CONFIG.urlEspec.indexOf('?') < 0 ? '?' : '&') + 'page=' + pagina + '&limit=1000',
    method: 'get', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
  };
}

function _reEspecGeralResultadoCache_(skus, cache) {
  var products = [];
  skus.forEach(function(sku) { products = products.concat(cache.porSku[sku]); });
  return { products: products, dividirPesoPor100000: cache.dividirPesoPor100000 };
}

function _reEspecGeralConsultar_(skus) {
  var solicitados = skus.filter(function(sku, i) { return skus.indexOf(sku) === i; });
  var cache = _reEspecGeralCacheExecucao_;
  if (cache && solicitados.every(function(sku) {
    return Object.prototype.hasOwnProperty.call(cache.porSku, sku);
  })) return _reEspecGeralResultadoCache_(solicitados, cache);
  var coletar = solicitados.slice();
  if (cache) Object.keys(cache.porSku).forEach(function(sku) {
    if (coletar.indexOf(sku) < 0) coletar.push(sku);
  });
  var token;
  try { token = _spgLogin_(); } catch (e) { throw new Error('Nao foi possivel autenticar no SPG para consultar a especificacao geral.'); }
  if (!token) throw new Error('Nao foi possivel autenticar no SPG para consultar a especificacao geral.');
  var base = null, encontrados = [], assinaturas = Object.create(null);
  var pesoDefinido = false, dividirPesoPor100000 = false;
  function adicionar(json, pagina) {
    var p = _reEspecGeralPagina_(json, pagina, base);
    if (!base) base = p;
    var assinatura = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
      JSON.stringify(json.products), Utilities.Charset.UTF_8).join(',');
    if (assinaturas[assinatura]) throw new Error('O SPG repetiu uma pagina da especificacao geral. Tente novamente.');
    assinaturas[assinatura] = true;
    json.products.forEach(function(r) {
      if (!pesoDefinido) {
        var peso = _numeroSPG_(r.pesoKg);
        if (peso !== null) { pesoDefinido = true; dividirPesoPor100000 = Math.abs(peso) >= 1000; }
      }
      if (coletar.indexOf(_reEspecGeralTexto_(r.skuAtual)) >= 0) encontrados.push(r);
    });
  }
  var primeira = _reEspecGeralRequisicao_(1, token), resposta;
  try { resposta = UrlFetchApp.fetch(primeira.url,
    { method: primeira.method, headers: primeira.headers, muteHttpExceptions: true }); }
  catch (e) { throw new Error('Nao foi possivel consultar a especificacao geral no SPG. Tente novamente.'); }
  adicionar(_reEspecGeralLerResposta_(resposta), 1);
  // A API ignora filtros de SKU. Lotes pequenos cobrem todo o catalogo e detectam duplicatas.
  for (var pagina = 2; pagina <= base.totalPages; pagina += 4) {
    var requisicoes = [];
    for (var n = pagina; n < pagina + 4 && n <= base.totalPages; n++) {
      requisicoes.push(_reEspecGeralRequisicao_(n, token));
    }
    var respostas;
    try { respostas = UrlFetchApp.fetchAll(requisicoes); }
    catch (e) { throw new Error('Nao foi possivel consultar todas as paginas da especificacao geral no SPG. Tente novamente.'); }
    if (!Array.isArray(respostas) || respostas.length !== requisicoes.length) {
      throw new Error('O SPG nao retornou todas as paginas da especificacao geral. Tente novamente.');
    }
    respostas.forEach(function(r, indice) { adicionar(_reEspecGeralLerResposta_(r), pagina + indice); });
  }
  // Publicar a memoria somente apos validar todas as paginas. Arrays vazios
  // registram ausencias confirmadas, e uma falha de consulta nao grava nada.
  var porSku = Object.create(null);
  coletar.forEach(function(sku) { porSku[sku] = []; });
  encontrados.forEach(function(r) { porSku[_reEspecGeralTexto_(r.skuAtual)].push(r); });
  _reEspecGeralCacheExecucao_ = { porSku: porSku, dividirPesoPor100000: dividirPesoPor100000 };
  return _reEspecGeralResultadoCache_(solicitados, _reEspecGeralCacheExecucao_);
}

function _reEspecGeralAutorizacaoValida_(item) {
  return item && (typeof item.chamado === 'string' || typeof item.chamado === 'number') &&
    /^\d{2,8}$/.test(_reEspecGeralTexto_(item.chamado)) && typeof item.sku === 'string' &&
    _reEspecGeralTexto_(item.sku) !== '' && _reEspecGeralTexto_(item.sku) !== '-' &&
    item.sku.length <= 200 && !/[\u0000-\u001f\u007f]/.test(item.sku);
}

function _reSpCompletarEspecLegado_(numero, espec, tiragem, especificacoesGerais) {
  if (!espec || !tiragem || !Array.isArray(espec.headers) || !Array.isArray(espec.linhas) ||
      !Array.isArray(tiragem.headers) || !Array.isArray(tiragem.linhas)) {
    throw new Error('Matrizes do chamado legado invalidas para consultar a especificacao geral.');
  }
  var idxEspec = _reEspecGeralIndice_(espec.headers, '1. ID_Codigo SKU');
  var idxTiragem = _reEspecGeralIndice_(tiragem.headers, '1. ID_Codigo SKU');
  var presentes = Object.create(null), faltantes = [];
  espec.linhas.forEach(function(r) { presentes[_reEspecGeralTexto_(r[idxEspec])] = true; });
  tiragem.linhas.forEach(function(r) {
    var sku = _reEspecGeralTexto_(r[idxTiragem]);
    if (sku && sku !== '-' && !presentes[sku] && faltantes.indexOf(sku) < 0) faltantes.push(sku);
  });
  if (!faltantes.length) return { espec: espec, skusEspecificacaoGeral: [], skusIgnorados: [], avisos: [] };
  var autorizado = Object.create(null), candidatos = faltantes.slice();
  (Array.isArray(especificacoesGerais) ? especificacoesGerais : []).forEach(function(a) {
    if (!_reEspecGeralAutorizacaoValida_(a)) return;
    var sku = _reEspecGeralTexto_(a.sku);
    if (_reEspecGeralTexto_(a.chamado) === _reEspecGeralTexto_(numero)) autorizado[sku] = true;
    // Pre-carregar outros SKUs escolhidos no lote reduz consultas, sem autorizar
    // sua substituicao em um chamado diferente do aceite original.
    if (candidatos.indexOf(sku) < 0) candidatos.push(sku);
  });
  var pendentes = faltantes.filter(function(sku) { return !autorizado[sku]; });
  if (pendentes.length) throw _reEspecGeralErroAusentes_(pendentes);

  var consulta = _reEspecGeralConsultar_(candidatos);
  var produtosDoChamado = consulta.products.filter(function(r) {
    return faltantes.indexOf(_reEspecGeralTexto_(r.skuAtual)) >= 0;
  });
  var referencias = _criarMapaReferenciasTiragem_([tiragem.headers].concat(tiragem.linhas));
  var mapeadas = _mapearEspec(_emularDadosDaPlanilha(produtosDoChamado), CABECALHO_ESPEC,
    { dicSKU: {}, modeTechGraf: {} }, referencias, consulta.dividirPesoPor100000);
  var idxSku = CABECALHO_ESPEC.indexOf('1. ID_Código SKU'), porSku = Object.create(null);
  for (var i = 1; i < mapeadas.length; i++) {
    var linha = mapeadas[i], sku = _reEspecGeralTexto_(linha[idxSku]);
    if (porSku[sku] && JSON.stringify(porSku[sku]) !== JSON.stringify(linha)) {
      throw new Error('O SPG retornou especificacoes gerais diferentes para o SKU ' + sku + '. Confira o cadastro antes de gerar.');
    }
    porSku[sku] = linha;
  }
  var naoEncontrados = faltantes.filter(function(sku) { return !porSku[sku]; });
  var encontrados = faltantes.filter(function(sku) { return !!porSku[sku]; });

  var headers = espec.headers.slice(), existentes = Object.create(null);
  headers.forEach(function(h) { if (_reEspecGeralChave_(h)) existentes[_reEspecGeralChave_(h)] = true; });
  CABECALHO_ESPEC.forEach(function(h) {
    var chave = _reEspecGeralChave_(h);
    if (!existentes[chave]) { headers.push(h); existentes[chave] = true; }
  });
  var indices = Object.create(null);
  CABECALHO_ESPEC.forEach(function(h, i) { indices[_reEspecGeralChave_(h)] = i; });
  var linhas = espec.linhas.map(function(r) {
    return headers.map(function(h, i) { return i < espec.headers.length && r[i] != null ? r[i] : ''; });
  });
  encontrados.forEach(function(sku) {
    linhas.push(headers.map(function(h) {
      var indice = indices[_reEspecGeralChave_(h)];
      return indice === undefined ? '' : porSku[sku][indice];
    }));
  });
  return {
    espec: { headers: headers, linhas: linhas },
    skusEspecificacaoGeral: encontrados,
    skusIgnorados: naoEncontrados,
    avisos: encontrados.map(function(sku) {
      return 'SKU ' + sku + ': especificacao tecnica geral do SPG utilizada por escolha do usuario; ausente na alocacao antiga.';
    }).concat(naoEncontrados.map(function(sku) {
      return 'SKU ' + sku + ': não encontrado no catálogo geral do SPG; não será incluído na geração.';
    }))
  };
}
