// Envio de formalizacao via SPG, independente das planilhas operacionais.
// O fluxo historico continua isolado em Menu_Emails.js durante a transicao.
var EMAIL_SPG_BASE = 'https://spg-api.prod.arcocv.co';
var EMAIL_SPG_DESTINATARIOS_PROP = 'EMAIL_SPG_DESTINATARIOS_GRAFICAS';
var EMAIL_SPG_CADASTRO_EDITADO_PROP = 'EMAIL_SPG_CADASTRO_EDITADO';

function _emailSpgId_(id) {
  var valor = String(id == null ? '' : id).trim();
  if (!/^\d{1,12}$/.test(valor) || Number(valor) < 1) throw new Error('Informe um numero de chamado valido.');
  return valor;
}

function _emailSpgNormalizar_(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/[^a-z0-9]/g, '');
}

function _emailSpgCampo_(origem, nomes) {
  var alvos = {};
  nomes.forEach(function(n) { alvos[_emailSpgNormalizar_(n)] = true; });
  var fila = [{valor: origem, nivel: 0}], vistos = 0;
  while (fila.length && vistos++ < 400) {
    var item = fila.shift(), obj = item.valor;
    if (!obj || typeof obj !== 'object' || item.nivel > 5) continue;
    if (Array.isArray(obj)) {
      obj.slice(0, 10).forEach(function(v) { fila.push({valor: v, nivel: item.nivel + 1}); });
      continue;
    }
    Object.keys(obj).forEach(function(k) {
      var valor = obj[k];
      if (alvos[_emailSpgNormalizar_(k)] && valor != null &&
          typeof valor !== 'object' && String(valor).trim()) {
        item.encontrado = item.encontrado || String(valor).trim();
      }
    });
    if (item.encontrado) return item.encontrado;
    Object.keys(obj).forEach(function(k) {
      var v = obj[k];
      if (v && typeof v === 'object') fila.push({valor: v, nivel: item.nivel + 1});
    });
  }
  return '';
}


function _emailSpgDescricao_(json) {
  // Nunca escolher a descricao de um SKU apenas por existir em "items".
  var principal = _emailSpgCampo_(json, [
    'indexadorTiragem', 'indexador', 'descricaoProducao',
    'descricaoOP', 'descricaoChamado', 'op']);
  if (principal) return principal;
  var possiveis = [json, json && json.chamado, json && json.solicitacao,
    json && json.data, json && json.reentrada, json && json.dados];
  for (var i = 0; i < possiveis.length; i++) {
    var objeto = possiveis[i];
    if (!objeto || typeof objeto !== 'object' || Array.isArray(objeto)) continue;
    var chaves = Object.keys(objeto);
    for (var j = 0; j < chaves.length; j++) {
      if (_emailSpgNormalizar_(chaves[j]) !== 'descricao') continue;
      var valor = objeto[chaves[j]];
      if (valor != null && typeof valor !== 'object' && String(valor).trim()) {
        return String(valor).trim();
      }
    }
  }
  return '';
}

function _emailSpgApi_(path, token) {
  var resposta = UrlFetchApp.fetch(EMAIL_SPG_BASE + path, {
    method: 'get',
    headers: {Authorization: 'Bearer ' + token},
    muteHttpExceptions: true,
    followRedirects: false
  });
  return resposta;
}

function _emailSpgJson_(resposta) {
  var codigo = resposta.getResponseCode();
  if (codigo !== 200) return null;
  var texto = resposta.getContentText();
  try { return JSON.parse(texto); } catch (e) { return null; }
}

function _emailSpgData_(valor) {
  if (valor == null || valor === '') return '';
  var str = String(valor).trim(), match;
  if ((match = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(str))) {
    return _emailSpgDataValida_(+match[1], +match[2], +match[3]) ? match[1] + '/' + match[2] + '/' + match[3] : '';
  }
  if ((match = /^(\d{4})-(\d{2})-(\d{2})/.exec(str))) {
    return _emailSpgDataValida_(+match[3], +match[2], +match[1]) ? match[3] + '/' + match[2] + '/' + match[1] : '';
  }
  return '';
}

function _emailSpgDataValida_(dia, mes, ano) {
  var data = new Date(ano, mes - 1, dia, 12);
  return ano >= 2000 && ano <= 2100 && data.getDate() === dia &&
    data.getMonth() === mes - 1 && data.getFullYear() === ano;
}

function _emailSpgColeta_(dataDesejada) {
  var normal = _emailSpgData_(dataDesejada);
  if (!normal) return '';
  var p = normal.split('/').map(Number);
  var d = new Date(p[2], p[1] - 1, p[0], 12);
  d.setDate(d.getDate() - 2);
  // Replica a regra historica: nao programar coleta sexta, sabado ou domingo.
  if (d.getDay() === 5) d.setDate(d.getDate() - 1);
  else if (d.getDay() === 6) d.setDate(d.getDate() - 2);
  else if (d.getDay() === 0) d.setDate(d.getDate() - 3);
  return ('0' + d.getDate()).slice(-2) + '/' +
    ('0' + (d.getMonth() + 1)).slice(-2) + '/' + d.getFullYear();
}

function _emailSpgHex_(bytes) {
  return bytes.map(function(b) { var x = (b + 256) % 256; return ('0' + x.toString(16)).slice(-2); }).join('');
}

function _emailSpgHash_(blob) {
  // O ZIP pode ter timestamps diferentes. Inclui textos e relacionamentos que
  // determinam o conteudo do Excel, excluindo apenas docProps dinamicos.
  var partes = Utilities.unzip(blob.copyBlob().setContentType('application/zip'))
    .filter(function(p) { return /^(?:xl\/.*\.(?:xml|rels)|\[Content_Types\]\.xml|_rels\/\.rels)$/i.test(p.getName()); })
    .sort(function(a, b) { return a.getName().localeCompare(b.getName()); })
    .map(function(p) { return p.getName() + ':' + p.getDataAsString('UTF-8'); });
  if (!partes.length) throw new Error('Excel sem planilhas para validar.');
  return _emailSpgHex_(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, partes.join('\n')));
}

function _emailSpgErroAnexo_(mensagem) { return new Error(mensagem); }

function _emailSpgChecarXlsx_(arquivo, versaoId) {
  var bytes = arquivo.getBytes();
  if (bytes.length < 4 || (bytes[0] & 255) !== 80 || (bytes[1] & 255) !== 75) throw _emailSpgErroAnexo_('O SPG não retornou um Excel válido para este chamado.', 'anexo', 'Excel de produção');
  if (bytes.length > 18 * 1024 * 1024) throw new Error('O Excel retornado pelo SPG é grande demais para este envio.');
  var partes;
  try { partes = Utilities.unzip(arquivo.copyBlob().setContentType('application/zip')); }
  catch (e) { throw _emailSpgErroAnexo_('O SPG não retornou um Excel válido para este chamado.', 'anexo', 'Excel de produção'); }
  var arquivos = {};
  var corresponde = false;
  partes.forEach(function (parte) {
    if (/\.(xml|rels)$/.test(parte.getName())) {
      var xml = parte.getDataAsString();
      arquivos[parte.getName()] = xml;
      if (versaoId && new RegExp('https://spg\\.prod\\.arcocv\\.co/(?:op|tiragem)/' + versaoId + '(?:[?/#\\x22&]|$)', 'i').test(xml)) corresponde = true;
    }
  });
  if (!arquivos['xl/workbook.xml'] || !arquivos['xl/_rels/workbook.xml.rels']) throw _emailSpgErroAnexo_('O Excel de produção está incompleto no SPG.', 'anexo', 'Excel de produção');
  if (versaoId && !corresponde) throw new Error('Não foi possível confirmar que o anexo corresponde à versão deste chamado.');
  var raiz = _emailSpgXlsxRaiz_(arquivos['xl/workbook.xml']);
  var ns = raiz.getNamespace();
  var folhas = raiz.getChild('sheets', ns);
  var relacoes = _emailSpgXlsxRaiz_(arquivos['xl/_rels/workbook.xml.rels']);
  var caminhos = {};
  var textos = [];
  relacoes.getChildren('Relationship', relacoes.getNamespace()).forEach(function (relacao) {
    var modo = relacao.getAttribute('TargetMode');
    if (modo && modo.getValue() === 'External') return;
    var id = relacao.getAttribute('Id');
    var alvo = relacao.getAttribute('Target');
    var tipo = relacao.getAttribute('Type');
    if (!id || !alvo || !tipo) return;
    var caminho = _emailSpgXlsxCaminho_(alvo.getValue());
    if (/\/worksheet$/.test(tipo.getValue())) caminhos[id.getValue()] = caminho;
    if (/\/sharedStrings$/.test(tipo.getValue()) && arquivos[caminho]) {
      var strings = _emailSpgXlsxRaiz_(arquivos[caminho]);
      textos = strings.getChildren('si', strings.getNamespace()).map(_emailSpgXlsxTexto_);
    }
  });
  var abas = {};
  var nsRelacao = XmlService.getNamespace('http://schemas.openxmlformats.org/officeDocument/2006/relationships');
  (folhas ? folhas.getChildren('sheet', ns) : []).forEach(function (folha) {
    var nome = folha.getAttribute('name');
    var id = folha.getAttribute('id', nsRelacao);
    if (nome && id) abas[nome.getValue()] = caminhos[id.getValue()];
  });
  var sku = '1. ID_Código SKU';
  var kit = '1. ID_Código Kit';
  var prioridades = [
    { nome: 'Tiragem (legado)', colunas: [sku, kit, '15. DES_CD envio', '16. TIR_SKU'], quantidade: true },
    { nome: 'Esp. Técnica (legado)', colunas: [sku] },
    { nome: 'Árvore (legado)', colunas: [sku, kit] }
  ];
  prioridades.forEach(function (aba) {
    var xml = arquivos[abas[aba.nome]];
    if (!xml || !_emailSpgXlsxTemDados_(xml, textos, aba.colunas, aba.quantidade)) {
      throw _emailSpgErroAnexo_('O Excel do SPG não contém dados de produção na aba necessária: ' + aba.nome + '.', 'anexo', 'Dados da aba ' + aba.nome);
    }
  });
  return true;
}

function _emailSpgXlsxRaiz_(xml) {
  try { return XmlService.parse(xml).getRootElement(); }
  catch (e) { throw _emailSpgErroAnexo_('A estrutura do Excel de produção está incompleta ou inválida no SPG.', 'anexo', 'Excel de produção'); }
}

function _emailSpgXlsxCaminho_(alvo) {
  var partes = (alvo.charAt(0) === '/' ? alvo.substring(1) : 'xl/' + alvo).split('/');
  var caminho = [];
  partes.forEach(function (parte) {
    if (parte === '..') caminho.pop();
    else if (parte && parte !== '.') caminho.push(parte);
  });
  return caminho.join('/');
}

function _emailSpgXlsxTexto_(elemento) {
  var ns = elemento.getNamespace();
  var texto = elemento.getChild('t', ns);
  if (texto) return texto.getText();
  return elemento.getChildren('r', ns).map(function (trecho) {
    var t = trecho.getChild('t', ns);
    return t ? t.getText() : '';
  }).join('');
}

function _emailSpgXlsxTemDados_(xml, textos, colunas, temQuantidade) {
  var raiz = _emailSpgXlsxRaiz_(xml);
  var ns = raiz.getNamespace();
  var dados = raiz.getChild('sheetData', ns);
  if (!dados) return false;
  var cabecalho = null;
  return dados.getChildren('row', ns).some(function (linha) {
    var valores = {};
    linha.getChildren('c', ns).forEach(function (celula) {
      var referencia = celula.getAttribute('r');
      var coluna = referencia && /^([A-Z]+)\d+$/.exec(referencia.getValue());
      if (!coluna) return;
      var tipo = celula.getAttribute('t');
      tipo = tipo ? tipo.getValue() : '';
      var v = celula.getChild('v', ns);
      var valor = v ? v.getText() : '';
      if (tipo === 's') valor = /^\d+$/.test(valor) ? textos[Number(valor)] || '' : '';
      else if (tipo === 'inlineStr') {
        var inline = celula.getChild('is', ns);
        valor = inline ? _emailSpgXlsxTexto_(inline) : '';
      } else if (tipo === 'e') valor = '';
      valores[coluna[1]] = String(valor).trim();
    });
    if (!cabecalho) {
      var encontradas = colunas.map(function (nome) {
        return Object.keys(valores).filter(function (coluna) { return valores[coluna] === nome; })[0];
      });
      if (encontradas.every(function (coluna) { return !!coluna; })) cabecalho = encontradas;
      return false;
    }
    return cabecalho.every(function (coluna, indice) {
      var valor = valores[coluna] || '';
      if (temQuantidade && indice === cabecalho.length - 1) return valor !== '' && isFinite(Number(valor)) && Number(valor) >= 0;
      return /^[a-z0-9][a-z0-9._\/-]*$/i.test(valor);
    });
  });
}

function _emailSpgBlobDeResposta_(resposta, id, versaoId) {
  if (resposta.getResponseCode() !== 200) {
    throw new Error('PPG nao encontrada no SPG (HTTP ' + resposta.getResponseCode() + ').');
  }
  var bytes = resposta.getContent();
  var blob;
  if (bytes.length >= 2 && ((bytes[0] + 256) % 256) === 80 &&
      ((bytes[1] + 256) % 256) === 75) {
    blob = Utilities.newBlob(bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  } else {
    var json = _emailSpgJson_(resposta);
    var base64 = json && _emailSpgCampo_(json, [
      'base64', 'arquivoBase64', 'fileBase64', 'contentBase64', 'excelBase64', 'arquivo']);
    if (!base64 || !/^[A-Za-z0-9+\/=\s]+$/.test(base64)) {
      throw new Error('O endpoint da PPG nao retornou o Excel completo.');
    }
    try {
      blob = Utilities.newBlob(Utilities.base64Decode(base64.replace(/\s/g, '')),
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    } catch (e) { throw new Error('O Excel retornado pelo SPG esta corrompido.'); }
  }
  blob.setName('Chamado ' + id + '.xlsx');
  _emailSpgChecarXlsx_(blob, versaoId);
  return blob;
}

function _emailSpgUuid_(valor) {
  var id = String(valor || '');
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id)) {
    throw new Error('O SPG nao retornou uma identificacao valida para o chamado ou sua producao.');
  }
  return id;
}

function _emailSpgLerJson_(path, token) {
  var resposta = _emailSpgApi_(path, token);
  var dados = _emailSpgJson_(resposta);
  if (!dados) throw new Error('Nao foi possivel consultar os dados do SPG (HTTP ' + resposta.getResponseCode() + ').');
  return dados;
}

function _emailSpgResolverId_(numero, token) {
  var ids = [], paginas = 1;
  // O nome da rota e "sheets", mas esta consulta e diretamente na API.
  // O relatorio so resolve a identidade; dados do email vem do detalhe do chamado.
  for (var pagina = 1; pagina <= paginas; pagina++) {
    var r = _emailSpgLerJson_('/api/reentradas/sheets/relatorio?page=' + pagina + '&limit=100', token);
    var totalPaginas = Number(r.totalPages);
    if (!Array.isArray(r.rows) || !isFinite(totalPaginas) || totalPaginas < 1 ||
        totalPaginas % 1 !== 0 || totalPaginas > 1000 || Number(r.page) !== pagina ||
        (pagina > 1 && totalPaginas !== paginas)) {
      throw new Error('O SPG retornou um relatorio ou paginacao invalidos.');
    }
    paginas = totalPaginas;
    r.rows.forEach(function(linha) {
      if (String(linha.Chamado) !== numero) return;
      var id = _emailSpgUuid_(linha['ID Chamado']);
      if (ids.indexOf(id) === -1) ids.push(id);
    });
    if (ids.length > 1) throw new Error('O SPG retornou identificacoes diferentes para o mesmo chamado.');
  }
  if (!ids.length) {
    var ausente = new Error('As informacoes do chamado nao foram encontradas no SPG.');
    ausente.codigo = 'CHAMADO_NAO_ENCONTRADO';
    throw ausente;
  }
  return ids[0];
}

function _emailSpgStatus_(valor) {
  return typeof valor === 'string' ? valor.trim() : '';
}

function _emailSpgCancelado_(status) {
  return /^CANCELAD[OA]$/.test(_emailSpgStatus_(status).toUpperCase());
}

function _emailSpgTipoFormalizacao_(detalhe) {
  var tipo = _emailSpgNormalizar_(detalhe.tipoSolicitacao);
  var motivo = _emailSpgStatus_(detalhe.justificativa).normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '').toLowerCase();
  if (tipo === 'reentrada') return 'reentrada';
  if (tipo === 'inclusivos' || (tipo === 'spot' && /(^|[^a-z0-9])ampliad[oa]s?($|[^a-z0-9])/.test(motivo))) return 'inclusivo';
  return '';
}

function identificarTipoEmailSPG(idChamado) {
  var numero = String(idChamado == null ? '' : idChamado).trim();
  try {
    numero = _emailSpgId_(numero);
    var token = _spgLogin_();
    if (!token) throw new Error('Autenticacao do SPG indisponivel.');
    var id = _emailSpgResolverId_(numero, token);
    var detalhe = _emailSpgLerJson_('/api/reentradas/' + id, token).reentrada;
    if (!detalhe || detalhe.id !== id || String(detalhe.numero) !== numero) {
      throw new Error('Nao foi possivel confirmar a identificacao do chamado no SPG.');
    }
    // A selecao usa apenas o tipo e a justificativa do chamado confirmado.
    // Dados de produtos, descricao da producao e anexo nao identificam o modelo.
    var tipoSolicitacao = _emailSpgStatus_(detalhe.tipoSolicitacao);
    var tipoEmail = _emailSpgTipoFormalizacao_(detalhe);
    var statusChamado = _emailSpgStatus_(detalhe.status);
    var statusProducao = '';
    var versaoId = String(detalhe.versaoId || '');
    if (/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(versaoId) &&
        detalhe.versao && detalhe.versao.id === versaoId) {
      statusProducao = _emailSpgStatus_(detalhe.versao.status);
    }
    return {encontrado: true, idChamado: numero, tipoEmail: tipoEmail,
      tipoSolicitacao: tipoSolicitacao, statusChamado: statusChamado,
      statusProducao: statusProducao,
      cancelado: _emailSpgCancelado_(statusChamado) || _emailSpgCancelado_(statusProducao)};
  } catch (e) {
    return {encontrado: false, idChamado: numero, erro: String(e.message || e)};
  }
}

function _emailSpgPacote_(idChamado, apenasMetadados, tipoEmail) {
  var statusChamado = '', statusProducao = '';
  try {
    var numero = _emailSpgId_(idChamado), token = _spgLogin_();
    if (!token) throw new Error('Autenticacao do SPG indisponivel.');
    var id = _emailSpgResolverId_(numero, token);
    var detalhe = _emailSpgLerJson_('/api/reentradas/' + id, token).reentrada;
    if (!detalhe || detalhe.id !== id || String(detalhe.numero) !== numero) {
      throw new Error('Nao foi possivel confirmar a identificacao do chamado no SPG.');
    }
    statusChamado = _emailSpgStatus_(detalhe.status);
    if (detalhe.versao && detalhe.versao.id === detalhe.versaoId) {
      statusProducao = _emailSpgStatus_(detalhe.versao.status);
    }
    if (_emailSpgCancelado_(statusChamado)) throw new Error('Este chamado está cancelado no SPG. O envio de formalização está bloqueado.');
    var versaoId = _emailSpgUuid_(detalhe.versaoId);
    if (!detalhe.versao || detalhe.versao.id !== versaoId) {
      throw new Error('A versao de producao do chamado esta incompleta ou inconsistente no SPG.');
    }
    var versao = _emailSpgLerJson_('/api/tiragem/' + versaoId, token).versao;
    if (!versao || versao.id !== versaoId) throw new Error('Nao foi possivel confirmar a versao da producao no SPG.');
    statusProducao = _emailSpgStatus_(versao.status);
    if (_emailSpgCancelado_(statusProducao)) {
      throw new Error('A versão de produção deste chamado está cancelada no SPG. O envio de formalização está bloqueado.');
    }
    if (tipoEmail === 'inclusivo' && _emailSpgTipoFormalizacao_(detalhe) !== 'inclusivo') {
      throw new Error('O chamado não está identificado como inclusivo no SPG. Confira o tipo e a justificativa.');
    }
    // Mesmo a escolha de historico precisa confirmar a situacao atual no SPG.
    // O Excel atual nao e necessario quando o usuario pediu uma versao antiga.
    if (apenasMetadados) return {id: numero, statusChamado: statusChamado, statusProducao: statusProducao};
    var respostaOp = _emailSpgApi_('/api/reentradas/' + id + '/op', token);
    var blob = _emailSpgBlobDeResposta_(respostaOp, numero, versaoId);
    // Le somente os campos do chamado e da versao confirmada, nunca de seus SKUs.
    var data = _emailSpgData_(detalhe.dataDesejadaEntrega);
    var descricao = String(versao.descricao || '').trim();
    var op = String(versao.opCode || '').trim();
    var opParcial = versao.opParcial === true;
    // O mesmo marcador protege o anexo, os dados e a situacao mostrada na revisao.
    var hash = _emailSpgHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
      JSON.stringify([id, numero, versaoId, data, descricao, op, opParcial,
        statusChamado, statusProducao, detalhe.tipoSolicitacao, detalhe.justificativa, _emailSpgHash_(blob)])));
    return {id: numero, idSpg: id, versaoId: versaoId, dataDesejada: data,
      descricao: descricao, op: op, opParcial: opParcial, hash: hash, blob: blob,
      statusChamado: statusChamado, statusProducao: statusProducao};
  } catch (e) {
    // Preserva os status confirmados mesmo se outro dado ou o Excel estiver incompleto.
    e.statusChamado = statusChamado;
    e.statusProducao = statusProducao;
    e.cancelado = _emailSpgCancelado_(statusChamado) || _emailSpgCancelado_(statusProducao);
    throw e;
  }
}

function consultarPacoteEmailSPG(idChamado, opcaoVersao, tipoEmail) {
  if (tipoEmail && tipoEmail !== 'reentrada' && tipoEmail !== 'inclusivo') {
    return {status: 'incompleto', anexoValido: false,
      erro: 'Tipo de formalizacao nao reconhecido.'};
  }
  var id;
  try {
    id = _emailSpgId_(idChamado);
    var p = _emailSpgPacote_(id, opcaoVersao === 'manual', tipoEmail);
    // Historico por versao somente quando o SPG expuser selecao inequívoca.
    if (opcaoVersao === 'manual') return {
      status: 'incompleto', idChamado: p.id, anexoValido: false,
      statusChamado: p.statusChamado, statusProducao: p.statusProducao, cancelado: false,
      erro: 'O SPG fornece a PPG atual. Para selecionar uma versao antiga, utilize o fluxo antigo.'
    };
    var faltantes = [];
    if (!p.dataDesejada) faltantes.push('Data desejada de entrega');
    if (!p.descricao) faltantes.push('Descricao da producao');
    return {status: faltantes.length ? 'incompleto' : 'ok',
      idChamado: p.id, anexoValido: true, anexoHash: p.hash,
      statusChamado: p.statusChamado, statusProducao: p.statusProducao, cancelado: false,
      dataDesejada: p.dataDesejada, dataColeta: _emailSpgColeta_(p.dataDesejada),
      descricao: p.descricao, op: p.op, opParcial: p.opParcial, versaoId: p.versaoId, faltantes: faltantes};
  } catch (e) {
    return {status: e.cancelado ? 'cancelado' : 'incompleto', idChamado: id || '',
      statusChamado: e.statusChamado || '', statusProducao: e.statusProducao || '',
      cancelado: e.cancelado === true, anexoValido: false, erro: e.message, faltantes: []};
  }
}

function _emailSpgDestinos_(raw) {
  if (!Array.isArray(raw) || !raw.length || raw.length > 80) throw new Error('Destinatarios invalidos.');
  var grupos = [], total = 0;
  raw.forEach(function(grupo) {
    var emails = [], vistos = {};
    String(grupo || '').split(/[;,]/).forEach(function(email) {
      var e = email.trim().toLowerCase();
      if (!e) return;
      if (!/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(e)) {
        throw new Error('E-mail de destinatario invalido: ' + e);
      }
      if (!vistos[e]) { vistos[e] = true; emails.push(e); }
    });
    if (!emails.length) throw new Error('Informe os destinatarios de cada grafica.');
    // Cada grupo representa uma grafica, mesmo com destinatarios iguais.
    total += emails.length;
    grupos.push(emails.join(', '));
  });
  if (!grupos.length || total > 80) throw new Error('Informe de 1 a 80 destinatarios.');
  return grupos;
}

function _emailSpgSobrescrita_(pacote, dados) {
  var m = dados.sobrescrita || {}, data = pacote.dataDesejada,
      descricao = pacote.descricao;
  if (!data) {
    if (m.omitirData === true) data = '';
    else {
      data = _emailSpgData_(m.dataDesejada || '');
      if (!data) throw new Error('A data desejada precisa ser preenchida ou omitida explicitamente.');
    }
  }
  if (!descricao) {
    if (m.omitirDescricao === true) descricao = '';
    else {
      descricao = String(m.descricao || '').trim();
      if (!descricao) throw new Error('Preencha a descricao ou escolha seguir sem ela.');
    }
  }
  return {data: data, descricao: descricao, dataColeta: _emailSpgColeta_(data)};
}

function dispararEmailPacoteSPG(dados) {
  try {
    if (!dados || typeof dados !== 'object') throw new Error('Dados do envio invalidos.');
    if (dados.tipoEmail && dados.tipoEmail !== 'reentrada' && dados.tipoEmail !== 'inclusivo') {
      throw new Error('Tipo de formalização não reconhecido.');
    }
    var id = _emailSpgId_(dados.idChamado), emails = _emailSpgDestinos_(dados.arrayDestinos);
    var assunto = String(dados.assunto || '').trim(), corpo = String(dados.corpo || '').trim();
    if (!assunto || !corpo) throw new Error('Assunto e corpo sao obrigatorios.');
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) throw new Error('Ha outro envio em andamento. Tente novamente.');
    try {
      // Refaz consulta e validacao: nunca confia no ID de arquivo ou origem informados pelo navegador.
      var p = _emailSpgPacote_(id, false, dados.tipoEmail);
      if (p.hash !== String(dados.anexoHash || '')) {
        throw new Error('A PPG ou os dados do chamado foram alterados no SPG. Retorne para a revisao antes de enviar.');
      }
      _emailSpgSobrescrita_(p, dados);
      var cache = CacheService.getUserCache();
      var chaveEnvio = 'emailSpg_' + _emailSpgHex_(Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        id + '|' + assunto + '|' + corpo + '|' + p.hash + '|' + emails.join(';'))).substring(0, 48);
      if (cache.get(chaveEnvio)) {
        throw new Error('Este envio ja foi iniciado recentemente. Verifique os e-mails enviados antes de repetir.');
      }
      cache.put(chaveEnvio, 'INICIADO', 600);
      var enviados = 0;
      try {
        // Mantem o alias de remetente do envio historico, sem alterar as permissoes.
        emails.forEach(function(destino) {
          GmailApp.sendEmail(destino, assunto, corpo, {
            attachments: [p.blob.copyBlob()],
            from: 'joao.alberti@arcoeducacao.com.br',
            name: 'Joao Victor Caporal Alberti'
          });
          enviados++;
        });
        cache.put(chaveEnvio, 'CONCLUIDO', 600);
      } catch (err) {
        // Nao liberar a chave: uma repeticao automatica duplicaria envios parciais.
        throw new Error('Falha no envio. ' + enviados + ' de ' + emails.length +
          ' grupos de destinatarios concluidos. Confira os enviados antes de tentar novamente. Detalhe: ' + err.message);
      }
      return {sucesso: true, mensagem: 'Enviado com sucesso!', enviados: enviados};
    } finally { lock.releaseLock(); }
  } catch (e) { return {sucesso: false, erro: e.message}; }
}

function obterDadosGraficasEmailSPG() {
  var props = PropertiesService.getScriptProperties();
  var salvo = props.getProperty(EMAIL_SPG_DESTINATARIOS_PROP);
  if (salvo) {
    try {
      var cadastro = JSON.parse(salvo);
      if (Object.keys(cadastro).length || props.getProperty(EMAIL_SPG_CADASTRO_EDITADO_PROP) === '1') return cadastro;
    } catch (e) {}
  }
  // Lista inicial do proprio portal. Alteracoes salvas continuam compartilhadas.
  var internos = ['joao.alberti@arcoeducacao.com.br', 'leonardo.pereira@arcoeducacao.com.br', 'priscillabianchi@arcoeducacao.com.br'];
  return {
    FormaCerta: internos.concat(['alessandra@formacerta.com.br','atendimento3@formacerta.com.br','fabricio@formacerta.com.br','neto@formacerta.com.br','orcamentos3@formacerta.com.br']),
    Logprint: internos.concat(['arnaldo.fortes@logprint.com.br','gilberto.moura@logprint.com.br','henrique.aragao@logprint.com.br','luciano.villaca@logprint.com.br','marcos.bolina@logprint.com.br','mateus.pereira@logprint.com.br','pedidosorcamentos@logprint.com.br','valter.almeida@logprint.com.br']),
    Meta: internos.concat(['andre.agnesini@metabrasil.com.br','andre.teixeira@metabrasil.com.br','ensino@metabrasil.com.br','orcamento@metabrasil.com.br']),
    Posigraf: internos.concat(['gcfurtado@positivo.com.br','hjuliano@positivo.com.br','htelma@positivo.com.br','jamilani@positivo.com.br','jcastro@positivo.com.br','joferreira@positivo.com.br']),
    Reproset: internos.concat(['daniel.esteban@reproset.com.br','igor@reproset.com.br','nedson@reproset.com.br','orcamento@reproset.com.br']),
    Rona: internos.concat(['edicao@ronaeditora.com.br','producao@ronaeditora.com.br','rafael@ronaeditora.com.br','simoney@ronaeditora.com.br']),
    'Casa da Grafica': internos.concat(['frankie@casadagrafica.com','rafaelrodrigues@casadagrafica.com']),
    Star7: internos.concat(['alexandre.gameiro@star-7.com','tiago.pereira@star-7.com']),
    Prints: internos.concat(['felipe@graficaprints.com.br']),
    Teste: ['leonardo.pereira@arcoeducacao.com.br']
  };
}

function salvarDestinatariosEmailSPG(nome, emails) {
  var grafica = String(nome || '').trim();
  if (!grafica || grafica.length > 100) throw new Error('Informe o nome da grafica.');
  var lista = _emailSpgDestinos_([emails])[0].split(', ');
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var atual = obterDadosGraficasEmailSPG();
    atual[grafica] = lista;
    PropertiesService.getScriptProperties().setProperty(
      EMAIL_SPG_DESTINATARIOS_PROP, JSON.stringify(atual));
    return atual;
  } finally { lock.releaseLock(); }
}


// Cadastro unico para a interface de email, usado nos fluxos SPG e antigo.
// Controle de versao impede que dois usuarios sobrescrevam alteracoes simultaneas.
function _emailSpgVersaoCadastro_(catalogo) {
  // Ordem de exibicao nao e uma alteracao de destinatarios.
  var conteudo = Object.keys(catalogo || {}).sort().map(function(nome) {
    return [nome, catalogo[nome].map(function(email) { return String(email).trim().toLowerCase(); }).sort()];
  });
  return _emailSpgHex_(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, JSON.stringify(conteudo)));
}

function obterCadastroEmailsSPG() {
  var graficas = obterDadosGraficasEmailSPG();
  return {graficas: graficas, versao: _emailSpgVersaoCadastro_(graficas),
    vazioIntencional: !Object.keys(graficas).length &&
      PropertiesService.getScriptProperties().getProperty(EMAIL_SPG_CADASTRO_EDITADO_PROP) === '1'};
}

function _emailSpgValidarCadastro_(proposto) {
  if (!proposto || typeof proposto !== 'object' || Array.isArray(proposto)) {
    throw new Error('O cadastro de destinatarios deve conter graficas e emails.');
  }
  var nomes = Object.keys(proposto);
  if (nomes.length > 100) throw new Error('Limite de graficas excedido.');
  var saida = Object.create(null);
  var graficaVistas = Object.create(null);
  nomes.forEach(function(original) {
    var nome = String(original).trim();
    if (!nome || nome.length > 100 || /[\x00-\x1f]/.test(nome) ||
        nome === '__proto__' || nome === 'constructor' || nome === 'prototype') {
      throw new Error('Nome de grafica invalido: ' + original);
    }
    var chave = nome.toLowerCase();
    if (graficaVistas[chave]) throw new Error('Grafica duplicada: ' + nome);
    graficaVistas[chave] = true;
    var entradas = proposto[original];
    if (!Array.isArray(entradas) || !entradas.length || entradas.length > 80) {
      throw new Error('A grafica ' + nome + ' deve possuir de 1 a 80 emails.');
    }
    // Cada elemento e um email individual, nao uma string de lista.
    var vistos = Object.create(null);
    var emails = [];
    entradas.forEach(function(valor) {
      var email = String(valor == null ? '' : valor).trim().toLowerCase();
      if (email.indexOf(';') !== -1 || email.indexOf(',') !== -1 ||
          !/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(email)) {
        throw new Error('Email invalido da grafica ' + nome + '.');
      }
      if (!vistos[email]) {vistos[email] = true;emails.push(email);}
    });
    saida[nome] = emails;
  });
  var texto = JSON.stringify(saida);
  // ScriptProperties limita o tamanho de cada propriedade (9 KB).
  if (Utilities.newBlob(texto, 'text/plain').getBytes().length > 8000) {
    throw new Error('Cadastro muito grande para salvar. Reduza registros antigos.');
  }
  return saida;
}

function salvarCadastroEmailsSPG(proposto, versaoEsperada) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Outro usuario esta alterando os emails. Tente novamente.');
  try {
    var atual = obterDadosGraficasEmailSPG();
    if (!versaoEsperada || _emailSpgVersaoCadastro_(atual) !== versaoEsperada) {
      throw new Error('O cadastro foi alterado por outra pessoa. Atualize a página para carregar a lista atual antes de salvar.');
    }
    var novo = _emailSpgValidarCadastro_(proposto);
    PropertiesService.getScriptProperties().setProperty(
      EMAIL_SPG_DESTINATARIOS_PROP, JSON.stringify(novo));
    PropertiesService.getScriptProperties().setProperty(EMAIL_SPG_CADASTRO_EDITADO_PROP, '1');
    return {graficas: novo, versao: _emailSpgVersaoCadastro_(novo)};
  } finally {lock.releaseLock();}
}
