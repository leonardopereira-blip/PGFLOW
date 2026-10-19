// Envio de formalizacao via SPG, independente das planilhas operacionais.
// O fluxo historico continua isolado em Menu_Emails.js durante a transicao.
var EMAIL_SPG_BASE = 'https://spg-api.prod.arcocv.co';
var EMAIL_SPG_DESTINATARIOS_PROP = 'EMAIL_SPG_DESTINATARIOS_GRAFICAS';

function _emailSpgId_(id) {
  var valor = String(id == null ? '' : id).trim();
  if (!/^\d{1,12}$/.test(valor)) throw new Error('Informe um numero de chamado valido.');
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
  // O endpoint pode reconstruir o ZIP com timestamps diferentes a cada GET.
  // Comparamos o conteudo das planilhas, nao os metadados do arquivo ZIP.
  var partes = Utilities.unzip(blob).filter(function(p) {
    return /^xl\/worksheets\/[^\/]+\.xml$/i.test(p.getName());
  }).sort(function(a, b) { return a.getName().localeCompare(b.getName()); })
    .map(function(p) { return p.getName() + ':' + p.getDataAsString('UTF-8'); });
  if (!partes.length) throw new Error('Excel sem planilhas para validar.');
  return _emailSpgHex_(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, partes.join('\n')));
}

function _emailSpgXmlAtributo_(trecho, nome) {
  var re = new RegExp('(?:^|\\s)' + nome.replace(':', '\\:') + '\\s*=\\s*["\\x27]([^"\\x27]*)["\\x27]', 'i');
  var m = re.exec(trecho);
  return m ? m[1] : '';
}

function _emailSpgChecarXlsx_(blob) {
  var bytes = blob.getBytes();
  if (bytes.length < 100 || bytes.length > 20 * 1024 * 1024 ||
      ((bytes[0] + 256) % 256) !== 80 || ((bytes[1] + 256) % 256) !== 75) {
    throw new Error('O SPG nao retornou um arquivo Excel valido.');
  }
  var entradas;
  try { entradas = Utilities.unzip(blob); }
  catch (e) { throw new Error('Nao foi possivel abrir o Excel retornado pelo SPG.'); }
  var arquivos = {};
  entradas.forEach(function(b) { arquivos[b.getName().replace(/^\/+/, '')] = b; });
  if (!arquivos['xl/workbook.xml'] || !arquivos['xl/_rels/workbook.xml.rels']) {
    throw new Error('Excel do SPG sem estrutura de planilha reconhecida.');
  }
  var livro = arquivos['xl/workbook.xml'].getDataAsString('UTF-8');
  var rels = arquivos['xl/_rels/workbook.xml.rels'].getDataAsString('UTF-8');
  var caminhoPorId = {};
  var relTag = /<Relationship\b[^>]*\/?>/gi, achou;
  while ((achou = relTag.exec(rels))) {
    var rid = _emailSpgXmlAtributo_(achou[0], 'Id');
    var alvo = _emailSpgXmlAtributo_(achou[0], 'Target').replace(/\\/g, '/');
    if (alvo.indexOf('../') !== -1 || alvo.indexOf('://') !== -1) continue;
    caminhoPorId[rid] = alvo.charAt(0) === '/' ? alvo.substring(1) :
      (alvo.indexOf('xl/') === 0 ? alvo : 'xl/' + alvo);
  }
  var presentes = {tiragem: false, arvore: false, especificacao: false};
  var sheetTag = /<sheet\b[^>]*\/?>/gi;
  while ((achou = sheetTag.exec(livro))) {
    var nome = _emailSpgNormalizar_(_emailSpgXmlAtributo_(achou[0], 'name'));
    var tipo = nome.indexOf('tiragem') >= 0 ? 'tiragem' :
      nome.indexOf('arvore') >= 0 ? 'arvore' :
      (nome.indexOf('espec') >= 0 ? 'especificacao' : '');
    if (!tipo) continue;
    var ridSheet = _emailSpgXmlAtributo_(achou[0], 'r:id');
    var alvoSheet = caminhoPorId[ridSheet];
    if (!alvoSheet || !arquivos[alvoSheet]) continue;
    var xml = arquivos[alvoSheet].getDataAsString('UTF-8');
    // A aba precisa conter ao menos uma linha de dados alem do cabecalho.
    var linhas = xml.match(/<row\b[^>]*>[\s\S]*?<\/row>/gi) || [];
    if (linhas.length > 1 && linhas.slice(1).some(function(l) {
      return /<(?:v|t)\b[^>]*>[^<]+<\/(?:v|t)>/i.test(l);
    })) presentes[tipo] = true;
  }
  var faltando = Object.keys(presentes).filter(function(k) { return !presentes[k]; });
  if (faltando.length) {
    throw new Error('Anexo do SPG sem dados nas abas: ' + faltando.join(', ') + '.');
  }
  return true;
}

function _emailSpgBlobDeResposta_(resposta, id) {
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
  _emailSpgChecarXlsx_(blob);
  return blob;
}

function _emailSpgPacote_(idChamado) {
  var id = _emailSpgId_(idChamado), token = _spgLogin_();
  if (!token) throw new Error('Autenticacao do SPG indisponivel.');
  var respostaDados = _emailSpgApi_('/api/reentradas/' + encodeURIComponent(id), token);
  var json = _emailSpgJson_(respostaDados);
  var responseOp = _emailSpgApi_('/api/reentradas/' + encodeURIComponent(id) + '/op', token);
  var blob = _emailSpgBlobDeResposta_(responseOp, id);
  var numeroApi = json ? _emailSpgCampo_(json, [
    'numeroChamado', 'numeroDoChamado', 'idChamado', 'chamadoNumero']) : '';
  if (numeroApi && numeroApi !== id) {
    throw new Error('O numero do chamado devolvido pelo SPG nao corresponde ao solicitado.');
  }
  var data = json ? _emailSpgCampo_(json, [
    'dataDesejada', 'dataDesejadaEntrega', 'dataDesejadaDeEntrega',
    'dataEntregaDesejada', 'dataDesejadaCD']) : '';
  var descricao = json ? _emailSpgDescricao_(json) : '';
  return {
    id: id, dataDesejada: _emailSpgData_(data), descricao: descricao,
    hash: _emailSpgHash_(blob), blob: blob
  };
}

function consultarPacoteEmailSPG(idChamado, opcaoVersao, tipoEmail) {
  if (tipoEmail === 'inclusivo') {
    // Rota especifica de inclusivos ainda nao homologada.
    // Nunca consultar /reentradas para obter um anexo inclusivo.
    return {status: 'incompleto', anexoValido: false,
      erro: 'O anexo de inclusivos deve ser obtido pelo fluxo antigo ate a rota especifica do SPG ser validada.'};
  }
  if (tipoEmail && tipoEmail !== 'reentrada') {
    return {status: 'incompleto', anexoValido: false,
      erro: 'Tipo de formalizacao nao reconhecido.'};
  }
  var id;
  try {
    id = _emailSpgId_(idChamado);
    // Historico por versao somente quando o SPG expuser selecao inequívoca.
    if (opcaoVersao === 'manual') return {
      status: 'incompleto', anexoValido: false,
      erro: 'O SPG fornece a PPG atual. Para selecionar uma versao antiga, utilize o fluxo antigo.'
    };
    var p = _emailSpgPacote_(id);
    var faltantes = [];
    if (!p.dataDesejada) faltantes.push('Data desejada de entrega');
    if (!p.descricao) faltantes.push('Descricao da producao');
    return {status: faltantes.length ? 'incompleto' : 'ok',
      idChamado: p.id, anexoValido: true, anexoHash: p.hash,
      dataDesejada: p.dataDesejada, dataColeta: _emailSpgColeta_(p.dataDesejada),
      descricao: p.descricao, faltantes: faltantes};
  } catch (e) {
    return {status: 'incompleto', idChamado: id || '',
      anexoValido: false, erro: e.message, faltantes: []};
  }
}

function _emailSpgDestinos_(raw) {
  if (!Array.isArray(raw)) throw new Error('Destinatarios invalidos.');
  var emails = [], vistos = {};
  raw.forEach(function(grupo) {
    String(grupo || '').split(/[;,]/).forEach(function(email) {
      var e = email.trim().toLowerCase();
      if (!e) return;
      if (!/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(e)) {
        throw new Error('E-mail de destinatario invalido: ' + e);
      }
      if (!vistos[e]) { vistos[e] = true; emails.push(e); }
    });
  });
  if (!emails.length || emails.length > 80) throw new Error('Informe de 1 a 80 destinatarios.');
  return emails;
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
    if (dados.tipoEmail === 'inclusivo' ||
        (dados.tipoEmail && dados.tipoEmail !== 'reentrada')) {
      throw new Error('Esta rota de PPG atende somente reentradas. Inclusivos devem usar o fluxo antigo.');
    }
    var id = _emailSpgId_(dados.idChamado), emails = _emailSpgDestinos_(dados.arrayDestinos);
    var assunto = String(dados.assunto || '').trim(), corpo = String(dados.corpo || '').trim();
    if (!assunto || !corpo) throw new Error('Assunto e corpo sao obrigatorios.');
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) throw new Error('Ha outro envio em andamento. Tente novamente.');
    try {
      // Refaz consulta e validacao: nunca confia no ID de arquivo ou origem informados pelo navegador.
      var p = _emailSpgPacote_(id);
      if (p.hash !== String(dados.anexoHash || '')) {
        throw new Error('A PPG foi alterada no SPG. Retorne para a revisao antes de enviar.');
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
          ' destinatarios concluidos. Confira os enviados antes de tentar novamente. Detalhe: ' + err.message);
      }
      return {sucesso: true, mensagem: 'Enviado com sucesso!', enviados: enviados};
    } finally { lock.releaseLock(); }
  } catch (e) { return {sucesso: false, erro: e.message}; }
}

function obterDadosGraficasEmailSPG() {
  var props = PropertiesService.getScriptProperties();
  var salvo = props.getProperty(EMAIL_SPG_DESTINATARIOS_PROP);
  if (salvo) {
    try { return JSON.parse(salvo); } catch (e) {}
  }
  // Migracao unica do cadastro de destinatarios, sem leitura recorrente do Sheets.
  var antigos = {};
  try { antigos = obterDadosGraficas(); } catch (e) { antigos = {}; }
  // Web Apps podem nao ter planilha ativa; migrar pelo ID ja utilizado no PGFLOW.
  if (!Object.keys(antigos).length) {
    try {
      var sheet = SpreadsheetApp.openById(ID_PLANILHA_LOG).getSheetByName('Emails_Envio');
      if (sheet) {
        var values = sheet.getDataRange().getDisplayValues();
        for (var i = 1; i < values.length; i++) {
          var nome = String(values[i][0] || '').trim();
          var email = String(values[i][1] || '').trim();
          if (!nome || !email) continue;
          if (!antigos[nome]) antigos[nome] = [];
          antigos[nome].push(email);
        }
      }
    } catch (e) { /* Cadastro pode ser criado pela tela. */ }
  }
  props.setProperty(EMAIL_SPG_DESTINATARIOS_PROP, JSON.stringify(antigos));
  return antigos;
}

function salvarDestinatariosEmailSPG(nome, emails) {
  var grafica = String(nome || '').trim();
  if (!grafica || grafica.length > 100) throw new Error('Informe o nome da grafica.');
  var lista = _emailSpgDestinos_([emails]);
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
  return _emailSpgHex_(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, JSON.stringify(catalogo || {})));
}

function obterCadastroEmailsSPG() {
  var graficas = obterDadosGraficasEmailSPG();
  return {graficas: graficas, versao: _emailSpgVersaoCadastro_(graficas)};
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
      throw new Error('O cadastro foi alterado por outra pessoa. Feche e abra a lista para atualizar antes de salvar.');
    }
    var novo = _emailSpgValidarCadastro_(proposto);
    PropertiesService.getScriptProperties().setProperty(
      EMAIL_SPG_DESTINATARIOS_PROP, JSON.stringify(novo));
    return {graficas: novo, versao: _emailSpgVersaoCadastro_(novo)};
  } finally {lock.releaseLock();}
}
