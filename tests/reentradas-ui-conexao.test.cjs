// Exercise the real read-only RPC callbacks without a browser or Google services.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const html = fs.readFileSync(path.join(__dirname, '..', 'ValidaInterface.html'), 'utf8');
const source = html.match(/<script>([\s\S]*?)<\/script>/)[1];
const plain = value => JSON.parse(JSON.stringify(value));
const network = () => new Error('NetworkError: Falha na conexão devido a HTTP 0');
const valid = id => ({ chamado: id, origem: 'spg', totalSkus: 1, tiragemTotal: 250,
  marcas: 'SAS', series: '', papeisCapa: '', papeisMiolo: '', cdsDestino: 'FOR' });

test('resultado somente tecnico mostra De-Para sem link de calculadora inexistente',()=>{
  const s=screen();
  s.context.renderResumo([{chamado:'2932',grafica:'Reproset',skus:0,tiragem:0,
    depara:'https://drive.google.com/file/d/tecnico/view',calculadora:'',somenteEspecificacao:true}],'resumoGerado');
  const html=s.el('resumoGerado').innerHTML;
  assert.match(html,/Abrir De-Para/);
  assert.match(html,/Somente especificação técnica: nenhum SKU calculável/);
  assert.doesNotMatch(html,/Abrir calculadora|undefined/);
});

function screen() {
  const elements = new Map();
  function el(id) {
    if (!elements.has(id)) {
      const classes = new Set();
      elements.set(id, { value: '', disabled: false, innerHTML: '', textContent: '', style: {},
        classList: { add: name => classes.add(name), remove: name => classes.delete(name), contains: name => classes.has(name) } });
    }
    return elements.get(id);
  }
  const document = { getElementById: el, querySelector: selector => el(selector), querySelectorAll: () => [] };
  const requests = [], timers = [];
  const google = { script: {} };
  Object.defineProperty(google.script, 'run', { get() {
    const chain = { withSuccessHandler(fn) { this.success = fn; return this; }, withFailureHandler(fn) { this.failure = fn; return this; } };
    for (const method of ['listarPastasGraficasReentradasSPG', 'listarUltimasReentradasSPG', 'buscarResumoReentradasSPG', 'validarReentradasSPG', 'gerarValidacaoReentradasSPG']) {
      chain[method] = (...args) => requests.push({ method, args: plain(args), success: chain.success, failure: chain.failure });
    }
    return chain;
  } });
  const context = { document, google, setTimeout: (fn, delay) => timers.push({ fn, delay }), clearTimeout() {}, clearInterval() {}, setInterval: () => 1 };
  vm.createContext(context); vm.runInContext(source, context);
  el('txtChamados').value = '2632'; el('selectGraficas').value = 'REPROSET';
  el('telaFase3').style.display = 'none';
  context.verificarLiberacaoBotao();
  const next = () => el('#botoesResumo .btn-continuar');
  function runRetry() {
    const index = timers.findIndex(timer => timer.delay === 1500);
    assert(index !== -1, 'Expected one retry after 1500 ms');
    timers.splice(index, 1)[0].fn();
  }
  function request(index = requests.length - 1) { return requests[index]; }
  return { context, requests, timers, el, next, runRetry, request };
}

test('resumo repete HTTP 0 uma vez e conserva argumentos e aceites por chamado/SKU', () => {
  const s = screen();
  s.context.origensLegadoReentradas = ['2632'];
  s.context.especificacoesGeraisReentradas = [{ chamado: '2632', sku: 'SKU-1' }];
  s.context.abrirResumoAntesDeAvancar(); const first = s.request();
  first.failure(network());
  assert.equal(s.requests.length, 1);
  assert.match(s.el('conteudoResumoLoader').textContent, /Consultando novamente/);
  assert.equal(s.el('txtChamados').disabled, true);
  s.runRetry();
  assert.equal(s.requests.length, 2);
  assert.deepEqual(s.request().args, ['2632', ['2632'], [{ chamado: '2632', sku: 'SKU-1' }]]);
  s.request().success([valid('2632')]);
  assert.equal(s.context._resumoProntoReentradas, true);
  assert.equal(s.next().style.display, '');
  assert.match(s.el('conteudoResumoData').innerHTML, /Chamado 2632/);
  assert.equal(s.timers.filter(timer => timer.delay === 1500).length, 0);
});

test('segunda falha de transporte libera campos e oferece nova consulta sem avançar', () => {
  const s = screen(); s.context.abrirResumoAntesDeAvancar();
  s.request().failure(network()); s.runRetry(); s.request().failure(network());
  assert.equal(s.requests.length, 2);
  assert.equal(s.el('txtChamados').disabled, false);
  assert.equal(s.el('selectGraficas').disabled, false);
  assert.equal(s.el('btnAvancar').disabled, false);
  assert.equal(s.context._resumoProntoReentradas, false);
  assert.equal(s.next().style.display, 'none');
  assert.equal(s.el('btnTentarResumo').style.display, '');
  assert.match(s.el('conteudoResumoLoader').innerHTML, /A comunicação com o portal foi interrompida\. Tente consultar novamente\./);
  s.context.continuarPosResumo();
  assert.equal(s.requests.length, 2);
  assert.equal(s.timers.filter(timer => timer.delay === 1500).length, 0);
  s.context.abrirResumoAntesDeAvancar();
  assert.equal(s.requests.length, 3);
  assert.equal(s.el('btnTentarResumo').style.display, 'none');
});

test('Voltar durante a espera cancela retry e ignora respostas da consulta abandonada', () => {
  const s = screen(); s.context.abrirResumoAntesDeAvancar(); const old = s.request();
  old.failure(network()); s.context.cancelarResumo(); s.runRetry();
  assert.equal(s.requests.length, 1);
  old.success([valid('2632')]); old.failure(network());
  assert.equal(s.context._resumoProntoReentradas, false);
  assert.equal(s.el('painelResumo').classList.contains('active'), false);
  assert.equal(s.el('txtChamados').disabled, false);
});

test('alterar chamado durante espera impede nova consulta e respostas tardias de trocar dados', () => {
  const s = screen(); s.context.abrirResumoAntesDeAvancar(); const old = s.request();
  old.failure(network());
  s.el('txtChamados').value = '3001'; s.context.verificarLiberacaoBotao(); s.runRetry();
  assert.equal(s.requests.length, 1);
  s.context.abrirResumoAntesDeAvancar(); s.request().success([valid('3001')]);
  old.success([valid('2632')]); old.failure(network());
  assert.match(s.el('conteudoResumoData').innerHTML, /Chamado 3001/);
  assert.doesNotMatch(s.el('conteudoResumoData').innerHTML, /Chamado 2632/);
});

test('callbacks duplicados de tentativa anterior não alteram retry nem resultado vigente', () => {
  const s = screen(); s.context.abrirResumoAntesDeAvancar(); const first = s.request();
  first.failure(network()); first.failure(network());
  assert.equal(s.timers.filter(timer => timer.delay === 1500).length, 1);
  first.success([valid('INDEVIDO')]);
  assert.equal(s.context._resumoProntoReentradas, false);
  s.runRetry(); const second = s.request();
  first.success([valid('INDEVIDO')]); first.failure(network());
  second.success([valid('2632')]); second.failure(network());
  assert.equal(s.context._resumoProntoReentradas, true);
  assert.doesNotMatch(s.el('conteudoResumoData').innerHTML, /INDEVIDO/);
  assert.equal(s.requests.length, 2);
  assert.equal(s.timers.filter(timer => timer.delay === 1500).length, 0);
});

test('erro de negócio retornado com sucesso não dispara retry nem oferece legado indevido', () => {
  const s = screen(); s.context.abrirResumoAntesDeAvancar();
  s.request().success([{ chamado: '2632', erro: 'HTTP 0 informado pelo serviço', totalSkus: 0, tiragemTotal: 0 }]);
  assert.equal(s.context._resumoProntoReentradas, false);
  assert.equal(s.requests.length, 1);
  assert.equal(s.timers.filter(timer => timer.delay === 1500).length, 0);
  assert.doesNotMatch(s.el('conteudoResumoData').innerHTML, /versão antiga de solicitação/);
});

test('somente HTTP 0 e NetworkError são repetidos; erros de API e permissão falham diretamente', () => {
  for (const message of ['HTTP 401', 'HTTP 403', 'HTTP 500', 'HTTP 00', 'TimeoutError', 'Falha na conexão com a API']) {
    const s = screen(); s.context.abrirResumoAntesDeAvancar();
    s.request().failure(new Error(message));
    assert.equal(s.requests.length, 1, message);
    assert.equal(s.timers.filter(timer => timer.delay === 1500).length, 0, message);
    assert.match(s.el('conteudoResumoLoader').innerHTML, /Erro ao consultar o SPG/);
    assert.equal(s.el('txtChamados').disabled, false);
  }
});

function startValidation(s) {
  s.context.abrirResumoAntesDeAvancar(); s.request().success([valid('2632')]);
  s.context.continuarPosResumo();
  assert.equal(s.request().method, 'validarReentradasSPG');
}

test('validação repete leitura uma vez com os mesmos argumentos sem chamar geração', () => {
  const s = screen(); startValidation(s);
  const first = s.request(); first.failure(new Error('NetworkError'));
  assert.match(s.el('mensagemFase1').innerHTML, /Validando novamente/);
  s.runRetry();
  assert.deepEqual(s.request().args, first.args);
  let result;
  s.context.processarResultadoFase1App = value => { result = value; };
  const response = { validos: ['2632'], invalidos: [], colunasManuais: [] };
  first.success({ validos: ['VELHO'] }); s.request().success(response);
  assert.equal(result, response);
  assert.equal(s.requests.some(req => req.method === 'gerarValidacaoReentradasSPG'), false);
  assert.equal(s.el('telaFase3').style.display, 'none');
});

test('falha final da validação fica na primeira fase e permite reconsultar preservando aceites', () => {
  const s = screen(); startValidation(s);
  s.request().failure(network()); s.runRetry(); s.request().failure(network());
  assert.equal(s.context._resumoProntoReentradas, false);
  assert.equal(s.el('telaFase3').style.display, 'none');
  assert.match(s.el('mensagemFase1').innerHTML, /A comunicação com o portal foi interrompida/);
  assert.equal(s.el('btnTentarResumo').style.display, '');
  assert.equal(s.el('txtChamados').disabled, false);
  assert.equal(s.el('selectGraficas').disabled, false);
  assert.equal(s.requests.some(req => req.method === 'gerarValidacaoReentradasSPG'), false);
  s.context.abrirResumoAntesDeAvancar();
  assert.equal(s.request().method, 'buscarResumoReentradasSPG');
});

test('cancelamento invalida retry de validação e sua resposta tardia', () => {
  const s = screen(); startValidation(s); const old = s.request();
  old.failure(network()); s.context.cancelarResumo(); s.runRetry();
  assert.equal(s.requests.length, 2);
  old.success({ validos: ['2632'], invalidos: [] });
  old.failure(network());
  assert.equal(s.el('telaFase2').style.display, undefined);
  assert.equal(s.el('telaFase3').style.display, 'none');
  assert.equal(s.el('painelResumo').classList.contains('active'), false);
});

test('validação com resposta funcional erro não repete nem abre tela de geração', () => {
  const s = screen(); startValidation(s);
  s.request().success({ erro: 'Nenhum chamado com SKUs disponíveis para gerar.' });
  assert.equal(s.timers.filter(timer => timer.delay === 1500).length, 0);
  assert.match(s.el('mensagemFase1').innerHTML, /Nenhum chamado com SKUs disponíveis/);
  assert.equal(s.el('telaFase3').style.display, 'none');
  assert.equal(s.el('txtChamados').disabled, false);
  assert.equal(s.el('selectGraficas').disabled, false);
  assert.equal(s.requests.length, 2);
});

test('geração continua sem repetição automática após HTTP 0', () => {
  const s = screen();
  s.context.iniciarTracker = () => {};
  let error;
  s.context.mostrarErroCustos = err => { error = err; };
  s.context.executarFase2();
  assert.equal(s.request().method, 'gerarValidacaoReentradasSPG');
  s.request().failure(network());
  assert.match(error, /HTTP 0/);
  assert.equal(s.requests.length, 1);
  assert.equal(s.timers.filter(timer => timer.delay === 1500).length, 0);
});
