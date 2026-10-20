// DOM and Apps Script are simulated. No browser, API, Drive or Sheets is used.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const html = fs.readFileSync(path.join(__dirname, '..', 'ValidaInterface.html'), 'utf8');
const source = html.match(/<script>([\s\S]*?)<\/script>/)[1];

function makeScreen(ids = '3100, 2567') {
  function matches(el, selector) {
    if (selector === '*') return true;
    const tag = selector.match(/^[\w-]+/);
    if (tag && el.tagName !== tag[0]) return false;
    const id = selector.match(/#([\w-]+)/);
    if (id && el.id !== id[1]) return false;
    return Array.from(selector.matchAll(/\.([\w-]+)/g)).every(([, cls]) => el.className.split(/\s+/).includes(cls));
  }
  function fits(el, selector) {
    const parts = selector.trim().split(/\s+/);
    if (!matches(el, parts.pop())) return false;
    let ancestor = el.parent;
    while (parts.length) {
      const part = parts.pop();
      while (ancestor && !matches(ancestor, part)) ancestor = ancestor.parent;
      if (!ancestor) return false;
      ancestor = ancestor.parent;
    }
    return true;
  }
  const decode = value => String(value).replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
  class Element {
    constructor(tag = '', attrs = {}) {
      this.tagName = tag; this.attrs = attrs; this.children = []; this.parent = null;
      this.id = attrs.id || ''; this.className = attrs.class || ''; this._value = attrs.value;
      this._text = ''; this._html = ''; this.disabled = 'disabled' in attrs; this.style = {};
      for (const part of (attrs.style || '').split(';')) {
        const at = part.indexOf(':');
        if (at !== -1) this.style[part.slice(0, at).trim()] = part.slice(at + 1).trim();
      }
      this.classList = {
        contains: name => this.className.split(/\s+/).includes(name),
        add: name => { if (!this.classList.contains(name)) this.className += ' ' + name; },
        remove: name => { this.className = this.className.split(/\s+/).filter(x => x !== name).join(' '); },
        toggle: name => this.classList.contains(name) ? this.classList.remove(name) : this.classList.add(name)
      };
    }
    appendChild(child) { child.parent = this; this.children.push(child); return child; }
    get value() {
      if (this._value !== undefined) return this._value;
      if (this.tagName === 'select') {
        const options = this.children.filter(child => child.tagName === 'option');
        const selected = options.find(option => 'selected' in option.attrs) || options[0];
        return selected ? selected.value : '';
      }
      return this.tagName === 'option' ? this.textContent : '';
    }
    set value(value) { this._value = String(value); }
    get textContent() { return this._text + this.children.map(child => child.textContent).join(''); }
    set textContent(value) { this._text = String(value); this.children = []; }
    get innerHTML() { return this._html; }
    set innerHTML(value) {
      this._html = String(value); this._text = ''; this.children = [];
      if (this.tagName === 'select') this._value = undefined;
      parse(this._html, this);
    }
    getAttribute(key) { return this.attrs[key] === undefined ? null : this.attrs[key]; }
    querySelectorAll(selector) {
      const all = [];
      function walk(el) { el.children.forEach(child => { all.push(child); walk(child); }); }
      walk(this);
      return all.filter(el => selector.split(',').some(part => fits(el, part)));
    }
    querySelector(selector) { return this.querySelectorAll(selector)[0] || null; }
    insertAdjacentHTML(_position, value) { parse(value, this); }
  }
  function parse(text, parent) {
    const stack = [parent];
    const voidTags = new Set(['input', 'meta', 'base', 'br', 'hr', 'link', 'img', 'path']);
    for (const token of text.matchAll(/<\/?([a-z][\w-]*)\b([^>]*)>|([^<]+)/gi)) {
      if (!token[1]) { stack.at(-1)._text += decode(token[3]); continue; }
      const tag = token[1].toLowerCase();
      if (token[0].startsWith('</')) {
        const at = stack.findLastIndex(el => el.tagName === tag);
        if (at > 0) stack.length = at;
        continue;
      }
      const attrs = {};
      for (const [, key, double, single] of token[2].matchAll(/([\w-]+)(?:="([^"]*)"|='([^']*)')?/g)) attrs[key] = decode(double ?? single ?? '');
      const el = new Element(tag, attrs); stack.at(-1).appendChild(el);
      if (!voidTags.has(tag) && !token[0].endsWith('/>')) stack.push(el);
    }
  }
  const document = new Element('document');
  document.createElement = tag => new Element(tag);
  document.getElementById = id => document.querySelectorAll('*').find(el => el.id === id) || null;
  parse(html.replace(/<script>[\s\S]*?<\/script>/g, '').replace(/<style>[\s\S]*?<\/style>/g, ''), document);
  const requests = [];
  const google = { script: {} };
  Object.defineProperty(google.script, 'run', { get() {
    const chain = {
      withSuccessHandler(fn) { this.success = fn; return this; },
      withFailureHandler(fn) { this.failure = fn; return this; }
    };
    ['listarPastasGraficasReentradasSPG', 'listarUltimasReentradasSPG', 'buscarResumoReentradasSPG', 'validarReentradasSPG', 'gerarValidacaoReentradasSPG'].forEach(method => {
      chain[method] = (...args) => requests.push({ method, args, success: chain.success, failure: chain.failure });
    });
    return chain;
  }});
  const timers = [];
  const context = { document, google, setTimeout: fn => timers.push(fn), clearTimeout() {}, setInterval: () => 1, clearInterval() {} };
  vm.createContext(context); vm.runInContext(source, context);
  const el = id => { const result = document.getElementById(id); assert(result, 'Missing element ' + id); return result; };
  function pending(method) { const result = requests.find(req => req.method === method && !req.done); assert(result, 'Expected pending ' + method); return result; }
  function reply(method, result, error = false) { const req = pending(method); req.done = true; (error ? req.failure : req.success)(result); return req; }
  function dispatch(element, event = 'click') {
    if (typeof element === 'string') element = el(element);
    if (element.disabled) return;
    const code = element.getAttribute('on' + event);
    if (code) vm.runInContext(code, context);
  }
  const next = () => document.querySelector('#botoesResumo .btn-continuar');
  function button(text) { const result = document.querySelectorAll('button').find(btn => btn.textContent.trim() === text); assert(result, 'Missing button ' + text); return result; }
  context.preencherPastasGraficasApp([{ id: 'META', nome: 'Meta' }, { id: 'LOGPRINT', nome: 'Logprint' }]);
  el('selectGraficas').value = 'META'; el('txtChamados').value = ids; dispatch('txtChamados', 'input');
  return { context, document, el, requests, pending, reply, dispatch, next, button,
    runTimers() { timers.splice(0).forEach(fn => fn()); } };
}
const ok = (chamado, origem = 'spg') => ({ chamado, origem, totalSkus: 2, tiragemTotal: 300,
  marcas: 'SAS', series: '3ª Série', papeisCapa: 'Cartão', papeisMiolo: 'Offset', cdsDestino: 'FOR',
  ...(origem === 'legado' ? { legadoTemporario: true } : {}) });
const missing = chamado => ({ chamado, origem: 'spg', naoEncontrado: true, erro: 'Chamado inexistente.', totalSkus: 0, tiragemTotal: 0 });
const plain = value => JSON.parse(JSON.stringify(value));

test('oferta de legado é por chamado ausente e exige confirmação antes de consultar alocação', () => {
  const s = makeScreen(); s.dispatch('btnAvancar');
  assert.deepEqual(plain(s.pending('buscarResumoReentradasSPG').args), ['3100, 2567', []]);
  s.reply('buscarResumoReentradasSPG', [ok('3100'), missing('2567')]);
  assert.match(s.el('conteudoResumoData').textContent, /Esse chamado está na versão antiga de solicitação\?/);
  assert.match(s.el('conteudoResumoData').textContent, /somente as especificações, a tiragem e a árvore/);
  assert.equal(s.next().style.display, 'none');
  s.context.continuarPosResumo();
  assert.equal(s.requests.some(r => r.method === 'validarReentradasSPG'), false);
  s.dispatch(s.button('Sim, usar legado temporário'));
  assert.deepEqual(plain(s.pending('buscarResumoReentradasSPG').args), ['3100, 2567', ['2567']]);
});

test('escolher não mantém SPG, bloqueia avanço e não consulta legado', () => {
  const s = makeScreen('2567'); s.dispatch('btnAvancar');
  s.reply('buscarResumoReentradasSPG', [missing('2567')]);
  s.dispatch(s.button('Não, continuar no SPG'));
  assert.match(s.el('conteudoResumoData').textContent, /Mantido no SPG/);
  assert.deepEqual(plain(s.context.origensLegadoReentradas), []);
  assert.equal(s.requests.length, 1);
  assert.equal(s.next().style.display, 'none');
  s.context.continuarPosResumo();
  assert.equal(s.requests.length, 1);
});

test('cada chamado antigo na mesma lista exige a sua própria confirmação', () => {
  const s = makeScreen('2567, 2448'); s.dispatch('btnAvancar');
  s.reply('buscarResumoReentradasSPG', [missing('2567'), missing('2448')]);
  s.context.selecionarLegadoTemporarioReentradas(0, true);
  assert.deepEqual(plain(s.pending('buscarResumoReentradasSPG').args[1]), ['2567']);
  s.reply('buscarResumoReentradasSPG', [ok('2567', 'legado'), missing('2448')]);
  assert.equal(s.next().style.display, 'none');
  s.context.selecionarLegadoTemporarioReentradas(1, true);
  assert.deepEqual(plain(s.pending('buscarResumoReentradasSPG').args[1]), ['2567', '2448']);
  s.reply('buscarResumoReentradasSPG', [ok('2567', 'legado'), ok('2448', 'legado')]);
  assert.equal(s.next().style.display, '');
});

test('mistura SPG e legado preserva gráfica, campos manuais e fornecimento até geração nativa', () => {
  const s = makeScreen(); s.dispatch('btnAvancar');
  s.reply('buscarResumoReentradasSPG', [ok('3100'), missing('2567')]);
  s.dispatch(s.button('Sim, usar legado temporário'));
  s.reply('buscarResumoReentradasSPG', [ok('3100'), ok('2567', 'legado')]);
  assert.match(s.el('conteudoResumoData').textContent, /Legado temporário/);
  assert.equal(s.next().style.display, ''); s.dispatch(s.next());
  assert.deepEqual(plain(s.pending('validarReentradasSPG').args), ['META', '3100, 2567', ['2567']]);
  s.reply('validarReentradasSPG', { validos: ['3100', '2567'], invalidos: [],
    colunasManuais: [{ index: 51, titulo: 'SHRINK', opcoes: ['Não', 'Sim'] }] });
  assert.equal(s.el('telaFase2').style.display, 'block');
  assert.match(s.el('painelChips').textContent, /2567 · Legado temporário/);
  const manual = s.document.querySelector('.inp-manual'); assert(manual); manual.value = 'Sim';
  s.el('selTecnologia').value = 'PLANA'; s.el('selFornecedor').value = 'ATHOS';
  s.dispatch('btnConcluir');
  const args = plain(s.pending('gerarValidacaoReentradasSPG').args);
  assert.equal(args[0], 'META'); assert.deepEqual(args[1], ['3100', '2567']);
  assert.deepEqual(args[4], [{ index: '51', valor: 'Sim' }]);
  assert.equal(args[5].tecnologia, 'PLANA'); assert.equal(args[5].fornecedor, 'ATHOS');
  assert.equal(args[5].elementos.DIGITAL.capa, 'GRÁFICA'); assert.deepEqual(args[6], ['2567']);
  assert.match(s.el('step1').textContent, /origens de dados confirmadas/);
  s.reply('gerarValidacaoReentradasSPG', { sucesso: true, arquivos: ['file'], resumo: [] });
  assert.equal(s.el('finishArea').style.display, 'block');
  s.dispatch('btnNovaCalc'); assert.deepEqual(plain(s.context.origensLegadoReentradas), []);
});

test('não oferece legado por erro de API, autorização, cancelamento, timeout ou dados incompletos', () => {
  for (const erro of ['HTTP 500', 'HTTP 401', 'Timeout', 'Chamado cancelado', 'Especificação não encontrada', 'Chamado não encontrado dentro do anexo']) {
    const s = makeScreen('2567'); s.dispatch('btnAvancar');
    s.reply('buscarResumoReentradasSPG', [{ chamado: '2567', origem: 'spg', erro, totalSkus: 0, tiragemTotal: 0 }]);
    assert.doesNotMatch(s.el('conteudoResumoData').textContent, /versão antiga de solicitação/);
    assert.equal(s.next().style.display, 'none');
    s.context.selecionarLegadoTemporarioReentradas(0, true);
    assert.deepEqual(plain(s.context.origensLegadoReentradas), []);
    assert.equal(s.requests.length, 1);
  }
  const s = makeScreen('2567'); s.dispatch('btnAvancar');
  s.reply('buscarResumoReentradasSPG', new Error('HTTP 500'), true);
  assert.match(s.el('conteudoResumoLoader').textContent, /Erro ao consultar o SPG/);
  assert.doesNotMatch(s.el('conteudoResumoLoader').textContent, /versão antiga de solicitação/);
  assert.equal(s.next().style.display, 'none');
});

test('alterar chamados limpa escolhas e respostas atrasadas não reabrem nem trocam o resumo', () => {
  const s = makeScreen('2567'); s.dispatch('btnAvancar');
  s.reply('buscarResumoReentradasSPG', [missing('2567')]);
  s.dispatch(s.button('Sim, usar legado temporário'));
  const old = s.pending('buscarResumoReentradasSPG'); old.done = true;
  s.context.cancelarResumo(); s.el('txtChamados').value = '3100'; s.dispatch('txtChamados', 'input');
  assert.deepEqual(plain(s.context.origensLegadoReentradas), []);
  s.dispatch('btnAvancar');
  old.success([ok('2567', 'legado')]);
  assert.equal(s.el('conteudoResumoData').style.display, 'none');
  old.failure(new Error('HTTP 500 atrasado'));
  assert.doesNotMatch(s.el('conteudoResumoLoader').textContent, /atrasado/);
  assert.deepEqual(plain(s.pending('buscarResumoReentradasSPG').args), ['3100', []]);
  s.reply('buscarResumoReentradasSPG', [ok('3100')]);
  assert.match(s.el('conteudoResumoData').textContent, /Chamado 3100/);
  assert.doesNotMatch(s.el('conteudoResumoData').textContent, /2567|Legado temporário/);
});

test('Voltar invalida consulta pendente e nova tentativa na mesma lista ignora resposta anterior', () => {
  const s = makeScreen('2567'); s.dispatch('btnAvancar');
  const old = s.pending('buscarResumoReentradasSPG'); old.done = true;
  s.context.cancelarResumo(); s.dispatch('btnAvancar');
  old.success([missing('2567')]);
  assert.equal(s.el('conteudoResumoData').style.display, 'none');
  s.reply('buscarResumoReentradasSPG', [ok('2567')]);
  assert.doesNotMatch(s.el('conteudoResumoData').textContent, /versão antiga de solicitação/);
  s.context.cancelarResumo(); s.context.continuarPosResumo();
  assert.equal(s.requests.some(r => r.method === 'validarReentradasSPG'), false);
});

test('falha no legado confirmado não permite nova migração nem avanço', () => {
  const s = makeScreen('2567'); s.dispatch('btnAvancar');
  s.reply('buscarResumoReentradasSPG', [missing('2567')]);
  s.dispatch(s.button('Sim, usar legado temporário'));
  s.reply('buscarResumoReentradasSPG', [{ chamado: '2567', origem: 'legado', legadoTemporario: true,
    naoEncontrado: true, totalSkus: 0, tiragemTotal: 0, erro: 'Dados antigos indisponíveis.' }]);
  assert.match(s.el('conteudoResumoData').textContent, /Falha ao carregar dados da alocação antiga/);
  assert.doesNotMatch(s.el('conteudoResumoData').textContent, /versão antiga de solicitação/);
  assert.equal(s.next().style.display, 'none');
  s.context.selecionarLegadoTemporarioReentradas(0, true); assert.equal(s.requests.length, 2);
});

test('abertura consulta últimas reentradas uma vez sem bloquear entrada manual nem consultar legado', () => {
  const s = makeScreen('');
  assert.match(s.el('reentradasRecentesStatus').textContent, /Carregando últimas reentradas/);
  assert.equal(s.el('txtChamados').disabled, false);
  s.runTimers(); s.runTimers();
  assert.equal(s.requests.filter(r => r.method === 'listarUltimasReentradasSPG').length, 1);
  assert.deepEqual(plain(s.pending('listarUltimasReentradasSPG').args), []);
  assert.equal(s.requests.some(r => r.method === 'buscarResumoReentradasSPG'), false);
  s.reply('listarUltimasReentradasSPG', { chamados: [
    { chamado: '3201', status: 'Alocado', marca: 'SAS', envio: '20/10/2026', alocadoEm: 'Meta' },
    { chamado: '3200', status: 'Em análise', marca: 'Pleno' }
  ] });
  assert.match(s.el('reentradasRecentesStatus').textContent, /2 chamado\(s\)/);
  assert.equal(s.el('reentradasRecentesLista').querySelectorAll('tbody tr').length, 2);
  assert.match(s.el('reentradasRecentesLista').textContent, /ChamadoStatusMarcaEnvioAlocado em/);
  assert.match(s.el('reentradasRecentesLista').textContent, /20\/10\/2026Meta/);
  assert.match(s.el('reentradasRecentesLista').textContent, /—/);
  s.dispatch(s.button('3201'));
  assert.equal(s.el('txtChamados').value, '3201');
  assert.equal(s.el('btnAvancar').disabled, false);
  s.dispatch('btnAvancar');
  assert.deepEqual(plain(s.pending('buscarResumoReentradasSPG').args), ['3201', []]);
  s.dispatch(s.button('3200'));
  assert.equal(s.el('txtChamados').value, '3201', 'Lista não muda entrada durante consulta ativa');
});

test('lista recente tem estados vazio e erro independentes da validação manual', () => {
  for (const response of [{ chamados: [] }, { erro: 'HTTP 500' }, null]) {
    const s = makeScreen('2567'); s.runTimers();
    s.reply('listarUltimasReentradasSPG', response);
    assert.match(s.el('reentradasRecentesStatus').textContent, response && Array.isArray(response.chamados) ? /Nenhuma solicitação/ : /Você pode informar o chamado acima/);
    assert.equal(s.el('reentradasRecentesLista').querySelectorAll('table').length, 0);
    assert.equal(s.el('txtChamados').disabled, false);
    assert.equal(s.el('btnAvancar').disabled, false);
    s.dispatch('btnAvancar'); assert(s.pending('buscarResumoReentradasSPG'));
  }
  const s = makeScreen(); s.runTimers();
  s.reply('listarUltimasReentradasSPG', new Error('timeout'), true);
  assert.match(s.el('reentradasRecentesStatus').textContent, /Você pode informar o chamado acima/);
  assert.equal(s.el('btnAvancar').disabled, false);
});

test('lista recente escapa todos os campos e nunca inventa linhas para completar trinta', () => {
  const s = makeScreen(); s.runTimers();
  s.reply('listarUltimasReentradasSPG', { chamados: [{ chamado: '<bad>', status: '<img>',
    marca: '<script>', envio: '<envio>', alocadoEm: '<grafica>' }] });
  const content = s.el('reentradasRecentesLista').innerHTML;
  for (const tag of ['bad', 'img', 'script', 'envio', 'grafica']) assert.match(content, new RegExp('&lt;' + tag + '&gt;'));
  assert.doesNotMatch(content, /<(?:bad|img|script|envio|grafica)>/);
  assert.equal(s.el('reentradasRecentesLista').querySelectorAll('tbody tr').length, 1);
  s.context.renderizarUltimasReentradasSPG({ chamados: Array.from({ length: 35 }, (_, index) => ({ chamado: String(4000 - index) })) });
  assert.equal(s.el('reentradasRecentesLista').querySelectorAll('tbody tr').length, 30);
  assert.match(s.el('reentradasRecentesStatus').textContent, /30 chamado\(s\)/);
});
