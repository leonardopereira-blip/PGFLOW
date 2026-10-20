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
const plain = value => JSON.parse(JSON.stringify(value));
const sku = 'MS26VF331P1SP';
const ok = (chamado, origem = 'legado') => ({ chamado, origem, totalSkus: 2, tiragemTotal: 300,
  marcas: 'SAS', series: '3ª Série', papeisCapa: 'Cartão', papeisMiolo: 'Offset', cdsDestino: 'FOR' });
const absent = (chamado = '2632', skus = [sku]) => ({ chamado, origem: 'legado',
  erro: 'SKU sem especificação na alocação antiga: ' + skus.join(', '),
  especificacoesAusentes: skus, totalSkus: 0, tiragemTotal: 0 });
function enterLegacy(s) {
  s.dispatch('btnAvancar');
  s.reply('buscarResumoReentradasSPG', [{ chamado: '2632', origem: 'spg', naoEncontrado: true,
    erro: 'Chamado inexistente.', totalSkus: 0, tiragemTotal: 0 }]);
  s.dispatch(s.button('Sim, usar legado temporário'));
}
function choose(s, index = 0) {
  const buttons = s.document.querySelectorAll('button').filter(btn => btn.textContent.trim() === 'Usar especificação geral deste SKU');
  assert(buttons[index], 'Missing SKU confirmation'); s.dispatch(buttons[index]);
}

test('especificação geral de SKU legado exige aceite e acompanha resumo, validação e geração', () => {
  const s = makeScreen('2632'); enterLegacy(s);
  s.reply('buscarResumoReentradasSPG', [absent()]);
  assert.match(s.el('conteudoResumoData').textContent, /SKU sem especificação na alocação antiga/);
  assert.match(s.el('conteudoResumoData').textContent, /MS26VF331P1SP/);
  assert.match(s.el('conteudoResumoData').textContent, /A tiragem e a árvore continuam na origem já confirmada/);
  assert.doesNotMatch(s.el('conteudoResumoData').textContent, /versão antiga de solicitação/);
  assert.equal(s.next().style.display, 'none'); s.context.continuarPosResumo();
  assert.equal(s.requests.some(req => req.method === 'validarReentradasSPG'), false);
  assert.deepEqual(plain(s.context.especificacoesGeraisReentradas), []);
  choose(s);
  const choices = [{ chamado: '2632', sku }];
  assert.deepEqual(plain(s.pending('buscarResumoReentradasSPG').args), ['2632', ['2632'], choices]);
  s.reply('buscarResumoReentradasSPG', [{ ...ok('2632'), skusEspecificacaoGeral: [sku], avisos: ['SKU ' + sku + ': especificação técnica geral utilizada.'] }]);
  assert.match(s.el('conteudoResumoData').textContent, /Consulta geral autorizada/);
  assert.match(s.el('conteudoResumoData').textContent, /Especificação técnica geral utilizada/);
  assert.match(s.el('conteudoResumoData').textContent, /especificação técnica geral utilizada/);
  assert.equal(s.next().style.display, ''); s.dispatch(s.next());
  assert.deepEqual(plain(s.pending('validarReentradasSPG').args), ['META', '2632', ['2632'], choices]);
  s.reply('validarReentradasSPG', { validos: ['2632'], invalidos: [], colunasManuais: [] });
  assert.match(s.el('painelChips').textContent, /2632 · Legado temporário · Especificação geral/);
  s.dispatch('btnConcluir');
  assert.deepEqual(plain(s.pending('gerarValidacaoReentradasSPG').args.slice(6)), [['2632'], choices]);
  s.reply('gerarValidacaoReentradasSPG', { sucesso: true, arquivos: ['file'], resumo: [],
    avisos: ['Chamado 2632: especificação geral utilizada para ' + sku] });
  assert.equal(s.el('finishArea').style.display, 'block');
  assert.match(s.el('erroFase3').textContent, /especificação geral utilizada/);
  s.dispatch('btnNovaCalc'); assert.deepEqual(plain(s.context.especificacoesGeraisReentradas), []);
});

test('cada SKU ausente do legado precisa de autorização própria sem escolha global', () => {
  const s = makeScreen('2632'); enterLegacy(s);
  s.reply('buscarResumoReentradasSPG', [absent('2632', [sku, 'SKU-B'])]);
  choose(s);
  assert.deepEqual(plain(s.pending('buscarResumoReentradasSPG').args[2]), [{ chamado: '2632', sku }]);
  s.reply('buscarResumoReentradasSPG', [absent('2632', ['SKU-B'])]);
  assert.equal(s.next().style.display, 'none');
  assert.match(s.el('conteudoResumoData').textContent, /Consulta geral autorizada/);
  assert.match(s.el('conteudoResumoData').textContent, /MS26VF331P1SP/);
  choose(s);
  assert.deepEqual(plain(s.pending('buscarResumoReentradasSPG').args[2]), [
    { chamado: '2632', sku }, { chamado: '2632', sku: 'SKU-B' }
  ]);
  s.reply('buscarResumoReentradasSPG', [ok('2632')]);
  assert.equal(s.next().style.display, '');
});

test('opção não aparece para OP do SPG nem por texto de erro sem lista tipada', () => {
  for (const response of [
    { ...absent(), origem: 'spg' },
    { ...absent(), especificacoesAusentes: undefined },
    { ...absent(), especificacoesAusentes: 'MS26VF331P1SP' },
    { ...absent(), especificacoesAusentes: [null, {}, '', 123] }
  ]) {
    const s = makeScreen('2632'); enterLegacy(s); s.reply('buscarResumoReentradasSPG', [response]);
    assert.doesNotMatch(s.el('conteudoResumoData').textContent, /Usar especificação geral deste SKU/);
    s.context.selecionarEspecificacaoGeralReentradas(0, 0);
    assert.deepEqual(plain(s.context.especificacoesGeraisReentradas), []);
    assert.equal(s.requests.length, 2); assert.equal(s.next().style.display, 'none');
  }
});

test('alterar os chamados limpa autorizações e ignora resumo e falha atrasados da especificação geral', () => {
  const s = makeScreen('2632'); enterLegacy(s); s.reply('buscarResumoReentradasSPG', [absent()]); choose(s);
  const old = s.pending('buscarResumoReentradasSPG'); old.done = true;
  s.context.cancelarResumo(); s.el('txtChamados').value = '3100'; s.dispatch('txtChamados', 'input');
  assert.deepEqual(plain(s.context.especificacoesGeraisReentradas), []);
  s.dispatch('btnAvancar'); old.success([ok('2632')]); old.failure(new Error('falha atrasada'));
  assert.equal(s.el('conteudoResumoData').style.display, 'none');
  assert.doesNotMatch(s.el('conteudoResumoLoader').textContent, /atrasada/);
  assert.deepEqual(plain(s.pending('buscarResumoReentradasSPG').args), ['3100', [], []]);
  s.reply('buscarResumoReentradasSPG', [ok('3100', 'spg')]);
  assert.doesNotMatch(s.el('conteudoResumoData').textContent, /2632|Consulta geral autorizada/);
});

test('especificação geral mantém alerta escapado e ausência tipada permanece pendente mesmo com totais válidos', () => {
  const s = makeScreen('2632'); enterLegacy(s);
  s.reply('buscarResumoReentradasSPG', [{ ...absent('2632', ['<sku onclick="bad">']),
    totalSkus: 2, tiragemTotal: 300, erro: '', avisos: ['Aviso <script>'] }]);
  const content = s.el('conteudoResumoData').innerHTML;
  assert.match(content, /&lt;sku onclick=&quot;bad&quot;&gt;/);
  assert.match(content, /Aviso &lt;script&gt;/);
  assert.doesNotMatch(content, /<sku|<script>/);
  assert.equal(s.next().style.display, 'none');
  s.context.cancelarResumo();
  s.context.selecionarEspecificacaoGeralReentradas(0, 0);
  assert.deepEqual(plain(s.context.especificacoesGeraisReentradas), []);
});

const ignored = (chamado = '2632', skus = [sku]) => ({ ...ok(chamado), totalSkus: 0, tiragemTotal: 0,
  skusIgnorados: skus, skusEspecificacaoGeral: [], semSkusParaGerar: true,
  avisos: ['Chamado ' + chamado + ': SKU ' + skus.join(', ') + ' não encontrado no cadastro geral.'] });

test('SKU ausente no cadastro geral é avisado e excluído sem bloquear os demais SKUs do mesmo chamado', () => {
  const s = makeScreen('2632'); enterLegacy(s);
  s.reply('buscarResumoReentradasSPG', [absent()]); choose(s);
  s.reply('buscarResumoReentradasSPG', [{ ...ok('2632'), totalSkus: 1, tiragemTotal: 150,
    skusIgnorados: [sku], skusEspecificacaoGeral: [], avisos: ['SKU ' + sku + ' ignorado.'] }]);
  const summary = s.el('conteudoResumoData').textContent;
  assert.match(summary, /Consulta geral autorizada/);
  assert.match(s.el('conteudoResumoData').innerHTML, /<strong>SKUs ignorados:<\/strong> MS26VF331P1SP/);
  assert.match(summary, /Serão excluídos da geração/);
  assert.match(s.el('conteudoResumoData').innerHTML, /Qtd\. SKU's Únicos:<\/strong> 1/);
  assert.match(s.el('conteudoResumoData').innerHTML, /Tiragem Válida:<\/strong> 150/);
  assert.doesNotMatch(summary, /Especificação técnica geral utilizada|Falha ao carregar/);
  assert.equal(s.next().style.display, ''); s.dispatch(s.next());
  s.reply('validarReentradasSPG', { validos: ['2632'], invalidos: [], colunasManuais: [], avisos: ['SKU ' + sku + ' ignorado.'] });
  assert.match(s.el('painelChips').textContent, /2632 · Legado temporário/);
  assert.doesNotMatch(s.el('painelChips').textContent, /Especificação geral/);
  assert.match(s.el('painelInvalidos').textContent, /SKU MS26VF331P1SP ignorado/);
  s.dispatch('btnConcluir');
  s.reply('gerarValidacaoReentradasSPG', { sucesso: true, arquivos: ['file'], resumo: [] });
  assert.match(s.el('erroFase3').textContent, /SKUs ignorados: MS26VF331P1SP/);
});

test('chamado sem SKUs geráveis não bloqueia outro chamado e seu aviso permanece na conclusão', () => {
  const s = makeScreen('2632, 3100'); enterLegacy(s);
  s.reply('buscarResumoReentradasSPG', [ignored(), ok('3100', 'spg')]);
  assert.equal(s.next().style.display, '');
  assert.match(s.el('conteudoResumoData').textContent, /Este chamado ficou sem SKUs para gerar e será ignorado/);
  assert.doesNotMatch(s.el('conteudoResumoData').textContent, /Falha ao carregar dados/);
  s.dispatch(s.next());
  s.reply('validarReentradasSPG', { validos: ['3100'], invalidos: ['2632: nenhum SKU para gerar'], colunasManuais: [] });
  assert.equal(s.el('painelChips').textContent, '3100');
  assert.match(s.el('painelInvalidos').textContent, /2632: nenhum SKU para gerar/);
  assert.doesNotMatch(s.el('painelInvalidos').textContent, /não encontrados\)/);
  s.dispatch('btnConcluir');
  assert.deepEqual(plain(s.pending('gerarValidacaoReentradasSPG').args[1]), ['3100']);
  s.reply('gerarValidacaoReentradasSPG', { sucesso: true, arquivos: ['file'], resumo: [], avisos: ['Outro aviso'] });
  assert.match(s.el('erroFase3').textContent, /Chamado 2632: SKU MS26VF331P1SP não encontrado/);
  assert.match(s.el('erroFase3').textContent, /Outro aviso/);
});

test('todos os chamados sem SKUs geráveis exibem motivo claro e não seguem para arquivos vazios', () => {
  const s = makeScreen('2632'); enterLegacy(s);
  s.reply('buscarResumoReentradasSPG', [ignored()]);
  assert.equal(s.next().style.display, 'none');
  assert.match(s.el('conteudoResumoData').textContent, /Nenhum chamado possui SKUs para gerar\. Nenhum arquivo será criado/);
  assert.doesNotMatch(s.el('conteudoResumoData').textContent, /Falha ao carregar dados/);
  s.context.continuarPosResumo();
  assert.equal(s.requests.some(req => req.method === 'validarReentradasSPG'), false);
  s.context.processarResultadoFase1App({ validos: [], invalidos: ['2632: nenhum SKU'], colunasManuais: [] });
  assert.match(s.el('mensagemFase1').textContent, /Nenhum chamado possui SKUs para gerar/);
  assert.equal(s.el('telaFase2').style.display, 'none');
});

test('flag sem SKUs não oculta erro real, especificação pendente ou chamado não encontrado', () => {
  for (const bad of [
    { ...ignored(), erro: 'API indisponível' },
    { ...ignored(), especificacoesAusentes: ['SKU-PENDENTE'] },
    { ...ignored(), naoEncontrado: true }
  ]) {
    const s = makeScreen('2632, 3100'); enterLegacy(s);
    s.reply('buscarResumoReentradasSPG', [bad, ok('3100', 'spg')]);
    assert.equal(s.next().style.display, 'none');
    s.context.continuarPosResumo();
    assert.equal(s.requests.some(req => req.method === 'validarReentradasSPG'), false);
  }
});

test('SKU ignorado e motivos aparecem escapados, sem aviso de uso indevido de especificação geral', () => {
  const s = makeScreen('2632'); enterLegacy(s);
  s.reply('buscarResumoReentradasSPG', [{ ...ok('2632'), skusIgnorados: ['<sku onclick="bad">'],
    skusEspecificacaoGeral: [], avisos: ['Aviso <script>'] }]);
  const summary = s.el('conteudoResumoData').innerHTML;
  assert.match(summary, /&lt;sku onclick=&quot;bad&quot;&gt;/);
  assert.doesNotMatch(summary, /<sku|<script>/);
  s.dispatch(s.next());
  s.reply('validarReentradasSPG', { validos: ['2632'], invalidos: ['<img onerror="bad">'], colunasManuais: [] });
  const warnings = s.el('painelInvalidos').innerHTML;
  assert.match(warnings, /&lt;img onerror=&quot;bad&quot;&gt;/);
  assert.match(warnings, /&lt;sku onclick=&quot;bad&quot;&gt;/);
  assert.doesNotMatch(warnings, /<img|<sku|<script>/);
});

test('novo conjunto de chamados limpa os avisos de exclusão da validação anterior', () => {
  const s = makeScreen('2632'); enterLegacy(s);
  s.reply('buscarResumoReentradasSPG', [{ ...ok('2632'), skusIgnorados: [sku], skusEspecificacaoGeral: [] }]);
  s.dispatch(s.next());
  s.reply('validarReentradasSPG', { validos: ['2632'], invalidos: [], colunasManuais: [] });
  assert(s.context._avisosValidacaoReentradas.length > 0);
  s.context.resetarParaNovaCalculadora();
  assert.deepEqual(plain(s.context._avisosValidacaoReentradas), []);
});
