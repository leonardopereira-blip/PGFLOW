// DOM and Apps Script calls are simulated. No browser, API, Drive or email is used.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const html = fs.readFileSync(path.join(__dirname, '..', 'TelaEmail.html'), 'utf8');
const source = html.match(/<script>([\s\S]*?)<\/script>/)[1];

function makeScreen() {
  let doc;
  function matches(el, selector) {
    if (selector.includes(':checked') && !el.checked) return false;
    const tag = selector.match(/^[\w-]+/);
    if (tag && el.tagName !== tag[0]) return false;
    const id = selector.match(/#([\w-]+)/);
    if (id && el.id !== id[1]) return false;
    for (const [, name] of selector.matchAll(/\.([\w-]+)/g)) {
      if (!el.className.split(/\s+/).includes(name)) return false;
    }
    return Array.from(selector.matchAll(/\[([\w-]+)(?:="([^"]*)")?\]/g)).every(([, key, val]) => {
      const actual = ['name', 'value', 'type', 'id'].includes(key) ? el[key] : el.getAttribute(key);
      return val === undefined ? actual != null : actual === val;
    });
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
  class Element {
    constructor(tag = '', attrs = {}) {
      this.tagName = tag; this.attrs = attrs; this.children = []; this.parent = null;
      this.id = attrs.id || ''; this.name = attrs.name || ''; this.type = attrs.type || '';
      this.className = attrs.class || ''; this.value = attrs.value || ''; this._text = '';
      this._checked = 'checked' in attrs; this.hidden = 'hidden' in attrs;
      this.disabled = 'disabled' in attrs; this.style = {}; this.dataset = {};
      this.classList = {
        contains: name => this.className.split(/\s+/).includes(name),
        toggle: (name, on) => {
          const tokens = this.className.split(/\s+/).filter(Boolean).filter(x => x !== name);
          if (on) tokens.push(name);
          this.className = tokens.join(' ');
        }
      };
    }
    appendChild(child) { child.parent = this; this.children.push(child); return child; }
    set textContent(value) {
      this._text = String(value);
      this.children.forEach(child => { child.parent = null; }); this.children = [];
    }
    get textContent() { return this._text + this.children.map(child => child.textContent).join(''); }
    get checked() { return this._checked; }
    set checked(on) {
      if (on && this.type === 'radio' && this.name && doc) {
        doc.querySelectorAll('input').filter(el => el.name === this.name).forEach(el => { el._checked = false; });
      }
      this._checked = !!on;
    }
    setAttribute(key, val) { this.attrs[key] = String(val); }
    getAttribute(key) { return this.attrs[key] === undefined ? null : this.attrs[key]; }
    removeAttribute(key) { delete this.attrs[key]; }
    querySelectorAll(selectors) {
      const all = [];
      function walk(el) { el.children.forEach(child => { all.push(child); walk(child); }); }
      walk(this);
      return all.filter(el => selectors.split(',').some(selector => fits(el, selector)));
    }
    querySelector(selector) { return this.querySelectorAll(selector)[0] || null; }
  }
  doc = new Element('document');
  doc.createElement = tag => new Element(tag);
  doc.getElementById = id => doc.querySelectorAll('*').find(el => el.id === id) || null;
  const tree = html.replace(/<script>[\s\S]*?<\/script>/g, '').replace(/<style>[\s\S]*?<\/style>/g, '');
  const stack = [doc];
  const voidTags = new Set(['input', 'meta', 'base', 'br', 'hr', 'link', 'img']);
  for (const [token, rawTag, attributes] of tree.matchAll(/<\/?([a-z][\w-]*)\b([^>]*)>/gi)) {
    const tag = rawTag.toLowerCase();
    if (token.startsWith('</')) { if (stack.length > 1) stack.pop(); continue; }
    const attrs = {};
    for (const [, key, value] of attributes.matchAll(/([\w-]+)(?:="([^"]*)")?/g)) attrs[key] = value ?? '';
    const el = new Element(tag, attrs); stack.at(-1).appendChild(el);
    if (!voidTags.has(tag)) stack.push(el);
  }
  const requests = [], alerts = [];
  const google = { script: {} };
  Object.defineProperty(google.script, 'run', { get() {
    const chain = {
      withSuccessHandler(fn) { this.success = fn; return this; },
      withFailureHandler(fn) { this.failure = fn; return this; }
    };
    ['obterCadastroEmailsSPG', 'salvarCadastroEmailsSPG', 'consultarPacoteEmailSPG',
      'obterDadosChamado', 'dispararEmailPacoteSPG', 'dispararEmailFinal'].forEach(method => {
      chain[method] = (...args) => requests.push({ method, args, success: chain.success, failure: chain.failure });
    });
    return chain;
  }});
  const graficas = { Teste: ['test@example.com', 'shared@example.com'], Outra: ['other@example.com', 'shared@example.com'] };
  const context = { document: doc, google, alert: msg => alerts.push(msg), confirm: () => true };
  vm.createContext(context); vm.runInContext(source.replace('/*PF_EMAIL_CADASTRO_INICIAL*/null',
    JSON.stringify({graficas,versao:'v1'})), context);
  const el = id => { const value = doc.getElementById(id); assert(value, 'Missing element ' + id); return value; };
  function pending(method) {
    const result = requests.find(req => req.method === method && !req.done);
    assert(result, 'Expected pending ' + method); return result;
  }
  function reply(method, result, error = false) {
    const req = pending(method); req.done = true; (error ? req.failure : req.success)(result); return req;
  }
  function dispatch(element, event = 'click') {
    if (typeof element === 'string') element = el(element);
    if (element.disabled) return;
    if (typeof element['on' + event] === 'function') return element['on' + event]();
    const code = element.getAttribute('on' + event);
    if (code) vm.runInContext(code, context);
  }
  function graphic(name) { return el('containerGraficas').querySelectorAll('input').find(x => x.value === name); }
  function choose(names = ['Teste'], type = 'reentrada') {
    el('inputChamado').value = '3100';
    doc.querySelector('input[name="tipoEmail"][value="' + type + '"]').checked = true;
    el('containerGraficas').querySelectorAll('input').forEach(input => { input.checked = names.includes(input.value); });
  }
  function button(text, scope = doc) {
    const result = scope.querySelectorAll('button').find(x => x.textContent === text);
    assert(result, 'Missing button ' + text); return result;
  }
  choose();
  return { context, doc, el, requests, alerts, pending, reply, dispatch, graphic, choose, button, graficas };
}
const complete = { status: 'ok', idChamado: '3100', anexoValido: true,
  anexoHash: 'hash3100', dataDesejada: '15/11/2026', descricao: 'Produção 3100',
  statusChamado: 'SOLICITADO', statusProducao: 'PUBLICADA' };

test('revisão mostra o status atual do chamado sem confundir com resultado técnico ou versão', () => {
  const s = makeScreen(); s.context.avancarParaRevisao();
  s.reply('consultarPacoteEmailSPG', { ...complete, statusChamado: 'VALIDADO' });
  assert.equal(s.el('pf-status-spg').textContent, 'Status do chamado no SPG: Validado');
  assert.equal(s.el('pf-status-spg').style.display, 'block');
  assert.equal(s.el('pf-status-spg').getAttribute('role'), 'status');
  assert.equal(s.el('passo2').classList.contains('ativo'), true);
  assert.doesNotMatch(s.el('corpoEmail').value, /Status do chamado/);
});

test('chamado cancelado mostra alerta e não permite legado, dados manuais, revisão ou envio', () => {
  const s = makeScreen(); s.context.avancarParaRevisao();
  s.reply('consultarPacoteEmailSPG', { ...complete, status: 'cancelado', cancelado: true,
    statusChamado: 'CANCELADO', dataDesejada: '', anexoValido: false,
    erro: 'Este chamado está cancelado no SPG. O envio de formalização está indisponível.',
    faltantes: ['Data desejada de entrega'] });
  assert.match(s.el('pf-status-spg').textContent, /Status do chamado no SPG: Cancelado/);
  assert.match(s.el('pf-status-spg').textContent, /envio de formalização está indisponível/);
  assert.equal(s.el('pf-status-spg').getAttribute('role'), 'alert');
  assert.equal(s.el('pgflowOrigemSPG').style.display, 'none');
  assert.equal(s.el('pgflowOrigemSPG').textContent, '');
  assert.equal(s.el('passo1').classList.contains('ativo'), true);
  assert.equal(s.el('passo2').classList.contains('ativo'), false);
  s.context.enviarEmail();
  assert.equal(s.requests.some(r => /obterDadosChamado|dispararEmail/.test(r.method)), false);
  assert.equal(Object.keys(s.context.dadosGlobais).length, 0);
});

for (const statusProducao of ['CANCELADO', 'CANCELADA']) test('produção ' + statusProducao + ' preserva a distinção do status ativo do chamado', () => {
  const s = makeScreen(); s.context.avancarParaRevisao();
  s.reply('consultarPacoteEmailSPG', { ...complete, status: 'cancelado', cancelado: true,
    statusChamado: 'LIBERADO', statusProducao, anexoValido: false });
  assert.match(s.el('pf-status-spg').textContent, /Status do chamado no SPG: Liberado/);
  assert.match(s.el('pf-status-spg').textContent, /Status da produção no SPG: Cancelada/);
  assert.doesNotMatch(s.el('pf-status-spg').textContent, /Status do chamado no SPG: Cancelado/);
  assert.equal(s.el('passo2').classList.contains('ativo'), false);
});

test('status real CANCELADA da produção bloqueia mesmo sem marcador técnico de cancelamento', () => {
  const s = makeScreen(); s.context.avancarParaRevisao();
  s.reply('consultarPacoteEmailSPG', { ...complete, statusChamado: 'VALIDADO', statusProducao: 'CANCELADA' });
  assert.equal(s.context.pfChamadoCancelado, true);
  assert.match(s.el('pf-status-spg').textContent, /Status do chamado no SPG: Validado/);
  assert.match(s.el('pf-status-spg').textContent, /Status da produção no SPG: Cancelada/);
  assert.equal(s.el('passo2').classList.contains('ativo'), false);
  assert.equal(s.el('pgflowOrigemSPG').textContent, '');
});

test('status desconhecido é exibido como texto seguro e ausente não é inventado', () => {
  const s = makeScreen(); s.context.avancarParaRevisao();
  s.reply('consultarPacoteEmailSPG', { ...complete, statusChamado: '<img src=x onerror=alert(1)>' });
  assert.equal(s.el('pf-status-spg').textContent, 'Status do chamado no SPG: <img src=x onerror=alert(1)>');
  assert.equal(s.el('pf-status-spg').querySelectorAll('img').length, 0);
  s.context.pfResetar();
  assert.equal(s.el('pf-status-spg').textContent, '');
  assert.equal(s.el('pf-status-spg').style.display, 'none');
  s.choose(['Teste'], 'inclusivo'); s.context.avancarParaRevisao();
  s.reply('obterDadosChamado', { idExcel: 'excel3100', valorF: '3100', dataColeta: '12/11/2026' });
  assert.equal(s.el('pf-status-spg').style.display, 'none');
});

test('faltantes mantém status durante a pergunta de origem e preenchimento manual', () => {
  const s = makeScreen(); s.context.avancarParaRevisao();
  s.reply('consultarPacoteEmailSPG', { ...complete, status: 'incompleto', statusChamado: 'ALOCADO',
    dataDesejada: '', faltantes: ['Data desejada de entrega'] });
  assert.equal(s.el('pf-status-spg').textContent, 'Status do chamado no SPG: Alocado');
  assert.match(s.el('pgflowOrigemSPG').textContent, /fluxo antigo/);
  s.dispatch(s.button('Não, continuar no SPG'));
  assert.equal(s.el('pf-status-spg').textContent, 'Status do chamado no SPG: Alocado');
  s.el('pgflowOrigemSPG').querySelector('input[type="checkbox"]').checked = true;
  s.dispatch(s.button('Continuar para revisão'));
  assert.equal(s.el('pf-status-spg').textContent, 'Status do chamado no SPG: Alocado');
  assert.equal(s.el('passo2').classList.contains('ativo'), true);
});

test('opção explícita pelo legado conserva apenas o status já consultado no SPG', () => {
  const s = makeScreen(); s.context.avancarParaRevisao();
  s.reply('consultarPacoteEmailSPG', { ...complete, status: 'incompleto', statusChamado: 'SOLICITADO',
    dataDesejada: '', faltantes: ['Data desejada de entrega'] });
  s.dispatch(s.button('Sim, fluxo antigo'));
  s.reply('obterDadosChamado', { idExcel: 'excel3100', valorF: '3100', dataColeta: '12/11/2026' });
  assert.equal(s.el('pf-status-spg').textContent, 'Status do chamado no SPG: Solicitado');
  assert.equal(s.context.dadosGlobais.origem, 'antigo');
});

test('troca do chamado e cancelar apagam status, inclusive cancelado, sem aceitar resposta atrasada', () => {
  const s = makeScreen(); s.context.avancarParaRevisao();
  s.reply('consultarPacoteEmailSPG', { ...complete, status: 'cancelado', cancelado: true, statusChamado: 'CANCELADO' });
  s.el('inputChamado').value = '2961'; s.dispatch('inputChamado', 'input');
  assert.equal(s.el('pf-status-spg').textContent, '');
  assert.equal(s.el('pf-status-spg').style.display, 'none');
  assert.equal(s.context.pfChamadoCancelado, false);
  s.context.avancarParaRevisao();
  const late = s.pending('consultarPacoteEmailSPG');
  s.context.pfResetar();
  late.done = true; late.success({ ...complete, idChamado: '2961', statusChamado: 'CANCELADO', cancelado: true });
  assert.equal(s.el('pf-status-spg').textContent, '');
  assert.equal(s.context.pfChamadoCancelado, false);
});

test('consulta bloqueia os campos e mantém snapshot do chamado, tipo, versão e gráficas', () => {
  const s = makeScreen(); s.context.avancarParaRevisao();
  assert.equal(s.el('inputChamado').disabled, true);
  assert.equal(s.graphic('Teste').disabled, true);
  assert.equal(s.pending('consultarPacoteEmailSPG').args.join('|'), '3100|ultima|reentrada');
  // Even a programmatic mutation without an input event cannot relabel this response.
  s.el('inputChamado').value = '2961'; s.graphic('Outra').checked = true;
  s.reply('consultarPacoteEmailSPG', complete);
  assert.equal(s.context.dadosGlobais.idChamado, '3100');
  assert.equal(s.context.pfGruposSelecionados.join(','), 'Teste');
  assert.match(s.el('assuntoEmail').value, /Chamado 3100/);
  assert.equal(s.el('passo2').classList.contains('ativo'), true);
});

test('inclusivo mantém arquivo e identificação da consulta antiga capturada', () => {
  const s = makeScreen(); s.choose(['Teste'], 'inclusivo'); s.context.avancarParaRevisao();
  assert.equal(s.requests.some(r => r.method === 'consultarPacoteEmailSPG'), false);
  assert.equal(s.pending('obterDadosChamado').args[0], '3100');
  s.el('inputChamado').value = '2961';
  s.reply('obterDadosChamado', { idExcel: 'excel3100', valorF: '3100', dataColeta: '12/11/2026' });
  assert.equal(s.context.dadosGlobais.idChamado, '3100');
  assert.equal(s.context.dadosGlobais.idArquivo, 'excel3100');
  assert.equal(s.context.dadosGlobais.tipoEmail, 'inclusivo');
});

test('Cancelar invalida resposta atrasada; nova consulta não recebe dados da anterior', () => {
  const s = makeScreen(); s.context.avancarParaRevisao();
  const first = s.pending('consultarPacoteEmailSPG');
  s.context.pfResetar(); assert.equal(s.el('inputChamado').disabled, false);
  s.choose(); s.el('inputChamado').value = '2961'; s.context.avancarParaRevisao();
  first.done = true; first.success(complete);
  assert.equal(s.el('passo2').classList.contains('ativo'), false);
  s.reply('consultarPacoteEmailSPG', { ...complete, idChamado: '2961', anexoHash: 'hash2961' });
  assert.equal(s.context.dadosGlobais.idChamado, '2961');
});

test('alterar o chamado depois da resposta incompleta invalida botões de decisão antigos', () => {
  const s = makeScreen(); s.context.avancarParaRevisao();
  s.reply('consultarPacoteEmailSPG', { ...complete, status: 'incompleto', faltantes: ['Data desejada de entrega'], dataDesejada: '' });
  const legacyButton = s.button('Sim, fluxo antigo');
  s.el('inputChamado').value = '2961'; s.dispatch('inputChamado', 'input');
  s.dispatch(legacyButton);
  assert.equal(s.requests.some(r => r.method === 'obterDadosChamado'), false);
  assert.equal(s.el('pgflowOrigemSPG').textContent, '');
  assert.equal(s.context.pfConsultaAtual, null);
});

test('fluxo incompleto oferece escolha antiga e permite omitir só o dado faltante', () => {
  const s = makeScreen(); s.context.avancarParaRevisao();
  s.reply('consultarPacoteEmailSPG', { ...complete, status: 'incompleto', faltantes: ['Data desejada de entrega'], dataDesejada: '' });
  s.dispatch(s.button('Não, continuar no SPG'));
  const skip = s.el('pgflowOrigemSPG').querySelector('input[type="checkbox"]'); skip.checked = true;
  s.dispatch(s.button('Continuar para revisão'));
  assert.equal(s.context.dadosGlobais.sobrescrita.omitirData, true);
  assert.equal(s.context.dadosGlobais.idChamado, '3100');
  assert.doesNotMatch(s.el('corpoEmail').value, /5\. Data da coleta desejada/);
});

test('resposta de outro chamado é bloqueada na revisão', () => {
  const s = makeScreen(); s.context.avancarParaRevisao();
  s.reply('consultarPacoteEmailSPG', { ...complete, idChamado: '9999' });
  assert.equal(s.el('passo2').classList.contains('ativo'), false);
  assert.match(s.el('mensagem-erro').textContent, /não corresponde/);
});

test('seleção e cancelamento do email ficam isolados da aba de custos', () => {
  const s = makeScreen();
  const cost = s.doc.createElement('input'); cost.type = 'checkbox'; cost.className = 'chk-grafica';
  cost.value = 'Outra'; cost.checked = true; s.doc.appendChild(cost);
  assert.equal(s.context.pfSelecionadas().join(','), 'Teste');
  assert.equal(s.doc.querySelectorAll('.chk-grafica:checked').length, 1);
  s.context.pfResetar();
  assert.equal(cost.checked, true); assert.equal(s.graphic('Teste').checked, false);
});

test('versão e bloco manual do email não alteram os controles de versão dos custos', () => {
  const s = makeScreen();
  const costLatest = s.doc.createElement('input');
  costLatest.type = 'radio'; costLatest.name = 'opcaoVersao'; costLatest.value = 'ultima';
  const costManual = s.doc.createElement('input');
  costManual.type = 'radio'; costManual.name = 'opcaoVersao'; costManual.value = 'selecionar';
  const costBlock = s.doc.createElement('div'); costBlock.id = 'blocoVersoesManuais';
  costBlock.style.display = 'block';
  s.doc.appendChild(costLatest); s.doc.appendChild(costManual); s.doc.appendChild(costBlock);
  costManual.checked = true;
  const emailLatest = s.doc.querySelector('input[name="pfEmailVersao"][value="ultima"]');
  const emailManual = s.doc.querySelector('input[name="pfEmailVersao"][value="manual"]');
  assert.equal(emailLatest.checked, true);
  emailManual.checked = true; s.dispatch(emailManual, 'change');
  assert.equal(costManual.checked, true);
  assert.equal(s.el('pfEmailBlocoVersoesManuais').style.display, 'block');
  assert.equal(s.el('blocoVersoesManuais').style.display, 'block');
  s.context.pfResetar();
  assert.equal(emailLatest.checked, true); assert.equal(costManual.checked, true);
  assert.equal(s.el('pfEmailBlocoVersoesManuais').style.display, 'none');
  assert.equal(s.el('blocoVersoesManuais').style.display, 'block');
});

test('fragmento do email não repete IDs nem grupos de radio das outras telas incluídas', () => {
  const markup = text => text.replace(/<script>[\s\S]*?<\/script>/g, '').replace(/<style>[\s\S]*?<\/style>/g, '');
  const values = (text, attr) => new Set(Array.from(markup(text).matchAll(new RegExp(attr + '="([^"]+)"', 'g')), m => m[1]));
  const emailIds = values(html, 'id');
  const emailNames = values(html, 'name');
  for (const file of ['Index.html', 'ValidaInterface.html']) {
    const other = fs.readFileSync(path.join(__dirname, '..', file), 'utf8');
    assert.deepEqual(Array.from(values(other, 'id')).filter(id => emailIds.has(id)), [], file + ' shared IDs');
    assert.deepEqual(Array.from(values(other, 'name')).filter(name => emailNames.has(name)), [], file + ' shared input names');
  }
});

test('contatos comuns aparecem em cada gráfica e seleção individual é preservada por grupo', () => {
  const s = makeScreen(); s.choose(['Teste', 'Outra']); s.context.avancarParaRevisao();
  s.reply('consultarPacoteEmailSPG', complete);
  const shared = s.el('containerDestinatarios').querySelectorAll('input').filter(x => x.value === 'shared@example.com');
  assert.equal(shared.length, 2); shared.find(x => x.dataset.grafica === 'Teste').checked = false;
  s.context.pfRenderDestinatarios();
  const byGroup = s.el('containerDestinatarios').querySelectorAll('input').filter(x => x.value === 'shared@example.com');
  assert.equal(byGroup.find(x => x.dataset.grafica === 'Teste').checked, false);
  assert.equal(byGroup.find(x => x.dataset.grafica === 'Outra').checked, true);
  assert.deepEqual(Array.from(s.context.pfDestinatariosGrupos()).sort(), ['other@example.com, shared@example.com', 'test@example.com']);
});

test('salvamento concluído depois de Cancelar não restaura seleções nem reabre cadastro', () => {
  const s = makeScreen(); s.context.pfAbrirCadastro();
  s.context.pfSalvarCadastro(); s.context.pfResetar();
  s.reply('salvarCadastroEmailsSPG', { graficas: s.graficas, versao: 'v2' });
  assert.equal(s.context.pfSelecionadas().length, 0);
  assert.equal(s.el('pf-modal').hidden, true);
  assert.equal(s.el('inputChamado').value, '');
  assert.equal(s.el('passo1').classList.contains('ativo'), true);
});

test('conflito mantém rascunho e impede sobrescrever cadastro sem atualizar a página', () => {
  const s = makeScreen(); s.context.pfAbrirCadastro();
  s.context.pfRascunhoCadastro.Teste[0] = 'draft@example.com'; s.context.pfSalvarCadastro();
  s.reply('salvarCadastroEmailsSPG', { message: 'O cadastro foi alterado por outra pessoa.' }, true);
  assert.equal(s.context.pfRascunhoCadastro.Teste[0], 'draft@example.com');
  s.context.pfFecharCadastro(); s.context.pfAbrirCadastro();
  assert.equal(s.el('pf-salvar-cadastro').disabled, true);
  assert.match(s.el('pf-erro-cadastro').textContent, /Atualize a página/);
  s.context.pfSalvarCadastro();
  assert.equal(s.requests.filter(r => r.method === 'salvarCadastroEmailsSPG').length, 1);
});

test('lista e popup abrem sem consultar função do servidor; cancelar limpa seleções', () => {
  const s = makeScreen(); s.context.pfAbrirCadastro(); s.context.pfResetar();
  assert.equal(s.requests.length, 0);
  assert.equal(s.el('containerGraficas').querySelectorAll('input').length, 2);
  assert.equal(s.context.pfSelecionadas().length, 0);
  assert.equal(s.el('pf-modal').hidden, true);
  assert.equal(s.context.pfRascunhoCadastro, null);
});

test('novo envio limpa a etapa concluída e permite preparar outro chamado', () => {
  const s = makeScreen(); s.context.avancarParaRevisao(); s.reply('consultarPacoteEmailSPG', complete);
  s.context.enviarEmail();
  const send = s.pending('dispararEmailPacoteSPG'); assert.equal(send.args[0].idChamado, '3100');
  s.reply('dispararEmailPacoteSPG', { sucesso: true });
  assert.equal(s.el('passo3').classList.contains('ativo'), true);
  s.context.enviarNovoEmail();
  assert.equal(s.el('inputChamado').value, ''); assert.equal(s.el('assuntoEmail').value, '');
  assert.equal(s.el('corpoEmail').value, ''); assert.equal(s.context.pfSelecionadas().length, 0);
  assert.equal(s.el('passo1').classList.contains('ativo'), true);
});

test('cadastro permite adicionar gráfica e editar, adicionar e excluir emails no popup', () => {
  const s = makeScreen(); s.context.pfAbrirCadastro();
  s.el('pf-nova-grafica').value = 'Nova'; s.context.pfAdicionarGrafica();
  let card = s.el('pf-lista-cadastro').querySelectorAll('.pf-manage-graphic').find(x => x.textContent.includes('Nova'));
  let input = card.querySelector('input'); input.value = 'new@example.com'; s.dispatch(input, 'input');
  s.dispatch(s.button('+ Adicionar e-mail', card));
  card = s.el('pf-lista-cadastro').querySelectorAll('.pf-manage-graphic').find(x => x.textContent.includes('Nova'));
  input = card.querySelectorAll('input')[1]; input.value = 'second@example.com'; s.dispatch(input, 'input');
  s.dispatch(s.button('Excluir', card));
  assert.equal(s.context.pfRascunhoCadastro.Nova.join(','), 'second@example.com');
  s.context.pfSalvarCadastro();
  assert.equal(s.pending('salvarCadastroEmailsSPG').args[0].Nova.join(','), 'second@example.com');
});
