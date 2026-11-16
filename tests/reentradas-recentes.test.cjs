const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '..', 'ReentradasRecentes.js'), 'utf8');
const plain = value => JSON.parse(JSON.stringify(value));
const uuid = n => '10000000-0000-4000-8000-' + String(n).padStart(12, '0');

function row(numero, extra = {}) {
  return { Chamado: numero, 'ID Chamado': uuid(numero), 'Tipo': 'REENTRADA',
    Status: 'ALOCADO', Marca: 'SAS', Envio: 'V2', 'Alocado Em': '20/10/2026 14:30', ...extra };
}

function without(objeto, ...campos) {
  for (const campo of campos) delete objeto[campo];
  return objeto;
}

function harness(pages, details = {}) {
  const calls = [];
  const context = { _spgLogin_: () => 'token', _emailSpgLerJson_: (route, token) => {
    assert.equal(token, 'token');
    calls.push(route);
    const page = route.match(/^\/api\/reentradas\/sheets\/relatorio\?page=(\d+)&limit=100$/);
    if (page) {
      const response = pages[Number(page[1]) - 1];
      if (response instanceof Error) throw response;
      return response;
    }
    const id = route.match(/^\/api\/reentradas\/([0-9a-f-]+)$/)?.[1];
    if (!id || !details[id]) throw new Error('Rota inesperada: ' + route);
    if (details[id] instanceof Error) throw details[id];
    return { reentrada: details[id] };
  }};
  vm.createContext(context);
  vm.runInContext(source, context);
  return { context, calls, run: () => plain(context.listarUltimasReentradasSPG()) };
}

test('lista recente consulta somente a API, sem arquivos, planilhas ou alteracoes', () => {
  assert.doesNotMatch(source, /\b(?:SpreadsheetApp|DriveApp|GmailApp|MailApp)\b|\/op\b|OPXLSX|\/api\/tiragem\//);
});

test('filtra exclusivamente REENTRADA e limita os 30 maiores numeros mesmo com paginas fora de ordem', () => {
  const rows = Array.from({ length: 45 }, (_, i) => row(1000 + i));
  rows.push(row(2000, { 'Tipo': 'SPOT' }), row(2001, { 'Tipo': 'INCLUSIVOS' }));
  const h = harness([
    { page: 1, totalPages: 2, rows: rows.slice(20).reverse() },
    { page: 2, totalPages: 2, rows: rows.slice(0, 20) }
  ]);
  const result = h.run();
  assert.equal(result.erro, undefined);
  assert.deepEqual(result.chamados.map(r => r.chamado), Array.from({ length: 30 }, (_, i) => String(1044 - i)));
  assert.equal(h.calls.length, 2, 'os campos completos do relatorio dispensam detalhes');
  assert.deepEqual(result.chamados[0], { chamado: '1044', status: 'ALOCADO', marca: 'SAS', envio: 'V2', alocadoEm: '20/10/2026 14:30' });
});

test('devolve menos de 30 e agrega as informacoes das linhas duplicadas do mesmo chamado', () => {
  const h = harness([
    { page: 1, totalPages: 2, rows: [row(30), row(5), row(30)] },
    { page: 2, totalPages: 2, rows: [row(30, { Marca: 'Geekie', Envio: 'V3' })] }
  ]);
  assert.deepEqual(h.run().chamados, [
    { chamado: '30', status: 'ALOCADO', marca: 'SAS, Geekie', envio: 'V2, V3', alocadoEm: '20/10/2026 14:30' },
    { chamado: '5', status: 'ALOCADO', marca: 'SAS', envio: 'V2', alocadoEm: '20/10/2026 14:30' }
  ]);
});

test('relatorio vazio devolve lista vazia tanto com zero quanto com uma pagina', () => {
  for (const totalPages of [0, 1]) {
    const h = harness([{ page: 1, totalPages, rows: [] }]);
    assert.deepEqual(h.run(), { chamados: [] });
    assert.equal(h.calls.length, 1);
  }
});

test('normaliza somente o tipo explicito e nao usa justificativa, descricao ou produtos como filtro', () => {
  const h = harness([{ page: 1, totalPages: 1, rows: [
    row(30, { 'Tipo': ' reentrada ' }),
    row(29, { 'Tipo': 'SPOT', Justificativa: 'REENTRADA Ampliados' }),
    row(28, { 'Tipo': 'INCLUSIVOS', Descricao: 'REENTRADA' }),
    row(27, { 'Tipo': 'REENTRADA ESPECIAL' })
  ] }]);
  assert.deepEqual(h.run().chamados.map(r => r.chamado), ['30']);
  assert.equal(h.calls.length, 1);
});

test('consulta o detalhe apenas para tipo ou metadados ausentes e confirma a identidade', () => {
  const h = harness([{ page: 1, totalPages: 1, rows: [
    without(row(30), 'Tipo', 'Status', 'Marca', 'Envio', 'Alocado Em'),
    without(row(29), 'Marca'), without(row(28, { Tipo: 'SPOT' }), 'Marca'),
    without(row(27), 'Tipo')
  ] }], {
    [uuid(30)]: { id: uuid(30), numero: 30, tipoSolicitacao: 'REENTRADA', status: 'SOLICITADO',
      brand: { name: 'Nave' }, versao: { envio: 'V1' }, alocadoEm: '' },
    [uuid(29)]: { id: uuid(29), numero: 29, tipoSolicitacao: 'REENTRADA', brand: { name: 'Geekie' } },
    [uuid(27)]: { id: uuid(27), numero: 27, tipoSolicitacao: 'INCLUSIVOS',
      items: [{ tipoSolicitacao: 'REENTRADA' }] }
  });
  assert.deepEqual(h.run().chamados, [
    { chamado: '30', status: 'SOLICITADO', marca: 'Nave', envio: 'V1', alocadoEm: '' },
    { chamado: '29', status: 'ALOCADO', marca: 'Geekie', envio: 'V2', alocadoEm: '20/10/2026 14:30' }
  ]);
  assert.equal(h.calls.length, 4);
  assert.ok(!h.calls.includes('/api/reentradas/' + uuid(28)), 'SPOT explicito nunca busca metadados');
});

test('para de consultar detalhes quando confirmar 30 reentradas', () => {
  const rows = Array.from({ length: 36 }, (_, i) => without(row(i + 1), 'Tipo'));
  const details = Object.fromEntries(rows.map(r => [r['ID Chamado'], {
    id: r['ID Chamado'], numero: r.Chamado, tipoSolicitacao: r.Chamado > 33 ? 'SPOT' : 'REENTRADA'
  }]));
  const h = harness([{ page: 1, totalPages: 1, rows }], details);
  assert.deepEqual(h.run().chamados.map(r => r.chamado), Array.from({ length: 30 }, (_, i) => String(33 - i)));
  assert.equal(h.calls.length, 34, 'uma pagina, tres SPOT e trinta reentradas');
  assert.ok(!h.calls.includes('/api/reentradas/' + uuid(3)));
});

test('identidade ou tipo ausente no detalhe nao produz lista parcial nem inclui outro fluxo', () => {
  const base = { id: uuid(30), numero: 30, tipoSolicitacao: 'REENTRADA' };
  for (const detalhe of [null, { ...base, id: uuid(31) }, { ...base, numero: 31 },
    { ...base, tipoSolicitacao: '', versao: { tipoSolicitacao: 'REENTRADA' } }]) {
    const h = harness([{ page: 1, totalPages: 1, rows: [without(row(30), 'Tipo')] }], {
      [uuid(30)]: detalhe || new Error('Detalhe ausente')
    });
    const result = h.run();
    assert.deepEqual(result.chamados, []);
    assert.ok(result.erro);
  }
  const h = harness([{ page: 1, totalPages: 1, rows: [without(row(30, { 'ID Chamado': '' }), 'Tipo')] }]);
  assert.match(h.run().erro, /identificacao valida/);
  assert.equal(h.calls.length, 1);
});

test('colunas presentes e vazias nao consultam detalhe nem usam grafica como data de alocacao', () => {
  const h = harness([{ page: 1, totalPages: 1, rows: [
    row(30, { 'Alocado Em': '', Grafica: 'Logprint' }),
    row(29, { Marca: '', Envio: '', Status: '', 'Alocado Em': null }),
    row(28, { Tipo: '', Grafica: 'Meta' })
  ] }]);
  assert.deepEqual(h.run(), { chamados: [
    { chamado: '30', status: 'ALOCADO', marca: 'SAS', envio: 'V2', alocadoEm: '' },
    { chamado: '29', status: '', marca: '', envio: '', alocadoEm: '' }
  ] });
  assert.equal(h.calls.length, 1);
});

test('respeita todas as linhas reais da pagina mesmo quando ultrapassam limit=100', () => {
  const rows = Array.from({ length: 241 }, (_, i) => row(i < 200 ? 2969 - i % 9 : 4000 + i,
    { Tipo: i < 200 ? 'REENTRADA' : 'SPOT', 'Alocado Em': '' }));
  const h = harness([{ page: 1, totalPages: 1, limit: 100, rows }]);
  assert.equal(h.run().chamados.length, 9);
  assert.equal(h.calls.length, 1);
});

test('falhas de autenticacao e transporte aparecem como erro sem lista parcial', () => {
  const h = harness([{ page: 1, totalPages: 2, rows: [row(30)] }, new Error('SPG indisponivel')]);
  assert.deepEqual(h.run(), { chamados: [], erro: 'SPG indisponivel' });
  h.context._spgLogin_ = () => null;
  assert.match(h.run().erro, /Autenticacao/);
  h.context._spgLogin_ = () => { throw new Error('Falha de login'); };
  assert.deepEqual(h.run(), { chamados: [], erro: 'Falha de login' });
});

test('rejeita paginacao incorreta ou inconsistente', () => {
  for (const data of [
    { page: 1, totalPages: 1001, rows: [] }, { page: 1, totalPages: -1, rows: [] },
    { page: 1, totalPages: 1.5, rows: [] }, { page: 2, totalPages: 2, rows: [] },
    { page: 1, totalPages: 'nao', rows: [] }, { page: 1, totalPages: 1, rows: {} },
    { page: 1, totalPages: null, rows: [] }, { page: 1, totalPages: false, rows: [] },
    { page: 1, totalPages: 0, rows: [row(30)] }, null
  ]) {
    assert.match(harness([data]).run().erro, /relatorio ou paginacao/);
  }
  const h = harness([{ page: 1, totalPages: 2, rows: [row(30)] }, { page: 2, totalPages: 3, rows: [] }]);
  assert.match(h.run().erro, /relatorio ou paginacao/);
});

test('rejeita numeros de chamado ou identidades divergentes no mesmo relatorio', () => {
  for (const invalid of [0, -1, 1.5, '', '1e4', '9999999999999', null]) {
    assert.match(harness([{ page: 1, totalPages: 1, rows: [row(invalid)] }]).run().erro, /numero de chamado/);
  }
  const h = harness([{ page: 1, totalPages: 1, rows: [row(30), row(30, { 'ID Chamado': uuid(31) })] }]);
  assert.match(h.run().erro, /identificacoes diferentes/);
  const types = harness([{ page: 1, totalPages: 1, rows: [row(30), row(30, { 'Tipo': 'SPOT' })] }]);
  assert.match(types.run().erro, /tipos diferentes/);
});
