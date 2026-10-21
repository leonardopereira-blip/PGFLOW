const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const contexto = {};
vm.createContext(contexto);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'ReentradasElegibilidade.js'), 'utf8'), contexto);
const separar = (espec, tiragem) => JSON.parse(JSON.stringify(contexto._reSpSkusNaoCalculaveis_(espec, tiragem)));
const matriz = (headers, linhas) => ({ headers, linhas });
const sku = '1. ID_Código SKU';

test('rotulos SPG excluem somente os SKUs nao graficos ou nao editoriais presentes na tiragem', () => {
  const espec = matriz([sku, 'Grafico', 'Editorial'], [
    ['LIVRO', 'Grafico', 'Editorial'],
    ['NAO-GRAFICO', 'Não Grafico', 'Editorial'],
    ['NAO-EDITORIAL', 'Grafico', 'Não Editorial'],
    ['AMBOS', 'Não Grafico', 'Não Editorial'],
    ['FORA-DA-OP', 'Não Grafico', 'Não Editorial']
  ]);
  const tiragem = matriz([sku], [['NAO-EDITORIAL'], ['LIVRO'], ['AMBOS'], ['NAO-GRAFICO']]);
  assert.deepEqual(separar(espec, tiragem), [
    { sku: 'NAO-EDITORIAL', motivos: ['Não editorial'] },
    { sku: 'AMBOS', motivos: ['Não gráfico', 'Não editorial'] },
    { sku: 'NAO-GRAFICO', motivos: ['Não gráfico'] }
  ]);
});

test('legado e especificacao geral: negativos dedicados aceitam caixa acentos hifens e booleanos explicitos', () => {
  const negativos = [false, ' FALSE ', ' Não ', 'nAo', 'não-grÁfico', 'NÃO GRÁFICO', 'nao–grafico'];
  for (const negativo of negativos) {
    const espec = matriz([sku, 'Gráfico'], [['200713648', negativo]]);
    assert.deepEqual(separar(espec, matriz([sku], [[200713648]])), [
      { sku: '200713648', motivos: ['Não gráfico'] }
    ], String(negativo));
  }
  const tiragem = matriz([sku, 'Editorial'], [['MS26VF331P1SP', ' NÃO—EDITORIAL ']]);
  assert.deepEqual(separar(matriz([sku], [['MS26VF331P1SP']]), tiragem), [
    { sku: 'MS26VF331P1SP', motivos: ['Não editorial'] }
  ]);
});

test('ausencia, numeros e textos livres nao permitem inferir inelegibilidade', () => {
  const desconhecidos = [undefined, null, '', '-', 0, '0', 1, '1', true, 'true', 'Sim',
    'Grafico', 'Editorial', 'Não Gráfico (digital)', 'Não gráfico?', 'indefinido'];
  const espec = matriz([sku, 'Grafico', 'Editorial', '2. INF_Classificação do produto', '1. ID_Descrição'],
    desconhecidos.map((valor, i) => ['SKU-' + i, valor, valor, 'Não Editorial', 'Não Grafico']));
  const tiragem = matriz([sku], espec.linhas.map(linha => [linha[0]]));
  assert.deepEqual(separar(espec, tiragem), []);
  assert.deepEqual(separar(matriz([sku, '2. INF_Classificação do produto'], [['SEM-FLAGS', 'Não Editorial']]),
    matriz([sku], [['SEM-FLAGS']])), []);
  assert.deepEqual(separar(undefined, matriz([sku], [['SEM-ESPEC']])), []);
});

test('rotulo negativo de outro campo nao e interpretado na coluna errada', () => {
  assert.deepEqual(separar(matriz([sku, 'Grafico', 'Editorial'], [
    ['INVERTIDO', 'Não Editorial', 'Não Gráfico']
  ]), matriz([sku], [['INVERTIDO']])), []);
});

test('consolida motivos das fontes sem duplicar SKU e preserva codigos e matrizes originais', () => {
  const espec = matriz([sku, 'Grafico', 'Editorial'], [
    ['000123', false, 'Editorial'], ['MS26VF331P1SP', 'Não Grafico', 'Editorial']
  ]);
  const tiragem = matriz([sku, 'Grafico', 'Editorial'], [
    [' MS26VF331P1SP ', 'Grafico', false], ['000123', 'Não Grafico', 'Editorial'],
    ['MS26VF331P1SP', 'Grafico', 'Não Editorial'], [123, 'Grafico', 'Editorial']
  ]);
  const copia = JSON.stringify([espec, tiragem]);
  assert.deepEqual(separar(espec, tiragem), [
    { sku: 'MS26VF331P1SP', motivos: ['Não gráfico', 'Não editorial'] },
    { sku: '000123', motivos: ['Não gráfico'] }
  ]);
  assert.equal(JSON.stringify([espec, tiragem]), copia);
});

test('quando todos sao nao calculaveis devolve cada motivo para o fluxo manter apenas o tecnico', () => {
  assert.deepEqual(separar(matriz([sku, 'Editorial'], [['A', 'Não Editorial'], ['B', false]]),
    matriz([sku], [['A'], ['B']])), [
    { sku: 'A', motivos: ['Não editorial'] },
    { sku: 'B', motivos: ['Não editorial'] }
  ]);
});
