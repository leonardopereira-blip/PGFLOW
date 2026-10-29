const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const root = path.join(__dirname, '..');
const principal = fs.readFileSync(path.join(root, 'ValidaPrincipal.js'), 'utf8');
const index = fs.readFileSync(path.join(root, 'Index.html'), 'utf8');

test('Principal cria um novo De-Para mesmo com nome repetido sem apagar o anterior', () => {
  const fluxo = principal.slice(principal.indexOf('function processarValidaPrincipal('));
  assert.match(fluxo, /makeCopy\(nomeArquivo, pastaDest\)/);
  assert.doesNotMatch(fluxo, /removerHomonimos\s*\(|setTrashed\s*\(/);
  assert.match(fluxo, /depara:\s*arquivoCopia\.getUrl\(\)/);
});

test('Filtros e geracao em lote da Principal permanecem em uso', () => {
  assert.match(principal, /payload\.skus/);
  assert.match(principal, /payload\.graficas/);
  assert.match(principal, /for\s*\(var graf in lotesPorGrafica\)/);
  assert.match(principal, /gravarNaFila\(metadadosPacote\)/);
  assert.match(principal, /MailApp\.sendEmail\(EMAIL_FILA_GMAIL/);
});

test('Resumo da Principal tem renderizador isolado e link do De-Para', () => {
  const start = index.indexOf('function renderResumoPrincipal(');
  const end = index.indexOf('function fmtTime(s)', start);
  assert.ok(start >= 0 && end > start);
  const functionText = index.slice(start, end);
  const host = {innerHTML:'', style:{display:'none'}};
  const render = new Function('document', functionText + '; return renderResumoPrincipal;')({
    getElementById:()=>host
  });
  render([{grafica:'Logprint',marcas:'SAS',series:'1ª série',skus:2,tiragem:100,
    depara:'https://docs.google.com/spreadsheets/d/id/edit',arquivo:'CALC.xlsx'}], 'resumoGeradoPrin');
  assert.match(host.innerHTML, /Abrir De-Para/);
  assert.match(host.innerHTML, /CALC\.xlsx/);
  assert.equal(host.style.display, 'block');
  assert.match(index, /renderResumoPrincipal\(_resumoPrin, 'resumoGeradoPrin'\)/);
  assert.doesNotMatch(index, /renderResumo\(_resumoPrin, 'resumoGeradoPrin'\)/);
});