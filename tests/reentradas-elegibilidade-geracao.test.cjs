const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const mx = require('../tools/molde-node.cjs');

const SKU = '1. ID_Código SKU';
const KIT = '1. ID_Código Kit';
const CODIGOS = ['200713648', 'MS26VF331P1SP', 'SKU-NAO-GRAFICO'];
const sha = texto => crypto.createHash('sha256').update(texto).digest('hex');
const simples = valor => JSON.parse(JSON.stringify(valor));

// Captura os blobs enviados ao empacotador sem comprimir o molde grande nem
// criar arquivos/consultar o Drive, como a verificacao de geracao por grafica.
function blobLeve(dados, tipo, nome) {
  return {
    getName: () => nome, getContentType: () => tipo,
    getBytes: () => Buffer.isBuffer(dados) ? dados : Buffer.from(dados, 'utf8'),
    getDataAsString: () => Buffer.isBuffer(dados) ? dados.toString('utf8') : dados,
    copyBlob: () => blobLeve(dados, tipo, nome),
    setContentType(novoTipo) { tipo = novoTipo; return this; }
  };
}

function ambiente(todosExcluidos) {
  const criados = [], pastas = [], consultas = [];
  const app = mx.carregarApps(['MoldeXlsx.js', 'MoldeLogprintDepara.js', 'MoldeReprosetCalc.js',
    'ValidaReentradasSPG.js', 'ReentradasElegibilidade.js'], {
    DriveApp: {
      getFileById() { throw Error('Nao pode ler dados/modelos externos'); },
      getFolderById(id) {
        pastas.push(id);
        return { createFile(blob) {
          const arquivo = { blob, getUrl: () => 'arquivo:' + blob.getName(), getId: () => 'id:' + blob.getName(),
            setName(nome) { assert.equal(nome, blob.getName()); return arquivo; } };
          criados.push(arquivo); return arquivo;
        } };
      }
    },
    SpreadsheetApp: new Proxy({}, { get() { throw Error('Nao pode ler Google Sheets'); } }),
    UrlFetchApp: new Proxy({}, { get() { throw Error('Nao pode consultar fontes externas'); } })
  });
  app.Utilities.newBlob = blobLeve;
  app.Utilities.zip = (blobs, nome) => {
    const mapa = Object.fromEntries(blobs.map(blob => [blob.getName(), blob]));
    assert.equal(Object.keys(mapa).length, blobs.length);
    return { mapa, getName: () => nome, setContentType(tipo) {
      assert.equal(tipo, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      return this;
    } };
  };
  const molde = app._reSpMolde_('REPROSET', 'depara');
  const cab = nome => app._reSpMatriz_(molde, nome).headers;
  // Os indicadores ficam alem das 80/20 colunas do molde: o filtro precisa
  // consultar a origem completa, antes que o mapeamento remova essas colunas.
  const espH = [...cab('PPG - Especificações'), 'Grafico', 'Editorial'];
  const arvH = cab('PPG - Arvore de Produto');
  const tirH = [...cab('PPG - Tiragem'), 'Gráfico', 'Editorial'];
  const row = (headers, valores) => headers.map(h => valores[h] === undefined ? '' : valores[h]);
  const arvore = [], tiragem = [], especificacoes = [];
  CODIGOS.forEach((sku, i) => {
    especificacoes.push(row(espH, {
      [SKU]: sku, '1. ID_Descrição': 'DESCRICAO-' + sku,
      Grafico: 'Grafico', Editorial: i === 1 || todosExcluidos ? 'Não Editorial' : 'Editorial'
    }));
    arvore.push(row(arvH, { [KIT]: sku, '13. MAN_Total de itens colecionados': 1 }));
    arvore.push(row(arvH, { [KIT]: sku, [SKU]: sku }));
    tiragem.push(row(tirH, { [KIT]: sku, [SKU]: sku,
      '15. DES_CD envio': 'CITA', '16. TIR_SKU': [125, 5000, 9000][i],
      'Gráfico': i === 2 ? 'Não Gráfico' : 'Grafico', Editorial: 'Editorial' }));
  });
  const dados = { pacote: { id: '2932', versaoId: 'fixture-eligibilidade', origem: 'spg' },
    espec: { headers: espH, linhas: especificacoes }, arvore: { headers: arvH, linhas: arvore },
    tiragem: { headers: tirH, linhas: tiragem }, avisos: ['Aviso da origem preservado.'] };
  const antes = JSON.stringify(dados);
  app._reSpPrepararDados_ = (numero, grafica) => { consultas.push([numero, grafica]); return dados; };
  return { app, criados, pastas, consultas, dados, antes };
}

function ler(app, arquivo, nome) {
  return arquivo.mapa[app._reSpAbaPorNome_(arquivo, nome)].getDataAsString();
}

function linhas(app, arquivo, nome, primeiraLinha) {
  const caminho = app._reSpAbaPorNome_(arquivo, nome);
  const numero = +/sheet(\d+)\.xml$/.exec(caminho)[1];
  const strings = app._reSpStrings_(arquivo.mapa['xl/sharedStrings.xml'].getDataAsString());
  return app._reSpLerAba_(arquivo.mapa, numero, strings).filter(row => row.numero >= primeiraLinha);
}

function conferirTecnicoCompleto(app, depara) {
  for (const [aba, coluna] of [['PPG - Especificações', SKU], ['PPG - Tiragem', SKU]]) {
    const matriz = app._reSpMatriz_(depara, aba);
    assert.equal(matriz.linhas.length, 3, aba);
    assert.deepEqual(Array.from(matriz.linhas, row => String(app._reSpBuscar_(row, matriz.headers, coluna))), CODIGOS);
  }
  const arvore = app._reSpMatriz_(depara, 'PPG - Arvore de Produto');
  assert.equal(arvore.linhas.length, 6, 'estrutura completa da arvore foi alterada');
  const tiragem = app._reSpMatriz_(depara, 'PPG - Tiragem');
  assert.equal(tiragem.linhas.reduce((total, row) => total + app._reSpBuscar_(row, tiragem.headers, '16. TIR_SKU'), 0), 14125);
  const ncEspec = linhas(app, depara, 'NC - PPG Especificações', 6);
  assert.equal(ncEspec.length, 3);
  assert.deepEqual(Array.from(ncEspec, row => String(row.cells[2])), CODIGOS);
}

test('geracao conserva tres especificacoes tecnicas e calcula apenas o SKU grafico editorial', () => {
  const { app, criados, pastas, consultas, dados, antes } = ambiente(false);
  try {
    const resultado = app._reSpGerarUm_('2932', 'REPROSET', [], { tecnologia: 'DIGITAL' });
    assert.deepEqual(consultas, [['2932', 'REPROSET']]);
    assert.equal(criados.length, 2);
    assert.deepEqual(pastas, [app.RE_SP_CONFIG.pastaDestino, app.RE_SP_CONFIG.pastaDestino]);
    assert.equal(resultado.skus, 1);
    assert.equal(resultado.tiragem, 125, 'tiragens dos produtos nao calculaveis contaminaram o resultado');
    assert.deepEqual(simples(resultado.skusSomenteEspecificacao), [
      { sku: CODIGOS[1], motivos: ['Não editorial'] },
      { sku: CODIGOS[2], motivos: ['Não gráfico'] }
    ]);
    assert.equal(resultado.avisos[0], 'Aviso da origem preservado.');
    assert.equal(resultado.avisos.length, 3);
    assert.match(resultado.avisos.join(' '), /excluido dos custos/);
    assert.deepEqual(simples(resultado.skusIgnorados), []);
    assert.equal(JSON.stringify(dados), antes, 'filtro alterou as fontes tecnicas');
    const depara = criados.find(f => /^PLANILHA DE PARA/.test(f.blob.getName())).blob;
    const calc = criados.find(f => /^CALCULADORA/.test(f.blob.getName())).blob;
    assert.equal(resultado.depara, 'arquivo:' + depara.getName());
    assert.equal(resultado.calculadora, 'arquivo:' + calc.getName());
    assert.match(calc.getName(), /\.xlsx$/);
    assert.equal(calc.mapa['xl/vbaProject.bin'], undefined);
    conferirTecnicoCompleto(app, depara);
    const ncGeral = linhas(app, depara, 'NC - CALC Geral', 4);
    assert.equal(ncGeral.length, 1, 'NC Geral preservou linhas antigas ou produtos nao calculaveis');
    assert.equal(ncGeral[0].cells[1], 200713648);
    assert.equal(ncGeral[0].cells[2], 125);
    assert.equal(ncGeral[0].cells[11], 200713648);
    assert.equal(ncGeral[0].cells[13], 125);
    assert.equal(ncGeral[0].cells[18], 125);
    const ppg = linhas(app, calc, 'PPG Especificações', 5);
    assert.equal(ppg.length, 1);
    assert.equal(ppg[0].cells[2], 200713648, 'chave da consulta deve continuar numerica');
    const geral = linhas(app, calc, 'CALC Geral', 6);
    assert.equal(geral.length, 1, 'CALC Geral preservou SKUs/formulas alem do ultimo calculavel');
    assert.equal(geral[0].cells[3], 200713648);
    assert.equal(geral[0].cells[5], 125);
    for (const coluna of [6, 7, 8]) assert.equal(geral[0].cells[coluna], 0);
    const modelo = app._reSpMolde_('REPROSET', 'calc');
    for (const nome of ['CALC Digital', 'CALC Plana', 'CALC Rotativa', 'CALC Acabamento',
      'TAB Regras', 'TAB Papel', 'TAB Impressão', 'TAB Acabamento']) {
      const original = modelo.mapa[app._reSpAbaPorNome_(modelo, nome)].getDataAsString();
      assert.equal(sha(ler(app, calc, nome)), sha(original), nome + ': formulas ou tarifas foram alteradas');
    }
  } finally { criados.length = 0; app._reSpMoldesCache_ = {}; }
});

test('todos nao calculaveis geram somente o DePara tecnico completo sem calculadora ou custos antigos', () => {
  const { app, criados, pastas, dados, antes } = ambiente(true);
  try {
    const resultado = app._reSpGerarUm_('2932', 'REPROSET', [], {});
    assert.equal(criados.length, 1);
    assert.deepEqual(pastas, [app.RE_SP_CONFIG.pastaDestino]);
    assert.equal(resultado.somenteEspecificacao, true);
    assert.equal(resultado.skus, 0);
    assert.equal(resultado.tiragem, 0);
    assert.equal(resultado.calculadora, '');
    assert.equal(resultado.skusSomenteEspecificacao.length, 3);
    assert.equal(resultado.avisos.length, 4);
    assert.equal(JSON.stringify(dados), antes);
    const depara = criados[0].blob;
    assert.match(depara.getName(), /^PLANILHA DE PARA.*\.xlsx$/);
    assert.equal(resultado.arquivo, depara.getName());
    assert.equal(resultado.depara, 'arquivo:' + depara.getName());
    conferirTecnicoCompleto(app, depara);
    assert.equal(linhas(app, depara, 'NC - CALC Geral', 4).length, 0,
      'DePara sem calculaveis manteve linhas/formulas de custo antigas do molde');
  } finally { criados.length = 0; app._reSpMoldesCache_ = {}; }
});
