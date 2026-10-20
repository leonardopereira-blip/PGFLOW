const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');

const SKU = 'MS26VF331P1SP';
const SKU_HEADER = '1. ID_Código SKU';
const REF_CHAPA = '11. REF_Referência troca de chapa';
const REF_PERS = '11. REF_Referência personalização';
const plain = value => JSON.parse(JSON.stringify(value));

function product(sku = SKU, extra = {}) {
  return { skuAtual: sku, descricao: 'Produto SPG', marca: 'Marca',
    capBaseMmAberto: 410.9, capAlturaMmAberto: 275.3, baseMm: 205.95,
    alturaMm: 275.8, espessuraMm: 2.78, pesoKg: 0.234,
    capTemCorteVinco: ' TrUe ', miolos: [{ corteVinco: 'FaLsE', serrilha: true }], ...extra };
}

function page(products, current = 1, total = products.length, limit = 500) {
  return { products, pagination: { total, page: current, limit, totalPages: Math.ceil(total / limit) } };
}

function harness(responses = [page([product()])]) {
  let cursor = 0, logins = 0;
  const requests = [], batches = [];
  function fetch(url, options) {
    requests.push({ url, options });
    const response = responses[cursor++];
    if (response instanceof Error) throw response;
    if (response === undefined) throw new Error('Unexpected request');
    return { getResponseCode: () => response.status ?? 200,
      getContentText: () => typeof response.body === 'string' ? response.body : JSON.stringify(response) };
  }
  const ctx = {
    Utilities: { DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
      computeDigest: (algorithm, value, encoding) => [...crypto.createHash(algorithm).update(value, encoding).digest()] },
    UrlFetchApp: { fetch, fetchAll(options) {
      batches.push(options.map(r => r.url));
      return options.map(r => fetch(r.url, r));
    } }
  };
  vm.createContext(ctx);
  for (const file of ['SPG.js', 'ReentradasEspecificacaoGeral.js']) {
    vm.runInContext(fs.readFileSync(path.resolve(__dirname, '..', file), 'utf8'), ctx);
  }
  ctx._spgLogin_ = () => { logins++; return 'test-token'; };
  return { ctx, requests, batches, get logins() { return logins; } };
}

function matrices(skus = [SKU]) {
  return {
    espec: { headers: ['ID (Coluna A)', SKU_HEADER, '1. ID_Descrição'], linhas: [] },
    tiragem: { headers: [SKU_HEADER, REF_CHAPA, REF_PERS, '16. TIR_SKU'],
      linhas: skus.map((sku, i) => [sku, 'CHAPA-' + i, 'PERSON-' + i, 200]) }
  };
}

function run(h, data = matrices(), authorizations = [{ chamado: '2632', sku: SKU }], chamado = '2632') {
  return h.ctx._reSpCompletarEspecLegado_(chamado, data.espec, data.tiragem, authorizations);
}

test('nao consulta API quando todas as especificacoes ja existem e preserva a matriz original', () => {
  const h = harness([]), data = matrices();
  data.espec.linhas = [[2632, SKU, 'Original']];
  const actual = run(h, data, []);
  assert.equal(actual.espec, data.espec);
  assert.deepEqual(plain(actual.skusEspecificacaoGeral), []);
  assert.equal(h.requests.length, 0);
  assert.equal(h.logins, 0);
});

test('aviso tipado exige escolha explicita por chamado e SKU antes de consultar a API', () => {
  for (const authorization of [[], ['2632'], [{ chamado: '2631', sku: SKU }],
    [{ chamado: '2632', sku: 'OUTRO-SKU' }], true, { chamado: '2632', sku: SKU }]) {
    const h = harness([]);
    assert.throws(() => run(h, matrices(), authorization), error => {
      assert.equal(error.codigo, 'ESPECIFICACAO_LEGADO_AUSENTE');
      assert.deepEqual(plain(error.especificacoesAusentes), [SKU]);
      return true;
    });
    assert.equal(h.requests.length, 0);
    assert.equal(h.logins, 0);
  }
});

test('a escolha de um SKU nao autoriza outro ausente no mesmo chamado', () => {
  const h = harness([]), data = matrices([SKU, 'SKU-2', SKU]);
  assert.throws(() => run(h, data), error => {
    assert.deepEqual(plain(error.especificacoesAusentes), ['SKU-2']);
    return true;
  });
  assert.equal(h.logins, 0);
});

test('usa o mapeamento SPG para dimensoes, booleanos e duas referencias da mesma tiragem', () => {
  const h = harness(), actual = run(h), row = actual.espec.linhas[0];
  const value = header => row[actual.espec.headers.indexOf(header)];
  assert.equal(value(SKU_HEADER), SKU);
  assert.equal(value('3,DIM_BxA(Aberto)(mm)'), '410X275');
  assert.equal(value('3.DIM_BXA(Fechado)'), '205X275');
  assert.equal(value('3. DIM_E (mm)'), 2.7);
  assert.equal(value('3.DIM_Dimensãoacabada(BxAxE)'), '205X275X2,7');
  assert.equal(value('3. DIM_Peso (Kg)'), 0.234);
  assert.equal(value('4. CAP_Corte e vinco'), 'Sim');
  assert.equal(value('5. MIO_PRIN_Corte e vinco'), 'Não');
  assert.equal(value('5. MIO_PRIN_Serrilha'), 'Sim');
  assert.equal(value(REF_CHAPA), 'CHAPA-0');
  assert.equal(value(REF_PERS), 'PERSON-0');
  assert.deepEqual(plain(actual.skusEspecificacaoGeral), [SKU]);
  assert.equal(actual.avisos.length, 1);
  assert.match(actual.avisos[0], /alocacao antiga/);
  assert.equal(h.requests.length, 1);
  assert.match(h.requests[0].url, /products\/spec-report\?page=1&limit=1000$/);
});

test('preserva dados e metadados legados, completa somente lacunas e nao modifica as entradas', () => {
  const h = harness([page([product('ANTIGO', { descricao: 'Nao usar' }), product()])]);
  const data = matrices(['ANTIGO', SKU]);
  data.espec.headers.push('Metadado externo', '3. DIM_BxA (Fechado)');
  data.espec.linhas = [[2632, 'ANTIGO', 'Descricao original', false, '123X456']];
  const before = plain(data), actual = run(h, data);
  assert.deepEqual(plain(data), before);
  assert.equal(actual.espec.linhas.length, 2);
  assert.deepEqual(plain(actual.espec.linhas[0].slice(0, 5)), before.espec.linhas[0]);
  assert.equal(actual.espec.linhas[1][3], '');
  assert.equal(actual.espec.linhas[1][4], '205X275');
  assert.equal(actual.espec.headers.filter(header => /Fechado/.test(header) && /BxA/i.test(header)).length, 1);
  const ref = actual.espec.headers.indexOf(REF_CHAPA);
  assert.equal(actual.espec.linhas[1][ref], 'CHAPA-1');
});

test('le todas as paginas declaradas e encontra o SKU mesmo depois da primeira', () => {
  const h = harness([page([product('OUTRO')], 1, 2, 1), page([product()], 2, 2, 1)]);
  assert.deepEqual(plain(run(h).skusEspecificacaoGeral), [SKU]);
  assert.equal(h.requests.length, 2);
  assert.match(h.requests[1].url, /page=2&limit=1000$/);
});

test('consulta paginas restantes em lotes de no maximo quatro sem interromper na primeira correspondencia', () => {
  const pages = Array.from({ length: 10 }, (_, i) => page([product(i ? 'SKU-' + i : SKU)], i + 1, 10, 1));
  const h = harness(pages);
  assert.equal(run(h).espec.linhas.length, 1);
  assert.deepEqual(h.batches.map(batch => batch.length), [4, 4, 1]);
  assert.equal(h.requests.length, 10);
});

test('preserva a escala de peso decidida pelo primeiro peso do catalogo inteiro, nao pelo SKU selecionado', () => {
  const h = harness([
    page([product('OUTRO', { pesoKg: 23400 })], 1, 2, 1),
    page([product(SKU, { pesoKg: 1200 })], 2, 2, 1)
  ]);
  const actual = run(h), index = actual.espec.headers.indexOf('3. DIM_Peso (Kg)');
  assert.equal(actual.espec.linhas[0][index], 0.012);
  const firstBlank = harness([
    page([product('SEM-PESO', { pesoKg: null })], 1, 3, 1),
    page([product('OUTRO', { pesoKg: 0.5 })], 2, 3, 1),
    page([product(SKU, { pesoKg: 1200 })], 3, 3, 1)
  ]);
  assert.equal(run(firstBlank).espec.linhas[0][index], 1200);
});

test('nao transforma falhas HTTP, de rede, JSON ou paginas parciais em ausencia de SKU', () => {
  const cases = [
    [{ status: 503, body: 'indisponivel' }], [new Error('secret-network-detail')],
    [{ body: 'not-json' }], [{ products: [] }],
    [page([product('OUTRO')], 1, 2, 1), { status: 401, body: 'secret-token' }],
    [page([product('OUTRO')], 1, 2, 1), page([], 2, 2, 1)],
    [page([product('OUTRO')], 1, 2, 1), page([product()], 2, 3, 1)],
    [page([product('OUTRO')], 1, 2, 1), page([product()], 1, 2, 1)]
  ];
  for (const responses of cases) {
    assert.throws(() => run(harness(responses)), error => {
      assert.doesNotMatch(error.message, /secret|nao encontrada/i);
      return true;
    });
  }
});

test('paginacao invalida nunca retorna um resultado parcial', () => {
  for (const mutate of [
    json => { json.pagination.totalPages = '1'; }, json => { json.pagination.total = -1; },
    json => { json.pagination.page = 2; }, json => { json.pagination.limit = 0; },
    json => { json.pagination.totalPages = 1001; }, json => { json.pagination.totalPages = 2; },
    json => { json.products = [null]; }, json => { json.products = 'invalid'; }
  ]) {
    const response = page([product()]); mutate(response);
    assert.throws(() => run(harness([response])), /invalida/);
  }
});

test('catalogo vazio e SKU ausente geram mensagem de cadastro acionavel', () => {
  for (const response of [page([]), page([product('OUTRO')])]) {
    assert.throws(() => run(harness([response])), new RegExp('nao encontrada.*' + SKU));
  }
});

test('duplicatas identicas sao aceitas mas especificacoes divergentes impedem escolha arbitraria', () => {
  const duplicate = harness([page([product(), product()])]);
  assert.equal(run(duplicate).espec.linhas.length, 1);
  const conflicting = harness([page([product(), product(SKU, { baseMm: 210 })])]);
  assert.throws(() => run(conflicting), /especificacoes gerais diferentes/);
});

test('nao aceita paginas repetidas com metadados novos nem expande autorizacao a SKU parecido', () => {
  const repeated = harness([page([product()], 1, 2, 1), page([product()], 2, 2, 1)]);
  assert.throws(() => run(repeated), /repetiu uma pagina/);
  assert.throws(() => run(harness([page([product(SKU + '-EXTRA')])])), /nao encontrada/);
});

test('falha de autenticacao nao expoe mensagem privada nem consulta produtos', () => {
  for (const login of [() => '', () => { throw new Error('secret-password'); }]) {
    const h = harness([]); h.ctx._spgLogin_ = login;
    assert.throws(() => run(h), error => {
      assert.match(error.message, /autenticar/);
      assert.doesNotMatch(error.message, /secret/);
      return true;
    });
    assert.equal(h.requests.length, 0);
  }
});

test('pre-carrega os SKUs escolhidos no lote e reutiliza uma unica leitura na mesma execucao', () => {
  const h = harness([page([product(), product('SKU-B', { baseMm: 210 })])]);
  const choices = [{ chamado: '2632', sku: SKU }, { chamado: '2633', sku: 'SKU-B' }];
  const first = run(h, matrices(), choices);
  const second = run(h, matrices(['SKU-B']), choices, '2633');
  assert.deepEqual(plain(first.skusEspecificacaoGeral), [SKU]);
  assert.deepEqual(plain(second.skusEspecificacaoGeral), ['SKU-B']);
  assert.equal(second.espec.linhas[0][second.espec.headers.indexOf('3.DIM_BXA(Fechado)')], '210X275');
  assert.equal(h.requests.length, 1);
  assert.equal(h.logins, 1);
  assert.throws(() => run(h, matrices(['SKU-B']), choices, '2632'), error => {
    assert.equal(error.codigo, 'ESPECIFICACAO_LEGADO_AUSENTE');
    return true;
  });
  assert.equal(h.requests.length, 1);
});

test('ausencia confirmada tambem fica em memoria sem transformar aceite de outro chamado em permissao', () => {
  const h = harness([page([product()])]);
  const choices = [{ chamado: '2632', sku: SKU }, { chamado: '2633', sku: 'SKU-AUSENTE' }];
  run(h, matrices(), choices);
  for (let i = 0; i < 2; i++) {
    assert.throws(() => run(h, matrices(['SKU-AUSENTE']), choices, '2633'), /nao encontrada.*SKU-AUSENTE/);
  }
  assert.equal(h.requests.length, 1);
  assert.equal(h.logins, 1);
});

test('falha de pagina nao grava produtos nem ausencias em memoria e a proxima consulta tenta novamente', () => {
  const h = harness([
    page([product()], 1, 2, 1), { status: 503, body: 'indisponivel' },
    page([product()])
  ]);
  assert.throws(() => run(h), /HTTP 503/);
  assert.equal(h.ctx._reEspecGeralCacheExecucao_, null);
  assert.equal(run(h).espec.linhas.length, 1);
  assert.equal(h.requests.length, 3);
  assert.equal(h.logins, 2);
});

test('SKU adicional nao previsto atualiza a memoria e a escala de peso como uma unica leitura valida', () => {
  const h = harness([
    page([product(SKU, { pesoKg: 23400 })]),
    page([product(SKU, { pesoKg: 0.2 }), product('SKU-B', { pesoKg: 0.5 })])
  ]);
  const first = run(h), weight = first.espec.headers.indexOf('3. DIM_Peso (Kg)');
  assert.equal(first.espec.linhas[0][weight], 0.234);
  const second = run(h, matrices(['SKU-B']), [{ chamado: '2633', sku: 'SKU-B' }], '2633');
  assert.equal(second.espec.linhas[0][weight], 0.5);
  assert.equal(run(h).espec.linhas[0][weight], 0.2);
  assert.equal(h.requests.length, 2);
  assert.equal(h.logins, 2);
});

test('somente escolhas com chamado e SKU validos participam do pre-carregamento', () => {
  const h = harness();
  run(h, matrices(), [{ chamado: '2632', sku: SKU },
    { chamado: 'nao-numero', sku: 'INVALIDO' }, { chamado: '2633', sku: '-' },
    { chamado: '2633', sku: '' }, { chamado: '2633', sku: true },
    { chamado: '2633', sku: 'controle\n' }, { chamado: {}, sku: 'INVALIDO' }]);
  assert.deepEqual(Object.keys(h.ctx._reEspecGeralCacheExecucao_.porSku), [SKU]);
  const separateExecution = harness();
  run(separateExecution);
  assert.equal(separateExecution.requests.length, 1);
  assert.equal(separateExecution.logins, 1);
});

test('completar especificacoes gerais nao depende de servicos de planilha ou Drive', () => {
  const source = fs.readFileSync(path.resolve(__dirname, '..', 'ReentradasEspecificacaoGeral.js'), 'utf8');
  assert.doesNotMatch(source, /SpreadsheetApp|DriveApp|Sheets\.|CacheService|PropertiesService/);
  const h = harness();
  assert.equal(run(h).espec.linhas.length, 1);
});
