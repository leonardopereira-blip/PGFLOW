// Run: node --test tests/email-spg.test.cjs
// Nao consulta a API nem envia e-mails: teste isolado das regras de seguranca.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function createContext(zip) {
  const context = {
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' },
      unzip: () => zip,
      computeDigest: () => [1, 2, 3],
      newBlob: (bytes) => ({getBytes:()=>bytes,setName(){return this}})
    },
    UrlFetchApp: {}, PropertiesService: {}, LockService: {},
    CacheService: {}, GmailApp: {}, SpreadsheetApp: {},
    _spgLogin_: ()=>'token',
    obterDadosGraficas: ()=>({}),
    ID_PLANILHA_LOG: 'test'
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.join(__dirname,'..','EmailSPG.js'),'utf8');
  vm.runInContext(source, context);
  return context;
}

const ctx = createContext([]);
test('normaliza data ISO e rejeita data inexistente', () => {
  assert.equal(ctx._emailSpgData_('2026-11-15T00:00:00Z'),'15/11/2026');
  assert.equal(ctx._emailSpgData_('31/02/2026'),'');
});
test('replica regra de coleta (nao sexta, sabado ou domingo)', () => {
  assert.equal(ctx._emailSpgColeta_('15/11/2026'),'12/11/2026');
});
test('busca campos aninhados sem usar equivalencias por aproximacao', () => {
  assert.equal(ctx._emailSpgCampo_({solicitacao:{dataDesejadaEntrega:'2026-11-15'}},
    ['dataDesejadaEntrega']),'2026-11-15');
  assert.equal(ctx._emailSpgCampo_({solicitacao:{dataPrevista:'2026-11-15'}},
    ['dataDesejadaEntrega']),'');
});
test('valida id e bloqueia entradas inesperadas', () => {
  assert.equal(ctx._emailSpgId_(' 3100 '),'3100');
  assert.throws(()=>ctx._emailSpgId_('../3100'));
});
test('destinatarios duplicados nao causam multiplos envios', () => {
  assert.equal(ctx._emailSpgDestinos_(['a@example.com, A@EXAMPLE.COM','b@example.com']).length,2);
  assert.throws(()=>ctx._emailSpgDestinos_(['invalido']));
});
function zip(terceiraValida) {
  const xml = valido => '<worksheet><sheetData><row r="1"><c><t>header</t></c></row>' +
    (valido?'<row r="2"><c><v>3</v></c></row>':'<row r="2"><c><v></v></c></row>') +
    '</sheetData></worksheet>';
  return [
    ['xl/workbook.xml','<workbook><sheets><sheet name="Tiragem" r:id="rId1"/>' +
      '<sheet name="Árvore de produto" r:id="rId2"/><sheet name="Especificação Técnica" r:id="rId3"/></sheets></workbook>'],
    ['xl/_rels/workbook.xml.rels','<Relationships>' +
      '<Relationship Id="rId1" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Target="worksheets/sheet2.xml"/>' +
      '<Relationship Id="rId3" Target="worksheets/sheet3.xml"/></Relationships>'],
    ['xl/worksheets/sheet1.xml',xml(true)],
    ['xl/worksheets/sheet2.xml',xml(true)],
    ['xl/worksheets/sheet3.xml',xml(terceiraValida)]
  ].map(([name,content])=>({getName:()=>name,getDataAsString:()=>content}));
}
test('excel completo com dados nas tres abas passa', () => {
  const c = createContext(zip(true));
  assert.equal(c._emailSpgChecarXlsx_({getBytes:()=>[80,75,...Array(110).fill(0)]}),true);
});
test('excel com aba somente de cabecalho e bloqueado', () => {
  const c = createContext(zip(false));
  assert.throws(()=>c._emailSpgChecarXlsx_({getBytes:()=>[80,75,...Array(110).fill(0)]}),
    /especificacao/);
});
test('dados opcionais so podem ser omitidos por escolha explicita', () => {
  const p = {dataDesejada:'',descricao:''};
  assert.throws(()=>ctx._emailSpgSobrescrita_(p,{sobrescrita:{}}));
  const ok=ctx._emailSpgSobrescrita_(p,{sobrescrita:{omitirData:true,omitirDescricao:true}});
  assert.equal(ok.data,'');
  assert.equal(ok.descricao,'');
});
