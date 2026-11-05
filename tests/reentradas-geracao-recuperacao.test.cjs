const test=require('node:test');
const assert=require('node:assert/strict');
const mx=require('../tools/molde-node.cjs');

function ambiente(arquivos){
  const app=mx.carregarApps(['ValidaReentradasSPG.js']);
  app.DriveApp={getFolderById:id=>{
    assert.equal(id,app.RE_SP_CONFIG.pastaDestino);
    return {getFilesByName(nome){
      const encontrados=arquivos.filter(a=>a.getName()===nome);let i=0;
      return {hasNext:()=>i<encontrados.length,next:()=>encontrados[i++]};
    }};
  }};
  return app;
}
function arquivo(nome,criado,id){return {getName:()=>nome,getDateCreated:()=>new Date(criado),
  getLastUpdated:()=>new Date(criado),getId:()=>id,getUrl:()=>`https://drive.google.com/file/d/${id}/view`,isTrashed:()=>false};}
function nomes(chamado){return [
  `PLANILHA DE PARA PPG REPROSET - CHAMADO ${chamado}.xlsx`,
  `CALCULADORA FORNECEDOR REPROSET - LIVROS REENTRADA 2026 - NOVA VERSÃO - CHAMADO ${chamado}.xlsx`
];}
test('apos falha de transporte recupera o par da sessao sem criar arquivos',()=>{
  const inicio=Date.now()-3000,[d,c]=nomes('3106');
  const app=ambiente([arquivo(d,inicio+100,'DEPARA-1'),arquivo(c,inicio+200,'CALC-1')]);
  const r=app.consultarGeracaoReentradasSPG('REPROSET',['3106'],inicio);
  assert.equal(r.sucesso,true,r.erro);
  assert.equal(r.arquivos.length,1);
  assert.equal(r.arquivos[0].calculadoraId,'CALC-1');
  assert.equal(r.arquivos[0].deparaId,'DEPARA-1');
  assert.equal(r.arquivos[0].atualizadoEm,inicio+200);
});
test('nao confunde arquivo antigo com nova geracao e preserva parcial',()=>{
  const inicio=Date.now()-3000,[d,c]=nomes('3106');
  const app=ambiente([arquivo(d,inicio+50,'DEPARA-1'),arquivo(c,inicio-60000,'CALC-ANTIGO')]);
  const r=app.consultarGeracaoReentradasSPG('REPROSET',['3106'],inicio);
  assert.equal(r.sucesso,false);
  assert.deepEqual(Array.from(r.pendentes),['3106']);
  assert.equal(r.arquivos[0].calculadoraId,'');
  assert.equal(r.arquivos[0].deparaId,'DEPARA-1');
});
test('arquivos duplicados nao sao escolhidos arbitrariamente',()=>{
  const inicio=Date.now()-3000,[d,c]=nomes('3106');
  const app=ambiente([arquivo(d,inicio+50,'DEPARA-1'),arquivo(d,inicio+100,'DEPARA-2'),arquivo(c,inicio+150,'CALC-1')]);
  const r=app.consultarGeracaoReentradasSPG('REPROSET',['3106'],inicio);
  assert.equal(r.sucesso,false);
  assert.equal(r.ambiguos.length,1);
});
