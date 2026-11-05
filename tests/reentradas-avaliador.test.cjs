const test=require('node:test');
const assert=require('node:assert/strict');
const mx=require('../tools/molde-node.cjs');

const app=()=>mx.carregarApps(['ValidaReentradasSPG.js','ReentradasAvaliador.js']);
function ambiente(){
  const valores={
    'CALC Geral':{6:{4:'SKU-1',6:800,7:0,8:230,9:0,10:1030,20:2060}},
    'PPG Especificações':{5:{3:'SKU-1',6:2.5},6:{3:'SKU-2',6:4}},
    'TAB Regras':{7:{45:'2.5 - DIGITAL',43:1.25}}
  };
  const ler=(aba,r,c)=>valores[aba]?.[r]?.[c]??'';
  const maxLinha=aba=>Math.max(0,...Object.keys(valores[aba]||{}).map(Number));
  return {aba:'CALC Geral',ler,maxLinha};
}
function avaliar(formula,ctx){const a=app();return a._rxAvExecutar_(a._rxAvCompilar_(formula),ctx||ambiente());}

test('soma os quatro tipos de tiragem e atualiza multiplicacoes dependentes',()=>{
  const contexto=ambiente();assert.equal(avaliar('SUM(F6:I6)',contexto),1030);
  assert.equal(avaliar('J6*2',contexto),2060);
  contexto.ler=(aba,r,c)=>aba==='CALC Geral'&&r===6&&c===10?1200:ambiente().ler(aba,r,c);
  assert.equal(avaliar('J6*2',contexto),2400);
});

test('consulta SKU entre abas, com fallback de IFNA sem aceitar erro desconhecido',()=>{
  const x='IF(D6&lt;&gt;"",_xlfn.XLOOKUP($D6,\'PPG Especificações\'!$C:$C,\'PPG Especificações\'!F:F),0)'.replace(/&lt;/g,'<').replace(/&gt;/g,'>');
  assert.equal(avaliar(x),2.5);
  assert.equal(avaliar('IFNA(XLOOKUP("nao existe",\'PPG Especificações\'!$C:$C,\'PPG Especificações\'!F:F),0)'),0);
  assert.throws(()=>avaliar('IFNA(FUNCAO_NAO_IMPLEMENTADA(D6),0)'),/Funcao ainda nao validada/);
});

test('INDEX, MATCH, IF, ROUND, CONCATENATE e COUNTIF usam valores atuais',()=>{
  assert.equal(avaliar('INDEX(\'PPG Especificações\'!F:F,MATCH(D6,\'PPG Especificações\'!C:C,0))'),2.5);
  assert.equal(avaliar('IF(J6>1000,ROUND(J6/3,1),0)'),343.3);
  assert.equal(avaliar('CONCATENATE(D6," - ","D")'),'SKU-1 - D');
  assert.equal(avaliar('COUNTIF(\'PPG Especificações\'!C:C,"SKU-1")'),1);
});
