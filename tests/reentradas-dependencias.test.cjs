const test=require('node:test');
const assert=require('node:assert/strict');
const mx=require('../tools/molde-node.cjs');

function app(){return mx.carregarApps(['MoldeXlsx.js','ValidaReentradasSPG.js','ReentradasDependencias.js','ReentradasAvaliador.js']);}
function modelo(a){
  const padroes=['SUM(F6:I6)','J6*2',"'CALC Geral'!T6+3",'IFNA(XLOOKUP(D6,\'TAB Regras\'!$AS:$AS,\'TAB Regras\'!$AQ:$AQ),0)'];
  return {abas:[{nome:'CALC Geral'},{nome:'CALC Digital'},{nome:'TAB Regras'}],
    padroesCalculo:padroes.map(p=>a._mxRel_(p,6)),posicoesCalculo:[
      {'9':[[6,7,0]],'19':[[6,7,1]],'10':[[6,8,3]]},
      {'1':[[6,7,2]]},{}
    ]};
}

test('uma mudanca em F6 percorre J6, T6 e outra aba, sem atingir a linha 7',()=>{
  const a=app(),r=a._rxCalculoAfetados_(modelo(a),[{aba:0,linha:6,coluna:6}]);
  assert.deepEqual([...r.celulas].sort(),['0:6:10','0:6:20','1:6:2']);
  assert.deepEqual([...r.avisos],[]);
});

test('alterar uma tabela de apoio propaga para referencias da coluna inteira',()=>{
  const a=app(),r=a._rxCalculoAfetados_(modelo(a),[{aba:2,linha:20,coluna:45}]);
  assert.deepEqual([...r.celulas].sort(),['0:6:11','0:7:11','0:8:11']);
});

test('os nomes de abas e os textos das formulas nao criam referencias falsas',()=>{
  const a=app(),p=a._mxRel_('IF(D6="F7 e H6",\'TAB Regras\'!$AQ$7,F6)',6);
  const refs=a._rxDepsReferencias_(p,'CALC Geral');
  assert.equal(refs.length,3);
  assert.deepEqual([...refs.map(x=>x.aba)],['CALC Geral','TAB Regras','CALC Geral']);
  assert.equal(refs[1].r1,7);assert.equal(refs[1].rel1,false);
});

test('avalia somente o encadeamento alterado e devolve valores de outra aba sem gravar',()=>{
  const a=app(),m=modelo(a);
  m.abas[0].linhas=[{numero:6,valores:{3:'SKU-1',5:800,6:0,7:230,8:0,9:1030,19:2060}}];
  m.abas[1].linhas=[{numero:6,valores:{1:2063}}];
  m.abas[2].linhas=[];
  const r=a._rxCalculoAplicar_(m,[{aba:0,linha:6,coluna:6,valor:1200}],100);
  assert.deepEqual([...r.resultados].map(x=>[x.aba,x.linha,x.coluna,x.valor]).sort((x,y)=>x[2]-y[2]),[
    [1,6,2,2863],[0,6,10,1430],[0,6,20,2860]]);
  assert.equal(r.pendentes.length,0);
  assert.equal(m.abas[0].linhas[0].valores[5],800,'recalcular nao salva nem altera a entrada');
});
