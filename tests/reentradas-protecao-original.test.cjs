const test=require('node:test');
const assert=require('node:assert/strict');
const mx=require('../tools/molde-node.cjs');

const atributo=(xml,nome)=>new RegExp('\\b'+nome+'="([^"]+)"').exec(xml||'')?.[1];
const protecao=xml=>(/<sheetProtection\b[^>]*?(?:\/>|>[\s\S]*?<\/sheetProtection>)/.exec(xml)||[])[0];
const celulas=xml=>new Map(Array.from(xml.matchAll(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g),
  m=>[atributo(m[0],'r'),m[0]]));

// Confere a protecao efetiva, incluindo celulas vazias que herdam os estilos
// da linha ou da coluna quando deixam de ter um elemento <c> no XML.
function mapaProtecao(xml,estilos){
  const xfs=Array.from(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(estilos)[1]
    .matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g),m=>m[0]);
  const colunas=Array.from(xml.matchAll(/<col\b[^>]*\/>/g),m=>({
    inicio:Number(atributo(m[0],'min')),fim:Number(atributo(m[0],'max')),
    estilo:atributo(m[0],'style')
  }));
  const linhas=new Map(Array.from(xml.matchAll(/<row\b[^>]*>/g),
    m=>[Number(atributo(m[0],'r')),m[0]]));
  const cells=celulas(xml);
  return {cells,consultar(ref){
    const letra=/^[A-Z]+/.exec(ref)[0],numero=Number(/\d+$/.exec(ref)[0]);
    let coluna=0;for(const c of letra)coluna=coluna*26+c.charCodeAt(0)-64;
    const faixa=colunas.find(c=>c.inicio<=coluna&&c.fim>=coluna),linha=linhas.get(numero);
    const estilo=cells.has(ref)?(atributo(cells.get(ref),'s')??'0'):
      ((atributo(linha,'customFormat')==='1'?(atributo(linha,'s')??'0'):undefined)??faixa?.estilo??'0');
    const tag=/<protection\b[^>]*\/>/.exec(xfs[Number(estilo)]||'')?.[0];
    return {locked:!/^(0|false)$/i.test(atributo(tag,'locked')||'1'),
      hidden:/^(1|true)$/i.test(atributo(tag,'hidden')||'0')};
  }};
}

test('protecao final conserva ausencias, senha e permissoes originais da aba',()=>{
  const app=mx.carregarApps(['ValidaReentradasSPG.js']);
  const prefixo='<worksheet><sheetData><row r="1"/></sheetData>';
  for(const tag of ['',
    '<sheetProtection algorithmName="SHA-512" hashValue="original" saltValue="original" spinCount="100000" sheet="1" objects="1" scenarios="1" autoFilter="0"/>',
    '<sheetProtection password="ABCD" sheet="1" selectLockedCells="1" sort="0"></sheetProtection>']){
    const original=prefixo+tag+'<extLst/></worksheet>';
    assert.equal(app._reSpProteger_(original),original);
  }
});

for(const nome of ['Logprint','CasaDaGrafica','FormaCerta','Meta','Reproset']){
  test(nome+': CALC Geral conserva a protecao original e todas as celulas editaveis do molde',()=>{
    const app=mx.carregarApps(['MoldeXlsx.js','Molde'+nome+'Calc.js','ValidaReentradasSPG.js']);
    const molde=app['_molde'+nome+'Calc_']();
    const rid=atributo(Array.from(molde.textos['xl/workbook.xml'].matchAll(/<sheet\b[^>]*\/>/g))
      .find(m=>atributo(m[0],'name')==='CALC Geral')[0],'r:id');
    const rel=Array.from(molde.textos['xl/_rels/workbook.xml.rels'].matchAll(/<Relationship\b[^>]*\/>/g))
      .find(m=>atributo(m[0],'Id')===rid)[0];
    const target=atributo(rel,'Target'),parte=target.startsWith('/')?target.slice(1):'xl/'+target;
    const original=app._mxGerarAba_(molde.abas[parte]);
    const skus=['200713648','MS26VF331P1SP','000123'];
    const dados=Object.fromEntries(skus.map(sku=>[sku,{avulso:100,kit:0,maleta:0}]));
    const preenchida=app._reSpPreencherCalcGeral_(original,skus,dados,'2932',{},nome.toUpperCase());
    const saida=app._reSpProteger_(preenchida);
    assert.equal(protecao(saida),protecao(original),'alterou a protecao/permissoes do molde');
    if(nome==='Logprint')assert.equal(protecao(saida),undefined,'a Logprint original tem abas livres');
    const antes=mapaProtecao(original,molde.textos['xl/styles.xml']);
    const depois=mapaProtecao(saida,molde.textos['xl/styles.xml']);
    const refs=new Set([...antes.cells.keys(),...depois.cells.keys()]);
    let editaveis=0,formulas=0;
    for(const ref of refs){
      if(Number(/\d+$/.exec(ref)[0])>skus.length+5)continue;
      const esperado=antes.consultar(ref);
      assert.deepEqual(depois.consultar(ref),esperado,ref+' mudou Locked/FormulaHidden');
      if(!esperado.locked)editaveis++;
      const c=antes.cells.get(ref)||'';
      if(/<f(?:\s|>)/.test(c)){
        assert.equal(depois.cells.get(ref),c,ref+' alterou uma formula protegida');
        formulas++;
      }
    }
    assert.ok(formulas>100,'nao conferiu as formulas reais da calculadora');
    if(nome!=='Logprint')assert.ok(editaveis>10,'nao encontrou as celulas desbloqueadas do molde');
  });
}
