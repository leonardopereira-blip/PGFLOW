const test=require('node:test'),assert=require('node:assert/strict');
const mx=require('../tools/molde-node.cjs');
const ctx=mx.carregarApps(['MoldeXlsx.js','MoldeRonaCalc.js','ValidaReentradasSPG.js']);
test('Rona: origem exata, 12 abas e 300 linhas de entrada',()=>{
  const m=ctx._moldeRonaCalc_();
  assert.equal(m.origem.sha256,'b689dfb5de110218d938f50379fdd2f3d5773cdeae3dbf6d75cdaa4dac7a2dcd');
  assert.equal(m.origem.bytes,15104235);assert.equal(m.ordem.length,62);
  assert.equal(Object.keys(m.abas).length,12);
  const geral=ctx._mxGerarAba_(m.abas['xl/worksheets/sheet2.xml']);
  assert.match(geral,/<row\b[^>]*\br="305"[^>]*>[\s\S]*?<f(?:\s|>)/);
  const skus=['200713648','SKU-RONA'];
  const porSku={'200713648':{avulso:10,kit:20,maleta:30},'SKU-RONA':{avulso:5,kit:0,maleta:0}};
  const blocos={DIGITAL:['CF','CK','CP','CU','CZ','DE','DI'],PLANA:['DN','DS','DX','EC','EH','EM','EQ'],ROTATIVA:['EV','FA','FF','FK','FP','FU','FY']};
  const elementos={};for(const t of Object.keys(blocos))elementos[t]={capa:'ARCO',forro:'GRÁFICA',miolo_principal:'ARCO TRIANG',miolo_professor:'ARCO',encartes:'GRÁFICA',adesivos:'ARCO',ajuste_papel:'ARCO'};
  const ler=(xml,ref)=>new RegExp('<c\\b[^>]*\\br="'+ref+'"[^>]*(?:/>|>[\\s\\S]*?</c>)').exec(xml)?.[0]||'';
  for(const t of Object.keys(blocos)){
    const xml=ctx._reSpPreencherCalcGeral_(geral,skus,porSku,'9876',{tecnologia:t,fornecedor:'ARCO',elementos},'RONA');
    assert.match(ler(xml,'D6'),/<v>200713648<\/v>/);assert.match(ler(xml,'D7'),/SKU-RONA/);
    assert.match(ler(xml,'Y6'),new RegExp(t));assert.match(ler(xml,'AB6'),/ARCO/);
    assert.match(ler(xml,'F6'),/<v>10<\/v>/);assert.match(ler(xml,'H6'),/<v>20<\/v>/);assert.match(ler(xml,'I6'),/<v>30<\/v>/);
    for(const k of Object.keys(blocos))for(const c of blocos[k]){
      const cel=ler(xml,c+'6');
      if(k===t)assert.match(cel,/ARCO|GRÁFICA/);else assert.doesNotMatch(cel,/<v>|<t>/);
    }
    assert.doesNotMatch(xml,/<row r="8"/);
    assert.equal(/<sheetProtection\b[^>]*\/>/.exec(xml)[0],/<sheetProtection\b[^>]*\/>/.exec(geral)[0]);
  }
});
