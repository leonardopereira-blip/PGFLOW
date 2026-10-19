const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const ctx={};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.resolve(__dirname,'..','SPG.js'),'utf8'),ctx);

test('Booleanos da API sao Sim/Nao sem afetar valores nao booleanos',()=>{
 assert.equal(ctx._boolPT(true),'Sim');
 assert.equal(ctx._boolPT(false),'Não');
 assert.equal(ctx._boolPT('TRUE'),'Sim');
 assert.equal(ctx._boolPT('false'),'Não');
 assert.equal(ctx._boolPT(' True '),'Sim');
 assert.equal(ctx._boolPT(' FALSE '),'Não');
 assert.equal(ctx._boolPT('Sim'),'Sim');
 assert.equal(ctx._boolPT('Não'),'Não');
 assert.equal(ctx._boolPT(null),null);
 assert.equal(ctx._boolPT(0),0);
 assert.equal(ctx._boolPT(1),1);
 assert.equal(ctx._boolPT('3,5'),'3,5');
});
test('Especificacao preserva dimensoes e utiliza referencias da tiragem',()=>{
 const headers=[
 '1. ID_Código SKU','3, DIM_B (Aberto)(mm)','3, DIM_A (Aberto)(mm)',
 '3,DIM_BxA(Aberto)(mm)','3, DIM_B (Fechado)(mm)','3, DIM_A (Fechado)(mm)',
 '3. DIM_E (mm)','3.DIM_BXA(Fechado)','3.DIM_Dimensãoacabada(BxAxE)',
 '3. DIM_Peso (Kg)','4. CAP_Corte e vinco','5. MIO_PRIN_Corte e vinco',
 '5. MIO_PRIN_Serrilha','6. MIO_PROF_Corte e vinco',
 '7. MIO_ENCA_Corte e vinco','8. MIO_ADES_Corte e vinco',
 '11. REF_Referência troca de chapa','11. REF_Referência personalização'
 ];
 const sku='SKU123';
 const raw=[{
  skuAtual:sku,capBaseMmAberto:410.9,capAlturaMmAberto:275.3,
  baseMm:205.95,alturaMm:275.8,espessuraMm:2.78,pesoKg:0.234,
  capTemCorteVinco:true,
  miolos:[{corteVinco:false,serrilha:'TRUE'},{corteVinco:'FALSE'}],
  encartes:[{corteVinco:' true '}],adesivos:[{corteVinco:'FALSE'}]
 }];
 const refs={[sku]:{trocaChapa:'TROCA-321',personalizacao:'PERSON-21'}};
 const rows=ctx._mapearEspec(raw,headers,{dicSKU:{},modeTechGraf:{}},refs);
 const row=rows[1],value=k=>row[headers.indexOf(k)];
 assert.equal(value('3,DIM_BxA(Aberto)(mm)'),'410X275');
 assert.equal(value('3.DIM_BXA(Fechado)'),'205X275');
 assert.equal(value('3. DIM_E (mm)'),2.7);
 assert.equal(value('3.DIM_Dimensãoacabada(BxAxE)'),'205X275X2,7');
 assert.equal(value('3. DIM_Peso (Kg)'),0.234);
 assert.equal(value('4. CAP_Corte e vinco'),'Sim');
 assert.equal(value('5. MIO_PRIN_Corte e vinco'),'Não');
 assert.equal(value('5. MIO_PRIN_Serrilha'),'Sim');
 assert.equal(value('6. MIO_PROF_Corte e vinco'),'Não');
 assert.equal(value('7. MIO_ENCA_Corte e vinco'),'Sim');
 assert.equal(value('8. MIO_ADES_Corte e vinco'),'Não');
 assert.equal(value('11. REF_Referência troca de chapa'),'TROCA-321');
 assert.equal(value('11. REF_Referência personalização'),'PERSON-21');
});
