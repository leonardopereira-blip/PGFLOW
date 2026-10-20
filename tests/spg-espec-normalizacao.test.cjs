const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const ctx={};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.resolve(__dirname,'..','SPG.js'),'utf8'),ctx);

test('Booleanos da API sao Sim/Nao sem afetar valores nao booleanos',()=>{
 for(const value of [true,'true','TRUE',' TrUe ','sim','sIM',' SIM ']) {
  assert.equal(ctx._spgBoolPT_(value),'Sim',String(value));
 }
 for(const value of [false,'false','FALSE',' FaLsE ','não','NãO','NAO',' nao ']) {
  assert.equal(ctx._spgBoolPT_(value),'Não',String(value));
 }
 for(const value of [undefined,null,0,1,'0','1','3,5','true/false','Sim, com ressalva','não informado','']) {
  assert.equal(ctx._spgBoolPT_(value),value,String(value));
 }
});

test('Entrada converte booleanos reais sem reescrever textos livres ou componentes JSON',()=>{
 const source={isGraphic:true,isEditorial:false,descricao:'TRUE',observacoes:'false',quantidade:123,
  miolos:[{corteVinco:true,serrilha:false}]};
 const row=ctx._emularDadosDaPlanilha([source])[0];
 assert.equal(row.isGraphic,'Sim');
 assert.equal(row.isEditorial,'Não');
 assert.equal(row.descricao,'TRUE');
 assert.equal(row.observacoes,'false');
 assert.equal(row.quantidade,'123');
 assert.deepEqual(JSON.parse(row.miolos),source.miolos);
 assert.equal(source.isGraphic,true);
});

test('Todas as opcoes booleanas da especificacao aceitam caixa variavel sem alterar campos livres',()=>{
 const booleanHeaders=[
  '4. CAP_Corte e vinco','5. MIO_PRIN_Corte e vinco','5. MIO_PRIN_Serrilha',
  '6. MIO_PROF_Corte e vinco','6. MIO_PROF_Serrilha','7. MIO_ENCA_Corte e vinco',
  '7. MIO_ENCA_Serrilha','8. MIO_ADES_Corte e vinco'
 ];
 for(const [value,expected] of [[true,'Sim'],['TRUE','Sim'],['sIM','Sim'],[false,'Não'],['FaLsE','Não'],['não','Não'],['NAO','Não']]) {
  const rows=ctx._mapearEspec([{
   skuAtual:'SKU123',descricao:'TRUE',obsNotas:'false',isPersonalizado:value,
   skuMaterialPadrao:'SKU-PADRAO',capTemOrelhas:value,capTemCorteVinco:value,
   miolos:[{corteVinco:value,serrilha:value},{corteVinco:value,serrilha:value}],
   encartes:[{corteVinco:value,serrilha:value}],adesivos:[{corteVinco:value}]
  }],ctx.CABECALHO_ESPEC,{dicSKU:{},modeTechGraf:{}},{});
  const row=rows[1],get=header=>row[ctx.CABECALHO_ESPEC.indexOf(header)];
  for(const header of booleanHeaders) assert.equal(get(header),expected,`${header}: ${value}`);
  assert.equal(get('2. INF_Cliente personalizado'),expected);
  assert.equal(get('4.Cap_orelha'),expected);
  assert.equal(get('11. REF_Tipo de personalização'),expected==='Sim'?'Capa personalizada':'Capa Padrão');
  assert.equal(get('1. ID_Descrição'),'TRUE');
  assert.equal(get('10. OBS_PROD_Observação para produção gráfica'),'false');
 }
 const rows=ctx._mapearEspec([
  {skuAtual:'COM-LARGURA',capTemOrelhas:true,capLarguraOrelha:90},
  {skuAtual:'SEM-INFORMACAO'}
 ],ctx.CABECALHO_ESPEC,{dicSKU:{},modeTechGraf:{}},{});
 assert.equal(rows[1][ctx.CABECALHO_ESPEC.indexOf('4.Cap_orelha')],90);
 assert.equal(rows[2][ctx.CABECALHO_ESPEC.indexOf('4.Cap_orelha')],'');
});

test('Motor, especificacao, tiragem e arvore preservam rotulos Grafico/Editorial para todas as formas booleanas',()=>{
 for(const [positive,negative] of [[true,false],['TRUE','FaLsE'],['sIM','nÃo'],[' Sim ',' NAO ']]) {
  const espec=ctx._emularDadosDaPlanilha([{skuAtual:'SKU123',isGraphic:positive,isEditorial:negative}]);
  const dicEspec={SKU123:espec[0]};
  const motor=ctx._mapearConsolidadoECriarMotor(ctx._emularDadosDaPlanilha([{
   sku:'SKU123',kitCode:'KIT123',marca:'SAS',tecnologia:'DIGITAL',classificacao:'NORMAL',
   envio:'1',grafica1:'LOGPRINT',liberado:true,observacao:'false'
  }]),dicEspec);
  const matrizes=[
   motor.matriz,
   ctx._mapearEspec(espec,ctx.CABECALHO_ESPEC,motor,{}),
   ctx._mapearTiragem([{sku:'SKU123',kitCode:'KIT123',quantidade:20}],dicEspec,ctx.CABECALHO_TIRAGEM,motor),
   ctx._mapearArvore([{kitCode:'KIT123',items:[{itemCode:'SKU123'}]}],ctx.CABECALHO_ARVORE,dicEspec,motor)
  ];
  for(const matriz of matrizes) {
   const row=matriz[matriz.length-1];
   assert.equal(row[matriz[0].indexOf('Grafico')],'Grafico',String(positive));
   assert.equal(row[matriz[0].indexOf('Editorial')],'Não Editorial',String(negative));
  }
  assert.equal(motor.matriz[1][motor.matriz[0].indexOf('liberado')],'Sim');
  assert.equal(motor.matriz[1][motor.matriz[0].indexOf('observacao')],'false');
 }
});

test('Mapeamento e formato da espessura nao dependem da ordem dos helpers de Ajusta_API',()=>{
 const spg=fs.readFileSync(path.resolve(__dirname,'..','SPG.js'),'utf8');
 const ajusta=fs.readFileSync(path.resolve(__dirname,'..','Ajusta_API.js'),'utf8');
 for(const files of [[spg,ajusta],[ajusta,spg]]) {
  const formats=[];
  const isolated={SpreadsheetApp:{flush(){}}};
  vm.createContext(isolated);
  files.forEach(source=>vm.runInContext(source,isolated));
  const headers=isolated.CABECALHO_ESPEC;
  const rows=isolated._mapearEspec([{
   skuAtual:'SKU123',isGraphic:'sIM',isEditorial:'nÃo',isPersonalizado:' TRUE ',
   skuMaterialPadrao:'PADRAO',capTemOrelhas:'false',capTemCorteVinco:' FaLsE ',
   miolos:[{corteVinco:'sIM',serrilha:'NAO'}],baseMm:205.95,alturaMm:275.8,espessuraMm:2.78
  }],headers,{dicSKU:{},modeTechGraf:{}},{SKU123:{trocaChapa:'TROCA',personalizacao:'PERSONAL'}});
  const get=header=>rows[1][headers.indexOf(header)];
  assert.equal(get('Grafico'),'Grafico');
  assert.equal(get('Editorial'),'Não Editorial');
  assert.equal(get('2. INF_Cliente personalizado'),'Sim');
  assert.equal(get('4.Cap_orelha'),'Não');
  assert.equal(get('4. CAP_Corte e vinco'),'Não');
  assert.equal(get('5. MIO_PRIN_Corte e vinco'),'Sim');
  assert.equal(get('5. MIO_PRIN_Serrilha'),'Não');
  assert.equal(get('3.DIM_Dimensãoacabada(BxAxE)'),'205X275X2,7');
  assert.equal(get('11. REF_Referência troca de chapa'),'TROCA');
  assert.equal(get('11. REF_Referência personalização'),'PERSONAL');
  const sheet={
   clear(){},getMaxRows(){return 100;},getMaxColumns(){return 100;},
   getRange(...args){return {setValues(){},setFontWeight(){},setNumberFormat(format){formats.push({args,format});}};}
  };
  isolated._spgGravarAba_({getSheetByName(){return sheet;}},isolated.SPG_CONFIG.abaEspec,rows);
  assert.deepEqual(formats,[{args:[2,headers.indexOf('3. DIM_E (mm)')+1,1,1],format:'0.0'}]);
 }
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
