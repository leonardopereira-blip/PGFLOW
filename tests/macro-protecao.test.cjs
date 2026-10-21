const test=require('node:test');
const assert=require('node:assert/strict');
const mx=require('../tools/molde-node.cjs');
const attrs=tag=>Object.fromEntries(Array.from((tag||'').matchAll(/([\w:]+)="([^"]*)"/g),m=>[m[1],m[2]]));
const cellTags=xml=>new Map(Array.from(xml.matchAll(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g),m=>[attrs(m[0]).r,m[0]]));
const colNum=c=>Array.from(c).reduce((n,ch)=>n*26+ch.charCodeAt(0)-64,0);
function partes(worksheet,styles,name='CALC Geral'){
  const files={
    'xl/workbook.xml':'<workbook><sheets><sheet name="'+name+'" sheetId="1" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels':'<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
    'xl/styles.xml':styles,'xl/worksheets/sheet1.xml':worksheet
  };
  return {mapa:Object.fromEntries(Object.entries(files).map(([n,xml])=>[n,{getDataAsString:()=>xml}]))};
}
function lerMapa(xml){
  const sheet=/<sheet\b[^>]*>([\s\S]*?)<\/sheet>/.exec(xml);
  if(!sheet)return null;
  const a=attrs(sheet[0].slice(0,sheet[0].indexOf('>')+1));
  const passos=Array.from(sheet[1].matchAll(/<(range|merge)\b[^>]*\/>/g),m=>{
    const p=attrs(m[0]);let start,end;
    if(/^[A-Z]+:[A-Z]+$/.test(p.address)){
      const [l,r]=p.address.split(':');start=[colNum(l),1];end=[colNum(r),1048576];
    }else if(/^\d+:\d+$/.test(p.address)){
      const [l,r]=p.address.split(':').map(Number);start=[1,l];end=[16384,r];
    }else{
      const [l,r]=p.address.split(':'),parse=x=>{const m=/^([A-Z]+)(\d+)$/.exec(x);assert.ok(m,x);return [colNum(m[1]),+m[2]];};
      start=parse(l);end=parse(r||l);
    }
    return {...p,start,end,tipo:m[1]};
  });
  const ranges=passos.filter(r=>r.tipo==='range'),merges=passos.filter(r=>r.tipo==='merge');
  return {attrs:a,ranges,merges,get(col,row){
    let locked=+a.defaultLocked,hidden=+a.defaultHidden;
    for(const r of passos)if(r.start[0]<=col&&r.end[0]>=col&&r.start[1]<=row&&r.end[1]>=row){
      locked=+r.locked;hidden=+r.hidden;
    }
    return {locked,hidden};
  }};
}
function conferirRangesSemMesclagens(map){
  for(const r of map.ranges)for(const m of map.merges)
    assert.ok(r.end[0]<m.start[0]||r.start[0]>m.end[0]||r.end[1]<m.start[1]||r.start[1]>m.end[1],
      r.address+' intersecta a mesclagem '+m.address);
}

const estilos='<styleSheet><cellStyleXfs count="2">'+
  '<xf/><xf><protection locked="0" hidden="1"/></xf></cellStyleXfs>'+
  '<cellXfs count="5"><xf xfId="0"/>'+
  '<xf xfId="0" applyProtection="1"><protection locked="0"/></xf>'+
  '<xf xfId="0" applyProtection="1"><protection hidden="1"/></xf>'+
  '<xf xfId="1" applyProtection="0"><protection locked="1"/></xf>'+
  '<xf xfId="1" applyProtection="1"/>'+
  '</cellXfs></styleSheet>';

test('mapa restaura defaults, colunas, linhas e excecoes de celula na ordem correta',()=>{
  const app=mx.carregarApps(['ValidaReentradasSPG.js','MacroProtecao.js']);
  const prefix='<worksheet><cols><col min="2" max="3" style="1"/>'+
    '<col min="4" max="4" style="1"/><col min="6" max="8" style="2"/></cols><sheetData>';
  const rows='<row r="2" s="0" customFormat="1"><c r="B2" s="1"/><c r="C2" s="0"/>'+
    '<c r="E2" s="2"/><c r="F2" s="0"/></row>'+
    '<row r="3" s="1" customFormat="0"><c r="C3" s="0"/><c r="E3" s="3"/></row>'+
    '<row r="4" s="1" customFormat="1"><c r="A4" s="0"/><c r="B4"/>'+
    '<c r="F4" s="2"/><c r="G4" s="4"/></row>'+
    '<row r="5"><c r="B5" s="0"><f>1+1</f><v>2</v></c><c r="C5" s="0"/></row>'+
    '<row r="6"><c r="B6" s="0"/><c r="C6" s="0"/></row>';
  const suffix='</sheetData><sheetProtection sheet="1" objects="1" scenarios="1" autoFilter="0" sort="0" '+
    'formatCells="0" selectLockedCells="1" selectUnlockedCells="0" password="ABCD" hashValue="segredo" saltValue="segredo"/></worksheet>';
  const xml=prefix+rows+suffix,original=partes(xml,estilos),antes=JSON.stringify(original.mapa['xl/worksheets/sheet1.xml'].getDataAsString());
  const mapaXml=app._reSpMapaProtecao_(original,{}),map=lerMapa(mapaXml);
  assert.match(mapaXml,/^<protection xmlns="urn:pgflow:protecao:v1" version="1">/);
  assert.doesNotMatch(mapaXml,/password|hashValue|saltValue|segredo|<f>|<v>/);
  assert.equal(map.attrs.autoFilter,'0');assert.equal(map.attrs.formatColumns,'1');
  assert.equal(map.attrs.selectLockedCells,'1');assert.equal(map.attrs.selectUnlockedCells,'0');
  assert.ok(map.ranges.some(r=>r.address==='B:D'&&r.locked==='0'),'colunas contiguas nao compactadas');
  assert.ok(map.ranges.some(r=>r.address==='B5:C6'&&r.locked==='1'),'retangulo de celulas nao compactado');
  for(const [col,row,locked,hidden] of [
    [1,1,1,0],[2,1,0,0],[4,100,0,0],[6,100,1,1],
    [2,2,0,0],[3,2,1,0],[4,2,1,0],[5,2,1,1],[6,2,1,0],
    [2,3,0,0],[3,3,1,0],[5,3,0,1],
    [1,4,1,0],[2,4,1,0],[5,4,0,0],[6,4,1,1],[7,4,1,0],
    [2,5,1,0],[3,6,1,0],[2,7,0,0]
  ])assert.deepEqual(map.get(col,row),{locked,hidden},'col '+col+' linha '+row);
  assert.equal(JSON.stringify(original.mapa['xl/worksheets/sheet1.xml'].getDataAsString()),antes,'alterou XML da calculadora');
});

test('snapshot usa linhas e estilos finais, mas exige protecao no molde original',()=>{
  const app=mx.carregarApps(['ValidaReentradasSPG.js','MacroProtecao.js']);
  const original='<worksheet><sheetData><row r="2"><c r="B2" s="1"/></row>'+
    '<row r="3"><c r="B3" s="1"/></row></sheetData><sheetProtection sheet="1" autoFilter="0"/></worksheet>';
  const final='<worksheet><sheetData><row r="2"><c r="B2" s="0"/><c r="C2" s="2"/></row></sheetData></worksheet>';
  const map=lerMapa(app._reSpMapaProtecao_(partes(original,estilos),{'xl/worksheets/sheet1.xml':final}));
  assert.deepEqual(map.get(2,2),{locked:1,hidden:0});assert.deepEqual(map.get(3,2),{locked:1,hidden:1});
  assert.deepEqual(map.get(2,3),{locked:1,hidden:0},'incluiu linha removida');
  assert.equal(map.attrs.autoFilter,'0','usou opcoes finais em vez do molde');
  for(const tag of ['','<sheetProtection sheet="0"/>']){
    const livre='<worksheet><sheetData></sheetData>'+tag+'</worksheet>';
    assert.equal(lerMapa(app._reSpMapaProtecao_(partes(livre,estilos),{'xl/worksheets/sheet1.xml':original})),null);
  }
});

test('celula explicita sem s usa XF0 e apenas coordenadas ausentes herdam linha ou coluna',()=>{
  const app=mx.carregarApps(['ValidaReentradasSPG.js','MacroProtecao.js']);
  const xml='<worksheet><cols><col min="2" max="4" style="1"/></cols><sheetData>'+
    '<row r="2"><c r="B2"/><c r="C2" t="inlineStr"><is><t>SKU</t></is></c>'+
    '<c r="D2"><f>1+1</f><v>2</v></c></row>'+
    '<row r="3"/><row r="4" s="1" customFormat="1"><c r="B4"/></row>'+
    '<row r="5" customFormat="1"/>'+
    '</sheetData><sheetProtection sheet="1"/></worksheet>';
  const map=lerMapa(app._reSpMapaProtecao_(partes(xml,estilos),{}));
  for(const [col,row,locked] of [[2,2,1],[3,2,1],[4,2,1],[2,3,0],[2,4,1],[3,4,0],[2,5,1]])
    assert.deepEqual(map.get(col,row),{locked,hidden:0},'col '+col+' linha '+row);
});

test('estilo zero desbloqueado e FormulaHidden sao defaults reais do mapa',()=>{
  const app=mx.carregarApps(['ValidaReentradasSPG.js','MacroProtecao.js']);
  const styles='<styleSheet><cellStyleXfs><xf/></cellStyleXfs><cellXfs>'+
    '<xf applyProtection="1"><protection locked="0" hidden="1"/></xf>'+
    '<xf applyProtection="1"><protection locked="1" hidden="0"/></xf></cellXfs></styleSheet>';
  const xml='<worksheet><cols><col min="3" max="4" style="1"/></cols><sheetData>'+
    '<row r="2" s="0" customFormat="1"><c r="D2" s="1"/></row>'+
    '</sheetData><sheetProtection sheet="1"/></worksheet>';
  const map=lerMapa(app._reSpMapaProtecao_(partes(xml,styles),{}));
  assert.equal(map.attrs.defaultLocked,'0');assert.equal(map.attrs.defaultHidden,'1');
  for(const [col,row,locked,hidden] of [[1,1,0,1],[3,1,1,0],[3,2,0,1],[4,2,1,0],[16384,1048576,0,1]])
    assert.deepEqual(map.get(col,row),{locked,hidden});
});

test('mesclagens guardam a protecao da ancora final por coluna, linha e celula',()=>{
  const app=mx.carregarApps(['ValidaReentradasSPG.js','MacroProtecao.js']);
  const pre='<worksheet><cols><col min="2" max="4" style="1"/></cols><sheetData>';
  const rows='<row r="11" s="2" customFormat="1"/>'+
    '<row r="12" s="2" customFormat="1"><c r="B12" s="1"/></row>'+
    '<row r="13"><c r="B13" s="2"/></row>';
  const post='</sheetData><mergeCells count="4"><mergeCell ref="B10:D10"/>'+
    '<mergeCell ref="B11:D11"/><mergeCell ref="B12:D12"/><mergeCell ref="B13:D13"/>'+
    '</mergeCells><sheetProtection sheet="1"/></worksheet>';
  const original=partes(pre+rows+post,estilos),final=pre+rows.replace('r="B13" s="2"','r="B13" s="0"')+post;
  const mapXml=app._reSpMapaProtecao_(original,{'xl/worksheets/sheet1.xml':final});
  const merges=Array.from(mapXml.matchAll(/<merge\b[^>]*\/>/g),m=>attrs(m[0]));
  assert.deepEqual(merges,[
    {address:'B10:D10',locked:'0',hidden:'0'},
    {address:'B11:D11',locked:'1',hidden:'1'},
    {address:'B12:D12',locked:'0',hidden:'0'},
    {address:'B13:D13',locked:'1',hidden:'0'}
  ]);
  assert.ok(mapXml.lastIndexOf('<range ')<mapXml.indexOf('<merge '),'mesclagens devem vir apos os ranges');
  const map=lerMapa(mapXml);
  conferirRangesSemMesclagens(map);
  assert.deepEqual(map.get(3,10),{locked:0,hidden:0});
  assert.deepEqual(map.get(4,11),{locked:1,hidden:1});
  assert.deepEqual(map.get(3,12),{locked:0,hidden:0});
  assert.deepEqual(map.get(4,13),{locked:1,hidden:0},'mesclagem deve seguir a ancora B13');
  assert.deepEqual(map.get(2,14),{locked:0,hidden:0},'nao deve perder as demais linhas da coluna');
});

function protecaoEfetiva(xml,styles){
  const xfs=Array.from(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)[1]
    .matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g),m=>m[0]);
  const cells=cellTags(xml),rows=new Map(Array.from(xml.matchAll(/<row\b[^>]*>/g),m=>[+attrs(m[0]).r,attrs(m[0])]));
  const cols=Array.from(xml.matchAll(/<col\b[^>]*\/>/g),m=>attrs(m[0]));
  return (ref,col,row)=>{
    const c=attrs((/^<c\b[^>]*>/.exec(cells.get(ref)||'')||[])[0]),r=rows.get(row)||{};
    const colstyle=cols.find(c=>+c.min<=col&&+c.max>=col)?.style;
    const id=cells.has(ref)?(c.s??'0'):((r.customFormat==='1'?(r.s??'0'):undefined)??colstyle??'0');
    const p=attrs((/<protection\b[^>]*\/>/.exec(xfs[+id])||[])[0]);
    return {locked:p.locked==='0'?0:1,hidden:p.hidden==='1'?1:0};
  };
}
for(const nome of ['Logprint','CasaDaGrafica','FormaCerta','Meta','Reproset']){
  test(nome+': mapa aplicado equivale a protecao efetiva da CALC Geral gerada',()=>{
    const app=mx.carregarApps(['MoldeXlsx.js','Molde'+nome+'Calc.js','ValidaReentradasSPG.js','MacroProtecao.js']);
    const m=app['_molde'+nome+'Calc_'](),s=attrs(Array.from(m.textos['xl/workbook.xml'].matchAll(/<sheet\b[^>]*>/g))
      .find(x=>attrs(x[0]).name==='CALC Geral')[0]);
    const rel=attrs(Array.from(m.textos['xl/_rels/workbook.xml.rels'].matchAll(/<Relationship\b[^>]*\/>/g))
      .find(x=>attrs(x[0]).Id===s['r:id'])[0]);
    const path=rel.Target.startsWith('/')?rel.Target.slice(1):'xl/'+rel.Target;
    const original=app._mxGerarAba_(m.abas[path]);
    const skus=['200713648','ALFA123','000123'],dados=Object.fromEntries(skus.map(s=>[s,{avulso:100,kit:0,maleta:0}]));
    const final=app._reSpPreencherCalcGeral_(original,skus,dados,'2932',{},nome.toUpperCase());
    const metadata=app._reSpMapaProtecao_(partes(original,m.textos['xl/styles.xml']),{'xl/worksheets/sheet1.xml':final});
    const map=lerMapa(metadata);
    if(nome==='Logprint'){assert.equal(map,null);return;}
    conferirRangesSemMesclagens(map);
    const esperado=protecaoEfetiva(final,m.textos['xl/styles.xml']);
    for(const row of [1,2,3,4,5,6,7,8,9,300])for(let col=1;col<=190;col++){
      const ref=app._reSpCol_(col)+row;
      assert.deepEqual(map.get(col,row),esperado(ref,col,row),nome+' '+ref);
    }
    assert.equal(map.attrs.autoFilter,'0');
    assert.ok(map.ranges.length<500,'mapa de tres SKUs nao foi compactado');
  });
}
