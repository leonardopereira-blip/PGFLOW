const test=require('node:test');
const assert=require('node:assert/strict');
const mx=require('../tools/molde-node.cjs');

const ctx=mx.carregarApps(['MoldeXlsx.js','MoldeFormaCertaCalc.js',
  'ValidaReentradasSPG.js','ReentradasFormaCerta.js']);
const forma=ctx._moldeFormaCertaCalc_();
const partes={mapa:Object.fromEntries(['xl/workbook.xml','xl/_rels/workbook.xml.rels']
  .map(nome=>[nome,{getDataAsString:()=>forma.textos[nome]}]))};
const caminho=nome=>ctx._reSpAbaPorNome_(partes,nome);
const gerarAba=nome=>ctx._mxGerarAba_(forma.abas[caminho(nome)]);
const cells=xml=>new Map(Array.from(xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g),
  m=>[/\br="([A-Z]+\d+)"/.exec(m[1])[1],m[0]]));
const row=(xml,n)=>new RegExp('<row\\b[^>]*\\br="'+n+'"[^>]*>[\\s\\S]*?<\\/row>').exec(xml)[0];
const quantidade={'SKU-NOVA':{avulso:123,kit:456,maleta:789}};
const elementos={capa:'ARCO',forro:'GRÁFICA',miolo_principal:'ARCO TRIANG',
  miolo_professor:'ARCO',encartes:'GRÁFICA',adesivos:'ARCO TRIANG',ajuste_papel:'ARCO'};
const fornecimento={tecnologia:'PLANA',fornecedor:'ARCO',elementos:{
  DIGITAL:{...elementos},PLANA:{...elementos},ROTATIVA:{...elementos}
}};
const valor=(mapa,ref)=>{
  const celula=mapa.get(ref)||'';
  const v=/<v>([\s\S]*?)<\/v>|<t>([\s\S]*?)<\/t>/.exec(celula);
  return v?ctx._reSpUnescape_(v[1]===undefined?v[2]:v[1]):undefined;
};

test('Forma Certa preenche contrato real G/H, X/AA e seis fornecimentos, preservando todas as formulas',()=>{
  // Reconstroi somente esta aba e usa suas duas primeiras linhas de dados.
  // As demais abas de calculo do molde (mais de 100 MB) nao sao expandidas.
  const geral=gerarAba('CALC Geral');
  const xml='<worksheet><sheetData>'+row(geral,6)+row(geral,7)+'</sheetData></worksheet>';
  const resultado=ctx._reSpPreencherCalcGeral_(xml,['SKU-NOVA'],quantidade,'2567',fornecimento,'Forma Certa');
  const antes=cells(xml),depois=cells(resultado);
  assert.equal(valor(depois,'C6'),'Chamado 2567');
  assert.equal(valor(depois,'D6'),'SKU-NOVA');
  for(const [ref,v] of Object.entries({F6:'123',G6:'456',H6:'789',X6:'PLANA',AA6:'ARCO'}))assert.equal(valor(depois,ref),v,ref);
  for(const ref of ['I6','Y6','AB6']){
    assert.match(antes.get(ref),/<f(?:\s|>)/);
    assert.equal(depois.get(ref),antes.get(ref),ref+' preservada integralmente');
  }
  for(const [ref,celula] of antes){
    if(ref.endsWith('6')&&/<f(?:\s|>)/.test(celula))assert.equal(depois.get(ref),celula,'formula alterada em '+ref);
  }
  assert.equal(depois.has('E7'),false,'formula nao utilizada permaneceu');
  const tecnologia={
    DIGITAL:['CA','CE','CI','CM','CQ','CU'],
    PLANA:['CY','DC','DG','DK','DO','DS'],
    ROTATIVA:['DW','EA','EE','EI','EM','EQ']
  };
  for(const [nome,bloco] of Object.entries(tecnologia)){
    bloco.forEach((col,i)=>assert.equal(valor(depois,col+'6'),
      nome==='PLANA'?Object.values(elementos)[i]:undefined,col));
  }
  for(const ref of ['C7','D7','F7','G7','H7','X7','AA7','CA7','CE7'])assert.equal(valor(depois,ref),undefined,'entrada antiga em '+ref);
  assert.doesNotMatch(resultado,/Chamado 2896/);
});

for(const grafica of ['Logprint','Casa da Gráfica','Meta','Reproset']){
  test(grafica+': conserva contrato original F/G/H/I, Y/AB e sete fornecimentos',()=>{
    const xml='<worksheet><sheetData><row r="6">'+
      '<c r="C6"><v>ANTIGO</v></c><c r="D6"><v>SKU-ANTIGA</v></c>'+
      '<c r="E6"><f>XLOOKUP(D6,D:D,E:E)</f><v>99</v></c>'+
      '<c r="F6"><v>0</v></c><c r="G6"><v>0</v></c><c r="H6"><v>0</v></c><c r="I6"><v>0</v></c>'+
      '<c r="X6"><f>SUM(F6:I6)</f><v>0</v></c>'+
      '<c r="Y6"><v>DIGITAL</v></c><c r="AA6"><f>SUM(AE6:AZ6)</f><v>0</v></c>'+
      '<c r="AB6"><v>GRÁFICA</v></c></row></sheetData></worksheet>';
    const resultado=ctx._reSpPreencherCalcGeral_(xml,['SKU-NOVA'],quantidade,'2567',fornecimento,grafica);
    const antes=cells(xml),depois=cells(resultado);
    for(const [ref,v] of Object.entries({F6:'123',G6:'0',H6:'456',I6:'789',Y6:'PLANA',AB6:'ARCO'}))assert.equal(valor(depois,ref),v,ref);
    for(const ref of ['E6','X6','AA6'])assert.equal(depois.get(ref),antes.get(ref),ref);
    const tecnologia={
      DIGITAL:['CF','CK','CP','CU','CZ','DE','DI'],
      PLANA:['DN','DS','DX','EC','EH','EM','EQ'],
      ROTATIVA:['EV','FA','FF','FK','FP','FU','FY']
    };
    for(const [nome,bloco] of Object.entries(tecnologia)){
      bloco.forEach((col,i)=>assert.equal(valor(depois,col+'6'),
        nome==='PLANA'?Object.values(elementos)[i]:undefined,col));
    }
    assert.doesNotMatch(resultado,/ANTIGO|SKU-ANTIGA/);
  });
}

test('IMPUT real Forma Certa atende 300 SKUs, resolve formulas compartilhadas e limita linhas a 305',()=>{
  const original=gerarAba('IMPUT Forma Certa');
  const resultado=ctx._reFcExpandirImput_(original,300);
  const final=cells(row(resultado,302));
  assert.match(final.get('D302'),/<f>'CALC Geral'!D305<\/f>/);
  assert.match(final.get('A302'),/<f>/);
  assert.doesNotMatch(row(resultado,302),/<v>|t="shared"/);
  for(let n=31;n<=302;n++){
    const linha=row(resultado,n);
    assert.match(cells(linha).get('D'+n),new RegExp("<f>'CALC Geral'!D"+(n+3)+'<\\/f>'));
    assert.doesNotMatch(linha,/<f\b[^>]*t="shared"|<v>/,'formula compartilhada ou cache herdado na linha '+n);
  }
  const numeros=Array.from(resultado.matchAll(/<row\b[^>]*\br="(\d+)"/g),m=>+m[1]);
  assert.equal(Math.max(...numeros),305);
  assert.equal(numeros.length,305);
  assert.match(resultado,/<dimension ref="A1:BW305"\/>/);
  assert.equal(row(resultado,30),row(original,30));
  assert.equal(row(resultado,303),row(original,303));
  assert.throws(()=>ctx._reFcExpandirImput_(original,301),/limite de 300/);
});
