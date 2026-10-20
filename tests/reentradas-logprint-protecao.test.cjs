const test=require('node:test');
const assert=require('node:assert/strict');
const mx=require('../tools/molde-node.cjs');

const attr=(x,n)=>new RegExp('\\b'+n+'="([^"]+)"').exec(x||'')?.[1];
const cells=x=>new Map(Array.from(x.matchAll(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g),m=>[attr(m[0],'r'),m[0]]));
const formulaOuValor=x=>x.replace(/<sheetProtection\b[^>]*\/>/g,'').replace(/<cols\b[^>]*>[\s\S]*?<\/cols>/g,'')
  .replace(/<(?:c|row)\b[^>]*>/g,tag=>tag.replace(/\s(?:s|customFormat)="[^"]*"/g,''));
const formato=x=>x.replace(/\sapplyProtection="[^"]*"/g,'').replace(/<protection\b[^>]*\/>/g,'').replace(/\/>$/, '></xf>');

function leitura(xml,estilos){
  const cs=cells(xml),cols=Array.from(xml.matchAll(/<col\b[^>]*\/>/g),m=>({min:+attr(m[0],'min'),max:+attr(m[0],'max'),s:+(attr(m[0],'style')||0)}));
  const rows=new Map(Array.from(xml.matchAll(/<row\b[^>]*>/g),m=>[+attr(m[0],'r'),m[0]]));
  const id=ref=>{
    let cn=0;for(const c of /^[A-Z]+/.exec(ref)[0])cn=cn*26+c.charCodeAt(0)-64;
    const row=rows.get(+/\d+$/.exec(ref)[0]),col=cols.find(c=>c.min<=cn&&c.max>=cn);
    return Number(cs.has(ref)?Number(attr(cs.get(ref),'s')||0):(attr(row,'customFormat')==='1'?attr(row,'s'):undefined)??col?.s??0);
  };
  return {cs,id,estilo:ref=>estilos[id(ref)]};
}
function colunasVisuais(xml){
  const cols=Array.from(xml.matchAll(/<col\b[^>]*\/>/g),m=>({min:+attr(m[0],'min'),max:+attr(m[0],'max'),
    valores:Object.fromEntries(Array.from(m[0].matchAll(/([A-Za-z]+)="([^"]*)"/g),a=>[a[1],a[2]]).filter(([k])=>!['min','max','style'].includes(k)))}));
  return {cols,valor:n=>cols.find(c=>c.min<=n&&c.max>=n)?.valores??{}};
}

test('Logprint recebe a mascara Casa/Reproset, mantendo cada formula, preco, valor e formato proprio',()=>{
  const app=mx.carregarApps(['MoldeXlsx.js','MoldeLogprintCalc.js','MoldeCasaDaGraficaCalc.js',
    'ValidaReentradasSPG.js','MacroProtecao.js','ProtecaoLogprint.js']);
  const molde=app._moldeLogprintCalc_(),referencia=app._moldeCasaDaGraficaCalc_(),partes=app._mxPartes_(molde),refParts=app._mxPartes_(referencia),alteradas={};
  const geral=app._reSpAbaPorNome_(partes,'CALC Geral'),ppg=app._reSpAbaPorNome_(partes,'PPG Especificações');
  const skus=['200713648','MS26VF331P1SP','000123'];
  alteradas[geral]=app._reSpPreencherCalcGeral_(partes.mapa[geral].getDataAsString(),skus,
    Object.fromEntries(skus.map(s=>[s,{avulso:100,kit:0,maleta:0}])),'2932',{},'LOGPRINT');
  alteradas[ppg]=app._reSpTrocarLinhas_(partes.mapa[ppg].getDataAsString(),5,skus.map((s,i)=>
    app._reSpLinha_(i+5,{2:app._reSpCodigoSku_(s)},app._reSpEstilos_(partes.mapa[ppg].getDataAsString(),5))));
  const antes={...alteradas},origXfs=Array.from(app._reMpXfs_(molde.textos['xl/styles.xml'],'cellXfs'));
  app._reSpAplicarProtecaoLogprint_(partes,alteradas,referencia);
  const novosXfs=Array.from(app._reMpXfs_(alteradas['xl/styles.xml'],'cellXfs'));
  assert.deepEqual(novosXfs.slice(0,origXfs.length),origXfs,'alterou um estilo original compartilhado');
  assert.ok(novosXfs.length>origXfs.length,'nao criou os estilos desbloqueados');
  for(const xf of novosXfs.slice(origXfs.length))assert.ok(origXfs.some(o=>formato(o)===formato(xf)),'clone alterou o formato visual');
  const estFinal=app._reMpEstilos_(alteradas['xl/styles.xml']),estRef=app._reMpEstilos_(referencia.textos['xl/styles.xml']);
  let formulas=0,editaveis=0;
  for(const sheet of app._reLpAbas_(molde.textos['xl/workbook.xml'],molde.textos['xl/_rels/workbook.xml.rels'])){
    const original=antes[sheet.caminho]??partes.mapa[sheet.caminho].getDataAsString(),saida=alteradas[sheet.caminho];
    assert.equal(formulaOuValor(saida),formulaOuValor(original),sheet.nome+' alterou conteudo/layout');
    assert.match(saida,new RegExp('<sheetProtection password="'+app._reSpSenhaHash_('PGFLOW2026')+'"'),sheet.nome+' sem protecao');
    const refXml=refParts.mapa[app._reSpAbaPorNome_(refParts,sheet.nome)].getDataAsString();
    formulas+=(saida.match(/<f(?:\s|>)/g)||[]).length;
    if(['PPG Especificações','CALC Geral','TAB Papel'].includes(sheet.nome)){
      const final=leitura(saida,estFinal),ref=leitura(refXml,estRef),anterior=leitura(original,origXfs);
      // Todas as entradas e tarifas reais, incluindo celulas materializadas
      // sem s e coordenadas vazias que dependem do estilo da coluna.
      for(const [coord,c] of final.cs){
        assert.deepEqual(final.estilo(coord),ref.estilo(coord),sheet.nome+'!'+coord+' divergiu da mascara');
        assert.equal(formato(novosXfs[final.id(coord)]),formato(origXfs[anterior.id(coord)]),sheet.nome+'!'+coord+' alterou formato visual/numerico');
        if(!final.estilo(coord).locked)editaveis++;
        if(/<f(?:\s|>)/.test(c))assert.equal(final.estilo(coord).locked,1,sheet.nome+'!'+coord+' formula livre');
      }
      for(const coord of ['A1','D1','C2','Y3','C6','D6','E6','F6','Y6','AB6','CF6','DN6','EV6','C10000','D10000','J6','R437','R438']){
        assert.deepEqual(final.estilo(coord),ref.estilo(coord),sheet.nome+'!'+coord+' heranca incorreta');
      }
      const colsAntes=colunasVisuais(original),colsDepois=colunasVisuais(saida);
      for(const c of [...colsAntes.cols,...colsDepois.cols])assert.deepEqual(colsDepois.valor(c.min),colsAntes.valor(c.min),sheet.nome+' mudou largura/visibilidade da coluna '+c.min);
    }else{
      assert.equal(saida.replace(/<sheetProtection\b[^>]*\/>/,''),original,sheet.nome+' mudou uma celula/estilo fora das entradas');
    }
    const originalTag=/<sheetProtection\b[^>]*>/.exec(refXml)[0],finalTag=/<sheetProtection\b[^>]*>/.exec(saida)[0];
    for(const k of Object.keys(app.RE_MP_OPCOES))assert.equal(+attr(finalTag,k),app._reMpBool_(attr(originalTag,k),app.RE_MP_OPCOES[k]),sheet.nome+' alterou permissao '+k);
  }
  assert.ok(formulas>375000,'nao conferiu as formulas reais: '+formulas);
  assert.ok(editaveis>1000,'nao conferiu as entradas e tarifas reais');
  assert.deepEqual(Object.keys(alteradas).sort(),['xl/styles.xml',...Object.keys(molde.abas)].sort(),'alterou partes fora das abas/estilos');
});

test('mascara Logprint rejeita uma formula na coordenada editavel antes de criar arquivos',()=>{
  const app=mx.carregarApps(['ValidaReentradasSPG.js','MacroProtecao.js','ProtecaoLogprint.js']);
  const xml='<worksheet><cols><col min="3" max="3"/></cols><sheetData><row r="6"><c r="C6"><f>1+1</f></c></row></sheetData></worksheet>';
  const locked={locked:1,hidden:0},unlocked={locked:0,hidden:0};
  const mask={padrao:locked,colunas:[{inicio:3,fim:3,estilo:1}],linhas:{},celulas:{},estilos:[locked,unlocked]};
  assert.throws(()=>app._reLpAplicarMascara_(xml,mask,()=>1),/desbloquearia formula Logprint: C6/);
});
