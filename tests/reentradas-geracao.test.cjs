const test=require('node:test');
const assert=require('node:assert/strict');
const crypto=require('node:crypto');
const mx=require('../tools/molde-node.cjs');

const GRAFICAS=[
  {id:'LOGPRINT',nome:'Logprint',molde:'Logprint'},
  {id:'CASADAGRAFICA',nome:'Casa da Gráfica',molde:'CasaDaGrafica'},
  {id:'FORMACERTA',nome:'Forma Certa',molde:'FormaCerta',skus:31},
  {id:'META',nome:'Meta',molde:'Meta'},
  {id:'REPROSET',nome:'Reproset',molde:'Reproset'}
];
const SEM_ENTRADA=['CALC Digital','CALC Plana','CALC Rotativa','CALC Acabamento',
  'TAB Regras','TAB Formatos','TAB Papel','TAB Impressão','TAB Acabamento','TAB Versões'];
const sha=texto=>crypto.createHash('sha256').update(texto).digest('hex');
const semProtecao=xml=>xml.replace(/<sheetProtection\b[^>]*(?:\/>|>[\s\S]*?<\/sheetProtection>)/g,'');
const linha=(xml,n)=>{
  const m=new RegExp('<row\\b[^>]*\\br="'+n+'"[^>]*>[\\s\\S]*?<\\/row>').exec(xml);
  assert.ok(m,'linha '+n+' ausente');return m[0];
};
const celulas=xml=>new Map(Array.from(xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g),
  m=>[/\br="([A-Z]+\d+)"/.exec(m[1])[1],m[0]]));
const valor=(app,mapa,ref)=>{
  const m=/<v>([\s\S]*?)<\/v>|<t>([\s\S]*?)<\/t>/.exec(mapa.get(ref)||'');
  return m?app._reSpUnescape_(m[1]===undefined?m[2]:m[1]):undefined;
};
function hashFormulas(xml){
  const hash=crypto.createHash('sha256');let total=0;
  for(const c of xml.matchAll(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)){
    if(/<f(?:\s|>)/.test(c[0])){hash.update(c[0]);total++;}
  }
  return {sha256:hash.digest('hex'),total};
}

// Mantem os XML como texto. O teste captura exatamente os blobs entregues ao
// empacotador, sem comprimir mais de 100 MB por calculadora nem salvar no Drive.
function blobLeve(dados,tipo,nome){
  return {getName:()=>nome,getContentType:()=>tipo,
    getBytes:()=>Buffer.isBuffer(dados)?dados:Buffer.from(dados,'utf8'),
    getDataAsString:()=>Buffer.isBuffer(dados)?dados.toString('utf8'):dados,
    copyBlob:()=>blobLeve(dados,tipo,nome),
    setContentType(t){tipo=t;return this;}};
}
function prepararMatrizes(app,grafica,skus){
  const dep=app._reSpMolde_(grafica,'depara');
  const cab=n=>app._reSpMatriz_(dep,n).headers;
  const espH=cab('PPG - Especificações'),arvH=cab('PPG - Arvore de Produto'),tirH=cab('PPG - Tiragem');
  const row=(h,obj)=>h.map(c=>obj[c]===undefined?'':obj[c]);
  const arvore=[],tiragem=[];
  skus.forEach((sku,i)=>{
    arvore.push(row(arvH,{'1. ID_Código Kit':sku,'13. MAN_Total de itens colecionados':1}));
    arvore.push(row(arvH,{'1. ID_Código Kit':sku,'1. ID_Código SKU':sku}));
    tiragem.push(row(tirH,{'1. ID_Código Kit':sku,'1. ID_Código SKU':sku,'15. DES_CD envio':'CITA','16. TIR_SKU':100+i}));
  });
  for(const [kit,emb,quant] of [['KIT-NOVO','Caixa',111],['MALETA-NOVA','Maleta',222]]){
    arvore.push(row(arvH,{'1. ID_Código Kit':kit,'13. MAN_Total de itens colecionados':2,'13. MAN_Embalagem':emb}));
    arvore.push(row(arvH,{'1. ID_Código Kit':kit,'1. ID_Código SKU':skus[0]}));
    tiragem.push(row(tirH,{'1. ID_Código Kit':kit,'1. ID_Código SKU':skus[0],'15. DES_CD envio':'CDSB','16. TIR_SKU':quant}));
  }
  return {pacote:{id:'9876',versaoId:'v-nova',origem:'spg'},
    espec:{headers:espH,linhas:skus.map(sku=>row(espH,{'1. ID_Código SKU':sku}))},
    arvore:{headers:arvH,linhas:arvore},tiragem:{headers:tirH,linhas:tiragem}};
}

for(const g of GRAFICAS){
  test(g.nome+': geracao completa de De-Para e calculadora usa seu molde, novas entradas e formulas originais',()=>{
    const criados=[],pastas=[];
    const app=mx.carregarApps(['MoldeXlsx.js','MoldeLogprintDepara.js','Molde'+g.molde+'Calc.js',
      'ValidaReentradasSPG.js','ReentradasFormaCerta.js'],{
      DriveApp:{getFileById(){throw Error('SPG nao pode ler modelo ou dados externos no Drive');},
        getFolderById(id){pastas.push(id);return {createFile(blob){
          const file={blob,setName(nome){assert.equal(nome,blob.getName());return file;},
            getUrl:()=>'arquivo:'+blob.getName()};criados.push(file);return file;
        }};}},
      SpreadsheetApp:new Proxy({},{get(){throw Error('SPG nao pode consultar Google Sheets');}}),
      UrlFetchApp:new Proxy({},{get(){throw Error('geracao nao pode buscar modelos externos');}})
    });
    app.Utilities.newBlob=blobLeve;
    app.Utilities.zip=(blobs,nome)=>{
      const mapa=Object.fromEntries(blobs.map(b=>[b.getName(),b]));
      assert.equal(Object.keys(mapa).length,blobs.length,'partes duplicadas no XLSX');
      return {mapa,getName:()=>nome,setContentType(tipo){assert.equal(tipo,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');return this;}};
    };
    const qtd=g.skus||2,skus=Array.from({length:qtd},(_,i)=>'SKU-NOVO-'+g.id+'-'+String(i+1).padStart(3,'0'));
    const dados=prepararMatrizes(app,g.id,skus),consultas=[];
    app._reSpPrepararDados_=(numero,grafica)=>{consultas.push([numero,grafica]);return dados;};
    const fn='_molde'+g.molde+'Calc_',carregarMolde=app[fn];let leiturasMolde=0;
    app[fn]=()=>{leiturasMolde++;return carregarMolde();};
    try{
      const resultado=app._reSpGerarUm_('9876',g.id,[],{tecnologia:'PLANA',fornecedor:'ARCO'});
      assert.deepEqual(consultas,[['9876',g.id]]);assert.equal(leiturasMolde,1,'molde escolhido incorreto');
      assert.equal(resultado.grafica,g.nome);assert.equal(resultado.origem,'spg');assert.equal(resultado.skus,qtd);
      assert.equal(resultado.tiragem,skus.reduce((total,_,i)=>total+100+i,333));
      assert.equal(criados.length,2);assert.deepEqual(pastas,[app.RE_SP_CONFIG.pastaDestino,app.RE_SP_CONFIG.pastaDestino]);
      const depara=criados.find(f=>/^PLANILHA DE PARA/.test(f.blob.getName())).blob;
      const calc=criados.find(f=>/^CALCULADORA/.test(f.blob.getName())).blob;
      for(const arquivo of [depara,calc]){
        assert.ok(arquivo.getName().includes(g.nome.toUpperCase()),'nome sem a grafica selecionada');
        assert.ok(arquivo.getName().endsWith('CHAMADO 9876.xlsx'));
        assert.match(arquivo.mapa['xl/workbook.xml'].getDataAsString(),/fullCalcOnLoad="1"/);
        assert.equal(arquivo.mapa['xl/calcChain.xml'],undefined);
      }
      const caminho=(arquivo,nome)=>app._reSpAbaPorNome_(arquivo,nome);
      const ler=(arquivo,nome)=>arquivo.mapa[caminho(arquivo,nome)].getDataAsString();
      const molde=app._reSpMoldesCache_[g.id+'|calc'];
      for(const nome of SEM_ENTRADA){
        const parte=caminho(calc,nome),saida=calc.mapa[parte].getDataAsString();
        assert.match(saida,/<sheetProtection password="[0-9A-F]{4}" sheet="1"/,nome+' sem protecao');
        assert.equal(sha(semProtecao(saida)),sha(semProtecao(app._mxGerarAba_(molde.abas[parte]))),nome+' teve preco/formula/layout alterado');
      }
      // Estilos, imagens e textos tambem precisam vir do molde desta grafica.
      for(const nome of ['xl/styles.xml','xl/sharedStrings.xml'])
        assert.equal(sha(calc.mapa[nome].getDataAsString()),molde.origem.hashes[nome],nome);
      const geral=ler(calc,'CALC Geral'),pGeral=caminho(calc,'CALC Geral');
      assert.deepEqual(hashFormulas(geral),hashFormulas(app._mxGerarAba_(molde.abas[pGeral])),
        'formulas da CALC Geral foram substituidas durante a geracao');
      const primeira=celulas(linha(geral,6)),forma=g.id==='FORMACERTA';
      const valores=forma?{C6:'Chamado 9876',D6:skus[0],F6:'100',G6:'111',H6:'222',X6:'PLANA',AA6:'ARCO'}:
        {C6:'Chamado 9876',D6:skus[0],F6:'100',G6:'0',H6:'111',I6:'222',Y6:'PLANA',AB6:'ARCO'};
      for(const [ref,v] of Object.entries(valores))assert.equal(valor(app,primeira,ref),v,ref);
      assert.equal(valor(app,celulas(linha(geral,qtd+5)),'D'+(qtd+5)),skus.at(-1));
      assert.equal(valor(app,celulas(linha(geral,qtd+6)),'D'+(qtd+6)),undefined,'SKU antigo permaneceu apos os novos');
      assert.ok(ler(calc,'PPG Especificações').includes(skus[0]),'especificacao nao chegou a calculadora');
      const t=app._reSpMatriz_(depara,'PPG - Tiragem');
      assert.equal(t.linhas.length,qtd+2);
      assert.equal(app._reSpBuscar_(t.linhas[0],t.headers,'1. ID_Código SKU'),skus[0]);
      assert.equal(app._reSpBuscar_(t.linhas[0],t.headers,'16. TIR_SKU'),'100');
      if(forma){
        for(const ref of ['I6','Y6','AB6'])assert.match(primeira.get(ref),/<f(?:\s|>)/,ref+' perdeu formula');
        const imput=ler(calc,'IMPUT Forma Certa'),ultima=qtd+2;
        assert.match(imput,/<sheetProtection password="[0-9A-F]{4}" sheet="1"/);
        assert.match(celulas(linha(imput,ultima)).get('D'+ultima),new RegExp("<f>'CALC Geral'!D"+(ultima+3)+'<\\/f>'));
        assert.doesNotMatch(linha(imput,ultima),/<v>|<f\b[^>]*t="shared"/,'expansao herdou cache ou formula sem mestre');
        assert.ok(celulas(linha(imput,ultima)).size>50,'IMPUT adicional ficou somente com AH');
      }
    }finally{
      // Nenhum contexto/modelo grande e mantido para o proximo caso.
      criados.length=0;app._reSpMoldesCache_={};
    }
  });
}
