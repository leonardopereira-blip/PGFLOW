const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const os=require('node:os');
const path=require('node:path');
const mx=require('../tools/molde-node.cjs');

const MOLDES=[
  {fn:'_moldeLogprintCalc_',arquivo:'MoldeLogprintCalc.js'},
  {fn:'_moldeLogprintDepara_',arquivo:'MoldeLogprintDepara.js'}
];
const ctx=mx.carregarApps(['MoldeXlsx.js',...MOLDES.map(m=>m.arquivo)]);
const moldes={};
for(const m of MOLDES)moldes[m.fn]=ctx[m.fn]();

test('formula relativa volta exatamente ao texto original',()=>{
  const casos=['<c r="D8" s="52"><f t="array" ref="D8">_xlfn.IFNA(_xlfn.XLOOKUP($C8,\'PPG\'!$C:$C,\'PPG\'!$D:$D),0)</f><v>\u0003</v></c>',
    '=SUM(A1:B2)+$A$1+A$1+$A1+LOG10(5)+"X12"+AB1048576+A0007','sem referencia'];
  for(const t of casos)for(const linha of [1,8,500])assert.equal(ctx._mxAbs_(ctx._mxRel_(t,linha),linha),t);
  assert.equal(ctx._mxRel_('$C8+E9',8),'$C\u00010\u0002+E\u00011\u0002');
});

for(const m of MOLDES){
  test(m.arquivo+': cada parte recriada do zero tem o mesmo SHA-256 do Excel modelo',()=>{
    const molde=moldes[m.fn],hashes=molde.origem.hashes;
    const gerado=ctx._mxGerarArquivos_(molde);
    assert.deepEqual(Array.from(gerado,a=>a.nome),Object.keys(hashes));
    for(const a of gerado)assert.equal(mx.sha256(mx.bytesDoArquivo(a)),hashes[a.nome],a.nome);
  });
}

// Comparacao direta com o Excel original quando ele esta disponivel na maquina.
const pastaFontes=process.env.MOLDE_FONTES||path.join(os.homedir(),'Downloads');
for(const m of MOLDES){
  const fonte=path.join(pastaFontes,moldes[m.fn].origem.arquivo);
  test(m.arquivo+': identico ao arquivo original '+path.basename(fonte),{skip:!fs.existsSync(fonte)&&'original ausente'},()=>{
    const buf=fs.readFileSync(fonte);
    assert.equal(mx.sha256(buf),moldes[m.fn].origem.sha256,'o arquivo local nao e a versao mapeada');
    const orig=new Map(mx.lerZip(buf).map(p=>[p.nome,p.dados]));
    const gerado=new Map(ctx._mxGerarArquivos_(moldes[m.fn]).map(a=>[a.nome,mx.bytesDoArquivo(a)]));
    for(const [nome,dados] of orig){
      if(nome==='xl/calcChain.xml'){assert.ok(!gerado.has(nome));continue;}
      const g=gerado.get(nome);
      if(nome==='[Content_Types].xml'||nome==='xl/_rels/workbook.xml.rels'){
        const semCalc=dados.toString('utf8').replace(/<Override\b[^>]*calcChain\.xml"[^>]*\/>|<Relationship\b[^>]*calcChain\.xml"[^>]*\/>/,'');
        assert.equal(g.toString('utf8'),semCalc,nome);
      }else assert.ok(g&&g.equals(dados),nome+' difere do original');
    }
    assert.equal(gerado.size,orig.size-(orig.has('xl/calcChain.xml')?1:0));
  });
}

test('geracao completa usa somente moldes e preserva abas sem entrada identicas ao modelo',()=>{
  const code=['MoldeXlsx.js','MoldeLogprintCalc.js','MoldeLogprintDepara.js','ValidaReentradasSPG.js',
    'MacroProtecao.js','ProtecaoLogprint.js','MoldeCasaDaGraficaCalc.js','ReentradasElegibilidade.js'];
  const criados=[];
  const app=mx.carregarApps(code,{
    DriveApp:{
      getFileById(){throw Error('nao pode ler planilha modelo no Drive');},
      getFolderById:()=>({createFile(b){const f={blob:b,setName(){return f;},getUrl:()=>'url:'+b.getName()};criados.push(f);return f;}})
    }
  });
  app.Utilities.zip=(blobs,nome)=>app.Utilities.newBlob(mx.escreverZip(blobs.map(b=>({nome:b.getName(),dados:b.getBytes()}))),'application/zip',nome);
  const dep=app._reSpMolde_('LOGPRINT','depara');
  const cab=n=>app._reSpMatriz_(dep,n).headers;
  const linha=(h,valores)=>h.map(c=>valores[c]===undefined?'':valores[c]);
  const espH=cab('PPG - Especificações'),arvH=cab('PPG - Arvore de Produto'),tirH=cab('PPG - Tiragem');
  const skus=['SKU-TESTE-1','SKU-TESTE-2'];
  const api={
    'Esp. Técnica (legado)':{headers:espH,linhas:skus.map(s=>linha(espH,{'1. ID_Código SKU':s}))},
    'Árvore (legado)':{headers:arvH,linhas:skus.flatMap(s=>[linha(arvH,{'1. ID_Código Kit':s,'13. MAN_Total de itens colecionados':1}),
      linha(arvH,{'1. ID_Código Kit':s,'1. ID_Código SKU':s})])},
    'Tiragem (legado)':{headers:tirH,linhas:skus.map(s=>linha(tirH,{'1. ID_Código Kit':s,'1. ID_Código SKU':s,'15. DES_CD envio':'CITA','16. TIR_SKU':100}))}
  };
  const pacote={id:'1234',versaoId:'v1',blob:{},statusChamado:'SOLICITADO',statusProducao:'PUBLICADA'};
  app._emailSpgPacote_=()=>pacote;
  const matrizOriginal=app._reSpMatriz_;
  app._reSpMatriz_=(partes,nome)=>partes===pacote.raw?api[nome]:matrizOriginal(partes,nome);
  app._reSpPartes_=()=>(pacote.raw={});
  const r=app._reSpGerarUm_('1234','LOGPRINT',[],{});
  assert.equal(r.skus,2);assert.equal(r.tiragem,200);assert.equal(criados.length,2);

  const calc=criados.find(f=>/CALCULADORA/.test(f.blob.getName())),zipCalc=calc.blob.getBytes();
  const ler=nome=>{const p=mx.lerZip(zipCalc,n=>n===nome)[0];return p&&p.dados.toString('utf8');};
  const molde=moldes._moldeLogprintCalc_,partesMolde=app._reSpMolde_('LOGPRINT','calc'),caminhos={};
  const semEntrada=['CALC Digital','CALC Plana','CALC Rotativa','CALC Acabamento','TAB Regras','TAB Formatos','TAB Papel','TAB Impressão','TAB Acabamento','TAB Versões'];
  for(const n of [...semEntrada,'CALC Geral'])caminhos[n]=app._reSpAbaPorNome_(partesMolde,n);
  for(const n of semEntrada){
    const x=ler(caminhos[n]),orig=ctx._mxGerarAba_(molde.abas[caminhos[n]]);
    const comparar=n==='TAB Papel'?s=>s.replace(/<sheetProtection\b[^>]*\/>/,'').replace(/<cols\b[^>]*>[\s\S]*?<\/cols>/,'').replace(/\s(?:s|customFormat)="\d+"/g,''):s=>s.replace(/<sheetProtection\b[^>]*\/>/,'');
    assert.equal(mx.sha256(Buffer.from(comparar(x))),mx.sha256(Buffer.from(comparar(orig))),n+' alterada alem da protecao autorizada da Logprint');
    assert.match(x,/<sheetProtection password="C02C"/);
  }
  assert.match(ler(caminhos['CALC Geral']),/<c r="D6"[^>]*t="inlineStr"><is><t>SKU-TESTE-1<\/t>/);
  assert.match(ler('xl/workbook.xml'),/fullCalcOnLoad="1"/);
  assert.equal(ler('xl/calcChain.xml'),undefined);
});
