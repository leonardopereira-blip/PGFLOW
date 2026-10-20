// Mapeia um Excel modelo para um arquivo de molde do Apps Script.
// Uso: node tools/mapear-molde.cjs <modelo.xlsx> <NomeDoMolde>
//   Ex.: node tools/mapear-molde.cjs "CALCULADORA ... LOGPRINT.xlsx" LogprintCalc
// Gera Molde<NomeDoMolde>.js na raiz com a funcao _molde<NomeDoMolde>_().
// O mapeamento so e gravado se a recriacao de cada parte for identica ao modelo.
const fs=require('node:fs');
const path=require('node:path');
const zlib=require('node:zlib');
const mx=require('./molde-node.cjs');

const [arquivo,nome]=process.argv.slice(2);
if(!arquivo||!/^[A-Z][A-Za-z0-9]*$/.test(nome||'')){
  console.error('Uso: node tools/mapear-molde.cjs <modelo.xlsx> <NomeDoMolde>');
  process.exit(1);
}
const buf=fs.readFileSync(arquivo);
const ctx=mx.carregarApps(['MoldeXlsx.js']);
const arquivos=mx.arquivosDoZip(buf);
const hashes={};
for(const a of ctx._mxSemCalcChain_(arquivos))hashes[a.nome]=mx.sha256(mx.bytesDoArquivo(a));
const origem={arquivo:path.basename(arquivo),bytes:buf.length,sha256:mx.sha256(buf),
  mapeadoEm:new Date().toISOString().slice(0,10),hashes};
const molde=ctx._mxMapear_(arquivos,origem);
const b64=zlib.gzipSync(Buffer.from(JSON.stringify(molde),'utf8'),{level:9}).toString('base64');
const linhas=b64.match(/.{1,20000}/g);
const fn='_molde'+nome+'_';
const js='// GERADO por tools/mapear-molde.cjs - NAO EDITAR A MAO.\n'+
  '// Modelo: '+origem.arquivo+' ('+origem.bytes+' bytes, sha256 '+origem.sha256+')\n'+
  '// Mapeado em '+origem.mapeadoEm+'. Para atualizar, rode a ferramenta com o novo Excel.\n'+
  'function '+fn+'() {\n  return _mxDecodificar_([\n'+
  linhas.map(l=>"'"+l+"'").join(',\n')+'\n  ].join(\'\'));\n}\n';
const saida=path.join(mx.RAIZ,'Molde'+nome+'.js');
fs.writeFileSync(saida,js);
console.log('OK',path.relative(mx.RAIZ,saida),(js.length/1024|0)+' KB',
  Object.keys(molde.abas).length+' abas',Object.keys(hashes).length+' partes conferidas');
