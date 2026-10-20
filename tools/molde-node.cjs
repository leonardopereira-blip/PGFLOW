// Suporte Node para MoldeXlsx.js: le/escreve .xlsx (zip) e carrega o codigo do
// Apps Script com um Utilities minimo. Usado pela ferramenta de mapeamento e testes.
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const zlib=require('node:zlib');
const crypto=require('node:crypto');

const RAIZ=path.join(__dirname,'..');

function lerZip(buf,filtro) {
  let i=buf.length-22;
  while(i>=0&&buf.readUInt32LE(i)!==0x06054b50)i--;
  if(i<0)throw Error('Arquivo nao e um .xlsx/zip valido');
  const total=buf.readUInt16LE(i+10);let p=buf.readUInt32LE(i+16);const out=[];
  for(let k=0;k<total;k++){
    const metodo=buf.readUInt16LE(p+10),tam=buf.readUInt32LE(p+20);
    const ln=buf.readUInt16LE(p+28),le=buf.readUInt16LE(p+30),lc=buf.readUInt16LE(p+32),off=buf.readUInt32LE(p+42);
    const nome=buf.toString('utf8',p+46,p+46+ln);
    const ini=off+30+buf.readUInt16LE(off+26)+buf.readUInt16LE(off+28);
    const d=buf.subarray(ini,ini+tam);
    if(!filtro||filtro(nome))out.push({nome,dados:metodo===0?Buffer.from(d):zlib.inflateRawSync(d)});
    p+=46+ln+le+lc;
  }
  return out;
}

function escreverZip(partes) {
  const locais=[],central=[];let off=0;
  for(const {nome,dados} of partes){
    const n=Buffer.from(nome,'utf8'),comp=zlib.deflateRawSync(dados,{level:6}),crc=zlib.crc32(dados);
    const h=Buffer.alloc(30);
    h.writeUInt32LE(0x04034b50,0);h.writeUInt16LE(20,4);h.writeUInt16LE(0x0800,6);h.writeUInt16LE(8,8);
    h.writeUInt32LE(crc,14);h.writeUInt32LE(comp.length,18);h.writeUInt32LE(dados.length,22);h.writeUInt16LE(n.length,26);
    const c=Buffer.alloc(46);
    c.writeUInt32LE(0x02014b50,0);c.writeUInt16LE(20,4);c.writeUInt16LE(20,6);c.writeUInt16LE(0x0800,8);c.writeUInt16LE(8,10);
    c.writeUInt32LE(crc,16);c.writeUInt32LE(comp.length,20);c.writeUInt32LE(dados.length,24);c.writeUInt16LE(n.length,28);c.writeUInt32LE(off,42);
    locais.push(h,n,comp);central.push(c,n);off+=30+n.length+comp.length;
  }
  const cd=Buffer.concat(central),fim=Buffer.alloc(22);
  fim.writeUInt32LE(0x06054b50,0);fim.writeUInt16LE(partes.length,8);fim.writeUInt16LE(partes.length,10);
  fim.writeUInt32LE(cd.length,12);fim.writeUInt32LE(off,16);
  return Buffer.concat([...locais,cd,fim]);
}

// Partes de texto viram string so se a volta para bytes for exata.
function arquivosDoZip(buf) {
  return lerZip(buf).map(({nome,dados})=>{
    if(/\.(xml|rels|vml)$/i.test(nome)){
      const texto=dados.toString('utf8');
      if(Buffer.from(texto,'utf8').equals(dados))return {nome,texto};
    }
    return {nome,base64:dados.toString('base64')};
  });
}

function bytesDoArquivo(a) {
  return a.base64!==undefined?Buffer.from(a.base64,'base64'):Buffer.from(a.texto,'utf8');
}
const sha256=b=>crypto.createHash('sha256').update(b).digest('hex');

function utilitiesNode() {
  const blob=(bytes,tipo,nome)=>({
    getBytes:()=>bytes,getName:()=>nome,getContentType:()=>tipo,
    getDataAsString:()=>bytes.toString('utf8'),
    copyBlob(){return blob(bytes,tipo,nome);},setContentType(t){tipo=t;return this;}
  });
  return {
    newBlob:(d,tipo,nome)=>blob(Buffer.isBuffer(d)?d:Array.isArray(d)?Buffer.from(d):Buffer.from(String(d),'utf8'),tipo,nome),
    base64Decode:s=>Buffer.from(s,'base64'),
    base64Encode:b=>Buffer.from(b).toString('base64'),
    gzip:b=>blob(zlib.gzipSync(b.getBytes(),{level:9}),'application/x-gzip'),
    ungzip:b=>blob(zlib.gunzipSync(b.getBytes()),'application/octet-stream')
  };
}

// Carrega arquivos .js do projeto Apps Script num mesmo contexto global.
function carregarApps(arquivos,extra) {
  const ctx=Object.assign({Utilities:utilitiesNode()},extra||{});
  vm.createContext(ctx);
  for(const f of arquivos)vm.runInContext(fs.readFileSync(path.join(RAIZ,f),'utf8'),ctx,{filename:f});
  return ctx;
}

module.exports={RAIZ,lerZip,escreverZip,arquivosDoZip,bytesDoArquivo,sha256,utilitiesNode,carregarApps};
