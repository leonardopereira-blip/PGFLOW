// Compacta somente as linhas repetidas alem da capacidade de 300 SKUs no
// modelo Forma Certa, sem alterar o Excel recebido. Qualquer dado inesperado
// interrompe a operacao. Todas as outras partes seguem identicas ao original
// (exceto calcChain, que o Excel reconstroi).
// Uso: node tools/mapear-forma-certa.cjs <original.xlsx> [copia-compactada.xlsx]
const fs=require('node:fs');
const path=require('node:path');
const zlib=require('node:zlib');
const {Readable}=require('node:stream');
const {StringDecoder}=require('node:string_decoder');
const crypto=require('node:crypto');
const {execFileSync}=require('node:child_process');
const mx=require('./molde-node.cjs');

const [entrada,saidaArg]=process.argv.slice(2);
if(!entrada)throw Error('Uso: node tools/mapear-forma-certa.cjs <original.xlsx> [copia-compactada.xlsx]');
const saida=saidaArg||path.join(mx.RAIZ,'work','forma-certa-compactado.xlsx');
if(path.resolve(entrada)===path.resolve(saida))throw Error('A copia compactada precisa ser um arquivo diferente da fonte');
const origem=fs.readFileSync(entrada);
let eocd=origem.length-22;
while(eocd>=0&&origem.readUInt32LE(eocd)!==0x06054b50)eocd--;
if(eocd<0)throw Error('Fonte nao e XLSX/ZIP');
let pos=origem.readUInt32LE(eocd+16);
const entradas=[];
for(let k=0;k<origem.readUInt16LE(eocd+10);k++){
  const ln=origem.readUInt16LE(pos+28),le=origem.readUInt16LE(pos+30),lc=origem.readUInt16LE(pos+32);
  const off=origem.readUInt32LE(pos+42),tam=origem.readUInt32LE(pos+20);
  const nome=origem.toString('utf8',pos+46,pos+46+ln);
  const inicio=off+30+origem.readUInt16LE(off+26)+origem.readUInt16LE(off+28);
  entradas.push({nome,metodo:origem.readUInt16LE(pos+10),dados:origem.subarray(inicio,inicio+tam)});
  pos+=46+ln+le+lc;
}

(async()=>{
  const parte='xl/worksheets/sheet2.xml',imput=entradas.find(e=>e.nome===parte);
  if(!imput||imput.metodo!==8)throw Error('Aba IMPUT nao corresponde ao modelo esperado');
  const decoder=new StringDecoder('utf8'),hash=crypto.createHash('sha256');
  let pendente='',inicio='',aberta=false,anterior=0,removidas=0;
  const linhas=[];
  const consumir=texto=>{
    pendente+=texto;
    if(!aberta){
      const i=pendente.indexOf('<sheetData>');if(i<0)return;
      inicio=pendente.slice(0,i+11);pendente=pendente.slice(i+11);aberta=true;
      if(!inicio.includes('<dimension ref="A1:BW1048576"/>'))throw Error('Dimensao IMPUT mudou; revisar o modelo');
    }
    const re=/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g;let m,fim=0;
    while((m=re.exec(pendente))){
      if(m.index!==fim)throw Error('Trecho nao reconhecido entre linhas de IMPUT');
      const row=m[0],rn=Number(/\br="(\d+)"/.exec(row)[1]);
      if(rn!==anterior+1)throw Error('Sequencia de linhas IMPUT mudou');anterior=rn;
      if(rn<=305)linhas.push(row);
      else{
        const esperado='<row r="'+rn+'" spans="34:34" ht="15" customHeight="1" x14ac:dyDescent="0.3"><c r="AH'+rn+'" s="680"><f>_xlfn.IFNA(_xlfn.XLOOKUP($D'+rn+',\'CALC Geral\'!$D:$D,\'CALC Geral\'!$CO:$CO),0)</f><v>0</v></c></row>';
        if(row!==esperado)throw Error('Dado inesperado na linha '+rn+'; fonte preservada, compactacao cancelada');
        removidas++;
      }
      fim=re.lastIndex;
    }
    pendente=pendente.slice(fim);
  };
  for await(const pedaco of Readable.from([imput.dados]).pipe(zlib.createInflateRaw())){
    hash.update(pedaco);consumir(decoder.write(pedaco));
  }
  consumir(decoder.end());
  if(anterior!==1048576||!pendente.startsWith('</sheetData>'))throw Error('Estrutura IMPUT mudou; revisar o modelo');
  const xml=inicio.replace('<dimension ref="A1:BW1048576"/>','<dimension ref="A1:BW305"/>')+linhas.join('')+pendente;
  const partes=mx.lerZip(origem,n=>n!==parte&&n!=='xl/calcChain.xml').map(p=>{
    if(p.nome==='[Content_Types].xml')return {...p,dados:Buffer.from(p.dados.toString('utf8').replace(/<Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/,''),'utf8')};
    if(p.nome==='xl/_rels/workbook.xml.rels')return {...p,dados:Buffer.from(p.dados.toString('utf8').replace(/<Relationship\b[^>]*Target="[^"]*calcChain\.xml"[^>]*\/>/,''),'utf8')};
    return p;
  });
  partes.push({nome:parte,dados:Buffer.from(xml,'utf8')});
  fs.mkdirSync(path.dirname(saida),{recursive:true});
  const compactado=mx.escreverZip(partes);fs.writeFileSync(saida,compactado);
  execFileSync(process.execPath,[path.join(__dirname,'mapear-molde.cjs'),saida,'FormaCertaCalc'],{stdio:'inherit'});
  const ctx=mx.carregarApps(['MoldeXlsx.js','MoldeFormaCertaCalc.js']);
  const molde=ctx._moldeFormaCertaCalc_();
  molde.origem.arquivo=path.basename(entrada);molde.origem.bytes=origem.length;molde.origem.sha256=mx.sha256(origem);
  molde.origem.normalizacao={parte,limiteLinha:305,linhasOriginais:anterior,linhasRemovidas:removidas,
    sha256Original:hash.digest('hex'),sha256Compactado:mx.sha256(Buffer.from(xml,'utf8')),
    motivo:'Somente formulas AH repetidas nas linhas 306:1048576, fora da capacidade de 300 SKUs. Linhas 1:305, estilos e demais partes preservados.'};
  const b64=zlib.gzipSync(Buffer.from(JSON.stringify(molde),'utf8'),{level:9}).toString('base64');
  const js='// GERADO por tools/mapear-forma-certa.cjs - NAO EDITAR A MAO.\n'+
    '// Modelo: '+molde.origem.arquivo+' ('+origem.length+' bytes, sha256 '+molde.origem.sha256+')\n'+
    '// IMPUT: preservadas linhas 1:305; '+removidas+' linhas repetidas AH fora da capacidade removidas na copia.\n'+
    '// Mapeado em '+molde.origem.mapeadoEm+'. Para atualizar, rode a ferramenta com o novo Excel.\n'+
    'function _moldeFormaCertaCalc_() {\n  return _mxDecodificar_([\n'+
    b64.match(/.{1,20000}/g).map(l=>"'"+l+"'").join(',\n')+'\n  ].join(\'\'));\n}\n';
  fs.writeFileSync(path.join(mx.RAIZ,'MoldeFormaCertaCalc.js'),js);
  console.log('Forma Certa: '+removidas+' formulas repetidas removidas da copia; original preservado e identificado por SHA-256.');
})().catch(e=>{console.error(e);process.exitCode=1;});
