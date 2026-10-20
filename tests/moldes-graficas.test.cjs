const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const os=require('node:os');
const crypto=require('node:crypto');
const zlib=require('node:zlib');
const {Readable}=require('node:stream');
const {StringDecoder}=require('node:string_decoder');
const mx=require('../tools/molde-node.cjs');

const MOLDES=['CasaDaGrafica','FormaCerta','Meta','Reproset'].map(n=>({
  nome:n,arquivo:'Molde'+n+'Calc.js',fn:'_molde'+n+'Calc_'
}));
// Sem DriveApp/SpreadsheetApp: os moldes recriam layout, precos e formulas
// usando apenas os dados incluidos no codigo.
const ctx=mx.carregarApps(['MoldeXlsx.js',...MOLDES.map(m=>m.arquivo)]);
for(const m of MOLDES)m.molde=ctx[m.fn]();

for(const m of MOLDES){
  test(m.nome+': todas as partes do Excel embutido conferem pelo SHA-256',()=>{
    const gerado=ctx._mxGerarArquivos_(m.molde),hashes=m.molde.origem.hashes;
    assert.deepEqual(Array.from(gerado,a=>a.nome),Object.keys(hashes));
    assert.ok(Object.keys(m.molde.abas).length>=12,'faltam abas da calculadora');
    assert.ok(hashes['xl/styles.xml'],'faltam os estilos originais');
    assert.ok(hashes['xl/sharedStrings.xml'],'faltam os textos originais');
    assert.ok(!hashes['xl/calcChain.xml'],'calcChain deve ser reconstruido pelo Excel');
    for(const parte of gerado)
      assert.equal(mx.sha256(mx.bytesDoArquivo(parte)),hashes[parte.nome],parte.nome+' mudou');
  });
}

const pastaFontes=process.env.MOLDE_FONTES||path.join(os.homedir(),'Desktop','Calculadoras_Reentrada');
for(const m of MOLDES){
  const fonte=path.join(pastaFontes,m.molde.origem.arquivo);
  test(m.nome+': formulas, precos, estilos e demais partes sao identicos ao Excel recebido',
    {skip:!fs.existsSync(fonte)&&'original ausente; defina MOLDE_FONTES para conferir contra a fonte'},()=>{
      const original=fs.readFileSync(fonte);
      assert.equal(mx.sha256(original),m.molde.origem.sha256,'fonte local difere da versao recebida');
      assert.equal(original.length,m.molde.origem.bytes);
      const normalizacao=m.molde.origem.normalizacao;
      // A aba IMPUT Forma Certa e conferida separadamente por streaming, para
      // nao manter seu XML original de 213 MB inteiro em memoria.
      const orig=mx.lerZip(original,n=>n!=='xl/calcChain.xml'&&n!==normalizacao?.parte);
      const gerado=new Map(ctx._mxGerarArquivos_(m.molde).map(a=>[a.nome,a]));
      for(const {nome,dados} of orig){
        const parte=gerado.get(nome);assert.ok(parte,nome+' ausente');
        let esperado=dados;
        if(nome==='[Content_Types].xml'||nome==='xl/_rels/workbook.xml.rels'){
          esperado=Buffer.from(dados.toString('utf8').replace(/<Override\b[^>]*calcChain\.xml"[^>]*\/>|<Relationship\b[^>]*calcChain\.xml"[^>]*\/>/,''),'utf8');
        }
        assert.equal(mx.sha256(mx.bytesDoArquivo(parte)),mx.sha256(esperado),nome+' difere do original');
      }
      assert.equal(gerado.size,orig.length+(normalizacao?1:0),'partes inesperadas no Excel recriado');
      assert.equal(gerado.has('xl/calcChain.xml'),false);
      if(m.nome!=='FormaCerta')assert.equal(normalizacao,undefined,'normalizacao nao autorizada para '+m.nome);
    });
}

// Localiza a entrada comprimida sem inflar os outros 110 MB de planilhas.
function entradaZip(buf,nome){
  let i=buf.length-22;while(i>=0&&buf.readUInt32LE(i)!==0x06054b50)i--;
  assert.ok(i>=0,'ZIP invalido');let p=buf.readUInt32LE(i+16);
  for(let k=0;k<buf.readUInt16LE(i+10);k++){
    const ln=buf.readUInt16LE(p+28),le=buf.readUInt16LE(p+30),lc=buf.readUInt16LE(p+32);
    const n=buf.toString('utf8',p+46,p+46+ln),off=buf.readUInt32LE(p+42);
    if(n===nome){
      assert.equal(buf.readUInt16LE(p+10),8);
      const inicio=off+30+buf.readUInt16LE(off+26)+buf.readUInt16LE(off+28);
      return buf.subarray(inicio,inicio+buf.readUInt32LE(p+20));
    }
    p+=46+ln+le+lc;
  }
  assert.fail('Parte ZIP ausente: '+nome);
}

const forma=MOLDES.find(m=>m.nome==='FormaCerta').molde;
const fonteForma=path.join(pastaFontes,forma.origem.arquivo);
test('Forma Certa: unica compactacao preserva cada byte das linhas 1:305 e remove somente AH repetida fora da capacidade',
  {skip:!fs.existsSync(fonteForma)&&'original ausente; defina MOLDE_FONTES'},async()=>{
    const n=forma.origem.normalizacao;
    assert.equal(n.parte,'xl/worksheets/sheet2.xml');assert.equal(n.limiteLinha,305);
    const decoder=new StringDecoder('utf8'),hash=crypto.createHash('sha256');
    let pendente='',prefixo='',entrou=false,anterior=0,removidas=0;
    const preservadas=[];
    const consumir=texto=>{
      pendente+=texto;
      if(!entrou){
        const i=pendente.indexOf('<sheetData>');if(i<0)return;
        prefixo=pendente.slice(0,i+11);pendente=pendente.slice(i+11);entrou=true;
      }
      const re=/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g;let m,pos=0;
      while((m=re.exec(pendente))){
        assert.equal(m.index,pos,'conteudo nao reconhecido entre linhas');
        const row=m[0],rn=Number(/\br="(\d+)"/.exec(row)[1]);
        assert.equal(rn,anterior+1);anterior=rn;
        if(rn<=305)preservadas.push(row);
        else{
          // Verifica texto, estilo, formula relativa e valor cacheado de TODAS
          // as 1.048.271 linhas removidas, sem presumir que estavam vazias.
          const esperado='<row r="'+rn+'" spans="34:34" ht="15" customHeight="1" x14ac:dyDescent="0.3"><c r="AH'+rn+'" s="680"><f>_xlfn.IFNA(_xlfn.XLOOKUP($D'+rn+',\'CALC Geral\'!$D:$D,\'CALC Geral\'!$CO:$CO),0)</f><v>0</v></c></row>';
          assert.equal(row,esperado,'linha removida '+rn+' contem outro dado');removidas++;
        }
        pos=re.lastIndex;
      }
      pendente=pendente.slice(pos);
    };
    const buf=fs.readFileSync(fonteForma),comprimido=entradaZip(buf,n.parte);
    for await(const pedaco of Readable.from([comprimido]).pipe(zlib.createInflateRaw())){
      hash.update(pedaco);consumir(decoder.write(pedaco));
    }
    consumir(decoder.end());
    assert.ok(pendente.startsWith('</sheetData>'));
    assert.equal(anterior,1048576);assert.equal(anterior,n.linhasOriginais);
    assert.equal(removidas,1048271);assert.equal(removidas,n.linhasRemovidas);
    assert.equal(hash.digest('hex'),n.sha256Original);
    assert.match(prefixo,/<dimension ref="A1:BW1048576"\/>/);
    const esperado=prefixo.replace('<dimension ref="A1:BW1048576"/>','<dimension ref="A1:BW305"/>')+
      preservadas.join('')+pendente;
    const reconstruido=ctx._mxGerarAba_(forma.abas[n.parte]);
    assert.equal(reconstruido,esperado,'as linhas preservadas, cabecalho ou configuracoes da aba mudaram');
    assert.equal(mx.sha256(Buffer.from(reconstruido,'utf8')),n.sha256Compactado);
    assert.equal(n.sha256Compactado,forma.origem.hashes[n.parte]);
  });
