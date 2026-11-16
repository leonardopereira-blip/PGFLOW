const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const mx=require('../tools/molde-node.cjs');

const ctx=mx.carregarApps(['MoldeXlsx.js','MoldePosigrafPrincipalCalc.js']);
const molde=ctx._moldePosigrafPrincipalCalc_();

test('Posigraf (Principal): cada parte reconstruida confere com o SHA-256 do molde recebido',()=>{
  const gerado=ctx._mxGerarArquivos_(molde),hashes=molde.origem.hashes;
  assert.deepEqual(Array.from(gerado,p=>p.nome),Object.keys(hashes));
  assert.equal(Object.keys(molde.abas).length,12,'faltam abas da calculadora principal');
  for(const nome of ['xl/styles.xml','xl/sharedStrings.xml','xl/workbook.xml'])
    assert.ok(hashes[nome],nome+' ausente');
  assert.equal(hashes['xl/calcChain.xml'],undefined,'o Excel deve reconstruir a cadeia de calculo');
  for(const parte of gerado)
    assert.equal(mx.sha256(mx.bytesDoArquivo(parte)),hashes[parte.nome],parte.nome+' mudou');
});

const fonte=process.env.POSIGRAF_FONTE||
  'G:/Drives compartilhados/Calculadoras Individuais/Posigraf/V1-27 validados/'+
  'CALCULADORA POSIGRAF - LIVROS PRINCIPAIS COL 2027.xlsx';
test('Posigraf (Principal): formulas, precos, estilos e demais partes sao identicos ao Excel original',
  {skip:!fs.existsSync(fonte)&&'original ausente; defina POSIGRAF_FONTE para conferir a fonte'},()=>{
    const original=fs.readFileSync(fonte);
    assert.equal(original.length,molde.origem.bytes,'tamanho da fonte difere do arquivo recebido');
    assert.equal(mx.sha256(original),molde.origem.sha256,'a fonte nao e a versao mapeada');
    const partes=mx.lerZip(original),gerado=new Map(ctx._mxGerarArquivos_(molde)
      .map(p=>[p.nome,mx.bytesDoArquivo(p)]));
    for(const {nome,dados} of partes){
      if(nome==='xl/calcChain.xml'){
        assert.equal(gerado.has(nome),false);
        continue;
      }
      let esperado=dados;
      if(nome==='[Content_Types].xml'||nome==='xl/_rels/workbook.xml.rels')
        esperado=Buffer.from(dados.toString('utf8').replace(
          /<Override\b[^>]*calcChain\.xml"[^>]*\/>|<Relationship\b[^>]*calcChain\.xml"[^>]*\/>/,''),'utf8');
      assert.ok(gerado.has(nome),nome+' ausente');
      assert.ok(gerado.get(nome).equals(esperado),nome+' difere do original');
    }
    assert.equal(gerado.size,partes.length-(partes.some(p=>p.nome==='xl/calcChain.xml')?1:0),
      'partes inesperadas no arquivo reconstruido');
    assert.equal(molde.origem.normalizacao,undefined,'nenhuma normalizacao adicional autorizada');
  });
