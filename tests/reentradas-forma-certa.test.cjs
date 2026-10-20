const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const ctx={};vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname,'..','ReentradasFormaCerta.js'),'utf8'),ctx);

function origem(){
  const abertura='<worksheet><sheetData>';
  const mestre='<row r="3"><c r="A3"><f t="shared" ref="A3:A30" si="5">IF(D3="","",ROW()-2)</f><v>1</v></c></row>';
  const modelo='<row r="30" ht="15"><c r="A30" s="1"><f t="shared" si="5"/><v>28</v></c><c r="D30" s="2"><f>\'CALC Geral\'!D33</f><v>ANTIGO</v></c><c r="AH30" s="3"><f>XLOOKUP($D30,\'CALC Geral\'!$D:$D,\'CALC Geral\'!$CO:$CO)</f><v>9</v></c><c r="B30" s="4"/></row>';
  const extras=Array.from({length:275},(_,i)=>'<row r="'+(31+i)+'"><c r="AH'+(31+i)+'"><f>0</f></c></row>').join('');
  return abertura+mestre+modelo+extras+'</sheetData><dataValidations><dataValidation sqref="AL3:AL103 AX3:AX103"/></dataValidations></worksheet>';
}

test('Forma Certa completa IMPUT para 300 SKUs com referencias relativas e formulas compartilhadas resolvidas',()=>{
  const res=ctx._reFcExpandirImput_(origem(),300);
  const final=res.match(/<row r="302"[^>]*>[\s\S]*?<\/row>/)[0];
  assert.match(final,/<c r="D302" s="2"><f>'CALC Geral'!D305<\/f>/);
  assert.match(final,/<c r="A302" s="1"><f>IF\(D302="","",ROW\(\)-2\)<\/f>/);
  assert.match(final,/XLOOKUP\(\$D302,'CALC Geral'!\$D:\$D,'CALC Geral'!\$CO:\$CO\)/);
  assert.doesNotMatch(final,/ANTIGO|<v>|t="shared"/);
  assert.match(res,/sqref="AL3:AL302 AX3:AX302"/);
  assert.equal(res.match(/<row r="30"[^>]*>[\s\S]*?<\/row>/)[0],origem().match(/<row r="30"[^>]*>[\s\S]*?<\/row>/)[0]);
});

test('ate 28 SKUs preserva IMPUT e fontes nao sao alteradas',()=>{
  const xml=origem();assert.equal(ctx._reFcExpandirImput_(xml,28),xml);
  const res=ctx._reFcExpandirImput_(xml,29);assert.notEqual(res,xml);
  assert.match(res,/<c r="D31" s="2"><f>'CALC Geral'!D34<\/f>/);
  assert.equal(ctx._reFcDeslocarFormula_('IF(D30="D30",$D$2,D$2+$D30)',5),'IF(D35="D30",$D$2,D$2+$D35)');
});

test('limite e origem compartilhada invalidos sao erros explicitos',()=>{
  assert.throws(()=>ctx._reFcExpandirImput_(origem(),301),/300/);
  assert.throws(()=>ctx._reFcExpandirImput_(origem().replace('si="5">','si="6">'),29),/compartilhada/);
});
