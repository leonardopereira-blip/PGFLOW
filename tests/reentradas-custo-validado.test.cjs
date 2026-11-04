const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const source=fs.readFileSync(path.join(__dirname,'..','ReentradasValidacaoFinal.js'),'utf8');
const id='ARQUIVO_CALCULADORA_1234567890';

function harness(){
  const out={writes:0,modeloPronto:true,version:1730000000000};
  const nome='CALCULADORA FORNECEDOR LOGPRINT - LIVROS REENTRADA 2026 - CHAMADO 3127.xlsx';
  const file={id,nome,getId:()=>id,getName:()=>nome,getMimeType:()=> 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    getLastUpdated:()=>new Date(out.version),isTrashed:()=>false};
  let blob=null;
  const storage={getFilesByName:()=>({hasNext:()=>!!blob,next:()=>({getId:()=> 'REGISTRO_1234567890',getBlob:()=>blob})}),
    createFile:nova=>{blob=nova;out.writes++;return {setName:()=>{}};}};
  const context={RX_PASTA:'PASTA_1234567890',RX_MIME:file.getMimeType(),
    _rxRevisaoPasta_:()=>storage,
    _rxRevisaoLer_:()=>{if(!out.modeloPronto)throw Error('Revisao nao preparada.');},
    _rxArquivo_:(fileId,versao)=>{if(fileId!==id)throw Error('Arquivo desconhecido');
      if(versao!==undefined&&Number(versao)!==out.version)throw Error('Arquivo alterado no Drive.');
      return {id,nome,atualizadoEm:out.version,calculadora:true};},
    Utilities:{newBlob:(content,type,name)=>({getDataAsString:()=>content,setName:()=>{},name,type})},
    Drive:{Files:{update:(_meta,_id,nova)=>{out.writes++;blob=nova;}}},
    DriveApp:{getFolderById:()=>({getFiles:()=>{let done=false;return {hasNext:()=>!done,next:()=>{done=true;return file;}};}})},
    LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock:()=>{}})}
  };
  vm.createContext(context);vm.runInContext(source,context);
  return {context,out,file};
}

test('geracao e abertura nao registram aceite; somente clique explicito salva versao e data',()=>{
  const h=harness();
  assert.deepEqual(Object.keys(h.context._rxCustoRegistros_().arquivos),[]);
  assert.equal(h.out.writes,0);
  const result=h.context.marcarCustoValidadoReentradas(id,h.out.version);
  assert.equal(result.sucesso,true);
  assert.equal(result.chamado,'3127');
  assert.ok(Number.isSafeInteger(result.validadoEm));
  assert.match(result.downloadUrl,/^https:\/\/drive\.google\.com\/uc\?export=download&id=/);
  assert.equal(h.out.writes,1);
  const registro=h.context._rxCustoRegistros_().arquivos[id];
  assert.equal(registro.atualizadoEm,h.out.version);
  assert.equal(registro.validadoEm,result.validadoEm);
});

test('lista so oferece download da versao validada; edicao invalida o aceite sem apagar historico',()=>{
  const h=harness();h.context.marcarCustoValidadoReentradas(id,h.out.version);
  let arquivos=[{id,nome:h.file.getName(),atualizadoEm:h.out.version,calculadora:true}];
  h.context._rxCustoListaDeArquivos_(arquivos);
  assert.ok(arquivos[0].downloadUrl);
  assert.ok(arquivos[0].validadoEm);
  h.out.version++;
  arquivos=[{id,nome:h.file.getName(),atualizadoEm:h.out.version,calculadora:true}];
  h.context._rxCustoListaDeArquivos_(arquivos);
  assert.equal(arquivos[0].downloadUrl,undefined);
  const resumo=h.context._rxCustoResumoChamados_();
  assert.ok(resumo['3127'].ultimaValidacaoEm);
  assert.equal(resumo['3127'].downloadUrl,'');
});

test('recusa marcar arquivo desatualizado ou sem revisao preparada e nao modifica o registro',()=>{
  const h=harness();
  assert.match(h.context.marcarCustoValidadoReentradas(id,h.out.version-1).erro,/alterado/);
  h.out.modeloPronto=false;
  assert.match(h.context.marcarCustoValidadoReentradas(id,h.out.version).erro,/nao preparada/);
  assert.equal(h.out.writes,0);
  assert.deepEqual(Object.keys(h.context._rxCustoRegistros_().arquivos),[]);
});
