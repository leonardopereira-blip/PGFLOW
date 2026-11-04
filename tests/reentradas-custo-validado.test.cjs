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
    getLastUpdated:()=>new Date(out.version),getDateCreated:()=>new Date(out.version-5000),isTrashed:()=>false};
  let blob=null;
  const storage={getFilesByName:()=>({hasNext:()=>!!blob,next:()=>({getId:()=> 'REGISTRO_1234567890',getBlob:()=>blob})}),
    createFile:nova=>{blob=nova;out.writes++;return {setName:()=>{}};}};
  const context={RX_PASTA:'PASTA_1234567890',RX_MIME:file.getMimeType(),
    _rxRevisaoPasta_:()=>storage,
    _rxRevisaoLer_:()=>{if(!out.modeloPronto)throw Error('Revisao nao preparada.');return {abas:[{
      nome:'CALC Geral',linhas:[
        {numero:3,valores:{3:'CÓDIGO SKU',5:'AVULSO ACAB',6:'AVULSO GRAF',7:'KIT S/ MALETA',8:'KIT C/ MALETA',9:'TOTAL',45:'SERVIÇO + PAPEL TOTAL'}},
        {numero:4,valores:{9:'Exemplares',46:'BRL / Total'}},
        {numero:6,valores:{3:'SKU-01',5:'800',6:'0',7:'0',8:'0',9:'800',46:'4920.5'}}]}]};},
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
  const gerado=h.context._rxCustoResumoChamados_()['3127'];
  assert.equal(gerado.ultimoArquivoId,id,'arquivo gerado aparece mesmo sem custo validado');
  assert.equal(gerado.ultimaValidacaoEm,undefined);
  const result=h.context.marcarCustoValidadoReentradas(id,h.out.version);
  assert.equal(result.sucesso,true);
  assert.equal(result.chamado,'3127');
  assert.ok(Number.isSafeInteger(result.validadoEm));
  assert.match(result.downloadUrl,/^https:\/\/drive\.google\.com\/uc\?export=download&id=/);
  assert.equal(h.out.writes,1);
  const registro=h.context._rxCustoRegistros_().arquivos[id];
  assert.equal(registro.atualizadoEm,h.out.version);
  assert.equal(registro.validadoEm,result.validadoEm);
  assert.equal(registro.tiragemValidada,800);
  assert.equal(registro.valorChamado,4920.5);
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
  assert.equal(resumo['3127'].ultimoGeradoEm,h.out.version-5000);
  assert.equal(resumo['3127'].ultimoArquivoId,id);
  assert.equal(resumo['3127'].tiragemValidada,800);
  assert.equal(resumo['3127'].valorChamado,4920.5);
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

test('não certifica tiragem incoerente ou preço final sem cálculo no XLSX',()=>{
  const h=harness(),ler=h.context._rxRevisaoLer_;
  h.context._rxRevisaoLer_=()=>{const m=ler();m.abas[0].linhas[2].valores[9]='30';return m};
  assert.match(h.context.marcarCustoValidadoReentradas(id,h.out.version).erro,/diverge da tiragem/);
  h.context._rxRevisaoLer_=()=>{const m=ler();m.abas[0].linhas[2].valores[46]='=AT6*J6';return m};
  assert.match(h.context.marcarCustoValidadoReentradas(id,h.out.version).erro,/nao possui resultado numerico/);
  assert.equal(h.out.writes,0);
});

test('molde Forma Certa soma tres tiragens sem confundir descricao com exemplar',()=>{
  const h=harness(),ler=h.context._rxRevisaoLer_;
  h.context._rxRevisaoLer_=()=>{const m=ler(),aba=m.abas[0];
    aba.linhas[0].valores={3:'CÓDIGO SKU',5:'AVULSO GRAF',6:'KIT S/ MALETA',7:'KIT C/ MALETA',8:'TOTAL',41:'SERVIÇO + PAPEL TOTAL'};
    aba.linhas[1].valores={8:'Exemplares',42:'BRL / Total'};
    aba.linhas[2].valores={3:'SKU-01',4:'Descrição do item',5:'200',6:'500',7:'100',8:'800',42:'7020.25'};
    return m;};
  const r=h.context.marcarCustoValidadoReentradas(id,h.out.version);
  assert.equal(r.sucesso,true,r.erro);
  assert.equal(r.tiragemValidada,800);
  assert.equal(r.valorChamado,7020.25);
});
