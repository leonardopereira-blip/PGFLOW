const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const source=fs.readFileSync(path.join(__dirname,'..','ReentradasValidacaoFinal.js'),'utf8');
const id='ARQUIVO_CALCULADORA_1234567890';

function harness(){
  const out={writes:0,modeloPronto:true,version:1730000000000,copies:[],savedFiles:[]};
  const nome='CALCULADORA FORNECEDOR LOGPRINT - LIVROS REENTRADA 2026 - CHAMADO 3127.xlsx';
  const file={id,nome,getId:()=>id,getName:()=>nome,getMimeType:()=> 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    getLastUpdated:()=>new Date(out.version),getDateCreated:()=>new Date(out.version-5000),isTrashed:()=>false,
    makeCopy:(n,folder)=>{out.copies.push(n);return {getId:()=> 'COPIA_CALCULADORA_12345'};}};
  const depara={getName:()=> 'PLANILHA DE PARA PPG LOGPRINT - CHAMADO 3127.xlsx',
    getMimeType:()=>file.getMimeType(),getDateCreated:()=>new Date(out.version-6000),isTrashed:()=>false,
    getId:()=> 'DEPARA_ORIGINAL_123456',makeCopy:n=>{out.copies.push(n);return {getId:()=> 'COPIA_DEPARA_123456'};}};
  let blob=null;
  const archive={getId:()=> 'PASTA_CONFIRMADA_12345',getUrl:()=> 'https://drive.google.com/drive/folders/PASTA_CONFIRMADA_12345',
    createFile:b=>out.savedFiles.push(b)};
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
      return {id,nome,arquivo:file,atualizadoEm:out.version,calculadora:true};},
    _reSpNormal_:s=>s.toUpperCase(),
    Utilities:{newBlob:(content,type,name)=>({getDataAsString:()=>content,setName:()=>{},name,type})},
    Drive:{Files:{update:(_meta,_id,nova)=>{out.writes++;blob=nova;}}},
    DriveApp:{getFolderById:folderId=>({getFiles:()=>{const entries=[file,depara];let at=0;return {
      hasNext:()=>at<entries.length,next:()=>entries[at++]};},createFolder:()=>archive})},
    Session:{getActiveUser:()=>({getEmail:()=> 'leonardo.pereira@arcoeducacao.com.br'}),
      getEffectiveUser:()=>({getEmail:()=> 'leonardo.pereira@arcoeducacao.com.br'})},
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
  assert.match(h.context.marcarCustoValidadoReentradas(id,h.out.version).erro,/selecao/i);
  const result=h.context.confirmarCustoComEmailReentradas(id,h.out.version,{texto:'Aprovamos os valores do chamado 3127',
    assunto:'Validação de custos',remetente:'grafica@fornecedor.com',confirmacaoExplicita:true,aceitouAnaliseManual:true});
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
  assert.equal(h.out.copies.length,2);
  assert.equal(h.out.savedFiles.length,2);
});

test('lista so oferece download da versao validada; edicao invalida o aceite sem apagar historico',()=>{
  const h=harness();h.context.confirmarCustoComEmailReentradas(id,h.out.version,{texto:'Aprovamos os custos do chamado 3127',
    remetente:'grafica@fornecedor.com',confirmacaoExplicita:true,aceitouAnaliseManual:true});
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
  assert.match(h.context.confirmarCustoComEmailReentradas(id,h.out.version-1,{texto:'Aprovamos os custos',
    remetente:'grafica@fornecedor.com',confirmacaoExplicita:true,aceitouAnaliseManual:true}).erro,/alterado/);
  h.out.modeloPronto=false;
  assert.match(h.context.confirmarCustoComEmailReentradas(id,h.out.version,{texto:'Aprovamos os custos',
    remetente:'grafica@fornecedor.com',confirmacaoExplicita:true,aceitouAnaliseManual:true}).erro,/resultados pendentes/);
  assert.equal(h.out.writes,0);
  assert.deepEqual(Object.keys(h.context._rxCustoRegistros_().arquivos),[]);
});

test('tiragem incoerente e preco sem calculo exigem reconhecimento explicito',()=>{
  const h=harness(),ler=h.context._rxRevisaoLer_;
  h.context._rxRevisaoLer_=()=>{const m=ler();m.abas[0].linhas[2].valores[9]='30';return m};
  assert.match(h.context.confirmarCustoComEmailReentradas(id,h.out.version,{texto:'Aprovamos os custos',
    remetente:'grafica@fornecedor.com',confirmacaoExplicita:true,aceitouAnaliseManual:true}).erro,/diverge da tiragem/);
  h.context._rxRevisaoLer_=()=>{const m=ler();m.abas[0].linhas[2].valores[46]='=AT6*J6';return m};
  assert.match(h.context.confirmarCustoComEmailReentradas(id,h.out.version,{texto:'Aprovamos os custos',
    remetente:'grafica@fornecedor.com',confirmacaoExplicita:true,aceitouAnaliseManual:true}).erro,/nao possui resultado numerico/);
  assert.equal(h.out.writes,0);
  const aceito=h.context.confirmarCustoComEmailReentradas(id,h.out.version,{texto:'Aprovamos os custos',
    remetente:'grafica@fornecedor.com',confirmacaoExplicita:true,aceitouPendencias:true,aceitouAnaliseManual:true});
  assert.equal(aceito.sucesso,true,aceito.erro);assert.match(aceito.pendencias,/nao possui resultado numerico/);
});

test('molde Forma Certa soma tres tiragens sem confundir descricao com exemplar',()=>{
  const h=harness(),ler=h.context._rxRevisaoLer_;
  h.context._rxRevisaoLer_=()=>{const m=ler(),aba=m.abas[0];
    aba.linhas[0].valores={3:'CÓDIGO SKU',5:'AVULSO GRAF',6:'KIT S/ MALETA',7:'KIT C/ MALETA',8:'TOTAL',41:'SERVIÇO + PAPEL TOTAL'};
    aba.linhas[1].valores={8:'Exemplares',42:'BRL / Total'};
    aba.linhas[2].valores={3:'SKU-01',4:'Descrição do item',5:'200',6:'500',7:'100',8:'800',42:'7020.25'};
    return m;};
  const r=h.context.confirmarCustoComEmailReentradas(id,h.out.version,{texto:'Aprovamos os custos',
    remetente:'grafica@fornecedor.com',confirmacaoExplicita:true,aceitouAnaliseManual:true});
  assert.equal(r.sucesso,true,r.erro);
  assert.equal(r.tiragemValidada,800);
  assert.equal(r.valorChamado,7020.25);
});

test('seleciona mensagem da caixa do usuario e arquiva conversa em EML sem interpretar aceite',()=>{
  const h=harness(),m={getId:()=> 'MSG12345',getFrom:()=> 'Contato da grafica <grafica@fornecedor.com>',
    getDate:()=>new Date(1730000000000),getSubject:()=> 'Validação de custos - Chamado 3127',getPlainBody:()=> 'Aprovamos os custos do chamado 3127',
    getRawContent:()=> 'From: grafica@fornecedor.com\r\nSubject: Chamado 3127\r\n\r\nCorpo'};
  const thread={getId:()=> 'THREAD_123456789',getMessages:()=> [m],getFirstMessageSubject:()=> 'Chamado 3127',
    getLastMessageDate:()=>new Date(1730000000000),getPermalink:()=> 'https://mail.google.com/mail/u/0/#all/abc'};
  m.getThread=()=>thread;
  const buscas=[];
  h.context.GmailApp={search:(q,inicio,limite)=>{buscas.push({q,inicio,limite});return [thread];},getMessageById:()=>m};
  const lista=h.context.obterConfirmacoesEmailCustoReentradas(id,h.out.version);
  assert.equal(lista.sucesso,true);assert.equal(lista.emails[0].mensagens[0].id,'MSG12345');
  assert.equal(lista.sugestao.threadId,'THREAD_123456789');
  assert.match(buscas[0].q,/"3127"/);
  const outra=h.context.obterConfirmacoesEmailCustoReentradas(id,h.out.version,'logprint',1);
  assert.equal(outra.pagina,1);assert.equal(buscas[2].q,'in:anywhere logprint');
  assert.equal(buscas[2].inicio,15);
  const registrado=h.context.confirmarCustoComEmailReentradas(id,h.out.version,{mensagemId:'MSG12345',
    threadId:'THREAD_123456789',texto:'Trecho escolhido pelo usuário',confirmacaoExplicita:true});
  assert.equal(registrado.sucesso,true,registrado.erro);
  assert.equal(h.out.savedFiles.length,3);
  assert.equal(h.out.savedFiles.at(-1).type,'message/rfc822');
  assert.match(h.out.savedFiles[0].getDataAsString(),/Trecho escolhido pelo usuário/);
});

test('ignorar sugestao registra uma unica vez por calculadora',()=>{
  const h=harness();
  assert.equal(h.context.ignorarSugestaoEmailCustoReentradas(id,h.out.version).sucesso,true);
  assert.equal(h.out.writes,1);
  assert.equal(h.context.ignorarSugestaoEmailCustoReentradas(id,h.out.version).sucesso,true);
  assert.equal(h.out.writes,1);
  assert.equal(h.context._rxCustoRegistros_().sugestoes[id].decisao,'ignorada');
});

test('examinar sugestao registra visualizacao sem validar custo ou repetir alerta',()=>{
  const h=harness();
  assert.equal(h.context.registrarSugestaoEmailCustoReentradas(id,h.out.version,'examinada').sucesso,true);
  assert.equal(h.context.registrarSugestaoEmailCustoReentradas(id,h.out.version,'examinada').sucesso,true);
  assert.equal(h.out.writes,1);
  const dados=h.context._rxCustoRegistros_();
  assert.equal(dados.sugestoes[id].decisao,'examinada');
  assert.equal(Object.keys(dados.arquivos).length,0);
});

test('permite enviar na thread escolhida manualmente mesmo sem citar o chamado',()=>{
  const h=harness(),enviados=[];
  h.file.getBlob=()=>({setName(){return this;}});
  h.context.GmailApp={getThreadById:()=>({getMessages:()=>[{getSubject:()=> 'Assunto livre do usuario'}],
    replyAll:(mensagem,opcoes)=>enviados.push({mensagem,opcoes})})};
  const r=h.context.enviarCustoPorEmailNaConversaReentradas(id,h.out.version,'THREAD_LIVRE_123456','Segue a calculadora');
  assert.equal(r.sucesso,true,r.erro);
  assert.equal(enviados.length,1);assert.equal(enviados[0].opcoes.attachments.length,1);
});

test('interpreta contexto com prioridade para negativas, pendencias e remetente externo',()=>{
  const h=harness(),avaliar=(texto,assunto='Validação de custos 3127',de='Gráfica <grafica@fornecedor.com>')=>
    h.context._rxCustoClassificarResposta_(texto,assunto,de);
  assert.equal(avaliar('Validamos os valores do chamado 3127, podem seguir.').situacao,'aprovacao');
  assert.equal(avaliar('Custos aprovados sem pendências.').situacao,'aprovacao');
  assert.equal(avaliar('Não aprovado. O valor está incorreto.').situacao,'rejeitada');
  assert.equal(avaliar('Ok, porém temos pendências na tiragem.').situacao,'pendente');
  assert.equal(avaliar('Aprovado, mas precisa corrigir o papel.').situacao,'pendente');
  assert.equal(avaliar('Ok, vamos analisar e retornar.').situacao,'pendente');
  assert.equal(avaliar('OK').situacao,'possivel');
  assert.equal(avaliar('OK','Pedido de prazo: 3127').situacao,'neutra');
  assert.equal(avaliar('Aprovamos os custos.','Validação de custos','Leonardo <leonardo.pereira@arcoeducacao.com.br>').situacao,'interna');
  assert.equal(avaliar('Precisamos revisar.\n\nEm 05/11, gráfica escreveu:\n> Custos aprovados').situacao,'pendente');
  assert.equal(avaliar('Custos aprovados.\n\nDe: gráfica@fornecedor.com\nNão aprovados').situacao,'aprovacao');
});

test('custo validado bloqueia negativas e pendencias mesmo com clique de confirmacao',()=>{
  const h=harness();
  for(const texto of ['Não aprovado para o chamado 3127','Ok, mas há divergência nos valores']){
    const r=h.context.confirmarCustoComEmailReentradas(id,h.out.version,{texto,remetente:'grafica@fornecedor.com',
      assunto:'Validação de custos 3127',confirmacaoExplicita:true,aceitouAnaliseManual:true});
    assert.match(r.erro,/nao confirma a aprovacao/i);
  }
  assert.equal(h.out.writes,0);
});

test('aprovacao isolada no contexto exige conferencia manual; sem ela nao registra',()=>{
  const h=harness();
  const r=h.context.confirmarCustoComEmailReentradas(id,h.out.version,{texto:'OK',
    assunto:'Validação de custos 3127',remetente:'grafica@fornecedor.com',confirmacaoExplicita:true});
  assert.match(r.erro,/nao e inequivoca/i);
  assert.equal(h.out.writes,0);
});

test('ultima resposta com pendencia desativa sugestao anterior de aceite',()=>{
  const h=harness(),data=new Date(1730000000000);
  const mensagem=(id,texto,from)=>({getId:()=>id,getPlainBody:()=>texto,getFrom:()=>from,
    getSubject:()=> 'Validação de custos - Chamado 3127',getDate:()=>data});
  const ms=[mensagem('APROVOU','Aprovamos os valores do chamado 3127','grafica@fornecedor.com'),
    mensagem('PENDENTE','Ajustar o custo de impressão, ainda está pendente.','grafica@fornecedor.com')];
  const thread={getId:()=> 'THREAD_123456789',getMessages:()=>ms,getFirstMessageSubject:()=> 'Validação de custos - Chamado 3127',
    getLastMessageDate:()=>data,getPermalink:()=> 'https://mail.google.com/mail/#all/x'};
  h.context.GmailApp={search:()=>[thread]};
  const r=h.context.obterConfirmacoesEmailCustoReentradas(id,h.out.version);
  assert.equal(r.sucesso,true,r.erro);
  assert.equal(r.sugestao,null);
  assert.equal(r.emails[0].mensagens[1].analise.situacao,'pendente');
});

test('nao sugere os e-mails de teste enviados pelo proprio usuario',()=>{
  const h=harness(),data=new Date(1730000000000);
  const m={getId:()=> 'ENVIADO_INTERNO',getPlainBody:()=> 'OK, valores aprovados para o chamado 3127',
    getFrom:()=> 'Leonardo <leonardo.pereira@arcoeducacao.com.br>',getSubject:()=> 'Validação de custos - Chamado 3127',getDate:()=>data};
  const t={getId:()=> 'THREAD_INTERNA',getMessages:()=>[m],getFirstMessageSubject:()=>m.getSubject(),
    getLastMessageDate:()=>data,getPermalink:()=> 'https://mail.google.com/mail/#all/y'};
  h.context.GmailApp={search:()=>[t]};
  const r=h.context.obterConfirmacoesEmailCustoReentradas(id,h.out.version);
  assert.equal(r.sucesso,true,r.erro);
  assert.equal(r.sugestao,null);
  assert.equal(r.emails[0].mensagens[0].analise.situacao,'interna');
});

test('na confirmacao final recusa aprovacao superada por resposta posterior negativa',()=>{
  const h=harness(),data=new Date(1730000000000);
  const mensagem=(msgId,body)=>({getId:()=>msgId,getFrom:()=> 'Gráfica <grafica@fornecedor.com>',
    getDate:()=>data,getSubject:()=> 'Validação de custos 3127',getPlainBody:()=>body});
  const aprovou=mensagem('MSG_APROVOU','Aprovamos os valores do chamado 3127');
  const revogou=mensagem('MSG_REVOGOU','Não aprovado. Existe divergência no papel.');
  const t={getId:()=> 'THREAD_123456789',getMessages:()=>[aprovou,revogou]};
  aprovou.getThread=()=>t;
  h.context.GmailApp={getMessageById:()=>aprovou};
  const r=h.context.confirmarCustoComEmailReentradas(id,h.out.version,{mensagemId:'MSG_APROVOU',
    threadId:'THREAD_123456789',confirmacaoExplicita:true});
  assert.match(r.erro,/resposta posterior/i);
  assert.equal(h.out.writes,0);
});
