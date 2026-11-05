// A validacao final e um aceite explicito do usuario, independente da geracao.
// O XLSX original permanece no Drive com o mesmo ID, nome, formulas e revisoes.
var RX_CUSTO_VALIDADO_SCHEMA = 1;
var RX_CUSTO_VALIDADO_NOME = 'custos-validados.json';
var RX_CUSTO_CONFIRMACOES_PASTA = '18jjrBkVvc0k7vqALJUtMSMuDbAYfiB7J';

function _rxCustoEmailSeguro_(valor, limite) {
  return String(valor === undefined || valor === null ? '' : valor).slice(0, limite || 3000);
}

function _rxCustoHtmlSeguro_(valor) {
  return _rxCustoEmailSeguro_(valor, 20000).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// O DEV executa como o usuario que acessa. Nunca consultar o Gmail do
// proprietario em nome de outro usuario: a consulta pessoal so e permitida
// quando o usuario ativo e a propria conta executora.
function _rxCustoGmailDaPessoa_() {
  var ativo = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  var executor = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  return !!ativo && ativo === executor;
}

// Le somente a resposta nova: citacoes da conversa anterior nao sao aceite.
function _rxCustoRespostaLimpa_(texto){
  var linhas=String(texto||'').split(/\r?\n/),novas=[];
  for(var i=0;i<linhas.length;i++){
    var l=linhas[i].trim();
    if(/^\s*(?:>|de:\s|from:\s|enviado:\s|sent:\s|assunto:\s|subject:\s|em .+ escreveu:|on .+ wrote:|[-_]{3,}\s*(?:mensagem|original))/i.test(l))break;
    if(/^--\s*$/.test(l))break;
    if(l)novas.push(l);
  }
  return novas.join(' ').slice(0,6000).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
}
function _rxCustoRemetenteExterno_(remetente){
  var atual=String(Session.getActiveUser().getEmail()||'').toLowerCase();
  var encontrado=/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i.exec(String(remetente||''));
  if(!atual||!encontrado)return false;
  return encontrado[0].toLowerCase().split('@')[1]!==atual.split('@')[1];
}
function _rxCustoClassificarResposta_(texto,assunto,remetente){
  var corpo=_rxCustoRespostaLimpa_(texto);
  var tema=_rxCustoRespostaLimpa_(assunto);
  var externo=_rxCustoRemetenteExterno_(remetente);
  if(!corpo)return {situacao:'sem_resposta',motivo:'Mensagem sem resposta nova identificável.',confianca:'nenhuma'};
  var textoSemExcecao=corpo.replace(/\bnao\s+(?:ha|existem|temos)\s+(?:pendencias|divergencias|ressalvas)\b/g,'');
  // Negacoes e condicoes prevalecem sobre palavras afirmativas na mesma mensagem.
  if(/\b(?:nao\s+(?:(?:foi|esta|estao)\s+)?(?:validado|validada|validamos|aprovo|aprovado|aprovada|aprovamos|confirmamos|autorizamos|concordamos|aceito|aceitamos)|nao\s+(?:ha|existe)\s+(?:aprovacao|validacao|aceite)|nao\s+(?:estamos|estou\s+)?de\s+acordo|sem\s+(?:aprovacao|validacao|aceite)|reprovad[oa]s?|recusad[oa]s?|discordamos|nao\s+(?:esta|estao)\s+corret[oa]s?|valores?\s+incorret[oa]s?|not\s+(?:approved|validated|accepted)|rejected)\b/.test(corpo)||
    /\bnao\b(?:\s+\w+){0,5}\s+\b(?:validad\w*|aprovad\w*|aceit\w*)\b/.test(textoSemExcecao))
    return {situacao:'rejeitada',motivo:'Resposta com negativa explícita. Não sugerir aprovação.',confianca:'alta'};
  var ressalvas=corpo.replace(/\bsem\s+(?:pendencias|ressalvas|divergencias)\b/g,'')
    .replace(/\bnao\s+(?:ha|existem|temos)\s+(?:pendencias|divergencias|ressalvas)\b/g,'');
  if(/\b(?:pendenci\w*|divergenci\w*|ressalv\w*|ajust\w*|revis\w*|corrig\w*|analis\w*|aguard\w*|falt\w*|condicion\w*|exceto|porem|mas\s+|parcialmente|ainda\s+nao|em\s+analise|sob\s+analise|nao\s+confere|sujeit\w*|provisori\w*|pending|awaiting|needs?\s+changes|conditional)\b/.test(ressalvas))
    return {situacao:'pendente',motivo:'Há ressalva, pendência ou pedido de ajuste na resposta.',confianca:'alta'};
  if(!externo)return {situacao:'interna',motivo:'Remetente interno ou não identificado; não é confirmação da gráfica.',confianca:'alta'};
  var afirmacao=/\b(?:validad[oa]s?|aprovad[oa]s?|aprovamos|validamos|confirmamos|de\s+acordo|aceitamos|aceito|autorizamos|liberado|liberamos|corret[oa]s?|conferid[oa]s?|podem\s+seguir|pode\s+seguir|ok|okay)\b/.test(corpo);
  var contexto=/\b(?:custos?|prec(?:o|os)|valores?|calculadora|orcamento|cotacao|planilha|proposta|tiragem|chamado)\b/.test(corpo);
  var contextoAssunto=/\b(?:validacao|aprovacao|confirmacao|custos?|calculadora|orcamento|cotacao)\b/.test(tema);
  if(!afirmacao)return {situacao:'neutra',motivo:'Não há aceite de custo reconhecível na resposta.',confianca:'nenhuma'};
  if(!contexto&&!contextoAssunto)return {situacao:'neutra',motivo:'A mensagem positiva não está vinculada à validação dos custos.',confianca:'nenhuma'};
  if(contexto&&/\b(?:aprovad[oa]s?|aprovamos|validad[oa]s?|validamos|confirmamos|de\s+acordo|corret[oa]s?|aceitamos)\b/.test(corpo))
    return {situacao:'aprovacao',motivo:'Resposta externa afirmativa sobre custos ou valores, sem ressalvas identificadas.',confianca:'alta'};
  return {situacao:'possivel',motivo:contextoAssunto?
    'Há um sinal positivo, mas o contexto da resposta precisa de conferência humana.':
    'A expressão positiva não identifica claramente a validação dos custos.',confianca:'baixa'};
}

function _rxCustoSugestoes_(dados) {
  if (!dados.sugestoes || typeof dados.sugestoes !== 'object' || Array.isArray(dados.sugestoes))
    dados.sugestoes = {};
  return dados.sugestoes;
}

function obterConfirmacoesEmailCustoReentradas(id, atualizadoEm, filtro, pagina) {
  try {
    var arquivo = _rxArquivo_(id, atualizadoEm);
    if (!arquivo.calculadora) throw Error('Selecione uma calculadora para conferir os e-mails.');
    var chamado = _rxCustoChamado_(arquivo.nome), dados = _rxCustoRegistros_();
    var paginaAtual=Number(pagina||0);
    if(!Number.isInteger(paginaAtual)||paginaAtual<0||paginaAtual>500)throw Error('Pagina de e-mails invalida.');
    var pesquisa=_rxCustoEmailSeguro_(filtro,100).trim(),limite=15;
    var retorno = {sucesso:true, chamado:chamado, pagina:paginaAtual, proximaPagina:false,
      sugestaoJaTratada:!!_rxCustoSugestoes_(dados)[id],
      confirmado:!!dados.arquivos[id], emails:[], sugestao:null, conexaoPessoalNecessaria:!_rxCustoGmailDaPessoa_()};
    if (retorno.conexaoPessoalNecessaria) return retorno;
    // A primeira lista prioriza o chamado; a busca explicita permite qualquer
    // conversa da caixa do usuario, ainda que nao mencione esse chamado.
    var consulta='in:anywhere '+(pesquisa||('"'+chamado+'"'));
    var threads = GmailApp.search(consulta, paginaAtual*limite, limite+1), vistos={};
    retorno.proximaPagina=threads.length>limite;
    threads.slice(0,limite).forEach(function(thread) {
      if (vistos[thread.getId()]) return;
      vistos[thread.getId()]=true;
      var mensagens=thread.getMessages(),assunto=thread.getFirstMessageSubject(),referencias=[];
      mensagens.slice(-4).forEach(function(m) {
        var analise=_rxCustoClassificarResposta_(m.getPlainBody(),m.getSubject(),m.getFrom());
        referencias.push({id:m.getId(),remetente:_rxCustoEmailSeguro_(m.getFrom(),300),
          data:m.getDate().getTime(),assunto:_rxCustoEmailSeguro_(m.getSubject(),300),
          trecho:_rxCustoEmailSeguro_(m.getPlainBody(),1200),analise:analise});
      });
      retorno.emails.push({id:thread.getId(),assunto:_rxCustoEmailSeguro_(assunto,300),
        ultimaData:thread.getLastMessageDate().getTime(),url:thread.getPermalink(),mensagens:referencias});
    });
    retorno.emails.sort(function(a,b){return b.ultimaData-a.ultimaData;});
    if(!retorno.sugestaoJaTratada&&chamado){
      var recomendadas=GmailApp.search('in:anywhere "'+chamado+'"',0,8);
      var opcoes=[];
      recomendadas.forEach(function(t){
        var ms=t.getMessages();if(!ms.length)return;
        var relevante=null,negaPosterior=false;
        ms.forEach(function(m){
          var analise=_rxCustoClassificarResposta_(m.getPlainBody(),m.getSubject(),m.getFrom());
          if(analise.situacao==='rejeitada'||analise.situacao==='pendente')negaPosterior=true;
          if(analise.situacao==='aprovacao'||analise.situacao==='possivel'){
            relevante={m:m,analise:analise};negaPosterior=false;
          }
        });
        if(!relevante||negaPosterior)return;
        var m=relevante.m,assunto=t.getFirstMessageSubject(),temChamado=new RegExp('(^|\\D)'+chamado+'(\\D|$)').test(assunto);
        var temChamadoNaResposta=new RegExp('(^|\\D)'+chamado+'(\\D|$)').test(_rxCustoRespostaLimpa_(m.getPlainBody()));
        if(!temChamado&&!temChamadoNaResposta)return;
        opcoes.push({threadId:t.getId(),mensagem:{id:m.getId(),remetente:_rxCustoEmailSeguro_(m.getFrom(),300),
          data:m.getDate().getTime(),assunto:_rxCustoEmailSeguro_(m.getSubject(),300),
          trecho:_rxCustoEmailSeguro_(m.getPlainBody(),1200),analise:relevante.analise},
          pontos:(relevante.analise.situacao==='aprovacao'?5:1)+(temChamado?2:0)});
      });
      opcoes.sort(function(a,b){return b.pontos-a.pontos||b.mensagem.data-a.mensagem.data;});
      if(opcoes.length)retorno.sugestao=opcoes[0];
    }
    return retorno;
  } catch(e) {return {erro:e.message || String(e)};}
}

function registrarSugestaoEmailCustoReentradas(id, atualizadoEm, decisao) {
  var lock=LockService.getScriptLock(),pegou=false;
  try {
    if(['ignorada','examinada'].indexOf(decisao)<0)throw Error('Decisao de sugestao invalida.');
    pegou=lock.tryLock(3000);if(!pegou)throw Error('Outra confirmacao esta em andamento.');
    var arquivo=_rxArquivo_(id,atualizadoEm);
    if(!arquivo.calculadora)throw Error('Selecione uma calculadora.');
    var dados=_rxCustoRegistros_(),avisos=_rxCustoSugestoes_(dados);
    if(!avisos[id]){avisos[id]={decisao:decisao,em:Date.now()};_rxCustoSalvarRegistros_(dados);}
    return {sucesso:true};
  }catch(e){return {erro:e.message||String(e)};}
  finally{if(pegou)lock.releaseLock();}
}
function ignorarSugestaoEmailCustoReentradas(id,atualizadoEm){
  return registrarSugestaoEmailCustoReentradas(id,atualizadoEm,'ignorada');
}

function _rxCustoDeparaDoArquivo_(calculadora, chamado) {
  var pasta=DriveApp.getFolderById(RX_PASTA),arquivos=pasta.getFiles(),candidatos=[];
  var grafica=calculadora.nome.replace(/^CALCULADORA FORNECEDOR /i,'').split(' - LIVROS REENTRADA')[0];
  while(arquivos.hasNext()){
    var a=arquivos.next(),m=/^PLANILHA DE PARA PPG (.+) - CHAMADO (\d+)\.xlsx$/i.exec(a.getName());
    if(!m||String(Number(m[2]))!==chamado||a.isTrashed()||a.getMimeType()!==RX_MIME)continue;
    if(_reSpNormal_(m[1])!==_reSpNormal_(grafica))continue;
    var tempo=a.getDateCreated().getTime(),delta=calculadora.arquivo.getDateCreated().getTime()-tempo;
    if(delta>=0 && delta<10*60*1000)candidatos.push({a:a,delta:delta});
  }
  candidatos.sort(function(a,b){return a.delta-b.delta;});
  if(!candidatos.length)throw Error('O De-Para da mesma geracao nao foi identificado. Nenhum custo foi validado.');
  return candidatos[0].a;
}

function confirmarCustoComEmailReentradas(id, atualizadoEm, prova) {
  var lock=LockService.getScriptLock(),pegou=false;
  try {
    if(!prova||typeof prova!=='object')throw Error('Selecione uma mensagem ou informe o comprovante.');
    var texto=_rxCustoEmailSeguro_(prova.texto,20000),emailId=_rxCustoEmailSeguro_(prova.mensagemId,100);
    if(!emailId && !texto.trim())throw Error('O comprovante selecionado esta vazio.');
    if(!prova.confirmacaoExplicita)throw Error('Confirme explicitamente a conclusao do custo.');
    pegou=lock.tryLock(5000);if(!pegou)throw Error('Outra validacao esta em andamento.');
    var arquivo=_rxArquivo_(id,atualizadoEm);
    if(!arquivo.calculadora)throw Error('O custo validado precisa ser uma calculadora.');
    var mensagem=null,thread=null;
    if(emailId){
      if(!_rxCustoGmailDaPessoa_())throw Error('A caixa Gmail do usuario ativo nao esta autorizada. Utilize um comprovante manual.');
      mensagem=GmailApp.getMessageById(emailId);
      if(!mensagem)throw Error('O e-mail selecionado nao esta disponivel na sua conta.');
      thread=mensagem.getThread();
      if(prova.threadId&&thread.getId()!==String(prova.threadId))throw Error('O e-mail nao pertence a conversa selecionada.');
      // O usuario escolhe livremente o trecho que vale como comprovante.
      // A mensagem original continua arquivada em .eml, sem interpretacao automatica.
      if(!texto.trim())texto=_rxCustoEmailSeguro_(mensagem.getPlainBody(),20000);
      prova.assunto=mensagem.getSubject();prova.remetente=mensagem.getFrom();prova.data=mensagem.getDate().getTime();
    }
    var analise=_rxCustoClassificarResposta_(emailId?mensagem.getPlainBody():texto,prova.assunto,prova.remetente);
    if(['rejeitada','pendente','interna','sem_resposta','neutra'].indexOf(analise.situacao)>=0)
      throw Error('Este comprovante nao confirma a aprovacao dos custos. '+analise.motivo);
    if((!emailId||analise.situacao!=='aprovacao')&&!prova.aceitouAnaliseManual)
      throw Error('A aprovacao nao e inequivoca ou foi informada manualmente. Confira o contexto e confirme conscientemente antes de continuar.');
    if(thread){
      var sequencia=thread.getMessages(),depois=false;
      sequencia.forEach(function(m){
        if(depois){
          var atual=_rxCustoClassificarResposta_(m.getPlainBody(),m.getSubject(),m.getFrom());
          if(atual.situacao==='rejeitada'||atual.situacao==='pendente')
            throw Error('Existe uma resposta posterior com negativa ou pendencia nesta conversa. A aprovacao anterior nao pode ser usada.');
        }
        if(m.getId()===emailId)depois=true;
      });
    }
    var chamado=_rxCustoChamado_(arquivo.nome),depara=_rxCustoDeparaDoArquivo_(arquivo,chamado);
    var dados=_rxCustoRegistros_(),registroAnterior=dados.arquivos[id];
    if(registroAnterior&&_rxCustoRegistroVigente_(registroAnterior,arquivo))
      throw Error('Esta calculadora ja foi validada. Consulte o arquivo confirmado.');
    var pendencias='',totais={tiragemValidada:null,valorChamado:null};
    try {
      var revisao=_rxRevisaoLer_(arquivo);
      if(revisao&&revisao.recalculoPendente)throw Error('Existem formulas ainda nao recalculadas.');
      totais=_rxCustoTotaisRevisados_(revisao);
    }catch(e){pendencias=e.message||String(e);if(!prova.aceitouPendencias)
      throw Error('Existem resultados pendentes: '+pendencias+' Confirme que deseja continuar assim mesmo.');}
    var emails=thread?thread.getMessages():[];
    if(emails.length>50)throw Error('A conversa possui mais de 50 mensagens. Utilize um comprovante manual para esta validacao.');
    var pasta=DriveApp.getFolderById(RX_CUSTO_CONFIRMACOES_PASTA),agora=Date.now();
    var sub=pasta.createFolder('Chamado '+chamado+' - '+agora),copia=arquivo.arquivo.makeCopy(arquivo.nome,sub);
    var copiaDepara=depara.makeCopy(depara.getName(),sub);
    var metadados={chamado:chamado,arquivoOriginalId:id,arquivoValidadoId:copia.getId(),
      deparaOriginalId:depara.getId(),deparaValidadoId:copiaDepara.getId(),
      comprovante:{origem:emailId?'gmail':'manual',threadId:_rxCustoEmailSeguro_(prova.threadId,100),
        mensagemId:emailId,assunto:_rxCustoEmailSeguro_(prova.assunto,300),
        remetente:_rxCustoEmailSeguro_(prova.remetente,300),data:prova.data||null,
        analise:analise,conferenciaManual:!!prova.aceitouAnaliseManual},
      registradoEm:agora,registradoPor:String(Session.getActiveUser().getEmail()||''),pendencias:pendencias};
    var html='<!doctype html><html lang="pt-br"><meta charset="utf-8"><title>Comprovante do chamado '+
      _rxCustoHtmlSeguro_(chamado)+'</title><body style="font:14px Arial;color:#13294b;max-width:900px;margin:40px auto">'+
      '<h1>ARCO · Confirmação de custos</h1><p><b>Chamado:</b> '+_rxCustoHtmlSeguro_(chamado)+
      '</p><p><b>Assunto:</b> '+_rxCustoHtmlSeguro_(prova.assunto)+
      '</p><p><b>Remetente:</b> '+_rxCustoHtmlSeguro_(prova.remetente)+
      '</p><p><b>Registrado por:</b> '+_rxCustoHtmlSeguro_(metadados.registradoPor)+
      '</p><pre style="white-space:pre-wrap;background:#f4f7fb;padding:20px">'+_rxCustoHtmlSeguro_(texto)+'</pre></body></html>';
    sub.createFile(Utilities.newBlob(html,'text/html','Comprovante_email.html'));
    sub.createFile(Utilities.newBlob(JSON.stringify(metadados),'application/json','Registro_validacao.json'));
    if(thread){
      emails.forEach(function(m,i){
        var origem=m.getRawContent();
        if(origem.length>18*1024*1024)throw Error('Uma das mensagens e grande demais para arquivamento seguro.');
        sub.createFile(Utilities.newBlob(origem,'message/rfc822',
          ('00'+(i+1)).slice(-3)+'_email.eml'));
      });
    }
    var registro={id:id,nome:arquivo.nome,chamado:chamado,atualizadoEm:arquivo.atualizadoEm,
      validadoEm:agora,tiragemValidada:totais.tiragemValidada,valorChamado:totais.valorChamado,
      arquivoValidadoId:copia.getId(),pastaId:sub.getId(),pendencias:pendencias,
      mensagemId:emailId,threadId:metadados.comprovante.threadId};
    dados.arquivos[id]=registro;
    _rxCustoSugestoes_(dados)[id]={decisao:'validada',em:agora};
    _rxCustoSalvarRegistros_(dados);
    return {sucesso:true,id:id,chamado:chamado,validadoEm:agora,atualizadoEm:arquivo.atualizadoEm,
      tiragemValidada:totais.tiragemValidada,valorChamado:totais.valorChamado,
      downloadUrl:_rxCustoUrlDownload_(copia.getId()),pastaUrl:sub.getUrl(),pendencias:pendencias};
  }catch(e){return {erro:e.message||String(e)};}
  finally{if(pegou)lock.releaseLock();}
}

function enviarCustoPorEmailNaConversaReentradas(id,atualizadoEm,threadId,mensagem) {
  try {
    if(!_rxCustoGmailDaPessoa_())throw Error('O Gmail deve pertencer ao usuario conectado.');
    var arquivo=_rxArquivo_(id,atualizadoEm),chamado=_rxCustoChamado_(arquivo.nome);
    var texto=String(mensagem||'').trim();
    if(!arquivo.calculadora||!texto||texto.length>12000)throw Error('Preencha o texto da validacao de custos.');
    if(!/^[A-Za-z0-9_-]{10,100}$/.test(String(threadId||'')))throw Error('Selecione uma conversa valida.');
    var conversa=GmailApp.getThreadById(threadId),emails=conversa&&conversa.getMessages();
    if(!emails||!emails.length)throw Error('A conversa nao esta disponivel na sua caixa.');
    // A escolha pode vir da busca livre de threads. O usuario confirma o
    // destinatario na interface; o numero do chamado nao restringe a resposta.
    conversa.replyAll(texto,{attachments:[arquivo.arquivo.getBlob().setName(arquivo.nome)]});
    return {sucesso:true,mensagem:'Validacao de custos enviada como resposta na conversa selecionada.'};
  }catch(e){return {erro:e.message||String(e)};}
}

function _rxCustoChamado_(nome) {
  var match = /\bCHAMADO\s+(\d+)\.xlsx$/i.exec(String(nome || ''));
  return match ? String(Number(match[1])) : '';
}

function _rxCustoRegistros_() {
  var pasta = _rxRevisaoPasta_(false);
  var arquivos = pasta && pasta.getFilesByName(RX_CUSTO_VALIDADO_NOME);
  if (!arquivos || !arquivos.hasNext()) return {schema: RX_CUSTO_VALIDADO_SCHEMA, arquivos: {}};
  var dados = JSON.parse(arquivos.next().getBlob().getDataAsString('UTF-8'));
  if (!dados || dados.schema !== RX_CUSTO_VALIDADO_SCHEMA || !dados.arquivos ||
      typeof dados.arquivos !== 'object' || Array.isArray(dados.arquivos)) {
    throw Error('O registro de validacoes precisa de verificacao. Nenhum aceite foi gravado.');
  }
  return dados;
}

function _rxCustoSalvarRegistros_(dados) {
  var pasta = _rxRevisaoPasta_(true);
  var arquivos = pasta.getFilesByName(RX_CUSTO_VALIDADO_NOME);
  var blob = Utilities.newBlob(JSON.stringify(dados), 'application/json', RX_CUSTO_VALIDADO_NOME);
  if (arquivos.hasNext()) {
    var existente = arquivos.next();
    Drive.Files.update({title: RX_CUSTO_VALIDADO_NOME, mimeType: 'application/json'},
      existente.getId(), blob, {newRevision: false, supportsAllDrives: true});
  } else {
    pasta.createFile(blob).setName(RX_CUSTO_VALIDADO_NOME);
  }
}

function _rxCustoUrlDownload_(id) {
  return 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(id);
}

function _rxCustoRegistroVigente_(registro, arquivo) {
  return !!(registro && registro.id === arquivo.id &&
    registro.atualizadoEm === arquivo.atualizadoEm &&
    registro.nome === arquivo.nome && arquivo.calculadora);
}

function _rxCustoValorNumero_(valor,descricao){
  var texto=String(valor===undefined?'':valor).trim();
  if(!/^-?\d+(?:\.\d+)?$/.test(texto))throw Error(descricao+' ainda nao possui resultado numerico calculado.');
  var n=Number(texto);if(!isFinite(n)||n<0)throw Error(descricao+' possui valor invalido.');
  return n;
}
// Le exclusivamente os CAMPOS FINAIS do CALC Geral, jamais valores do SPG
// anteriores a revisao. Confere o TOTAL de exemplares com as quatro tiragens.
function _rxCustoTotaisRevisados_(modelo){
  var abas=modelo&&modelo.abas||[],aba=abas.find(function(a){return /^CALC[\s-]*Geral$/i.test(a.nome);});
  if(!aba)throw Error('A etapa Resumo de custos nao foi encontrada para apurar o chamado.');
  var linhas=aba.linhas||[],cabecalho=(linhas.find(function(r){return r.numero===3;})||{}).valores||{};
  var subtitulo=(linhas.find(function(r){return r.numero===4;})||{}).valores||{};
  var colSku=Number(Object.keys(cabecalho).find(function(k){return /C[ÓO]DIGO\s*SKU/i.test(cabecalho[k]);}));
  var colTiragem=Number(Object.keys(cabecalho).find(function(k){return /^TOTAL$/i.test(String(cabecalho[k]).trim())&&
    /Exemplares/i.test(String(subtitulo[k]||''));}));
  var colPreco=Number(Object.keys(cabecalho).find(function(k){return /SERVI[CÇ]O\s*\+\s*PAPEL\s+TOTAL/i.test(String(cabecalho[k]));}));
  if(!Number.isInteger(colSku)||!Number.isInteger(colTiragem)||!Number.isInteger(colPreco)||
    colTiragem<=colSku+2||colTiragem>colSku+8||
    !Object.keys(cabecalho).length||!(''+cabecalho[colSku]).match(/SKU/i)||
    !/BRL\s*\/\s*Total/i.test(String(subtitulo[colPreco+1]||'')))
    throw Error('As colunas de tiragem e preco final deste molde precisam ser conferidas antes do aceite.');
  var tiragem=0,valor=0,quantidade=0;
  linhas.forEach(function(r){
    if(r.numero<=5)return;
    var v=r.valores||{},sku=String(v[colSku]||'').trim();if(!sku||/^=/.test(sku))return;
    var total=_rxCustoValorNumero_(v[colTiragem],'Total de exemplares do SKU '+sku);
    var partes=0;
    for(var c=colSku+2;c<colTiragem;c++)partes+=_rxCustoValorNumero_(v[c]===undefined?'0':v[c],'Tiragem do SKU '+sku);
    if(Math.abs(partes-total)>0.001)throw Error('O total de exemplares do SKU '+sku+
      ' diverge da tiragem informada. Recalcule a calculadora antes de validar o custo.');
    tiragem+=total;
    valor+=_rxCustoValorNumero_(v[colPreco+1],'Preco final do SKU '+sku);
    quantidade++;
  });
  if(!quantidade)throw Error('Nenhum SKU calculado foi encontrado para validar o custo.');
  return {tiragemValidada:tiragem,valorChamado:Math.round((valor+Number.EPSILON)*100)/100};
}

// Nunca e chamado durante a geracao nem ao simplesmente abrir o arquivo.
function _rxCustoAceiteLegadoDesativado_(id, atualizadoEm) {
  var lock = LockService.getScriptLock(), pegou = false;
  try {
    if (!/^\d{10,16}$/.test(String(atualizadoEm || '')) ||
        !Number.isSafeInteger(Number(atualizadoEm))) throw Error('Atualize a revisao antes de validar.');
    pegou = lock.tryLock(3000);
    if (!pegou) throw Error('Outra operacao esta em andamento. Tente novamente.');
    var arq = _rxArquivo_(id, atualizadoEm);
    if (!arq.calculadora) throw Error('O aceite de custo e exclusivo da calculadora revisada.');
    var chamado = _rxCustoChamado_(arq.nome);
    if (!chamado) throw Error('Nao foi possivel identificar o chamado desta calculadora.');
    var revisao=_rxRevisaoLer_(arq); // exige revisao preparada para a versao exata do XLSX
    if(revisao&&revisao.recalculoPendente)
      throw Error('Este custo foi alterado e ainda precisa recalcular os resultados das formulas antes do aceite final.');
    var totais=_rxCustoTotaisRevisados_(revisao);
    var dados = _rxCustoRegistros_(), confirmado = _rxArquivo_(id, atualizadoEm);
    var registro = {id: confirmado.id, nome: confirmado.nome, chamado: chamado,
      atualizadoEm: confirmado.atualizadoEm, validadoEm: Date.now(),
      tiragemValidada:totais.tiragemValidada,valorChamado:totais.valorChamado};
    dados.arquivos[id] = registro;
    _rxCustoSalvarRegistros_(dados);
    return {sucesso: true, id: id, chamado: chamado, validadoEm: registro.validadoEm,
      tiragemValidada:registro.tiragemValidada,valorChamado:registro.valorChamado,
      atualizadoEm: registro.atualizadoEm, downloadUrl: _rxCustoUrlDownload_(id)};
  } catch (e) { return {erro: e.message || String(e)}; }
  finally { if (pegou) lock.releaseLock(); }
}

function marcarCustoValidadoReentradas() {
  return {erro:'O Custo Validado agora exige a selecao manual do comprovante na etapa E-mails e confirmacao.'};
}

function _rxCustoListaDeArquivos_(arquivos) {
  var registros = _rxCustoRegistros_().arquivos;
  arquivos.forEach(function(arquivo) {
    var r = registros[arquivo.id];
    if (_rxCustoRegistroVigente_(r, arquivo)) {
      arquivo.validadoEm = r.validadoEm;
      arquivo.downloadUrl = _rxCustoUrlDownload_(r.arquivoValidadoId || arquivo.id);
      arquivo.pastaId = r.pastaId || '';
    }
  });
}

// Na lista de chamados mostra a ultima validacao, mesmo que o arquivo tenha
// mudado depois. Download so aparece se a versao ainda for exatamente a validada.
function _rxCustoResumoChamados_() {
  var registros = _rxCustoRegistros_().arquivos, lista = Object.keys(registros);
  var atuais = {}, porChamado = {}, iterator = DriveApp.getFolderById(RX_PASTA).getFiles();
  while (iterator.hasNext()) {
    var a = iterator.next(), id = a.getId();
    if(a.isTrashed()||a.getMimeType()!==RX_MIME)continue;
    var nome=a.getName(), chamadoArquivo=/^CALCULADORA FORNECEDOR .*\bCHAMADO\s+(\d+)\.xlsx$/i.exec(nome);
    if(chamadoArquivo){
      var chamadoGerado=String(Number(chamadoArquivo[1]));
      var dataCriacao=a.getDateCreated().getTime(),anteriorGerado=porChamado[chamadoGerado];
      if(!anteriorGerado||!anteriorGerado.ultimoGeradoEm||anteriorGerado.ultimoGeradoEm<dataCriacao){
        porChamado[chamadoGerado]=Object.assign({},anteriorGerado||{},
          {ultimoGeradoEm:dataCriacao,ultimoArquivoId:id});
      }
    }
    if (registros[id]) {
      atuais[id] = {id: id, nome: a.getName(), atualizadoEm: a.getLastUpdated().getTime(),
        calculadora: /^CALCULADORA FORNECEDOR /i.test(a.getName())};
    }
  }
  lista.forEach(function(id) {
    var r = registros[id];
    if (!r || !/^\d+$/.test(String(r.chamado)) || !Number.isSafeInteger(r.validadoEm)) return;
    var anterior = porChamado[r.chamado];
    if (!anterior || !anterior.ultimaValidacaoEm || anterior.ultimaValidacaoEm < r.validadoEm) {
      var atual = atuais[id];
      porChamado[r.chamado] = Object.assign({},anterior||{},{ultimaValidacaoEm: r.validadoEm,
        tiragemValidada:r.tiragemValidada,valorChamado:r.valorChamado,
        pastaId:r.pastaId||'',
        downloadUrl: atual && _rxCustoRegistroVigente_(r, atual) ? _rxCustoUrlDownload_(r.arquivoValidadoId || id) : ''});
    }
  });
  return porChamado;
}
