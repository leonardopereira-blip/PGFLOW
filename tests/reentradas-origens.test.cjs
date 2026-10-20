const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const fontes=['ValidaReentradasSPG.js','LegadoAlocacaoReentradas.js']
  .map(nome=>fs.readFileSync(path.join(__dirname,'..',nome),'utf8'));

function fixture(){
  const api=[],planilhas=[];
  const falhas=new Map();
  const headers={
    espec:['1. ID_Código SKU','2. INF_Marca'],
    arvore:['1. ID_Código Kit','1. ID_Código SKU'],
    tiragem:['1. ID_Código Kit','1. ID_Código SKU','15. DES_CD envio','16. TIR_SKU']
  };
  const dados=(id)=>({
    'Esp. Técnica (legado)':{headers:headers.espec,linhas:[['SKU-'+id,'SAS']]},
    'Árvore (legado)':{headers:headers.arvore,linhas:[['KIT-'+id,'SKU-'+id]]},
    'Tiragem (legado)':{headers:headers.tiragem,linhas:[['KIT-'+id,'SKU-'+id,'CITA',200]]}
  });
  const link='https://drive.google.com/ppg-legada';
  const abas=Object.fromEntries([
    ['Especificação Técnica',headers.espec,['SKU-2567','Pleno']],
    ['Árvore de produto',headers.arvore,['KIT-2567','SKU-2567']],
    ['Tiragem',headers.tiragem,['KIT-2567','SKU-2567','CITA',300]]
  ].map(([nome,cab,row])=>[nome,[['ID (Coluna A)','PPG utilizada (Último Link)',...cab,'Versão da PPG'],[2567,link,...row,1]]]));
  const context={SpreadsheetApp:{openById:id=>{
    planilhas.push(id);
    return {getSpreadsheetTimeZone:()=> 'America/Sao_Paulo',getSheetByName:nome=>{
      const rows=abas[nome];
      return {getLastRow:()=>rows.length,getLastColumn:()=>rows[0].length,
        getRange:(linha,coluna,quantidade=1,largura=1)=>({
          getValues:()=>rows.slice(linha-1,linha-1+quantidade).map(r=>r.slice(coluna-1,coluna-1+largura)),
          createTextFinder:numero=>{
            const finder={matchEntireCell:()=>finder,useRegularExpression:()=>finder,
              findAll:()=>rows.flatMap((row,i)=>i>=linha-1&&i<linha-1+quantidade&&String(row[coluna-1])===numero?[{getRow:()=>i+1}]:[])};
            return finder;
          }
        })};
    }};
  }}};
  vm.createContext(context);fontes.forEach(code=>vm.runInContext(code,context));
  context._emailSpgPacote_=(id,soMetadados,tipo)=>{
    api.push({id,soMetadados,tipo});
    const falha=falhas.get(id);
    if(falha instanceof Error)throw falha;
    const matrizes=dados(id);
    if(falha==='incompleto')matrizes['Árvore (legado)'].linhas=[];
    return {id,versaoId:'v1',blob:matrizes,statusChamado:falha==='cancelado'?'CANCELADO':'SOLICITADO',statusProducao:'PUBLICADA'};
  };
  context._reSpPartes_=blob=>blob;
  context._reSpMatriz_=(raw,nome)=>raw[nome];
  const erro=(mensagem,codigo)=>Object.assign(new Error(mensagem),codigo?{codigo}:{});
  falhas.set('2567',erro('Chamado não encontrado','CHAMADO_NAO_ENCONTRADO'));
  return {context,api,planilhas,falhas,erro};
}

test('fonte padrao usa a API e nunca abre a planilha de alocacao',()=>{
  const {context,api,planilhas}=fixture();
  const r=context._reSpDadosChamado_('1234','META');
  assert.equal(r.pacote.id,'1234');
  assert.equal(r.tiragem.linhas[0][3],200);
  assert.deepEqual(api,[{id:'1234',soMetadados:false,tipo:'reentrada'}]);
  assert.equal(planilhas.length,0);
});

test('ausencia na API pergunta pelo legado sem consultar a planilha automaticamente',()=>{
  const {context,planilhas}=fixture();
  const r=context.buscarResumoReentradasSPG('2567');
  assert.equal(r[0].naoEncontrado,true);
  assert.match(r[0].erro,/não encontrado/);
  assert.equal(planilhas.length,0);
  for(const selecao of [undefined,null,true,'2567',{2567:true},['9999']]){
    assert.throws(()=>context._reSpDadosChamado_('2567','META',selecao),err=>err.codigo==='CHAMADO_NAO_ENCONTRADO');
  }
  assert.equal(planilhas.length,0);
});

test('legado exige array com o chamado confirmado e erro de ausencia tipado da API',()=>{
  const {context,api,planilhas}=fixture();
  const r=context._reSpDadosChamado_('2567','Forma Certa',['2567']);
  assert.equal(r.pacote.origem,'legado');
  assert.equal(context._reSpBuscar_(r.tiragem.linhas[0],r.tiragem.headers,'16. TIR_SKU'),300);
  assert.deepEqual(api,[{id:'2567',soMetadados:true,tipo:'reentrada'}]);
  assert.deepEqual(planilhas,['1pR9a5-oogD7ZYHxImsE9ABNpHeVyCJaTcIab7_U_bI0']);
});

for(const mensagem of ['HTTP 401','HTTP 500','Tempo limite: timeout','Chamado não encontrado']){
  test(mensagem+': falha generica jamais seleciona o legado, mesmo com confirmacao',()=>{
    const {context,falhas,erro,planilhas}=fixture();
    falhas.set('2567',erro(mensagem));
    assert.throws(()=>context._reSpDadosChamado_('2567','META',['2567']),e=>e.message===mensagem);
    const resumo=context.buscarResumoReentradasSPG('2567',['2567']);
    assert.equal(resumo[0].naoEncontrado,false);
    assert.equal(planilhas.length,0);
  });
}

for(const [estado,padrao] of [['cancelado',/Chamado cancelado/],['incompleto',/sem as tres matrizes/]]){
  test(estado+': dados da API sao bloqueados sem permitir trocar por legado',()=>{
    const {context,falhas,planilhas}=fixture();
    falhas.set('2567',estado);
    assert.throws(()=>context._reSpDadosChamado_('2567','META'),padrao);
    assert.throws(()=>context._reSpDadosChamado_('2567','META',['2567']),/existe no SPG/);
    assert.equal(planilhas.length,0);
  });
}

test('chamado existente recusa a fonte antiga mesmo com selecao explicita',()=>{
  const {context,planilhas}=fixture();
  assert.throws(()=>context._reSpDadosChamado_('1234','META',['1234']),/existe no SPG/);
  assert.equal(planilhas.length,0);
});

test('resumo, validacao e geracao encaminham fontes por chamado em um lote misto',()=>{
  const {context,api,planilhas}=fixture();
  const selecao=['2567'];
  const resumo=context.buscarResumoReentradasSPG('1234, 2567',selecao);
  assert.deepEqual(Array.from(resumo,r=>[r.chamado,r.origem,r.tiragemTotal]),[
    ['1234','spg',200],['2567','legado',300]
  ]);
  assert.equal(resumo[1].legadoTemporario,true);
  const validacao=context.validarReentradasSPG('Forma Certa','1234,2567',selecao);
  assert.equal(validacao.sucesso,true);
  assert.deepEqual(Array.from(validacao.validos),['1234','2567']);
  assert.equal(validacao.invalidos.length,0);
  // Isola apenas a montagem de XLSX; preserva o encaminhamento e a consulta
  // efetiva de cada origem durante a etapa de geracao.
  context._reSpGerarUm_=(id,grafica,respostas,fornecimento,origens)=>{
    assert.equal(origens,selecao);assert.equal(grafica,'Forma Certa');
    const p=context._reSpDadosChamado_(id,grafica,origens);
    return {chamado:id,origem:p.pacote.origem||'spg'};
  };
  const geracao=context.gerarValidacaoReentradasSPG('Forma Certa',['1234','2567'],'spg','',[],{},selecao);
  assert.equal(geracao.sucesso,true);
  assert.deepEqual(Array.from(geracao.arquivos,r=>[r.chamado,r.origem]),[['1234','spg'],['2567','legado']]);
  assert.equal(geracao.avisos.length,0);
  assert.equal(planilhas.length,3);
  assert.equal(api.length,6);
  assert.ok(api.every(r=>r.tipo==='reentrada'&&r.soMetadados===(r.id==='2567')));
});
