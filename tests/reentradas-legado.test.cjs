const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const code=fs.readFileSync(path.join(__dirname,'..','LegadoAlocacaoReentradas.js'),'utf8');
const motor=fs.readFileSync(path.join(__dirname,'..','ValidaReentradasSPG.js'),'utf8');
const geral=fs.readFileSync(path.join(__dirname,'..','ReentradasEspecificacaoGeral.js'),'utf8');
const spg=fs.readFileSync(path.join(__dirname,'..','SPG.js'),'utf8');
const idFonte='1pR9a5-oogD7ZYHxImsE9ABNpHeVyCJaTcIab7_U_bI0';
const meta=['ID (Coluna A)','PPG utilizada (Último Link)'];
const link='https://drive.google.com/file/d/PPG-2567-ATUAL/view';

function fixture(){
  const data={
    'Especificação Técnica':[
      [...meta,'1. ID_Código SKU','2. INF_Marca','12. VER_Data verificação editorial','Versão da PPG','Tiragem'],
      [12567,'https://drive.google.com/OUTRO','OUTRO','OUTRA','',9,10],
      [2567,'https://drive.google.com/ANTIGO','SKU-ANTIGO','MARCA','',1,10],
      [2567,link,'SKU-A','SAS',new Date('2026-10-19T12:00:00Z'),2,500],
      [2567,link,'SKU-B','SAS','',2,100],
      [2567,link,'SKU-SEM-TIRAGEM','SAS','',2,100]
    ],
    'Árvore de produto':[
      [...meta,'Grupo','1. ID_Código Kit','1. ID_Código SKU','13. MAN_Total de itens colecionados','13. MAN_Embalagem','Versão da PPG'],
      [12567,'https://drive.google.com/OUTRO','','OUTRO','OUTRO',1,'',9],
      [2567,'https://drive.google.com/ANTIGO','','KIT-ANTIGO','SKU-ANTIGO',1,'',1],
      [2567,link,'','KIT-A','',2,'Maleta',2],
      [2567,link,'','KIT-A','SKU-A',2,'Maleta',2],
      [2567,link,'','KIT-A','SKU-B',2,'Maleta',2],
      [2567,link,'','KIT-A','SKU-SEM-TIRAGEM',2,'Maleta',2],
      [2567,link,'','OUTRO-KIT','SKU-A',1,'Caixa',2]
    ],
    'Tiragem':[
      [...meta,'1. ID_Código Kit','1. ID_Código SKU','15. DES_CD envio','16. TIR_SKU','17. VER_Data de alteração','Versão da PPG','Chave'],
      [12567,'https://drive.google.com/OUTRO','OUTRO','OUTRO','CITA',999,'',9,'OUTRO'],
      [2567,'https://drive.google.com/ANTIGO','KIT-ANTIGO','SKU-ANTIGO','CITA',50,'',1,'OLD'],
      [2567,link,'KIT-A','SKU-A','CITA',200,new Date('2026-10-20T12:00:00Z'),2,'A1'],
      [2567,link,'KIT-A','SKU-A','CJUN','1.000','',2,'A2'],
      [2567,link,'KIT-A','SKU-B','CITA',100,'',2,'B1']
    ]
  };
  const reads=[],opens=[],dateFormats=[];
  const timeZone='America/Sao_Paulo';
  function sheet(nome){
    const rows=data[nome];
    if(!rows)return null;
    return {
      getLastRow:()=>rows.length,
      getLastColumn:()=>rows[0].length,
      getRange:(linha,coluna,quantidade=1,largura=1)=>({
        getValues:()=>{
          reads.push({nome,linha,coluna,quantidade,largura});
          return rows.slice(linha-1,linha-1+quantidade).map(r=>r.slice(coluna-1,coluna-1+largura));
        },
        createTextFinder:numero=>{
          let inteira=false,regex=true;
          const finder={
            matchEntireCell:v=>{inteira=v;return finder;},
            useRegularExpression:v=>{regex=v;return finder;},
            findAll:()=>{
              assert.equal(inteira,true);assert.equal(regex,false);
              return rows.flatMap((r,i)=>i>=linha-1&&i<linha-1+quantidade&&String(r[coluna-1])===numero?[{getRow:()=>i+1}]:[]);
            }
          };
          return finder;
        }
      })
    };
  }
  const context={SpreadsheetApp:{openById:id=>{
    assert.equal(id,idFonte);opens.push(id);return {getSheetByName:sheet,getSpreadsheetTimeZone:()=>timeZone};
  }},Utilities:{formatDate:(date,zone,format)=>{
    assert.equal(zone,timeZone);assert.equal(format,"yyyy-MM-dd'T'HH:mm:ss.SSS");
    dateFormats.push({date,zone});
    const parts=Object.fromEntries(new Intl.DateTimeFormat('en-CA',{
      timeZone:zone,year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',
      minute:'2-digit',second:'2-digit',hourCycle:'h23'
    }).formatToParts(date).map(p=>[p.type,p.value]));
    return `${parts.year}-${parts.month}-${parts.day}T${parts.hour}:${parts.minute}:${parts.second}.${String(date.getUTCMilliseconds()).padStart(3,'0')}`;
  }}};
  vm.createContext(context);vm.runInContext(spg,context);vm.runInContext(motor,context);vm.runInContext(code,context);vm.runInContext(geral,context);
  return {context,data,reads,opens,dateFormats};
}

test('origem legada e inativa ao carregar e nunca escreve ou abre planilhas de links',()=>{
  const {opens}=fixture();
  assert.equal(opens.length,0);
  assert.doesNotMatch(code,/\b(?:setValues|setValue|clear|clearContent|insertSheet|createFile|copy|sendEmail|gravarNaFila)\s*\(/);
  assert.doesNotMatch(code,/getDataRange\(|DriveApp\.|UrlFetchApp\.|MailApp\.|GmailApp\./);
});

test('consulta o chamado exato e a ultima versao comum, preserva OP integral e tipos',()=>{
  const {context,data,reads,opens}=fixture();
  const antes=JSON.stringify(data);
  const r=context._reSpPrepararDadosLegado_('2567','META');
  assert.equal(opens.length,1);
  assert.equal(r.pacote.id,'2567');
  assert.equal(r.pacote.origem,'legado');
  assert.equal(r.pacote.versaoPpg,2);
  assert.equal(r.pacote.linkPpg,link);
  assert.equal(r.pacote.fontePlanilha,idFonte);
  assert.equal(r.espec.linhas.length,2);
  assert.equal(r.arvore.linhas.length,3);
  assert.equal(r.tiragem.linhas.length,3);
  assert.equal(r.tiragem.linhas.reduce((s,row)=>s+context._reSpQuant_(row[5]),0),1300);
  assert.equal(r.espec.linhas[0][4],46314.375);
  assert.equal(r.tiragem.linhas[0][6],46315.375);
  assert.equal(typeof r.tiragem.linhas[0][5],'number');
  assert.equal(typeof r.tiragem.linhas[1][5],'string');
  assert.ok(reads.every(r=>r.linha===1||r.linha>=3),'nenhuma leitura de dados do chamado 12567');
  assert.equal(JSON.stringify(data),antes);
});

test('preserva o modelo de custos escolhido sem exigir colunas de grafica na alocacao',()=>{
  const a=fixture().context._reSpPrepararDadosLegado_('2567','LOGPRINT');
  const b=fixture().context._reSpPrepararDadosLegado_('2567','FORMA CERTA');
  assert.equal(JSON.stringify(a),JSON.stringify(b));
});

test('datas usam o horario da planilha, incluindo meia-noite e horario de verao, sem alterar a fonte',()=>{
  const {context,data,dateFormats}=fixture();
  const madrugada=new Date('2026-10-20T01:30:15.125Z');
  const horarioVerao=new Date('2018-11-05T02:15:00Z');
  data['Especificação Técnica'][3][4]=madrugada;
  data.Tiragem[3][6]=horarioVerao;
  const antes=JSON.stringify(data);
  const r=context._reSpPrepararDadosLegado_('2567','META');
  const serial=(ano,mes,dia,hora,minuto,segundo=0,ms=0)=>
    (Date.UTC(ano,mes-1,dia,hora,minuto,segundo,ms)-Date.UTC(1899,11,30))/86400000;
  assert.equal(r.espec.linhas[0][4],serial(2026,10,19,22,30,15,125));
  assert.equal(r.tiragem.linhas[0][6],serial(2018,11,5,0,15));
  assert.ok(dateFormats.some(c=>c.date===madrugada&&c.zone==='America/Sao_Paulo'));
  assert.ok(dateFormats.some(c=>c.date===horarioVerao&&c.zone==='America/Sao_Paulo'));
  assert.equal(JSON.stringify(data),antes);
  assert.equal(data['Especificação Técnica'][3][4],madrugada);
});

test('numero com formato de data recupera o serial 20 e preserva os demais tipos',()=>{
  const {context,data}=fixture();
  // Sao Paulo ainda usava o deslocamento historico -03:06:28 em 1900.
  const vinteFormatadoComoData=new Date('1900-01-19T03:06:28Z');
  const extras=['Paginas com formato incorreto','Paginas numericas','Texto','Booleano','Zero'];
  data['Especificação Técnica'][0].push(...extras);
  data['Especificação Técnica'][3].push(vinteFormatadoComoData,20,'0020',false,0);
  const antes=JSON.stringify(data);
  const r=context._reSpPrepararDadosLegado_('2567','META');
  assert.deepEqual(Array.from(r.espec.linhas[0].slice(-5)),[20,20,'0020',false,0]);
  assert.equal(JSON.stringify(data),antes);
  assert.equal(data['Especificação Técnica'][3].at(-5),vinteFormatadoComoData);
});

test('nao retorna versao anterior quando a ultima esta incompleta em uma aba',()=>{
  const {context,data}=fixture();
  data['Árvore de produto']=data['Árvore de produto'].slice(0,3);
  assert.throws(()=>context._reSpPrepararDadosLegado_('2567','LOGPRINT'),/versão 2 incompleta.*Árvore de produto/);
});

test('bloqueia versoes ambiguas por link e valores de versao invalidos',()=>{
  const divergente=fixture();
  divergente.data.Tiragem[4][1]='https://drive.google.com/OUTRA-PPG';
  assert.throws(()=>divergente.context._reSpPrepararDadosLegado_('2567','LOGPRINT'),/links.*divergentes/);
  const invalida=fixture();invalida.data.Tiragem[4][7]='atual';
  assert.throws(()=>invalida.context._reSpPrepararDadosLegado_('2567','LOGPRINT'),/versão da PPG inválida/);
});

test('chamado ausente e cabecalho alterado produzem erro explicito',()=>{
  const ausente=fixture();
  assert.throws(()=>ausente.context._reSpPrepararDadosLegado_('9999','LOGPRINT'),/não encontrado na alocação antiga/);
  const cabecalho=fixture();cabecalho.data.Tiragem[0][0]='ID Solicitacao';
  assert.throws(()=>cabecalho.context._reSpPrepararDadosLegado_('2567','LOGPRINT'),/coluna ID \(Coluna A\) ausente/);
});

test('oferece especificacao geral para lacunas, mas exige arvore, CD e tiragem valida',()=>{
  const semEspec=fixture();semEspec.data['Especificação Técnica'].splice(3,1);
  assert.throws(()=>semEspec.context._reSpPrepararDadosLegado_('2567','LOGPRINT'),/SKU sem especificação.*SKU-A/);
  const semArvore=fixture();semArvore.data['Árvore de produto'].splice(4,1);
  assert.throws(()=>semArvore.context._reSpPrepararDadosLegado_('2567','LOGPRINT'),/Kit\/SKU sem árvore.*KIT-A \/ SKU-A/);
  const semCD=fixture();semCD.data.Tiragem[3][4]='';
  assert.throws(()=>semCD.context._reSpPrepararDadosLegado_('2567','LOGPRINT'),/sem SKU, kit ou CD/);
  const semTiragem=fixture();semTiragem.data.Tiragem[3][5]='';
  assert.throws(()=>semTiragem.context._reSpPrepararDadosLegado_('2567','LOGPRINT'),/Tiragem vazia/);
  const negativa=fixture();negativa.data.Tiragem[3][5]=-10;
  assert.throws(()=>negativa.context._reSpPrepararDadosLegado_('2567','LOGPRINT'),/Tiragem invalida/);
});

test('falta de toda especificacao da ultima versao oferece cada SKU sem reutilizar versao antiga',()=>{
  const {context,data}=fixture();
  data['Especificação Técnica']=data['Especificação Técnica'].slice(0,3);
  assert.throws(()=>context._reSpPrepararDadosLegado_('2567','LOGPRINT'),err=>{
    assert.equal(err.codigo,'ESPECIFICACAO_LEGADO_AUSENTE');
    assert.deepEqual(Array.from(err.especificacoesAusentes),['SKU-A','SKU-B']);
    return true;
  });
});

test('SKU ausente no cadastro geral sai da tiragem e da arvore sem impedir os demais SKUs',()=>{
  const {context,data}=fixture();
  data['Especificação Técnica'].splice(3,1);
  data['Árvore de produto'].splice(4,1); // A ausencia de arvore do SKU excluido tambem nao bloqueia B.
  context._reEspecGeralConsultar_=()=>({products:[],dividirPesoPor100000:false});
  const r=context._reSpPrepararDadosLegado_('2567','META',[{chamado:'2567',sku:'SKU-A'}]);
  assert.deepEqual(Array.from(r.skusIgnorados),['SKU-A']);
  assert.deepEqual(Array.from(r.skusEspecificacaoGeral),[]);
  assert.equal(r.tiragemIgnorada,1200);assert.equal(r.semSkusParaGerar,false);
  assert.equal(r.espec.linhas.length,1);
  assert.equal(context._reSpBuscar_(r.espec.linhas[0],r.espec.headers,'1. ID_Código SKU'),'SKU-B');
  assert.equal(r.tiragem.linhas.length,1);
  assert.equal(context._reSpBuscar_(r.tiragem.linhas[0],r.tiragem.headers,'16. TIR_SKU'),100);
  assert.deepEqual(Array.from(r.arvore.linhas,row=>context._reSpBuscar_(row,r.arvore.headers,'1. ID_Código SKU')),['','SKU-B']);
  assert.match(r.avisos.join(' '),/SKU-A.*não será incluído/);
});

test('kit sem componentes incluidos e removido e chamado totalmente sem SKUs fica identificado',()=>{
  const {context,data}=fixture();
  data['Especificação Técnica']=data['Especificação Técnica'].filter((r,i)=>i===0||!['SKU-A','SKU-B'].includes(r[2]));
  context._reEspecGeralConsultar_=()=>({products:[],dividirPesoPor100000:false});
  const r=context._reSpPrepararDadosLegado_('2567','META',[
    {chamado:'2567',sku:'SKU-A'},{chamado:'2567',sku:'SKU-B'}]);
  assert.equal(r.semSkusParaGerar,true);assert.equal(r.tiragemIgnorada,1300);
  assert.equal(r.espec.linhas.length,0);assert.equal(r.tiragem.linhas.length,0);assert.equal(r.arvore.linhas.length,0);
  assert.deepEqual(Array.from(r.skusIgnorados),['SKU-A','SKU-B']);
});

test('pai que repete o codigo do kit na coluna SKU nao consulta geral nem duplica tiragem',()=>{
  const {context,data}=fixture();
  data.Tiragem.push([2567,link,'KIT-A','KIT-A','CITA',1200,'',2,'PAI']);
  context._reEspecGeralConsultar_=()=>{throw Error('Pai estrutural nao pode consultar especificacao geral');};
  const r=context._reSpPrepararDadosLegado_('2567','META');
  assert.equal(r.espec.linhas.length,2);assert.equal(r.tiragem.linhas.length,3);
  assert.equal(r.tiragem.linhas.reduce((n,row)=>n+context._reSpQuant_(context._reSpBuscar_(row,r.tiragem.headers,'16. TIR_SKU')),0),1300);
  assert.deepEqual(Array.from(r.skusEspecificacaoGeral),[]);assert.deepEqual(Array.from(r.skusIgnorados),[]);
});

test('avulso SKU igual ao kit permanece quando arvore comprova esse produto, mesmo com outros filhos',()=>{
  const {context,data}=fixture();
  data.Tiragem.push([2567,link,'KIT-A','KIT-A','CITA',50,'',2,'AVULSO']);
  data['Árvore de produto'].push([2567,link,'','KIT-A','KIT-A',1,'Caixa',2]);
  data['Especificação Técnica'].push([2567,link,'KIT-A','SAS','',2,50]);
  const r=context._reSpPrepararDadosLegado_('2567','META');
  assert.equal(r.espec.linhas.length,3);assert.equal(r.tiragem.linhas.length,4);
  assert.equal(context._reSpBuscar_(r.tiragem.linhas.at(-1),r.tiragem.headers,'1. ID_Código SKU'),'KIT-A');
});

test('codigo igual ao kit so e considerado pai quando arvore e CD comprovam a estrutura',()=>{
  const outroCD=fixture();outroCD.data.Tiragem.push([2567,link,'KIT-A','KIT-A','CD-SEM-FILHO',50,'',2,'CD']);
  assert.throws(()=>outroCD.context._reSpPrepararDadosLegado_('2567','META'),err=>{
    assert.equal(err.codigo,'ESPECIFICACAO_LEGADO_AUSENTE');
    assert.ok(Array.from(err.especificacoesAusentes).includes('KIT-A'));return true;
  });
  const semPai=fixture();semPai.data['Árvore de produto'].splice(3,1);
  semPai.data.Tiragem.push([2567,link,'KIT-A','KIT-A','CITA',50,'',2,'SEM-PAI']);
  assert.throws(()=>semPai.context._reSpPrepararDadosLegado_('2567','META'),err=>{
    assert.equal(err.codigo,'ESPECIFICACAO_LEGADO_AUSENTE');return true;
  });
});

test('numeros invalidos sao recusados antes de ler a planilha',()=>{
  const {context,opens}=fixture();
  ['2567,2448','2567 OR 1','',null].forEach(id=>assert.throws(()=>context._reSpPrepararDadosLegado_(id,'LOGPRINT'),/Número do chamado inválido/));
  assert.equal(opens.length,0);
});

test('chamado 2567 real: pai estrutural nao dobra os 200 exemplares e mantem embalagem da arvore',()=>{
  const {context,data}=fixture();
  const ppg='https://drive.google.com/open?id=10Bv67LB8qVKuTV4HhMajieXBEBZJYoxB';
  data['Especificação Técnica']=[
    [...meta,'1. ID_Código SKU','2. INF_Marca','Versão da PPG'],
    [2567,ppg,'PL26C6A1AC30','Pleno',1]
  ];
  data['Árvore de produto']=[
    [...meta,'1. ID_Código Kit','1. ID_Código SKU','13. MAN_Total de itens colecionados','13. MAN_Embalagem','13. MAN_Observação de embalagem','Versão da PPG'],
    [2567,ppg,'KAF3PL26C6A2','-',1,'Caixa','Caixa parda coletiva fornecida pela Arco',1],
    [2567,ppg,'KAF3PL26C6A2','PL26C6A1AC30',1,'-','',1]
  ];
  data.Tiragem=[
    [...meta,'1. ID_Código Kit','1. ID_Código SKU','15. DES_CD envio','16. TIR_SKU','Versão da PPG'],
    [2567,ppg,'KAF3PL26C6A2','-','RIBEIRÃO PRETO',200,1],
    [2567,ppg,'KAF3PL26C6A2','PL26C6A1AC30','RIBEIRÃO PRETO',200,1]
  ];
  const r=context._reSpPrepararDadosLegado_('2567','META');
  assert.equal(r.espec.linhas.length,1);
  assert.equal(r.arvore.linhas.length,2);
  assert.equal(r.arvore.linhas[0][3],'-');
  assert.equal(r.tiragem.linhas.length,1);
  assert.equal(r.tiragem.linhas[0][5],200);
  const cab=['Status SKU 2026','1. ID_Código SKU'];
  const espec={headers:cab,linhas:context._reSpMontarBase_(r.espec,cab)};
  const nc=context._reSpConstrucaoNC_(espec,r.arvore,r.tiragem,[]);
  assert.equal(nc.especificacoes.length,1);
  assert.equal(nc.calc.length,1);
  assert.equal(nc.calc[0].sku,'PL26C6A1AC30');
  assert.equal(nc.calc[0].quant,200);
  assert.equal(nc.calc[0].emb,'Caixa');
  assert.equal(nc.calc[0].elementos,1);
  assert.equal(nc.porSku.PL26C6A1AC30.total,200);
  data.Tiragem[1][3]='';
  data['Árvore de produto'][1][3]='';
  const vazio=context._reSpPrepararDadosLegado_('2567','META');
  assert.equal(vazio.tiragem.linhas.length,1);
  assert.equal(vazio.arvore.linhas.length,2);
});

test('pai sem SKU so e descartado quando existe filho no mesmo kit e CD',()=>{
  const {context,data}=fixture();
  data.Tiragem.push([2567,link,'KIT-INCOMPLETO','-','CITA',200,'',2,'PAI']);
  assert.throws(()=>context._reSpPrepararDadosLegado_('2567','LOGPRINT'),/kit sem SKU no CD.*KIT-INCOMPLETO/);
  const outroCD=fixture();
  outroCD.data.Tiragem.push([2567,link,'KIT-A','-','CD-SEM-FILHO',200,'',2,'PAI']);
  assert.throws(()=>outroCD.context._reSpPrepararDadosLegado_('2567','LOGPRINT'),/kit sem SKU no CD.*KIT-A \/ CD-SEM-FILHO/);
});
