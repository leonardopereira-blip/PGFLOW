const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),path=require('node:path');
const c={};vm.createContext(c);vm.runInContext(fs.readFileSync(path.join(__dirname,'..','GestaoFormalizacoes.js'),'utf8'),c);
const catalogo={Logprint:['leonardo@arcoeducacao.com.br','a@logprint.com.br','shared@x.com.br'],Rona:['b@rona.com.br','shared@x.com.br'],Teste:['test@x.com.br']};
function msg(over={}){return {getFrom:()=> 'Joao <joao.alberti@arcoeducacao.com.br>',getSubject:()=> 'Solicitação de prazo: Chamado 3100 - exemplo',getAttachments:()=>[{getName:()=> 'Chamado 3100.xlsx'}],getTo:()=> 'a@logprint.com.br, leonardo@arcoeducacao.com.br',getDate:()=>new Date('2026-11-24T14:00:00Z'),...over};}
const read=(m)=>c._gfLerMensagem_(m,{'3100':true},['joao.alberti@arcoeducacao.com.br'],c._gfMapaDestinos_(catalogo));
test('gestao usa correspondencia de assunto, anexo, remetente e grafica',()=>{const x=read(msg());assert.equal(x.id,'3100');assert.equal(x.graficas[0],'Logprint');assert.equal(x.ambiguo,false);});
test('gestao nao conta resposta, anexo de outro chamado, remetente alheio ou teste interno',()=>{
  for(const m of [msg({getSubject:()=> 'Re: Solicitação de prazo: Chamado 3100'}),msg({getAttachments:()=>[{getName:()=> 'Chamado 31000.xlsx'}]}),msg({getFrom:()=> 'other@x.com.br'}),msg({getTo:()=> 'leonardo@arcoeducacao.com.br'}),msg({getTo:()=> 'test@x.com.br'})])assert.equal(read(m),null);
});
test('gestao nao inventa grafica para destinatario ambiguo ou fora do cadastro',()=>{
  for(const email of ['shared@x.com.br','unknown@x.com.br']){const x=read(msg({getTo:()=>email}));assert.equal(x.graficas.length,0);assert.equal(x.ambiguo,true);}
});
test('gestao separada preserva o script e a tela de formalizacao',()=>{
  const html=fs.readFileSync(path.join(__dirname,'..','TelaEmail.html'),'utf8');
  assert.match(html,/id="pf-gestao-layout"/);assert.match(html,/id="pf-email"/);assert.match(html,/id="pf-gestao"/);assert.match(html,/obterGestaoFormalizacoes\(page,query\)/);
  assert.equal((html.match(/<script>/g)||[]).length,2);
  assert.doesNotThrow(()=>new Function(html.match(/<script>([\s\S]*?)<\/script>/g)[1].replace(/^<script>|<\/script>$/g,'')));
});

test('gestao agrega envios e nao inventa status quando a consulta falha',()=>{
  c.Session={getActiveUser:()=>({getEmail:()=> 'leonardo.pereira@arcoeducacao.com.br'}),getEffectiveUser:()=>({getEmail:()=> 'leonardo.pereira@arcoeducacao.com.br'})};
  c._formalizacaoNormalizarEmail_=x=>String(x||'').toLowerCase();
  c._formalizacaoRemetentes_=()=>[{email:'joao.alberti@arcoeducacao.com.br'}];
  c._gfNumerosSPG_=()=>['3101','3100'];
  c.obterDadosGraficasEmailSPG=()=>catalogo;
  c.Utilities={formatDate:d=>d.toISOString()};
  let queries=[];
  c.GmailApp={search:(q)=>{queries.push(q);return ['thread'];},getMessagesForThreads:()=>[[
    msg({getId:()=> '1'}),msg({getId:()=> '2',getTo:()=> 'b@rona.com.br',getDate:()=>new Date('2026-11-24T15:00:00Z')}),
    msg({getId:()=> '3',getTo:()=> 'test@x.com.br'})
  ]]};
  const res=c.obterGestaoFormalizacoes(1,'');
  assert.equal(res.chamados[0].status,'nao_localizado');
  assert.equal(res.chamados[1].status,'formalizado');
  assert.equal(JSON.stringify(res.chamados[1].graficas.map(g=>g.nome)),JSON.stringify(['Logprint','Rona']));
  assert.equal(res.chamados[1].dataEnvio,'2026-11-24T15:00:00.000Z');
  assert.match(queries[0],/in:sent/);
  c.GmailApp.search=()=>{throw Error('indisponivel');};
  const failed=c.obterGestaoFormalizacoes(1,'3100');
  assert.equal(failed.chamados[0].status,'indeterminado');assert.equal(failed.completo,false);
});
