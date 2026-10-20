# Validacao de custos de reentradas

As calculadoras de Logprint, Casa da Grafica, Forma Certa, Meta e Reproset usam moldes incorporados ao codigo. O De-Para e comum. Cada calculadora preserva as formulas, precos, abas e estilos do arquivo recebido. O portal gera o De-Para XLSX e a calculadora XLSM com macro manual; o Excel recalcula as formulas ao abrir.

## Origem dos dados

O caminho normal usa o chamado e sua versao na API do SPG. Apenas a ausencia comprovada do chamado permite a pergunta sobre solicitacao antiga. A escolha e individual e explicita. Falhas de acesso, timeout, cancelamento e dados incompletos nao habilitam o legado.

O modulo temporario `LegadoAlocacaoReentradas.js` le somente as abas Especificacao Tecnica, Arvore de produto e Tiragem da alocacao antiga configurada. Usa a maior Versao da PPG e exige o mesmo link nas tres matrizes; nao abre o link nem grava na planilha. Pais estruturais nao duplicam a tiragem dos filhos. Isso inclui SKU vazio, traco ou o codigo do proprio kit: neste ultimo caso, a arvore precisa comprovar o pai, sem filho com aquele SKU, e outro componente real precisa estar na tiragem do mesmo kit e CD. Um avulso verdadeiro com SKU igual ao kit continua sendo produto. A origem escolhida muda apenas essas matrizes: acabamento, fornecimento, modelos e geracao continuam iguais.

Ao encerrar a transicao, remover o modulo temporario e a selecao de origem da interface e de `_reSpDadosChamado_`.

Quando faltar especificacao de um SKU no chamado legado, a interface conserva o alerta e oferece o uso da especificacao tecnica geral, com aceite individual por chamado e SKU. `ReentradasEspecificacaoGeral.js` busca esses produtos diretamente na API do SPG e usa o mesmo mapeador do relatorio `SPG.js`: dimensoes, peso e referencias de troca de chapa e personalizacao da tiragem do proprio chamado. Nenhuma planilha externa e usada para completar a especificacao geral. As especificacoes existentes permanecem preservadas; arvore e tiragem continuam as do legado. O resumo e o resultado indicam os SKUs completados. A opcao nao se aplica aos chamados da API atual.

Os indicadores booleanos do relatorio SPG usam Sim/Nao e aceitam true/false ou sim/nao sem diferenciar maiusculas de minusculas, preservando numeros e textos livres.

Se o SKU autorizado tambem nao existir no catalogo geral completo da API, ele e omitido com aviso no resumo e no resultado. Sua tiragem e seus vinculos de arvore sao retirados da geracao; os demais SKUs continuam. Chamados que ficarem sem nenhum SKU disponivel nao geram arquivos vazios nem impedem os outros chamados do lote. Falhas de acesso, paginacao incompleta e especificacoes conflitantes nao equivalem a ausencia de cadastro.

O usuario tambem pode escolher ignorar diretamente os SKUs indicados com erro, por SKU ou todos os pendentes do chamado. A escolha explicita e vinculada ao chamado/SKU, dispensa a consulta geral desses itens e conserva o aviso de exclusao no resultado.

## Forma Certa

O layout da CALC Geral tem entradas distintas e uma aba adicional IMPUT Forma Certa. A adaptacao preserva as formulas da calculadora e completa as formulas da IMPUT conforme os SKUs do chamado, ate 300.

O arquivo recebido tinha a formula da coluna AH repetida ate a linha 1.048.576. `tools/mapear-forma-certa.cjs` verifica individualmente o padrao e conserva as linhas 1 a 305 da IMPUT; nenhuma outra parte e reduzida. Os metadados do molde registram a origem e os hashes da normalizacao. A fonte original nao e modificada.

## Lista inicial

`ReentradasRecentes.js` consulta exclusivamente a API, filtra o tipo REENTRADA, agrupa linhas do mesmo chamado e apresenta ate 30 numeros mais recentes. Status, marca, envio e Alocado Em usam os campos do relatorio. Alocado Em e a data de alocacao, sem substituir pelo nome da grafica. A falta de alocacao permanece vazia na origem e aparece como traco na tela.

## Verificacao e envio

Os testes conferem a fidelidade dos moldes, layouts de entrada, isolamento da origem antiga, datas, ausencia comprovada na API, especificacao geral, interface e lista recente. Executar `node --test --test-concurrency=1 tests/reentradas-*.test.cjs tests/molde*.test.cjs tests/spg-espec-normalizacao.test.cjs`.

Antes de editar ou enviar, sincronizar GitHub e Apps Script conforme AGENTS.md. O workflow de testes confere a base lida antes do envio e rele todos os arquivos depois. Implantacoes versionadas de producao continuam sendo atualizadas somente por pedido explicito.

Falhas de transporte HTTP 0 ou NetworkError nas consultas de resumo e validacao permitem uma repeticao automatica e depois nova tentativa manual, preservando as escolhas e descartando respostas antigas. Nao comprovam ausencia na API e nao habilitam o legado por si. A geracao de arquivos nao e repetida automaticamente. Falhas de consulta da fase inicial permanecem nessa fase.

## Tecnologia exclusiva e formulas por SKU

A tela inicia em DIGITAL e exibe somente os campos da tecnologia ativa. PLANA ou ROTATIVA podem ser escolhidas, mas nunca sao preenchidas em paralelo. O gerador recebe e utiliza somente o bloco escolhido, mesmo que haja dados antigos no payload. Na CALC Geral, as formulas originais permanecem apenas nas linhas dos SKUs presentes (a partir da linha 6); linhas excedentes do molde sao excluidas. As outras abas e formulas dos SKUs gerados permanecem inalteradas.

## Tipos de SKU e protecao original

Os campos de codigo SKU/kit usam numero quando o codigo consiste em ate 15 digitos canonicos, sem zeros iniciais. Codigos alfanumericos, zeros iniciais e codigos longos permanecem texto. A regra vale para as entradas do De-Para, NC, PPG Especificacoes e CALC Geral D, mantendo o mesmo tipo nas consultas. ISBN, textos livres, formulas e tarifas nao sao convertidos por essa regra.

A geracao conserva as tags de protecao, permissoes e estilos Locked/Hidden dos moldes protegidos. A Logprint foi fornecida sem protecao para permitir edicao: por orientacao do usuario, recebe o mesmo mapa de entradas editaveis da Casa da Grafica/Reproset, com senha PGFLOW2026, preservando seus formatos, precos e formulas. Nao se cria bloqueio geral de celulas no final da geracao.

A calculadora inclui `RestaurarProtecaoOriginal`, disponivel em Desenvolvedor > Macros, sem botao nem execucao automatica. Depois de desproteger manualmente uma aba, a macro restaura o mapa de celulas editaveis e bloqueadas do arquivo gerado e reaplica as permissoes originais. Ela ignora abas que continuam protegidas e abas originalmente livres. A reprotecao usa a senha `PGFLOW2026`; a geracao inicial mantem as senhas originais dos moldes. Valores, formulas, formatos visuais e custos nao sao regravados pela macro.

O mapa fica em uma parte Custom XML do XLSM, sem novas abas ou formulas. Faixas da macro excluem celulas mescladas, restauradas sempre como uma regiao inteira. O projeto VBA e protegido para visualizacao com a mesma senha PGFLOW2026. A fonte auditavel esta em `tools/RestaurarProtecaoOriginal.bas`; `tools/compilar-macro.py` gera o binario incorporado em `MacroProtecaoVba.js` exclusivamente a partir de fontes, sem templates VBA externos. Alterar a fonte exige recompilar e conferir o hash de fonte/binario.
