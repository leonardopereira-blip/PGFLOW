# Validacao de custos de reentradas

As calculadoras de Logprint, Casa da Grafica, Forma Certa, Meta e Reproset usam moldes incorporados ao codigo. O De-Para e comum. Cada calculadora preserva as formulas, precos, abas e estilos do arquivo recebido. O portal preenche as entradas e gera os dois XLSX; o Excel recalcula as formulas ao abrir.

## Origem dos dados

O caminho normal usa o chamado e sua versao na API do SPG. Apenas a ausencia comprovada do chamado permite a pergunta sobre solicitacao antiga. A escolha e individual e explicita. Falhas de acesso, timeout, cancelamento e dados incompletos nao habilitam o legado.

O modulo temporario `LegadoAlocacaoReentradas.js` le somente as abas Especificacao Tecnica, Arvore de produto e Tiragem da alocacao antiga configurada. Usa a maior Versao da PPG e exige o mesmo link nas tres matrizes; nao abre o link nem grava na planilha. Pais estruturais sem SKU nao duplicam a tiragem dos filhos. A origem escolhida muda apenas essas matrizes: acabamento, fornecimento, modelos e geracao continuam iguais.

Ao encerrar a transicao, remover o modulo temporario e a selecao de origem da interface e de `_reSpDadosChamado_`.

## Forma Certa

O layout da CALC Geral tem entradas distintas e uma aba adicional IMPUT Forma Certa. A adaptacao preserva as formulas da calculadora e completa as formulas da IMPUT conforme os SKUs do chamado, ate 300.

O arquivo recebido tinha a formula da coluna AH repetida ate a linha 1.048.576. `tools/mapear-forma-certa.cjs` verifica individualmente o padrao e conserva as linhas 1 a 305 da IMPUT; nenhuma outra parte e reduzida. Os metadados do molde registram a origem e os hashes da normalizacao. A fonte original nao e modificada.

## Lista inicial

`ReentradasRecentes.js` consulta exclusivamente a API, filtra o tipo REENTRADA, agrupa linhas do mesmo chamado e apresenta ate 30 numeros mais recentes. Status, marca, envio e Alocado Em usam os campos do relatorio. Alocado Em e a data de alocacao, sem substituir pelo nome da grafica. A falta de alocacao permanece vazia na origem e aparece como traco na tela.

## Verificacao e envio

Os testes conferem a fidelidade dos moldes, layouts de entrada, isolamento da origem antiga, datas, ausencia comprovada na API, interface e lista recente. Executar `node --test --test-concurrency=1 tests/reentradas-*.test.cjs tests/molde*.test.cjs`.

Antes de editar ou enviar, sincronizar GitHub e Apps Script conforme AGENTS.md. O workflow de testes confere a base lida antes do envio e rele todos os arquivos depois. Implantacoes versionadas de producao continuam sendo atualizadas somente por pedido explicito.
