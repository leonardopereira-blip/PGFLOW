# Fluxo de trabalho do PGFLOW

- Antes de qualquer alteracao ou envio, atualize a main pelo GitHub e leia tambem o HEAD do Apps Script. Preserve as pendencias locais e uma copia das fontes antes de conciliar. Edicoes diretas no Google devem ser incorporadas antes de qualquer push de codigo.
- Preserve a Validacao Principal, emails e demais fluxos fora do pedido. O legado temporario de reentradas muda somente a origem das tres matrizes apos escolha explicita e ausencia comprovada na API.
- As calculadoras e o De-Para sao gerados pelos moldes internos. Nao reintroduza leitura de arquivos-modelo no Drive. Preserve as formulas, precos e layouts de cada grafica e os testes de fidelidade.
- Use verificacoes proporcionais e envie alteracoes implementadas para main. Acompanhe a execucao do workflow e confirme os arquivos por releitura do Google antes de afirmar atualizacao do ambiente de testes.
- Atualize `.github/pgflow-base-google.json` com os hashes SHA-256 (CRLF normalizado para LF) do HEAD efetivamente lido antes da alteracao; o envio deve conferir novamente essa base e interromper se houver novas edicoes externas.
- Nunca crie versao nem altere implantacoes de producao sem pedido explicito. Nao execute publicar.cmd como parte dos ajustes.
- Credenciais ficam somente em autenticacao local e GitHub Secrets, nunca em arquivos de codigo, mensagens ou logs. Nao abra navegador de previa/teste a menos que o usuario solicite.
