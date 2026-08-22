function buscarRelatorioAlocacaoSPG() {
  // ==========================================
  // PASSO 1: FAZER LOGIN E PEGAR O TOKEN
  // ==========================================
  var urlLogin = 'https://spg-api.prod.arcocv.co/api/auth/login';
  var credenciais = {
    "email": "custos@spg.arcoeducacao.com.br",
    "password": "DyyWUXzC0iRBIGVEs8qA0KrlX2KZyCNY"
  };
  
  var opcoesLogin = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(credenciais),
    'muteHttpExceptions': true
  };
  
  var respostaLogin = UrlFetchApp.fetch(urlLogin, opcoesLogin);
  var jsonLogin = JSON.parse(respostaLogin.getContentText());
  var token = jsonLogin.token;

  if (!token) {
    Logger.log("Erro ao pegar o token. Verifique as credenciais.");
    return;
  }

  // ==========================================
  // PASSO 2: LOOP DE PAGINAÇÃO (DE 500 EM 500)
  // ==========================================
  var limite = 500;
  var pagina = 1;
  var todosOsDados = [];
  var continuarBuscando = true;

  Logger.log("Iniciando a busca completa...");

  while (continuarBuscando) {
    Logger.log("Acessando a API: Página " + pagina + "...");
    
    // NOVO CAMINHO DA API (Tiragem / Relatório de Alocação)
    var urlRelatorio = 'https://spg-api.prod.arcocv.co/api/tiragem/relatorio-alocacao?page=' + pagina + '&limit=' + limite;
    var opcoesRelatorio = {
      'method': 'get',
      'headers': {
        'Authorization': 'Bearer ' + token 
      },
      'muteHttpExceptions': true
    };
    
    var respostaRelatorio = UrlFetchApp.fetch(urlRelatorio, opcoesRelatorio);
    var statusCode = respostaRelatorio.getResponseCode();
    
    if (statusCode !== 200) {
      Logger.log("Erro na API (Página " + pagina + "): Código " + statusCode);
      break; 
    }
    
    var json = JSON.parse(respostaRelatorio.getContentText());
    
    // BUSCA INTELIGENTE: Encontra automaticamente onde está o Array de dados no JSON
    var dadosDaPagina = null;
    if (Array.isArray(json)) {
      dadosDaPagina = json;
    } else if (typeof json === 'object') {
      for (var chave in json) {
        if (Array.isArray(json[chave])) {
          dadosDaPagina = json[chave];
          break; // Achou a lista, para de procurar
        }
      }
    }
    
    // Tratativa de segurança caso a API retorne algo inesperado
    if (!dadosDaPagina || !Array.isArray(dadosDaPagina)) {
      Logger.log("Atenção: A API não retornou uma lista reconhecível na página " + pagina + ".");
      break;
    }
    
    if (dadosDaPagina.length === 0) {
      break; // Fim da paginação (veio vazio)
    }
    
    todosOsDados = todosOsDados.concat(dadosDaPagina);
    
    // Se vieram menos itens do que o limite, significa que é a última página
    if (dadosDaPagina.length < limite) {
      continuarBuscando = false;
    } else {
      pagina++;
    }
  }

  if (todosOsDados.length === 0) {
    Logger.log("Nenhum dado foi baixado da API.");
    return;
  }

  // ==========================================
  // PASSO 3: ESCANEAR TODAS AS COLUNAS POSSÍVEIS
  // ==========================================
  Logger.log("Mapeando todas as colunas escondidas...");
  var mapeamentoColunas = {};
  
  for (var p = 0; p < todosOsDados.length; p++) {
    var chavesDoDado = Object.keys(todosOsDados[p]);
    for (var k = 0; k < chavesDoDado.length; k++) {
      mapeamentoColunas[chavesDoDado[k]] = true; 
    }
  }
  
  var cabecalhos = Object.keys(mapeamentoColunas);
  var matrizDados = [cabecalhos]; 

  Logger.log("Encontradas " + cabecalhos.length + " colunas diferentes.");

  // Preenche a tabela garantindo que a coluna certa receba o dado certo
  for (var i = 0; i < todosOsDados.length; i++) {
    var linha = [];
    for (var c = 0; c < cabecalhos.length; c++) {
      var nomeDaColuna = cabecalhos[c];
      var valor = todosOsDados[i][nomeDaColuna];
      
      // Tratamento de dados para evitar nulos e transformar objetos em texto
      if (valor === undefined || valor === null) {
        valor = "";
      } else if (typeof valor === 'object') {
        valor = JSON.stringify(valor);
      } else {
        valor = String(valor);
      }
      
      // TRAVA DE SEGURANÇA: Limite de 50.000 caracteres do Google Sheets
      if (typeof valor === 'string' && valor.length > 49000) {
        valor = valor.substring(0, 49000) + "... [CONTEÚDO TRUNCADO DEVIDO AO LIMITE DO GOOGLE SHEETS]";
      }
      
      linha.push(valor);
    }
    matrizDados.push(linha);
  }

  // ==========================================
  // PASSO 4: JOGAR TUDO NA PLANILHA DE DESTINO
  // ==========================================
  // NOVO ID da planilha destino
  var idPlanilhaDestino = "1j1RCyvofFLoSFc3eTdV8O3nf78Bxu4nIiXtWRjJnBKg";
  var ss = SpreadsheetApp.openById(idPlanilhaDestino);
  
  // NOVA Aba
  var nomeAba = "Tiragem";
  var aba = ss.getSheetByName(nomeAba);
  
  if (!aba) {
    aba = ss.insertSheet(nomeAba);
  } else {
    aba.clear();
  }
  
  // Grava tudo de uma vez
  aba.getRange(1, 1, matrizDados.length, cabecalhos.length).setValues(matrizDados);
  
  Logger.log("✅ Concluído! O relatório de alocação de tiragem foi sincronizado com sucesso.");
}