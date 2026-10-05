// ================== CONFIGURAÇÕES DO ALERTA ==================
var REPORT_CONFIG = {
  idPlanilha: '1aRUIzE1AA_qFGjZO9RcD51fggfok7-lTpf1IpRCCQ-w',
  
  abaConsolidado: 'ConsolidadoProducao',
  abaEspec: 'Espec',
  abaAntigo: 'Prod_Antigo',
  abaLog: 'Log_Consolidado', 
  
  slackEmail: 'teste_alteracoes_spg_-aaaavy4eh7d7k6wjwhceldpf5y@arco.org.slack.com'
};

var CABECALHO_ANTIGO = [
  'SKU', 'Marca', 'Grafica', 'CD Destino', 'Tiragem', 
  '4. CAP_Papel da capa', '4. CAP_Papel revestimento', 
  '4. CAP_Papel forro', '5. MIO_PRIN_Papel', '3,DIM_BxA(Aberto)(mm)', 'Grafico'
];

var CABECALHO_LOG = [
  'Data e Hora', 'Tipo de Alteração', 'Marca', 'SKU', 'Gráfica Atual', 
  'Gráfica Antiga / Origem', 'O que mudou (Detalhe)', 'Tiragem Atual', 'Diferença Tiragem'
];


// ======================================================================
// ETAPA 1: APENAS AUDITA, COMPARA E ENVIA O ALERTA (MODO DE TESTE LIVRE)
// ======================================================================
function testarEEnviarAlertasSlack() {
  var ss = SpreadsheetApp.openById(REPORT_CONFIG.idPlanilha);
  Logger.log('Iniciando auditoria de mudanças (Filtro: Somente Gráficos)...');
  
  var retratoHoje = _obterRetratoDeHoje(ss);
  if (!retratoHoje || retratoHoje.length === 0) return;

  var flatOntem = _lerAbaObjetos(ss, REPORT_CONFIG.abaAntigo);
  if (!flatOntem || flatOntem.length === 0) {
    Logger.log('Aviso: A aba Prod_Antigo está vazia. Rode a Etapa 2 primeiro.');
    return;
  }

  // Aplica o filtro de gráfico no Prod_Antigo (para garantir)
  var ontemFiltrado = [];
  for (var i = 0; i < flatOntem.length; i++) {
    var isG = String(flatOntem[i]['Grafico'] || flatOntem[i]['grafico'] || flatOntem[i]['GRAFICO'] || '').trim().toLowerCase();
    isG = isG.replace(/[áàãâä]/g, 'a').replace(/[éèẽêë]/g, 'e');
    if (isG === 'grafico') ontemFiltrado.push(flatOntem[i]);
  }

  if (ontemFiltrado.length === 0) {
    Logger.log('Aviso: A aba Prod_Antigo não tem nenhum item marcado como "Grafico". Rode a Etapa 2 primeiro.');
    return;
  }

  var layoutAntigoTemSpecs = (ontemFiltrado[0].hasOwnProperty('4. CAP_Papel da capa') || ontemFiltrado[0].hasOwnProperty('3,DIM_BxA(Aberto)(mm)'));

  var estadoHoje = _construirEstado(retratoHoje);
  var estadoOntem = _construirEstado(ontemFiltrado);

  Logger.log('Analisando as mudanças (Deltas)...');
  var alteracoes = _compararEstados(estadoHoje, estadoOntem, layoutAntigoTemSpecs);

  if (alteracoes.length > 0) {
    Logger.log('Encontradas ' + alteracoes.length + ' alterações reais! Gerando report e enviando para o Slack...');
    var dataHoraStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    
    _appendLog(ss, REPORT_CONFIG.abaLog, CABECALHO_LOG, alteracoes, dataHoraStr);
    
    // Captura o objeto da aba inteira (para cópia nativa sem estourar o timeout)
    var abaConsolidadoRaw = ss.getSheetByName(REPORT_CONFIG.abaConsolidado);

    var nomeNovoArquivo = "Report_Alteracoes_SPG_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");
    var urlPlanilhaAnexa = _criarPlanilhaDeReporte(nomeNovoArquivo, CABECALHO_LOG, alteracoes, dataHoraStr, abaConsolidadoRaw);
    
    _enviarAlertaSlackPorEmail(alteracoes, urlPlanilhaAnexa);
    Logger.log('Reporte enviado com sucesso!');
    
  } else {
    Logger.log('Nenhuma alteração na produção detectada hoje em relação ao Prod_Antigo.');
  }
}


// ======================================================================
// ETAPA 2: FECHA O DIA (SOBRESCREVE O PROD_ANTIGO PARA O DIA SEGUINTE)
// ======================================================================
function fecharODiaEAtualizarProdAntigo() {
  var ss = SpreadsheetApp.openById(REPORT_CONFIG.idPlanilha);
  Logger.log('Atualizando a base Prod_Antigo com o retrato do momento (Novo Layout)...');
  
  var retratoHoje = _obterRetratoDeHoje(ss);
  if (!retratoHoje || retratoHoje.length === 0) {
    Logger.log('⚠️ Falha ao atualizar Prod_Antigo: Não foi possível obter o retrato de hoje.');
    return;
  }
  
  var matriz = [CABECALHO_ANTIGO];
  for (var i = 0; i < retratoHoje.length; i++) {
    var r = retratoHoje[i];
    matriz.push([
      r['SKU'], r['Marca'], r['Grafica'], r['CD Destino'], r['Tiragem'],
      r['4. CAP_Papel da capa'], r['4. CAP_Papel revestimento'], r['4. CAP_Papel forro'], 
      r['5. MIO_PRIN_Papel'], r['3,DIM_BxA(Aberto)(mm)'], 'Grafico'
    ]);
  }
  
  _gravarAbaExata(ss, REPORT_CONFIG.abaAntigo, matriz);
  Logger.log('✅ Prod_Antigo atualizado com sucesso! Pronto para o próximo ciclo.');
}


// ================== MOTOR DE CRUZAMENTO (CRIA O RETRATO ATUAL) ==================

function _obterRetratoDeHoje(ss) {
  var rawConsolidado = _lerAbaObjetos(ss, REPORT_CONFIG.abaConsolidado); 
  var rawEspec = _lerAbaObjetos(ss, REPORT_CONFIG.abaEspec); 
  
  if (!rawConsolidado || rawConsolidado.length === 0) {
    Logger.log('ERRO: A aba ConsolidadoProducao não existe ou está fisicamente vazia (0 linhas).');
    return null;
  }
  
  Logger.log('Sucesso: Aba ConsolidadoProducao lida com ' + rawConsolidado.length + ' linhas. Aplicando filtro de Gráficos...');

  var dicEspec = {};
  for (var i = 0; i < rawEspec.length; i++) {
    var skuEspec = String(rawEspec[i]['1. ID_Código SKU'] || '').trim();
    if (skuEspec) {
      dicEspec[skuEspec] = {
        capa: rawEspec[i]['4. CAP_Papel da capa'] || '',
        rev: rawEspec[i]['4. CAP_Papel revestimento'] || '',
        forro: rawEspec[i]['4. CAP_Papel forro'] || '',
        miolo: rawEspec[i]['5. MIO_PRIN_Papel'] || '',
        formato: rawEspec[i]['3,DIM_BxA(Aberto)(mm)'] || rawEspec[i]['3. DIM_BxA (Aberto)(mm)'] || ''
      };
    }
  }

  var retrato = [];
  for (var c = 0; c < rawConsolidado.length; c++) {
    var r = rawConsolidado[c];
    
    var isG = String(r['Grafico'] || r['grafico'] || r['GRAFICO'] || '').trim().toLowerCase();
    isG = isG.replace(/[áàãâä]/g, 'a').replace(/[éèẽêë]/g, 'e'); 
    
    if (isG !== 'grafico') continue;
    
    var sku = String(r['SKU'] || r['sku'] || '').trim();
    var marca = r['Marca'] || r['marca'] || '';
    var destino = r['CD Destino'] || r['cdDestino'] || r['Destino'] || '';
    var grafica = r['Grafica PCP'] || r['GRAFICA'] || r['Grafica'] || '';
    var tiragem = parseFloat(r['Quantidade'] || r['Tiragem SKU'] || r['quantidade']) || 0;
    
    if (!sku) continue;

    var spc = dicEspec[sku] || { capa: '', rev: '', forro: '', miolo: '', formato: '' };
    
    retrato.push({
      'SKU': sku, 
      'Marca': marca, 
      'Grafica': grafica, 
      'CD Destino': destino, 
      'Tiragem': tiragem,
      '4. CAP_Papel da capa': spc.capa, 
      '4. CAP_Papel revestimento': spc.rev, 
      '4. CAP_Papel forro': spc.forro, 
      '5. MIO_PRIN_Papel': spc.miolo, 
      '3,DIM_BxA(Aberto)(mm)': spc.formato
    });
  }
  
  if (retrato.length === 0) {
    Logger.log('ERRO: Nenhuma linha gráfica encontrada.');
    return null;
  }
  
  Logger.log('Retrato montado com sucesso! Encontrados ' + retrato.length + ' itens gráficos válidos.');
  return retrato;
}


// ================== MOTOR DE COMPARAÇÃO (INTELIGÊNCIA) ==================

function _construirEstado(flatArray) {
  var estado = {};
  for (var i = 0; i < flatArray.length; i++) {
    var row = flatArray[i];
    
    var sku = _obterValorSeguro(row, ['SKU', 'sku', '1. ID_Código SKU', 'itemCode']);
    var marca = _obterValorSeguro(row, ['Marca', 'marca', '2. INF_Marca']);
    var graf = _obterValorSeguro(row, ['Grafica PCP', 'Grafica', 'GRAFICA', 'grafica']);
    var cd = _obterValorSeguro(row, ['CD Destino', 'cdDestino', 'Destino', '15. DES_CD envio']);
    var tir = parseFloat(_obterValorSeguro(row, ['Quantidade', 'Tiragem', 'quantidade', 'Tiragem SKU', '16. TIR_SKU'])) || 0;

    if (!sku) continue;

    if (!estado[sku]) {
      estado[sku] = {
        marca: marca,
        specs: {
          capa: _obterValorSeguro(row, ['4. CAP_Papel da capa']),
          rev: _obterValorSeguro(row, ['4. CAP_Papel revestimento']),
          forro: _obterValorSeguro(row, ['4. CAP_Papel forro']),
          miolo: _obterValorSeguro(row, ['5. MIO_PRIN_Papel']),
          formato: _obterValorSeguro(row, ['3,DIM_BxA(Aberto)(mm)', '3. DIM_BxA (Aberto)(mm)'])
        },
        graficas: {}
      };
    }
    
    if (!estado[sku].graficas[graf]) estado[sku].graficas[graf] = { tiragem: 0, cds: [] };
    
    estado[sku].graficas[graf].tiragem += tir;
    if (cd !== '' && estado[sku].graficas[graf].cds.indexOf(cd) === -1) {
      estado[sku].graficas[graf].cds.push(cd);
    }
  }
  return estado;
}

function _compararEstados(hj, ont, comparaSpecs) {
  var alt = [];

  for (var sku in hj) {
    var t = hj[sku];
    var y = ont[sku];

    var grafsHj = Object.keys(t.graficas);
    var strGrafHj = grafsHj.join(', ');
    var totalTirHj = _somarTiragem(t.graficas);

    if (!y) {
      alt.push(['ENTRADA DE SKU', t.marca, sku, strGrafHj, '-', 'Novo SKU na produção', totalTirHj, '+' + totalTirHj]);
      continue;
    }

    var grafsOnt = Object.keys(y.graficas);
    var strGrafOnt = grafsOnt.join(', ');

    var grafsEntraram = grafsHj.filter(function(g) { return grafsOnt.indexOf(g) === -1; });
    var grafsSairam = grafsOnt.filter(function(g) { return grafsHj.indexOf(g) === -1; });
    
    if (grafsEntraram.length > 0 || grafsSairam.length > 0) {
      alt.push(['TRANSFERÊNCIA DE GRÁFICA', t.marca, sku, strGrafHj, strGrafOnt, 'Transferido de ' + strGrafOnt + ' para ' + strGrafHj, totalTirHj, '-']);
    }

    var commonGrafs = grafsHj.filter(function(g) { return grafsOnt.indexOf(g) !== -1; });
    
    for (var i = 0; i < commonGrafs.length; i++) {
      var g = commonGrafs[i];
      var tG = t.graficas[g];
      var yG = y.graficas[g];

      if (tG.tiragem !== yG.tiragem) {
        var dif = tG.tiragem - yG.tiragem;
        var sinal = dif > 0 ? '+' : '';
        alt.push(['ALTERAÇÃO DE TIRAGEM', t.marca, sku, g, g, 'Tiragem mudou de ' + yG.tiragem + ' para ' + tG.tiragem, tG.tiragem, sinal + dif]);
      }

      var tCD = tG.cds.sort().join(', ') || 'Nenhum';
      var yCD = yG.cds.sort().join(', ') || 'Nenhum';
      if (tCD !== yCD) {
        alt.push(['ALTERAÇÃO DE CD DESTINO', t.marca, sku, g, g, 'Mudou de [' + yCD + '] para [' + tCD + ']', tG.tiragem, '-']);
      }
    }

    if (comparaSpecs) {
      var chavesSpec = ['capa', 'rev', 'forro', 'miolo', 'formato'];
      var nomesSpec = ['Papel da Capa', 'Papel Revestimento', 'Papel Forro', 'Papel do Miolo', 'Formato Aberto'];
      
      for (var s = 0; s < chavesSpec.length; s++) {
        var k = chavesSpec[s];
        if (t.specs[k] !== y.specs[k]) {
          var oldS = y.specs[k] || 'Em branco'; 
          var newS = t.specs[k] || 'Em branco'; 
          var grafRepresentante = grafsHj[0] || '-'; 
          alt.push(['ESPEC. TÉCNICA ALTERADA (' + nomesSpec[s] + ')', t.marca, sku, grafRepresentante, grafRepresentante, 'De: ' + oldS + ' Para: ' + newS, totalTirHj, '-']);
        }
      }
    }
  }

  for (var sku in ont) {
    if (!hj[sku]) {
      var y = ont[sku];
      var strGrafOnt = Object.keys(y.graficas).join(', ');
      var totalTirOnt = _somarTiragem(y.graficas);
      alt.push(['SAÍDA DE SKU', y.marca, sku, '-', strGrafOnt, 'SKU não aparece mais no consolidado', 0, '-' + totalTirOnt]);
    }
  }

  return alt;
}

function _somarTiragem(graficasObj) {
  var soma = 0;
  for (var g in graficasObj) soma += graficasObj[g].tiragem;
  return soma;
}


// ================== FUNÇÕES DE REPORT, LOG E E-MAIL BLINDADAS ==================

function _criarPlanilhaDeReporte(nomeArquivo, cabecalhos, linhasAlt, dataStr, abaConsolidadoOrigem) {
  var newSs = SpreadsheetApp.create(nomeArquivo);
  
  // 1. Aba principal: Alterações
  var aba = newSs.getSheets()[0];
  aba.setName('Alteracoes');
  
  var dados = [];
  dados.push(cabecalhos);
  
  for (var i = 0; i < linhasAlt.length; i++) {
    var l = linhasAlt[i].slice();
    l.unshift(dataStr); 
    var linhaSegura = [];
    for (var j = 0; j < cabecalhos.length; j++) {
      linhaSegura.push(_sanitizeVal(l[j]));
    }
    dados.push(linhaSegura);
  }
  
  var totalLinhas = dados.length;
  var totalCols = cabecalhos.length;

  if (aba.getMaxRows() < totalLinhas) aba.insertRowsAfter(aba.getMaxRows(), totalLinhas - aba.getMaxRows());
  if (aba.getMaxColumns() < totalCols) aba.insertColumnsAfter(aba.getMaxColumns(), totalCols - aba.getMaxColumns());
  
  SpreadsheetApp.flush(); 
  
  var chunkSize = 2000;
  for (var k = 0; k < totalLinhas; k += chunkSize) {
    var chunk = dados.slice(k, k + chunkSize);
    aba.getRange(k + 1, 1, chunk.length, totalCols).setValues(chunk);
  }
  aba.getRange(1, 1, 1, totalCols).setFontWeight("bold");
  
  // 2. Nova aba: Consolidado Atual Completo usando cópia nativa
  if (abaConsolidadoOrigem) {
    var novaAba = abaConsolidadoOrigem.copyTo(newSs);
    novaAba.setName('Consolidado Atual');
  }
  
  SpreadsheetApp.flush(); 
  
  try {
    DriveApp.getFileById(newSs.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {
    Logger.log("Aviso: Permissão ANYONE_WITH_LINK bloqueada. Tentando DOMAIN_WITH_LINK...");
    try {
      DriveApp.getFileById(newSs.getId()).setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(e2) {
      Logger.log("Aviso 2: Compartilhamento restrito aplicado devido às permissões corporativas.");
    }
  }
  
  return newSs.getUrl();
}

function _appendLog(ss, nomeAba, cabecalhos, linhasAlt, dataStr) {
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) {
    aba = ss.insertSheet(nomeAba);
    aba.appendRow(cabecalhos);
    aba.getRange(1, 1, 1, cabecalhos.length).setFontWeight("bold");
  }
  
  var dados = [];
  for (var i = 0; i < linhasAlt.length; i++) {
    var l = linhasAlt[i].slice();
    l.unshift(dataStr); 
    var linhaSegura = [];
    for (var j = 0; j < cabecalhos.length; j++) {
      linhaSegura.push(_sanitizeVal(l[j]));
    }
    dados.push(linhaSegura);
  }
  
  var linhaInicial = aba.getLastRow() + 1;
  var totalLinhas = dados.length;
  var totalNecessario = linhaInicial + totalLinhas - 1;
  
  if (aba.getMaxRows() < totalNecessario) aba.insertRowsAfter(aba.getMaxRows(), totalNecessario - aba.getMaxRows());
  
  SpreadsheetApp.flush(); 
  
  var chunkSize = 2000;
  for (var k = 0; k < totalLinhas; k += chunkSize) {
    var chunk = dados.slice(k, k + chunkSize);
    aba.getRange(linhaInicial + k, 1, chunk.length, cabecalhos.length).setValues(chunk);
  }
  SpreadsheetApp.flush();
}

// FORMATADOR DE NÚMEROS (15000 -> 15.000)
function _fmtNum(num) {
  if (isNaN(num)) return "0";
  return String(num).replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

function _enviarAlertaSlackPorEmail(alteracoes, urlPlanilha) {
  var resumoGeral = {
    entradas: { skus: 0, tiragem: 0 },
    saidas: { skus: 0, tiragem: 0 },
    transf: { skus: 0, tiragem: 0 },
    cd: { skus: 0, tiragem: 0 },
    tiragem: { skus: 0, varAbsoluta: 0 },
    espec: { skus: 0, tiragem: 0 }
  };
  
  var resumoMarca = {};
  var resumoGrafica = {};

  for (var i = 0; i < alteracoes.length; i++) {
    var t = alteracoes[i][0];
    
    var marca = alteracoes[i][1];
    if (!marca || String(marca).trim() === '' || marca === '-') marca = 'Em branco';
    
    var grafAtual = alteracoes[i][3];
    var grafAntiga = alteracoes[i][4];
    var grafica = (grafAtual && grafAtual !== '-') ? grafAtual : grafAntiga; 
    if (!grafica || String(grafica).trim() === '' || grafica === '-') grafica = 'Em branco';

    var volAtual = parseFloat(String(alteracoes[i][6]).replace(/[^\d.-]/g, '')) || 0;
    var difVal = parseFloat(String(alteracoes[i][7]).replace(/[^\d.-]/g, '')) || 0;

    if (!resumoMarca[marca]) resumoMarca[marca] = { entrada: 0, saida: 0, transfGraf: 0, altTiragem: 0, altCD: 0, altPapel: 0, altFormato: 0, totalTiragem: 0 };
    if (!resumoGrafica[grafica]) resumoGrafica[grafica] = { entrada: 0, saida: 0, transfGraf: 0, altTiragem: 0, altCD: 0, altPapel: 0, altFormato: 0, totalTiragem: 0 };

    var cat = null;
    var impactoLinha = 0;

    if (t.indexOf('ENTRADA') !== -1) {
      resumoGeral.entradas.skus++;
      resumoGeral.entradas.tiragem += volAtual;
      cat = 'entrada';
      impactoLinha = volAtual;
    } else if (t.indexOf('SAÍDA') !== -1) {
      resumoGeral.saidas.skus++;
      resumoGeral.saidas.tiragem += Math.abs(difVal); 
      cat = 'saida';
      impactoLinha = Math.abs(difVal);
    } else if (t.indexOf('TRANSFERÊNCIA') !== -1) {
      resumoGeral.transf.skus++;
      resumoGeral.transf.tiragem += volAtual; 
      cat = 'transfGraf';
      impactoLinha = volAtual;
    } else if (t.indexOf('TIRAGEM') !== -1) {
      resumoGeral.tiragem.skus++;
      resumoGeral.tiragem.varAbsoluta += Math.abs(difVal); 
      cat = 'altTiragem';
      impactoLinha = Math.abs(difVal);
    } else if (t.indexOf('CD') !== -1) {
      resumoGeral.cd.skus++;
      resumoGeral.cd.tiragem += volAtual;
      cat = 'altCD';
      impactoLinha = volAtual;
    } else if (t.indexOf('ESPEC') !== -1) {
      resumoGeral.espec.skus++;
      resumoGeral.espec.tiragem += volAtual;
      impactoLinha = volAtual;
      
      // Separando Papel de Formato Aberto
      if (t.indexOf('Formato Aberto') !== -1) {
        cat = 'altFormato';
      } else {
        cat = 'altPapel';
      }
    }

    if (cat) {
      resumoMarca[marca][cat] += impactoLinha;
      resumoMarca[marca].totalTiragem += impactoLinha;
      
      resumoGrafica[grafica][cat] += impactoLinha;
      resumoGrafica[grafica].totalTiragem += impactoLinha;
    }
  }
  
  function _gerarTabelaHTML(dados, tituloColuna1) {
    var chaves = Object.keys(dados).sort(function(a, b) { return dados[b].totalTiragem - dados[a].totalTiragem; });

    var html = "<table style='border-collapse: collapse; width: 100%; max-width: 900px; font-family: Arial, sans-serif; font-size: 13px; text-align: center; border: 1px solid #ddd;'>";
    html += "<tr style='background-color: #f4f4f4; border-bottom: 2px solid #ccc; font-size: 12px;'>" +
            "<th style='padding: 8px; text-align: left;'>" + tituloColuna1 + "</th>" +
            "<th style='padding: 8px;'>Entrada</th>" +
            "<th style='padding: 8px;'>Saida</th>" +
            "<th style='padding: 8px;'>Transf. Graf.</th>" +
            "<th style='padding: 8px;'>Alt. Tiragem</th>" +
            "<th style='padding: 8px;'>Alt. CD</th>" +
            "<th style='padding: 8px;'>Alt. Papel</th>" +
            "<th style='padding: 8px;'>Alt. Form. Aberto</th>" +
            "<th style='padding: 8px; background-color: #e9ecef;'>Total Tiragem Impactada</th>" +
            "</tr>";

    var limite = Math.min(chaves.length, 15);
    for (var k = 0; k < limite; k++) {
      var nome = chaves[k];
      var d = dados[nome];
      var bg = (k % 2 === 0) ? "#ffffff" : "#f9f9f9"; 
      
      html += "<tr style='background-color: " + bg + "; border-bottom: 1px solid #eee;'>" +
              "<td style='padding: 8px; text-align: left;'><b>" + nome + "</b></td>" +
              "<td style='padding: 8px;'>" + _fmtNum(d.entrada) + "</td>" +
              "<td style='padding: 8px;'>" + _fmtNum(d.saida) + "</td>" +
              "<td style='padding: 8px;'>" + _fmtNum(d.transfGraf) + "</td>" +
              "<td style='padding: 8px;'>" + _fmtNum(d.altTiragem) + "</td>" +
              "<td style='padding: 8px;'>" + _fmtNum(d.altCD) + "</td>" +
              "<td style='padding: 8px;'>" + _fmtNum(d.altPapel) + "</td>" +
              "<td style='padding: 8px;'>" + _fmtNum(d.altFormato) + "</td>" +
              "<td style='padding: 8px; background-color: #f1f3f5;'><b>" + _fmtNum(d.totalTiragem) + "</b></td>" +
              "</tr>";
    }
    html += "</table>";
    
    if (chaves.length > 15) {
      html += "<p style='font-size: 11px; color: #777; margin-top: 4px;'><i>* Exibindo o Top 15 de " + chaves.length + " " + tituloColuna1.toLowerCase() + "s. Veja os detalhes na planilha.</i></p>";
    }
    return html;
  }

  // Define dinamicamente o texto do comparativo
  var dataAtualStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  var textoComparativo = "";
  if (dataAtualStr === "03/09/2026") {
    textoComparativo = "📅 <b>Comparativo:</b> Hoje (" + dataAtualStr + ") vs 28/08/2026";
  } else {
    textoComparativo = "📅 <b>Comparativo:</b> Hoje (" + dataAtualStr + ") vs D-1 (Ontem)";
  }

  var resumo = "";
  if (resumoGeral.entradas.skus > 0) resumo += "🟢 <b>" + resumoGeral.entradas.skus + "</b> SKUs entraram na produção <span style='color: #666;'>(+ " + _fmtNum(resumoGeral.entradas.tiragem) + " na Tiragem)</span><br>";
  if (resumoGeral.saidas.skus > 0) resumo += "🔴 <b>" + resumoGeral.saidas.skus + "</b> SKUs saíram da produção <span style='color: #666;'>(- " + _fmtNum(resumoGeral.saidas.tiragem) + " na Tiragem)</span><br>";
  if (resumoGeral.transf.skus > 0) resumo += "🚚 <b>" + resumoGeral.transf.skus + "</b> SKUs mudaram de Gráfica <span style='color: #666;'>(Tiragem transf.: " + _fmtNum(resumoGeral.transf.tiragem) + ")</span><br>";
  if (resumoGeral.cd.skus > 0) resumo += "📦 <b>" + resumoGeral.cd.skus + "</b> SKUs mudaram de CD <span style='color: #666;'>(Tiragem impactada: " + _fmtNum(resumoGeral.cd.tiragem) + ")</span><br>";
  if (resumoGeral.tiragem.skus > 0) resumo += "📈 <b>" + resumoGeral.tiragem.skus + "</b> SKUs tiveram a tiragem ajustada <span style='color: #666;'>(Ajuste absoluto: " + _fmtNum(resumoGeral.tiragem.varAbsoluta) + " na Tiragem)</span><br>";
  if (resumoGeral.espec.skus > 0) resumo += "⚙️ <b>" + resumoGeral.espec.skus + "</b> SKUs tiveram especificações alteradas <span style='color: #666;'>(Tiragem impactada: " + _fmtNum(resumoGeral.espec.tiragem) + ")</span><br>";

  var assunto = "🚨 Relatório de Auditoria: " + alteracoes.length + " Alterações na Produção (SPG)";
  
  var corpoHtml = 
    "<div style='font-family: Arial, sans-serif; color: #333;'>" +
      "<h2>🚨 Relatório de Auditoria: Alterações na Produção (SPG)</h2>" +
      "<p style='font-size: 15px;'>" + textoComparativo + "</p>" +
      "<p>Foram encontradas <b>" + alteracoes.length + "</b> mudanças consolidadas na planilha.</p>" +
      
      "<h3 style='color: #444; border-bottom: 1px solid #ccc; padding-bottom: 4px;'>📊 Resumo Geral:</h3>" +
      "<p style='line-height: 1.6; font-size: 14px;'>" + resumo + "</p>" +
      
      "<h3 style='color: #444; border-bottom: 1px solid #ccc; padding-bottom: 4px; margin-top: 30px;'>🏭 Top Alterações (Volume de Tiragem) por Gráfica:</h3>" +
      _gerarTabelaHTML(resumoGrafica, "Gráfica") +

      "<h3 style='color: #444; border-bottom: 1px solid #ccc; padding-bottom: 4px; margin-top: 30px;'>🏷️ Top Alterações (Volume de Tiragem) por Marca:</h3>" +
      _gerarTabelaHTML(resumoMarca, "Marca") +

      "<br><hr style='border: 0; border-top: 1px solid #ccc; margin-top: 30px;'>" +
      "<p style='font-size: 15px;'>👉 <a href='" + urlPlanilha + "' style='color: #0056b3; text-decoration: none;'><b>Clique aqui para abrir a planilha com o Dê/Para e o Consolidado Atual Completo</b></a></p>" +
    "</div>";

  MailApp.sendEmail({
    to: REPORT_CONFIG.slackEmail,
    subject: assunto,
    htmlBody: corpoHtml
  });
}


// ================== FUNÇÕES UTILITÁRIAS BLINDADAS ==================

function _sanitizeVal(val) {
  if (val === null || val === undefined || (typeof val === 'number' && isNaN(val))) return 'Em branco';
  var str = String(val).replace(/[\x00-\x09\x0B-\x0C\x0E-\x1F\x7F-\x9F]/g, '');
  
  if (str.trim() === '') return 'Em branco';
  
  if (str.length > 40000) str = str.substring(0, 40000);
  if (str.charAt(0) === '+' || str.charAt(0) === '=' || str.charAt(0) === '-' || str.charAt(0) === '@') {
    return "'" + str;
  }
  return str;
}

function _obterValorSeguro(obj, possiveisChaves) {
  if (!obj || typeof obj !== 'object') return '';
  for (var i = 0; i < possiveisChaves.length; i++) {
    var k = possiveisChaves[i];
    if (obj[k] !== undefined && obj[k] !== null && String(obj[k]).trim() !== '') {
      return String(obj[k]).trim(); 
    }
  }
  return '';
}

function _lerAbaObjetos(ss, nomeAba) {
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) return [];
  var dados = aba.getDataRange().getValues();
  if (dados.length < 2) return []; 
  
  var cabecalho = dados[0];
  var objetos = [];
  
  for (var i = 1; i < dados.length; i++) {
    var linha = dados[i];
    var obj = {};
    for (var c = 0; c < cabecalho.length; c++) {
      var key = String(cabecalho[c]).trim();
      if (key !== '') obj[key] = linha[c];
    }
    objetos.push(obj);
  }
  return objetos;
}

function _gravarAbaExata(ss, nomeAba, dadosMatriz) {
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) aba = ss.insertSheet(nomeAba);
  else aba.clear(); 
  
  if (dadosMatriz.length > 0) {
    var totalLinhas = dadosMatriz.length;
    var totalCols = dadosMatriz[0].length;
    if (aba.getMaxRows() < totalLinhas) aba.insertRowsAfter(aba.getMaxRows(), totalLinhas - aba.getMaxRows());
    if (aba.getMaxColumns() < totalCols) aba.insertColumnsAfter(aba.getMaxColumns(), totalCols - aba.getMaxColumns());

    SpreadsheetApp.flush(); 

    var chunkSize = 2000;
    for (var k = 0; k < totalLinhas; k += chunkSize) {
      var chunk = dadosMatriz.slice(k, k + chunkSize);
      aba.getRange(k + 1, 1, chunk.length, totalCols).setValues(chunk);
    }
    
    aba.getRange(1, 1, 1, totalCols).setFontWeight("bold");
  }
  SpreadsheetApp.flush();
}