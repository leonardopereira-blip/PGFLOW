// ================== CONFIGURAÇÕES GERAIS ==================
var SPG_CONFIG = {
  idPlanilha: '1aRUIzE1AA_qFGjZO9RcD51fggfok7-lTpf1IpRCCQ-w', // ID EXATO DA PLANILHA
  urlLogin: 'https://spg-api.prod.arcocv.co/api/auth/login',
  
  urlTiragem: 'https://spg-api.prod.arcocv.co/api/tiragem/relatorio-alocacao',
  urlEspec: 'https://spg-api.prod.arcocv.co/api/products/spec-report',
  urlArvore: 'https://spg-api.prod.arcocv.co/api/bom/arvore-de-produtos',
  urlConsolidado: 'https://spg-api.prod.arcocv.co/api/tiragem/relatorio-consolidado',
  
  abaTiragem: 'Tiragem',
  abaEspec: 'Espec',
  abaArvore: 'Arvore',
  abaConsolidado: 'ConsolidadoProducao'
};

function _spgCredenciais_() {
  var props = PropertiesService.getScriptProperties();
  return {
    email: props.getProperty('SPG_EMAIL') || 'custos@spg.arcoeducacao.com.br',
    password: props.getProperty('SPG_PASSWORD') || 'DyyWUXzC0iRBIGVEs8qA0KrlX2KZyCNY'
  };
}

// ================== CABEÇALHOS HARDCODED ==================

var CABECALHO_TIRAGEM = [
  '1. ID_Código Kit', '1. ID_Descrição Kit', '1. ID_Código somente KIT', '1. ID_Código SKU', '1. ID_ISBN', 
  '1. ID_Descrição', '2. INF_Segmento', '2. INF_Série', '2. INF_Volume', '2. INF_Usuário', 
  '11. REF_Referência troca de chapa', '11. REF_Referência personalização', '14. EST_Posição de estoque', 
  '15. DES_CD envio', '16. TIR_SKU', '16. TIR_Item', '16. TIR_Miolo', '16. TIR_Capa avulsa', '16. TIR_Kit', 
  'MARCA', 'GRAFICA', 'Tecnologia', 'Grafico', 'Editorial'
];

var CABECALHO_ESPEC = [
  'Status SKU 2026', '1. ID_Código SKU', '1. ID_Descrição', '1. ID_ISBN', '1. ID_Código SKU anterior', 
  '2. INF_Marca', '2. INF_Grupo da marca', '2. INF_Segmento', '2. INF_Série', '2. INF_Volume', 
  '2. INF_Envio', '2. INF_Frequência', '2. INF_Usuário', '2. INF_Assunto/ Disciplina/ Área', 
  '2. INF_Tipo de material', '2. INF_Classificação do produto', '2. INF_Cliente personalizado', 
  '3, DIM_B (Aberto)(mm)', '3, DIM_A (Aberto)(mm)', '3,DIM_BxA(Aberto)(mm)', '3, DIM_B (Fechado)(mm)', 
  '3, DIM_A (Fechado)(mm)', '3. DIM_E (mm)', '3.DIM_BXA(Fechado)', '3.DIM_Dimensãoacabada(BxAxE)', 
  '3. DIM_Peso (Kg)', '3. DIM_Quantidade total de páginas', '4. CAP_Tipo de capa', '4.Cap_orelha', 
  '4. CAP_Papel da capa', '4. CAP_Cor capa', '4. CAP_Papel revestimento', '4. CAP_Papel forro', 
  '4. CAP_Acabamento da capa', '4. CAP_Acabamento interno capa', '4. CAP_Aproveitamento chapa capa (Cor)', 
  '4. CAP_Corte e vinco', '5. MIO_PRIN_Quantidade de páginas', '5. MIO_PRIN_Papel', '5. MIO_PRIN_Cor', 
  '5. MIO_PRIN_Aproveitamento chapa miolo (Cor)', '5. MIO_PRIN_Corte e vinco', '5. MIO_PRIN_Serrilha', 
  '5. MIO_PRIN_Obs miolo', '6. MIO_PROF_Quantidade de páginas', '6. MIO_PROF_Papel', '6. MIO_PROF_Cor', 
  '6. MIO_PROF_Aproveitamento chapa miolo (Cor)', '6. MIO_PROF_Corte e vinco', '6. MIO_PROF_Serrilha', 
  '6. MIO_PROF_Obs miolo', '7. MIO_ENCA_Quantidade de páginas', '7. MIO_ENCA_Papel', '7. MIO_ENCA_Cor', 
  '7. MIO_ENCA_Aproveitamento chapa encarte (Cor)', '7. MIO_ENCA_Corte e vinco', '7. MIO_ENCA_Serrilha', 
  '7. MIO_ENCA_Obs encarte', '8. MIO_ADES_Quantidade de páginas', '8. MIO_ADES_Papel', '8. MIO_ADES_Cor', 
  '8. MIO_ADES_Aproveitamento chapa adesivo (Cor)', '8. MIO_ADES_Corte e vinco', '8. MIO_ADES_Obs adesivo', 
  '9. ENC_Tipo acabamento', '9. ENC_Posição', '9.ENC_Bitola do espiral', '9. ENC_Cor do espiral', 
  '9. ENC_Obs de encadernação', '10. OBS_PROD_Observação para produção gráfica', '10. Orientação de montagem do livro', 
  '11. REF_Referência troca de chapa', '11. REF_Tipo de personalização', '11. REF_Quantidade de personalizados', 
  '11. REF_Referência personalização', '12. VER_Data verificação editorial', '12. VER_Responsável verificação editorial', 
  '12. VER_Data verificação engenharia', '12. VER_Responsável verificação engenharia', 'GRAFICA', 'Tecnologia', 'Grafico', 'Editorial'
];

var CABECALHO_ARVORE = [
  '1. ID_Código Kit', '1. ID_Descrição Kit', '1. ID_Código somente KIT', '1. ID_Código SKU', '1. ID_ISBN', 
  '1. ID_Descrição', '2. INF_Marca', '2. INF_Grupo da marca', '2. INF_Segmento', '2. INF_Série', 
  '2. INF_Volume', '2. INF_Envio', '2. INF_Frequência', '2. INF_Usuário', '2. INF_Assunto/ Disciplina/ Área', 
  '2. INF_Tipo de material', '2. INF_Classificação do produto', '2. INF_Cliente personalizado', '3. DIM_E (mm)', 
  '3. DIM_Peso (Kg)', '13. MAN_Total de itens colecionados', '13. MAN_Embalagem', '13. MAN_Observação de embalagem', 
  '13. MAN_Espessura do kit (mm)', '13. MAN_Peso do kit (kg)', '13.MAN_Código caixa', '13.MAN_Dimensõescaixaparda', 
  '13. MAN_Quantidade de kits por caixa parda', 'GRAFICA', 'Tecnologia', 'Grafico', 'Editorial'
];

// ================== FUNÇÃO EXECUTÁVEL PRINCIPAL ==================

function gerarRelatoriosCompletos() {
  var inicio = new Date().getTime();
  var ss = SpreadsheetApp.openById(SPG_CONFIG.idPlanilha);
  
  var token = _spgLogin_();
  if (!token) {
    Logger.log('ERRO: Falha ao autenticar.');
    return;
  }

  // 1. ESPECIFICAÇÕES (Dicionário Base)
  Logger.log('1. Baixando ESPECIFICAÇÕES...');
  var rawEspec = _emularDadosDaPlanilha(_fetchApiData(SPG_CONFIG.urlEspec, token, true));
  var dicEspec = {};
  for (var i = 0; i < rawEspec.length; i++) {
    var sku = String(rawEspec[i]['skuAtual']).trim();
    if (sku) dicEspec[sku] = rawEspec[i];
  }

  // 2. CONSOLIDADO (Geração do Motor de Cruzamento e Cálculo de Tecnologias)
  Logger.log('2. Baixando e processando CONSOLIDADO (Cérebro de Gráfica PCP e Tecnologia)...');
  var rawConsolidado = _emularDadosDaPlanilha(_fetchApiData(SPG_CONFIG.urlConsolidado, token, false));
  var motorConsolidado = _mapearConsolidadoECriarMotor(rawConsolidado, dicEspec);
  _gravarAba(ss, SPG_CONFIG.abaConsolidado, motorConsolidado.matriz);
  rawConsolidado = null; 

  // 3. ESPECIFICAÇÕES (Mapeamento com Clonagem por Gráfica)
  Logger.log('3. Mapeando ESPECIFICAÇÕES em Blocos por Gráfica PCP...');
  var matrizEspec = _mapearEspec(rawEspec, CABECALHO_ESPEC, motorConsolidado);
  _gravarAba(ss, SPG_CONFIG.abaEspec, matrizEspec);
  rawEspec = null; 

  // 4. TIRAGEM (Baixa, cruza com o Dicionário)
  Logger.log('4. Baixando e processando TIRAGEM...');
  var rawTiragem = _emularDadosDaPlanilha(_fetchApiData(SPG_CONFIG.urlTiragem, token, false));
  var matrizTiragem = _mapearTiragem(rawTiragem, dicEspec, CABECALHO_TIRAGEM, motorConsolidado);
  _gravarAba(ss, SPG_CONFIG.abaTiragem, matrizTiragem);
  rawTiragem = null;

  // 5. ÁRVORE (Baixa, desaninha e aplica lógicas de Pai/Filho)
  Logger.log('5. Baixando e processando ÁRVORE...');
  var rawArvore = _emularDadosDaPlanilha(_fetchApiData(SPG_CONFIG.urlArvore, token, true));
  var matrizArvore = _mapearArvore(rawArvore, CABECALHO_ARVORE, dicEspec, motorConsolidado);
  _gravarAba(ss, SPG_CONFIG.abaArvore, matrizArvore);
  rawArvore = null;

  var segundos = Math.round((new Date().getTime() - inicio) / 1000);
  Logger.log('✅ SUCESSO! Abas cruzadas perfeitamente em ' + segundos + 's.');
}

// ================== MOTORES DE MAPEAMENTO ==================

function _fetchApiData(urlBase, token, usaPaginacao) {
  var todosOsDados = [];
  var pagina = 1;
  var limite = 500;
  var assinaturaUltimaPagina = null;

  while (true) {
    var urlAtual = urlBase;
    if (usaPaginacao) {
      urlAtual += (urlAtual.indexOf('?') !== -1 ? '&' : '?') + 'page=' + pagina + '&limit=' + limite;
    }

    var resposta = UrlFetchApp.fetch(urlAtual, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (resposta.getResponseCode() !== 200) break;
    var json = JSON.parse(resposta.getContentText());
    var dadosDaPagina = _spgExtrairLista_(json);

    if (!dadosDaPagina || dadosDaPagina.length === 0) break;
    var assinaturaAtual = JSON.stringify(dadosDaPagina[0]);
    if (assinaturaUltimaPagina === assinaturaAtual) break;
    assinaturaUltimaPagina = assinaturaAtual;

    todosOsDados = todosOsDados.concat(dadosDaPagina);
    if (!usaPaginacao || dadosDaPagina.length < limite) break;
    pagina++;
  }
  return todosOsDados;
}

function _mapearConsolidadoECriarMotor(rawConsolidado, dicEspec) {
  var cabecalhos = [];
  var indiceColuna = {};
  var matriz = [];
  
  var dicSKU = {};       
  var dicKitSKU = {};    
  var freqTechGraf = {}; 
  var freqTechKit = {};  
  var freqGrafKit = {};

  var classificacoesBloqueadas = ['CONGELADOS', 'INCLUSIVOS', 'PREVIA', 'MONTAGEM-ATHOS-1A-JANELA-MODULARES', 'REENTRADA'];
  var graficasPermitidas = [
    'COAN', 'MIDIOGRAF', 'MAXIGRAFICA', 'RICARGRAF', 'RONA', 'REPROSET', 'IPSIS', 'PIFFERPRINT', 
    'BERCROM', 'STAR7', 'LOGPRINT', 'WALPRINT', 'POSIGRAF', 'META', 'IMOS', 'LEOGRAF', 
    'SAO FRANCISCO', 'OCEANO', 'FORMA CERTA', 'ZIT', 'MARGRAF'
  ];

  for (var i = 0; i < rawConsolidado.length; i++) {
    var registro = rawConsolidado[i];
    
    var classificacao = String(registro['classificacao'] || '').trim().toUpperCase();
    var envio = String(registro['envio'] || '').trim().toUpperCase();
    
    if (classificacoesBloqueadas.indexOf(classificacao) !== -1) continue; 
    if (envio.indexOf('1') === -1) continue;
    
    // Captura segura de Kit e SKU (considerando variações da API)
    var sku = _obterValorSeguro(registro, ['sku', 'SKU', 'skuAtual', 'itemCode']);
    var kit = _obterValorSeguro(registro, ['kitCode', 'kit', 'Kit']);
    var marca = _obterValorSeguro(registro, ['marca', 'Marca', 'brand']);
    var tech = _obterValorSeguro(registro, ['tecnologia', 'Tecnologia', 'tech']);
    
    var espec = dicEspec[sku] || {};
    var isG = String(espec['isGraphic']).trim().toLowerCase();
    var isE = String(espec['isEditorial']).trim().toLowerCase();
    
    var txtGrafico = (isG === 'true') ? 'Grafico' : (isG === 'false' ? 'Não Grafico' : '');
    var txtEditorial = (isE === 'true') ? 'Editorial' : (isE === 'false' ? 'Não Editorial' : '');
    
    registro['Grafico'] = txtGrafico;
    registro['Editorial'] = txtEditorial;

    var g1 = _obterValorSeguro(registro, ['grafica1', 'Gráfica 1', 'grafica1Atual', 'grafica_1']);
    var g2 = _obterValorSeguro(registro, ['grafica2', 'Gráfica 2', 'grafica2Atual', 'grafica_2']);
    
    var graficaPCP = '';

    if (txtGrafico === 'Não Grafico') {
      graficaPCP = '-';
    } else if (g1 === '') {
      graficaPCP = 'NAO_ALOCADO';
    } else {
      if (graficasPermitidas.indexOf(g1.toUpperCase()) === -1) {
        graficaPCP = (g2 !== '') ? g2 : 'NAO_ALOCADO';
      } else {
        graficaPCP = g1;
      }
    }
    registro['Grafica PCP'] = graficaPCP;

    // Alimentando o Cérebro de Cruzamento
    if (sku) {
      if (!dicSKU[sku]) dicSKU[sku] = {};
      if (!dicSKU[sku][graficaPCP] || tech !== '') dicSKU[sku][graficaPCP] = tech;
    }
    
    if (kit && sku) dicKitSKU[kit + '_' + sku] = { grafica: graficaPCP, tech: tech, marca: marca };
    
    if (kit && graficaPCP !== '' && graficaPCP !== '-' && graficaPCP !== 'NAO_ALOCADO') {
      if (!freqGrafKit[kit]) freqGrafKit[kit] = {};
      freqGrafKit[kit][graficaPCP] = (freqGrafKit[kit][graficaPCP] || 0) + 1;
    }

    if (tech !== '') {
      if (!freqTechGraf[graficaPCP]) freqTechGraf[graficaPCP] = {};
      freqTechGraf[graficaPCP][tech] = (freqTechGraf[graficaPCP][tech] || 0) + 1;

      if (kit) {
        if (!freqTechKit[kit]) freqTechKit[kit] = {};
        freqTechKit[kit][tech] = (freqTechKit[kit][tech] || 0) + 1;
      }
    }

    var linha = [];
    for (var chave in registro) {
      if (!Object.prototype.hasOwnProperty.call(registro, chave)) continue;
      if (indiceColuna[chave] === undefined) {
        indiceColuna[chave] = cabecalhos.length;
        cabecalhos.push(chave);
      }
      var val = registro[chave];
      if (typeof val === 'object' && val !== null) val = JSON.stringify(val); 
      linha[indiceColuna[chave]] = val;
    }
    matriz.push(linha);
  }

  for (var m = 0; m < matriz.length; m++) {
    for (var c = 0; c < cabecalhos.length; c++) {
      if (matriz[m][c] === undefined) matriz[m][c] = '';
    }
  }
  matriz.unshift(cabecalhos); 

  // Resolvendo as "Modas" (Valores que mais se repetem)
  var modeTechGraf = {};
  for (var gp in freqTechGraf) modeTechGraf[gp] = _getMode(freqTechGraf[gp]);

  var modeTechKit = {};
  for (var kt in freqTechKit) modeTechKit[kt] = _getMode(freqTechKit[kt]);

  var modeGrafKit = {};
  for (var k in freqGrafKit) modeGrafKit[k] = _getMode(freqGrafKit[k]);

  return {
    matriz: matriz,
    dicSKU: dicSKU,
    dicKitSKU: dicKitSKU,
    modeGrafKit: modeGrafKit,
    modeTechGraf: modeTechGraf,
    modeTechKit: modeTechKit
  };
}

function _mapearEspec(rawEspec, cabecalhos, motor) {
  var linhas = [];
  var idxGrafica = cabecalhos.indexOf('GRAFICA');
  var idxTech = cabecalhos.indexOf('Tecnologia');
  var idxGrafico = cabecalhos.indexOf('Grafico');
  var idxEditorial = cabecalhos.indexOf('Editorial');

  for (var i = 0; i < rawEspec.length; i++) {
    var r = rawEspec[i];
    var sku = String(r['skuAtual']).trim();
    
    var isG = String(r['isGraphic']).trim().toLowerCase();
    var isE = String(r['isEditorial']).trim().toLowerCase();
    var txtGrafico = (isG === 'true') ? 'Grafico' : (isG === 'false' ? 'Não Grafico' : '');
    var txtEditorial = (isE === 'true') ? 'Editorial' : (isE === 'false' ? 'Não Editorial' : '');

    var dimBxAAberto = (r['capBaseMmAberto'] && r['capAlturaMmAberto']) ? r['capBaseMmAberto'] + 'x' + r['capAlturaMmAberto'] : '';
    var dimBxAFechado = (r['baseMm'] && r['alturaMm']) ? r['baseMm'] + 'x' + r['alturaMm'] : '';
    var dimAcabada = (dimBxAFechado && r['espessuraMm']) ? dimBxAFechado + 'x' + String(r['espessuraMm']).replace('.', ',') : '';

    var m1 = _getArrItem(r['miolos'], 0);
    var m2 = _getArrItem(r['miolos'], 1); 
    var enc = _getArrItem(r['encartes'], 0); 
    var ad = _getArrItem(r['adesivos'], 0);  
    
    var isPers = String(r['isPersonalizado']).toLowerCase() === 'true';
    var tipoPersonalizacao = (isPers && r['skuMaterialPadrao']) ? 'Capa personalizada' : 'Capa Padrão';

    var linhaBase = [];
    for (var c = 0; c < cabecalhos.length; c++) {
      var col = String(cabecalhos[c]).trim();
      var val = '';
      
      if (col === 'Status SKU 2026') val = r['status'];
      else if (col === '1. ID_Código SKU') val = sku;
      else if (col === '1. ID_Descrição') val = r['descricao'];
      else if (col === '1. ID_ISBN') val = r['isbn'];
      else if (col === '1. ID_Código SKU anterior') val = r['skuReferencia'];
      else if (col === '2. INF_Marca') val = r['marca'];
      else if (col === '2. INF_Grupo da marca') val = r['grupoMarca'];
      else if (col === '2. INF_Segmento') val = r['segmento'];
      else if (col === '2. INF_Série') val = r['serie'];
      else if (col === '2. INF_Volume') val = r['volume'];
      else if (col === '2. INF_Envio') val = r['envio'];
      else if (col === '2. INF_Frequência') val = r['frequencia'];
      else if (col === '2. INF_Usuário') val = r['uso'];
      else if (col === '2. INF_Assunto/ Disciplina/ Área') val = r['assunto'];
      else if (col === '2. INF_Tipo de material') val = r['tipo'];
      else if (col === '2. INF_Classificação do produto') val = r['classificacao'];
      else if (col === '2. INF_Cliente personalizado') val = isPers ? 'Sim' : 'Não';
      else if (col === '3, DIM_B (Aberto)(mm)' || col === '3. DIM_B (Aberto)(mm)') val = r['capBaseMmAberto'];
      else if (col === '3, DIM_A (Aberto)(mm)' || col === '3. DIM_A (Aberto)(mm)') val = r['capAlturaMmAberto'];
      else if (col === '3,DIM_BxA(Aberto)(mm)' || col === '3. DIM_BxA (Aberto)(mm)') val = dimBxAAberto;
      else if (col === '3, DIM_B (Fechado)(mm)' || col === '3. DIM_B (Fechado)(mm)') val = r['baseMm'];
      else if (col === '3, DIM_A (Fechado)(mm)' || col === '3. DIM_A (Fechado)(mm)') val = r['alturaMm'];
      else if (col === '3. DIM_E (mm)') val = r['espessuraMm'];
      else if (col === '3.DIM_BXA(Fechado)' || col === '3. DIM_BxA (Fechado)') val = dimBxAFechado;
      else if (col === '3.DIM_Dimensãoacabada(BxAxE)' || col === '3. DIM_Dimensão acabada (BxAxE)') val = dimAcabada;
      else if (col === '3. DIM_Peso (Kg)') val = r['pesoKg'];
      else if (col === '3. DIM_Quantidade total de páginas') val = r['pgTotal'];
      else if (col === '4. CAP_Tipo de capa') val = r['capTipoCapa'];
      else if (col === '4.Cap_orelha' || col === '4. CAP_Orelha') val = r['capLarguraOrelha'] || (String(r['capTemOrelhas']).toLowerCase() === 'true' ? 'Sim' : '');
      else if (col === '4. CAP_Papel da capa') val = r['capPapel'];
      else if (col === '4. CAP_Cor capa') val = r['capCor'];
      else if (col === '4. CAP_Papel revestimento') val = r['capRevestimentoPapel'];
      else if (col === '4. CAP_Papel forro') val = r['capForroPapel'];
      else if (col === '4. CAP_Acabamento da capa') val = r['capRevestimentoTipo'];
      else if (col === '4. CAP_Acabamento interno capa') val = r['capAcabamentoObs'] || r['capRevestimentoAplicacao'];
      else if (col === '4. CAP_Aproveitamento chapa capa (Cor)') val = r['capCorChapa'];
      else if (col === '4. CAP_Corte e vinco') val = _boolPT(r['capTemCorteVinco']);
      else if (col === '5. MIO_PRIN_Quantidade de páginas') val = m1['paginacao'];
      else if (col === '5. MIO_PRIN_Papel') val = m1['papel'];
      else if (col === '5. MIO_PRIN_Cor') val = m1['cor'];
      else if (col === '5. MIO_PRIN_Aproveitamento chapa miolo (Cor)') val = m1['corChapa'];
      else if (col === '5. MIO_PRIN_Corte e vinco') val = _boolPT(m1['corteVinco']);
      else if (col === '5. MIO_PRIN_Serrilha') val = _boolPT(m1['serrilha']);
      else if (col === '5. MIO_PRIN_Obs miolo') val = m1['observacoes'];
      else if (col === '6. MIO_PROF_Quantidade de páginas') val = m2['paginacao'];
      else if (col === '6. MIO_PROF_Papel') val = m2['papel'];
      else if (col === '6. MIO_PROF_Cor') val = m2['cor'];
      else if (col === '6. MIO_PROF_Aproveitamento chapa miolo (Cor)') val = m2['corChapa'];
      else if (col === '6. MIO_PROF_Corte e vinco') val = _boolPT(m2['corteVinco']);
      else if (col === '6. MIO_PROF_Serrilha') val = _boolPT(m2['serrilha']);
      else if (col === '6. MIO_PROF_Obs miolo') val = m2['observacoes'];
      else if (col === '7. MIO_ENCA_Quantidade de páginas') val = enc['paginacao'];
      else if (col === '7. MIO_ENCA_Papel') val = enc['papel'];
      else if (col === '7. MIO_ENCA_Cor') val = enc['cor'];
      else if (col === '7. MIO_ENCA_Aproveitamento chapa encarte (Cor)') val = enc['corChapa'];
      else if (col === '7. MIO_ENCA_Corte e vinco') val = _boolPT(enc['corteVinco']);
      else if (col === '7. MIO_ENCA_Serrilha') val = _boolPT(enc['serrilha']);
      else if (col === '7. MIO_ENCA_Obs encarte') val = enc['observacoes'];
      else if (col === '8. MIO_ADES_Quantidade de páginas') val = ad['paginacao'];
      else if (col === '8. MIO_ADES_Papel') val = ad['papel'];
      else if (col === '8. MIO_ADES_Cor') val = ad['cor'];
      else if (col === '8. MIO_ADES_Aproveitamento chapa adesivo (Cor)') val = ad['corChapa'];
      else if (col === '8. MIO_ADES_Corte e vinco') val = _boolPT(ad['corteVinco']);
      else if (col === '8. MIO_ADES_Obs adesivo') val = ad['observacoes'];
      else if (col === '9. ENC_Tipo acabamento') val = r['capAcabamento'];
      else if (col === '9. ENC_Posição') val = r['capPosicaoAcabamento'];
      else if (col === '9.ENC_Bitola do espiral' || col === '9. ENC_Bitola do espiral') val = r['encBitolaEspiral'];
      else if (col === '9. ENC_Cor do espiral') val = r['encCorEspiral'];
      else if (col === '9. ENC_Obs de encadernação') val = r['obsEncadernacao'];
      else if (col === '10. OBS_PROD_Observação para produção gráfica') val = r['obsNotas'];
      else if (col === '10. Orientação de montagem do livro') val = r['obsMontagem'];
      else if (col === '11. REF_Referência troca de chapa') val = r['capSkuReferenciaChapa'] || r['skuReferenciaChapa'];
      else if (col === '11. REF_Tipo de personalização') val = tipoPersonalizacao;
      else if (col === '11. REF_Quantidade de personalizados') val = '0';
      else if (col === '11. REF_Referência personalização') val = r['skuMaterialPadrao'];
      else if (col === 'Grafico') val = txtGrafico;
      else if (col === 'Editorial') val = txtEditorial;
      
      linhaBase.push(val === undefined || val === null ? '' : val);
    }

    var graficasDoSku = motor.dicSKU[sku];
    
    if (!graficasDoSku || Object.keys(graficasDoSku).length === 0) {
      var linhaClone = linhaBase.slice();
      if (idxGrafica !== -1) linhaClone[idxGrafica] = '';
      if (idxTech !== -1) linhaClone[idxTech] = '';
      linhas.push(linhaClone);
    } else {
      for (var g in graficasDoSku) {
        var linhaClone = linhaBase.slice();
        var tech = graficasDoSku[g] || motor.modeTechGraf[g] || '';
        if (idxGrafica !== -1) linhaClone[idxGrafica] = g;
        if (idxTech !== -1) linhaClone[idxTech] = tech;
        linhas.push(linhaClone);
      }
    }
  }
  
  linhas.sort(function(a, b) {
    var gA = a[idxGrafica] || '';
    var gB = b[idxGrafica] || '';
    return gA.localeCompare(gB);
  });
  
  linhas.unshift(cabecalhos);
  return linhas;
}

function _mapearTiragem(rawTiragem, dicEspec, cabecalhos, motor) {
  var linhas = [];
  linhas.push(cabecalhos);

  for (var i = 0; i < rawTiragem.length; i++) {
    var r = rawTiragem[i];
    var sku = _obterValorSeguro(r, ['sku', 'SKU']);
    var kit = _obterValorSeguro(r, ['kitCode', 'kit', 'Kit']);
    var ks = kit + '_' + sku;
    
    var espec = dicEspec[sku] || {}; 
    var dataKS = motor.dicKitSKU[ks] || {};
    
    var graficaFinal = dataKS.grafica || motor.modeGrafKit[kit] || '';
    var techFinal = dataKS.tech || motor.modeTechKit[kit] || '';
    var marcaFinal = dataKS.marca || r['marca'] || espec['marca'] || '';
    
    var isG = String(espec['isGraphic']).trim().toLowerCase();
    var isE = String(espec['isEditorial']).trim().toLowerCase();
    var txtGrafico = (isG === 'true') ? 'Grafico' : (isG === 'false' ? 'Não Grafico' : '');
    var txtEditorial = (isE === 'true') ? 'Editorial' : (isE === 'false' ? 'Não Editorial' : '');

    var linha = [];
    for (var c = 0; c < cabecalhos.length; c++) {
      var col = String(cabecalhos[c]).trim();
      var val = '';

      if (col === '1. ID_Código Kit') val = kit;
      else if (col === '1. ID_Descrição Kit') val = r['kitDescricao'];
      else if (col === '1. ID_Código somente KIT') val = kit;
      else if (col === '1. ID_Código SKU') val = sku;
      else if (col === '1. ID_ISBN') val = espec['isbn'];
      else if (col === '1. ID_Descrição') val = r['skuDescricao'];
      else if (col === '2. INF_Segmento') val = espec['segmento'];
      else if (col === '2. INF_Série') val = espec['serie'];
      else if (col === '2. INF_Volume') val = espec['volume'];
      else if (col === '2. INF_Envio') val = r['envio'] || espec['envio'];
      else if (col === '2. INF_Frequência') val = espec['frequencia'];
      else if (col === '2. INF_Usuário') val = espec['uso'];
      else if (col === '11. REF_Referência troca de chapa') val = espec['capSkuReferenciaChapa'] || espec['skuReferencia'];
      else if (col === '11. REF_Referência personalização') val = espec['skuMaterialPadrao'];
      else if (col === '14. EST_Posição de estoque') val = r['estoque'];
      else if (col === '15. DES_CD envio') val = r['cdDestino'];
      else if (col === '16. TIR_SKU' || col === '16. TIR_Item' || col === '16. TIR_Kit' || col === '16. TIR_Miolo' || col === '16. TIR_Capa avulsa') val = r['quantidade']; 
      else if (col === 'MARCA') val = marcaFinal;
      else if (col === 'GRAFICA') val = graficaFinal;
      else if (col === 'Tecnologia') val = techFinal;
      else if (col === 'Grafico') val = txtGrafico;
      else if (col === 'Editorial') val = txtEditorial;

      linha.push(val === undefined || val === null ? '' : val);
    }
    linhas.push(linha);
  }
  return linhas;
}

function _mapearArvore(rawArvore, cabecalhos, dicEspec, motor) {
  var linhas = [];
  linhas.push(cabecalhos);

  for (var i = 0; i < rawArvore.length; i++) {
    var r = rawArvore[i];
    var kit = _obterValorSeguro(r, ['kitCode', 'kit', 'Kit']);
    
    var itens = _parse(r['items'], []);
    var boxType = _parse(r['boxType'], {});
    var brand = _parse(r['brand'], {});
    var segment = _parse(r['segment'], {});
    var builder = _parse(r['builderData'], {});
    
    var metrics = builder['metrics'] || {};
    var embalagemDesc = boxType['description'] || '';
    var caixaCodigo = boxType['sku'] || '';
    var caixaDimensoes = (boxType['baseMm'] && boxType['alturaMm'] && boxType['profundidadeMm']) ? boxType['baseMm'] + 'x' + boxType['alturaMm'] + 'x' + boxType['profundidadeMm'] : '';
    
    var marcaNome = brand['name'] || '';
    var marcaGrupo = brand['brandGroup'] || '';
    var segmentoGrupo = segment['segmentGroup'] || '';

    var graficaPai = motor.modeGrafKit[kit] || '';
    var techPai = motor.modeTechKit[kit] || '';

    var linhaPai = [];
    for (var c = 0; c < cabecalhos.length; c++) {
      var col = String(cabecalhos[c]).trim();
      var val = '';

      if (col === '1. ID_Código Kit') val = kit;
      else if (col === '1. ID_Descrição Kit') val = r['description'];
      else if (col === '1. ID_Código somente KIT') val = kit;
      else if (col === '2. INF_Marca') val = marcaNome;
      else if (col === '2. INF_Grupo da marca') val = marcaGrupo;
      else if (col === '2. INF_Segmento') val = segmentoGrupo;
      else if (col === '2. INF_Série') val = r['serie'];
      else if (col === '2. INF_Volume') val = r['volume'];
      else if (col === '2. INF_Envio') val = r['envio'];
      else if (col === '2. INF_Frequência') val = r['frequencia'];
      else if (col === '2. INF_Usuário') val = r['usoMaterial'];
      else if (col === '13. MAN_Total de itens colecionados') val = itens.length;
      else if (col === '13. MAN_Embalagem') val = embalagemDesc;
      else if (col === '13. MAN_Observação de embalagem') val = r['instrucaoManuseio'];
      else if (col === '13.MAN_Código caixa' || col === '13. MAN_Código caixa') val = caixaCodigo;
      else if (col === '13.MAN_Dimensõescaixaparda' || col === '13. MAN_Dimensões caixa parda') val = caixaDimensoes;
      else if (col === '13. MAN_Espessura do kit (mm)') val = metrics['totalThickness'];
      else if (col === '13. MAN_Peso do kit (kg)') val = metrics['totalWeight'];
      else if (col === '13. MAN_Quantidade de kits por caixa parda') val = metrics['itemsPerBox'];
      else if (col === 'GRAFICA') val = graficaPai;
      else if (col === 'Tecnologia') val = techPai;
      
      linhaPai.push(val === undefined || val === null ? '' : val);
    }
    linhas.push(linhaPai);

    for (var j = 0; j < itens.length; j++) {
      var filho = itens[j];
      var prodFilho = filho['product'] || {}; 
      var skuFilho = _obterValorSeguro(filho, ['itemCode', 'skuAtual']) || _obterValorSeguro(prodFilho, ['skuAtual']);
      
      var ks = kit + '_' + skuFilho;
      var dataKS = motor.dicKitSKU[ks] || {};
      
      var graficaFilho = dataKS.grafica || motor.modeGrafKit[kit] || '';
      var techFilho = dataKS.tech || motor.modeTechKit[kit] || '';
      
      var especFilho = dicEspec[skuFilho] || {};
      var isG = String(especFilho['isGraphic']).trim().toLowerCase();
      var isE = String(especFilho['isEditorial']).trim().toLowerCase();
      var txtGrafico = (isG === 'true') ? 'Grafico' : (isG === 'false' ? 'Não Grafico' : '');
      var txtEditorial = (isE === 'true') ? 'Editorial' : (isE === 'false' ? 'Não Editorial' : '');

      var linhaFilho = [];
      for (var c = 0; c < cabecalhos.length; c++) {
        var col = String(cabecalhos[c]).trim();
        var val = '';

        if (col === '1. ID_Código Kit') val = kit;
        else if (col === '1. ID_Descrição Kit') val = r['description'];
        else if (col === '1. ID_Código somente KIT') val = kit;
        else if (col === '1. ID_Código SKU') val = skuFilho;
        else if (col === '1. ID_Descrição') val = prodFilho['descricao'] || filho['description'] || filho['descricao'];
        else if (col === '1. ID_ISBN') val = prodFilho['isbn'] || '';
        else if (col === '2. INF_Marca') val = marcaNome;
        else if (col === '2. INF_Grupo da marca') val = marcaGrupo;
        else if (col === '2. INF_Segmento') val = segmentoGrupo;
        else if (col === '2. INF_Série') val = r['serie'];
        else if (col === '2. INF_Envio') val = r['envio'];
        else if (col === '2. INF_Frequência') val = r['frequencia'];
        else if (col === '2. INF_Usuário') val = r['usoMaterial'];
        else if (col === '3. DIM_E (mm)') val = prodFilho['espessuraMm'] || filho['espessuraMm'];
        else if (col === '3. DIM_Peso (Kg)') val = prodFilho['pesoKg'] || filho['pesoKg'];
        else if (col === 'GRAFICA') val = graficaFilho;
        else if (col === 'Tecnologia') val = techFilho;
        else if (col === 'Grafico') val = txtGrafico;
        else if (col === 'Editorial') val = txtEditorial;

        linhaFilho.push(val === undefined || val === null ? '' : val);
      }
      linhas.push(linhaFilho);
    }
  }
  return linhas;
}

// ================== FUNÇÕES AUXILIARES BLINDADAS ==================

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

function _getMode(freqObj) {
  var maxKey = '';
  var maxCount = 0;
  for (var k in freqObj) {
    if (freqObj[k] > maxCount) {
      maxCount = freqObj[k];
      maxKey = k;
    }
  }
  return maxKey;
}

function _emularDadosDaPlanilha(dadosMatrizObj) {
  var emulado = [];
  for (var i = 0; i < dadosMatrizObj.length; i++) {
    var original = dadosMatrizObj[i];
    var linhaEmulada = {};
    for (var k in original) {
      if (!Object.prototype.hasOwnProperty.call(original, k)) continue;
      var val = original[k];
      if (val === null || val === undefined) {
        linhaEmulada[k] = '';
      } else if (typeof val === 'object') {
        linhaEmulada[k] = JSON.stringify(val);
      } else {
        linhaEmulada[k] = String(val);
      }
    }
    emulado.push(linhaEmulada);
  }
  return emulado;
}

function _spgLogin_() {
  var cred = _spgCredenciais_();
  var resposta = UrlFetchApp.fetch(SPG_CONFIG.urlLogin, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ email: cred.email, password: cred.password }),
    muteHttpExceptions: true
  });
  if (resposta.getResponseCode() !== 200 && resposta.getResponseCode() !== 201) return null;
  return JSON.parse(resposta.getContentText()).token || null;
}

function _spgExtrairLista_(json) {
  if (Array.isArray(json)) return json;
  if (json && typeof json === 'object') {
    var chaves = ['products', 'kits', 'rows', 'versoes', 'items'];
    for (var i = 0; i < chaves.length; i++) {
      if (Array.isArray(json[chaves[i]])) return json[chaves[i]];
    }
    for (var chave in json) {
      if (Array.isArray(json[chave])) return json[chave];
    }
  }
  return [];
}

function _gravarAba(ss, nomeAba, dadosMatriz) {
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) aba = ss.insertSheet(nomeAba);
  else aba.clear(); 
  
  if (dadosMatriz.length > 0) {
    var totalLinhas = dadosMatriz.length;
    var totalCols = dadosMatriz[0].length;
    
    if (aba.getMaxRows() < totalLinhas) aba.insertRowsAfter(aba.getMaxRows(), totalLinhas - aba.getMaxRows());
    if (aba.getMaxColumns() < totalCols) aba.insertColumnsAfter(aba.getMaxColumns(), totalCols - aba.getMaxColumns());

    aba.getRange(1, 1, totalLinhas, totalCols).setValues(dadosMatriz);
    aba.getRange(1, 1, 1, totalCols).setFontWeight("bold");
  }
  SpreadsheetApp.flush();
}

function _boolPT(val) {
  if (val === 'true') return 'Sim';
  if (val === 'false') return 'Não';
  return val;
}

function _getArrItem(data, index) {
  if (!data) return {};
  var arr = data;
  if (typeof data === 'string') {
    try { arr = JSON.parse(data); } catch(e) { return {}; }
  }
  if (Array.isArray(arr) && arr.length > index) return arr[index] || {};
  return {};
}

function _parse(data, fallback) {
  if (!data) return fallback;
  if (typeof data !== 'string') return data;
  try { return JSON.parse(data) || fallback; } catch(e) { return fallback; }
}