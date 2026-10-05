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
  Logger.log('2. Baixando e processando CONSOLIDADO (Cérebro de Gráfica PCP, Marca e Tecnologia)...');
  var rawConsolidado = _emularDadosDaPlanilha(_fetchApiData(SPG_CONFIG.urlConsolidado, token, false));
  var motorConsolidado = _mapearConsolidadoECriarMotor(rawConsolidado, dicEspec);
  _gravarAba(ss, SPG_CONFIG.abaConsolidado, motorConsolidado.matriz);
  rawConsolidado = null;

  // 3. TIRAGEM primeiro, pois as referências da ESPEC passam a usar esta aba como fonte
  Logger.log('3. Baixando e processando TIRAGEM...');
  var rawTiragem = _emularDadosDaPlanilha(_fetchApiData(SPG_CONFIG.urlTiragem, token, false));
  var matrizTiragem = _mapearTiragem(rawTiragem, dicEspec, CABECALHO_TIRAGEM, motorConsolidado);
  _gravarAba(ss, SPG_CONFIG.abaTiragem, matrizTiragem);
  var referenciasTiragem = _criarMapaReferenciasTiragem_(matrizTiragem);
  rawTiragem = null;

  // 4. ESPECIFICAÇÕES (Mapeamento com clonagem por Gráfica + Marca)
  Logger.log('4. Mapeando ESPECIFICAÇÕES por Gráfica PCP + Marca...');
  var matrizEspec = _mapearEspec(rawEspec, CABECALHO_ESPEC, motorConsolidado, referenciasTiragem);
  _gravarAba(ss, SPG_CONFIG.abaEspec, matrizEspec);
  rawEspec = null;

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
  
  // Os dicionários abaixo guardam VARIANTES, sem sobrescrever marca/gráfica.
  var dicSKU = {};
  var dicKitSKU = {};
  var dicKit = {};
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
    
    var sku = _obterValorSeguro(registro, ['sku', 'SKU', 'skuAtual', 'itemCode']);
    var kit = _obterValorSeguro(registro, ['kitCode', 'kit', 'Kit']);
    var marca = _obterValorSeguro(registro, ['marca', 'Marca', 'brand']);
    var tech = _obterValorSeguro(registro, ['tecnologia', 'Tecnologia', 'tech']);
    var edicaoRef = _obterValorSeguro(registro, ['edicaoRef', 'EdicaoRef', 'ediçãoRef']);

    // Bradesco tem precedência sobre a marca original, independente de maiúsculas/minúsculas.
    if (String(edicaoRef).toLowerCase().indexOf('bradesco') !== -1) {
      marca = 'SAS-BRADESCO';
      registro['marca'] = marca;
    }
    // SAS ADAPT já vem correto no Consolidado e é apenas preservado.

    var espec = dicEspec[sku] || {};
    var isG = String(espec['isGraphic']).trim().toLowerCase();
    var isE = String(espec['isEditorial']).trim().toLowerCase();
    
    var txtGrafico = (isG === 'true') ? 'Grafico' : (isG === 'false' ? 'Não Grafico' : '');
    var txtEditorial = (isE === 'true') ? 'Editorial' : (isE === 'false' ? 'Não Editorial' : '');
    
    registro['Grafico'] = txtGrafico;
    registro['Editorial'] = txtEditorial;

    var g1 = _obterValorSeguro(registro, ['grafica1', 'Gráfica 1', 'grafica1Atual', 'grafica_1']);
    var g1Normalizada = String(g1 || '').trim().toUpperCase();

    // Regra única da Grafica PCP: só aceita a whitelist. Todo o resto vira "-".
    var graficaPCP = graficasPermitidas.indexOf(g1Normalizada) !== -1 ? g1Normalizada : '-';
    registro['Grafica PCP'] = graficaPCP;

    var variante = { grafica: graficaPCP, tech: tech, marca: marca };
    if (sku) _registrarVariante_(dicSKU, sku, variante);
    if (kit && sku) _registrarVariante_(dicKitSKU, kit + '_' + sku, variante);
    if (kit) _registrarVariante_(dicKit, kit, variante);

    if (kit && graficaPCP !== '-' && graficaPCP !== '') {
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
    dicKit: dicKit,
    modeGrafKit: modeGrafKit,
    modeTechGraf: modeTechGraf,
    modeTechKit: modeTechKit
  };
}

function _mapearEspec(rawEspec, cabecalhos, motor, referenciasTiragem) {
  var linhas = [];
  var idxMarca = cabecalhos.indexOf('2. INF_Marca');
  var idxGrafica = cabecalhos.indexOf('GRAFICA');
  var idxTech = cabecalhos.indexOf('Tecnologia');
  var dividirPesoPor100000 = _pesoEspecVemEmEscalaGrande_(rawEspec);

  for (var i = 0; i < rawEspec.length; i++) {
    var r = rawEspec[i];
    var sku = String(r['skuAtual']).trim();
    
    var isG = String(r['isGraphic']).trim().toLowerCase();
    var isE = String(r['isEditorial']).trim().toLowerCase();
    var txtGrafico = (isG === 'true') ? 'Grafico' : (isG === 'false' ? 'Não Grafico' : '');
    var txtEditorial = (isE === 'true') ? 'Editorial' : (isE === 'false' ? 'Não Editorial' : '');

    var bAberto = _arredondarParaBaixoInteiro_(r['capBaseMmAberto']);
    var aAberto = _arredondarParaBaixoInteiro_(r['capAlturaMmAberto']);
    var bFechado = _arredondarParaBaixoInteiro_(r['baseMm']);
    var aFechado = _arredondarParaBaixoInteiro_(r['alturaMm']);
    var espessura = _arredondarParaBaixoUmaCasa_(r['espessuraMm']);
    var peso = _tratarPesoEspec_(r['pesoKg'], dividirPesoPor100000);

    var dimBxAAberto = (bAberto !== '' && aAberto !== '') ? bAberto + 'X' + aAberto : '';
    var dimBxAFechado = (bFechado !== '' && aFechado !== '') ? bFechado + 'X' + aFechado : '';
    var dimAcabada = (dimBxAFechado && espessura !== '') ? dimBxAFechado + 'X' + _formatarUmaCasaTexto_(espessura) : '';

    var refs = referenciasTiragem[sku] || {};

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
      else if (col === '3, DIM_B (Aberto)(mm)' || col === '3. DIM_B (Aberto)(mm)') val = bAberto;
      else if (col === '3, DIM_A (Aberto)(mm)' || col === '3. DIM_A (Aberto)(mm)') val = aAberto;
      else if (col === '3,DIM_BxA(Aberto)(mm)' || col === '3. DIM_BxA (Aberto)(mm)') val = dimBxAAberto;
      else if (col === '3, DIM_B (Fechado)(mm)' || col === '3. DIM_B (Fechado)(mm)') val = bFechado;
      else if (col === '3, DIM_A (Fechado)(mm)' || col === '3. DIM_A (Fechado)(mm)') val = aFechado;
      else if (col === '3. DIM_E (mm)') val = espessura;
      else if (col === '3.DIM_BXA(Fechado)' || col === '3. DIM_BxA (Fechado)') val = dimBxAFechado;
      else if (col === '3.DIM_Dimensãoacabada(BxAxE)' || col === '3. DIM_Dimensão acabada (BxAxE)') val = dimAcabada;
      else if (col === '3. DIM_Peso (Kg)') val = peso;
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
      else if (col === '11. REF_Referência troca de chapa') val = refs.trocaChapa || '';
      else if (col === '11. REF_Tipo de personalização') val = tipoPersonalizacao;
      else if (col === '11. REF_Quantidade de personalizados') val = '0';
      else if (col === '11. REF_Referência personalização') val = refs.personalizacao || '';
      else if (col === 'Grafico') val = txtGrafico;
      else if (col === 'Editorial') val = txtEditorial;
      
      linhaBase.push(val === undefined || val === null ? '' : val);
    }

    var variantes = motor.dicSKU[sku] || [];
    
    if (variantes.length === 0) {
      var linhaClone = linhaBase.slice();
      if (idxGrafica !== -1) linhaClone[idxGrafica] = '';
      if (idxTech !== -1) linhaClone[idxTech] = '';
      linhas.push(linhaClone);
    } else {
      for (var v = 0; v < variantes.length; v++) {
        var variante = variantes[v];
        var linhaClone = linhaBase.slice();
        var g = variante.grafica || '';
        var tech = variante.tech || motor.modeTechGraf[g] || '';
        if (idxMarca !== -1) linhaClone[idxMarca] = variante.marca || r['marca'] || '';
        if (idxGrafica !== -1) linhaClone[idxGrafica] = g;
        if (idxTech !== -1) linhaClone[idxTech] = tech;
        linhas.push(linhaClone);
      }
    }
  }
  
  linhas.sort(function(a, b) {
    var gA = idxGrafica !== -1 ? (a[idxGrafica] || '') : '';
    var gB = idxGrafica !== -1 ? (b[idxGrafica] || '') : '';
    var cmp = String(gA).localeCompare(String(gB));
    if (cmp !== 0) return cmp;
    var mA = idxMarca !== -1 ? (a[idxMarca] || '') : '';
    var mB = idxMarca !== -1 ? (b[idxMarca] || '') : '';
    return String(mA).localeCompare(String(mB));
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
    var variantes = (motor.dicKitSKU[ks] || []).slice();
    var graficaRaw = _obterValorSeguro(r, ['grafica1Atual', 'grafica1', 'GRAFICA', 'grafica']);
    
    if (graficaRaw && variantes.length > 1) {
      var graficaNormalizada = String(graficaRaw).trim().toUpperCase();
      var filtradas = variantes.filter(function(v) {
        return String(v.grafica || '').trim().toUpperCase() === graficaNormalizada;
      });
      if (filtradas.length > 0) variantes = filtradas;
    }

    if (variantes.length === 0) {
      variantes = [{
        grafica: motor.modeGrafKit[kit] || '',
        tech: motor.modeTechKit[kit] || '',
        marca: r['marca'] || espec['marca'] || ''
      }];
    }
    
    var isG = String(espec['isGraphic']).trim().toLowerCase();
    var isE = String(espec['isEditorial']).trim().toLowerCase();
    var txtGrafico = (isG === 'true') ? 'Grafico' : (isG === 'false' ? 'Não Grafico' : '');
    var txtEditorial = (isE === 'true') ? 'Editorial' : (isE === 'false' ? 'Não Editorial' : '');

    for (var v = 0; v < variantes.length; v++) {
      var variante = variantes[v] || {};
      var graficaFinal = variante.grafica || motor.modeGrafKit[kit] || '';
      var techFinal = variante.tech || motor.modeTechKit[kit] || '';
      var marcaFinal = variante.marca || r['marca'] || espec['marca'] || '';

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
    
    var marcaNomeOriginal = brand['name'] || '';
    var marcaGrupo = brand['brandGroup'] || '';
    var segmentoGrupo = segment['segmentGroup'] || '';

    // A marca da Árvore vem do Consolidado por kitCode. Se houver SAS e uma marca especial
    // no mesmo kit, os dois blocos são mantidos em vez de um sobrescrever o outro.
    var marcasKit = _marcasUnicasDoKit_(motor.dicKit[kit] || []);
    if (marcasKit.length === 0) marcasKit = [marcaNomeOriginal];

    for (var mk = 0; mk < marcasKit.length; mk++) {
      var marcaKit = marcasKit[mk] || marcaNomeOriginal;
      var variantePai = _buscarVariantePorMarca_(motor.dicKit[kit] || [], marcaKit) || {};
      var graficaPai = variantePai.grafica || motor.modeGrafKit[kit] || '';
      var techPai = variantePai.tech || motor.modeTechKit[kit] || '';

      var linhaPai = [];
      for (var c = 0; c < cabecalhos.length; c++) {
        var col = String(cabecalhos[c]).trim();
        var val = '';

        if (col === '1. ID_Código Kit') val = kit;
        else if (col === '1. ID_Descrição Kit') val = r['description'];
        else if (col === '1. ID_Código somente KIT') val = kit;
        else if (col === '2. INF_Marca') val = marcaKit;
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
        var variantesFilho = motor.dicKitSKU[ks] || [];
        var varianteFilho = _buscarVariantePorMarca_(variantesFilho, marcaKit) || variantesFilho[0] || {};
        
        var graficaFilho = varianteFilho.grafica || motor.modeGrafKit[kit] || '';
        var techFilho = varianteFilho.tech || motor.modeTechKit[kit] || '';
        var marcaFilho = varianteFilho.marca || marcaKit || marcaNomeOriginal;
        
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
          else if (col === '2. INF_Marca') val = marcaFilho;
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
  }
  return linhas;
}

// ================== FUNÇÕES AUXILIARES BLINDADAS ==================

function _registrarVariante_(mapa, chave, variante) {
  if (!chave) return;
  if (!mapa[chave]) mapa[chave] = [];
  var lista = mapa[chave];
  var grafica = String(variante.grafica || '').trim();
  var marca = String(variante.marca || '').trim();

  for (var i = 0; i < lista.length; i++) {
    if (String(lista[i].grafica || '').trim() === grafica &&
        String(lista[i].marca || '').trim() === marca) {
      if (!lista[i].tech && variante.tech) lista[i].tech = variante.tech;
      return;
    }
  }
  lista.push({
    grafica: variante.grafica || '',
    tech: variante.tech || '',
    marca: variante.marca || ''
  });
}

function _buscarVariantePorMarca_(variantes, marca) {
  var alvo = String(marca || '').trim().toUpperCase();
  for (var i = 0; i < variantes.length; i++) {
    if (String(variantes[i].marca || '').trim().toUpperCase() === alvo) return variantes[i];
  }
  return null;
}

function _marcasUnicasDoKit_(variantes) {
  var saida = [];
  var vistos = {};
  for (var i = 0; i < variantes.length; i++) {
    var marca = String(variantes[i].marca || '').trim();
    if (!marca) continue;
    var chave = marca.toUpperCase();
    if (!vistos[chave]) {
      vistos[chave] = true;
      saida.push(marca);
    }
  }
  return saida;
}

function _criarMapaReferenciasTiragem_(matrizTiragem) {
  var mapa = {};
  if (!matrizTiragem || matrizTiragem.length < 2) return mapa;

  var cab = matrizTiragem[0];
  var idxSku = cab.indexOf('1. ID_Código SKU');
  var idxTroca = cab.indexOf('11. REF_Referência troca de chapa');
  var idxPers = cab.indexOf('11. REF_Referência personalização');
  if (idxSku === -1) return mapa;

  for (var i = 1; i < matrizTiragem.length; i++) {
    var linha = matrizTiragem[i];
    var sku = String(linha[idxSku] || '').trim();
    if (!sku) continue;
    if (!mapa[sku]) mapa[sku] = { trocaChapa: '', personalizacao: '' };

    var troca = idxTroca !== -1 ? linha[idxTroca] : '';
    var pers = idxPers !== -1 ? linha[idxPers] : '';

    if (!mapa[sku].trocaChapa && troca !== undefined && troca !== null && String(troca).trim() !== '') {
      mapa[sku].trocaChapa = troca;
    }
    if (!mapa[sku].personalizacao && pers !== undefined && pers !== null && String(pers).trim() !== '') {
      mapa[sku].personalizacao = pers;
    }
  }
  return mapa;
}

function _numeroSPG_(valor) {
  if (valor === undefined || valor === null || String(valor).trim() === '') return null;
  if (typeof valor === 'number') return isNaN(valor) ? null : valor;

  var txt = String(valor).trim().replace(/\s/g, '');
  if (txt.indexOf(',') !== -1 && txt.indexOf('.') !== -1) {
    if (txt.lastIndexOf(',') > txt.lastIndexOf('.')) txt = txt.replace(/\./g, '').replace(',', '.');
    else txt = txt.replace(/,/g, '');
  } else if (txt.indexOf(',') !== -1) {
    txt = txt.replace(',', '.');
  }
  var n = Number(txt);
  return isNaN(n) ? null : n;
}

function _arredondarParaBaixoInteiro_(valor) {
  var n = _numeroSPG_(valor);
  return n === null ? '' : Math.floor(n);
}

function _arredondarParaBaixoUmaCasa_(valor) {
  var n = _numeroSPG_(valor);
  return n === null ? '' : Math.floor(n * 10) / 10;
}

function _formatarUmaCasaTexto_(valor) {
  var n = _numeroSPG_(valor);
  if (n === null) return '';
  return n.toFixed(1).replace('.', ',');
}

function _pesoEspecVemEmEscalaGrande_(rawEspec) {
  for (var i = 0; i < rawEspec.length; i++) {
    var valor = rawEspec[i] ? rawEspec[i]['pesoKg'] : '';
    var n = _numeroSPG_(valor);
    if (n === null) continue;

    // O primeiro peso preenchido decide o padrão do lote inteiro.
    // Valores já em kg normalmente vêm como 0,x ou 1,x; valores na escala antiga vêm em milhares.
    return Math.abs(n) >= 1000;
  }
  return false;
}

function _tratarPesoEspec_(valor, dividirPor100000) {
  var n = _numeroSPG_(valor);
  if (n === null) return valor === undefined || valor === null ? '' : valor;
  return dividirPor100000 ? (n / 100000) : n;
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

    // Garante visualmente 1 casa decimal na espessura da aba Espec.
    if (nomeAba === SPG_CONFIG.abaEspec && totalLinhas > 1) {
      var idxEsp = dadosMatriz[0].indexOf('3. DIM_E (mm)');
      if (idxEsp !== -1) aba.getRange(2, idxEsp + 1, totalLinhas - 1, 1).setNumberFormat('0.0');
    }
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