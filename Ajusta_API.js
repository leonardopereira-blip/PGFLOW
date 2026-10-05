// ================== CONFIGURAÇÕES GERAIS ==================
var CONFIG_LAYOUT = {
  idPlanilha: '1aRUIzE1AA_qFGjZO9RcD51fggfok7-lTpf1IpRCCQ-w', // ID EXATO DA PLANILHA
  
  // Abas de Dados da API
  abaTiragemRaw: 'API_Tiragem',
  abaEspecRaw: 'API_ESPEC',
  abaArvoreRaw: 'API_ARVORE',
  
  // Abas que ditam o Layout Original (Templates)
  abaTiragemTemplate: 'Tiragem',
  abaEspecTemplate: 'Espec',
  abaArvoreTemplate: 'Arvore',
  
  // Abas Finais Ajustadas
  abaTiragemOut: 'Tiragem_API_AJUSTADA',
  abaEspecOut: 'Espec_API_AJUSTADA',
  abaArvoreOut: 'Arvore_API_AJUSTADA'
};

// ================== FUNÇÃO PRINCIPAL ==================
function gerarAbasAjustadas() {
  var inicio = new Date().getTime();
  var ss = SpreadsheetApp.openById(CONFIG_LAYOUT.idPlanilha);
  
  Logger.log('1. Lendo dados brutos da API...');
  var rawTiragem = _lerAbaComoObjetos(ss, CONFIG_LAYOUT.abaTiragemRaw);
  var rawEspec = _lerAbaComoObjetos(ss, CONFIG_LAYOUT.abaEspecRaw);
  var rawArvore = _lerAbaComoObjetos(ss, CONFIG_LAYOUT.abaArvoreRaw);
  
  Logger.log('2. Lendo layouts idênticos das abas originais...');
  var cabecalhoTiragem = _lerCabecalhoBase(ss, CONFIG_LAYOUT.abaTiragemTemplate);
  var cabecalhoEspec = _lerCabecalhoBase(ss, CONFIG_LAYOUT.abaEspecTemplate);
  var cabecalhoArvore = _lerCabecalhoBase(ss, CONFIG_LAYOUT.abaArvoreTemplate);

  if (!rawTiragem || !rawEspec || !rawArvore || cabecalhoEspec.length === 0) {
    Logger.log('ERRO: Falha ao ler os dados ou os cabeçalhos. Verifique os nomes das abas.');
    return;
  }

  Logger.log('3. Criando dicionário de Especificações para Cruzamento (JOIN)...');
  var dicEspec = {};
  for (var i = 0; i < rawEspec.length; i++) {
    var sku = String(rawEspec[i]['skuAtual']).trim();
    if (sku) dicEspec[sku] = rawEspec[i];
  }

  Logger.log('4. Mapeando ESPECIFICAÇÕES (Extraindo Miolos, Encartes e Adesivos aninhados)...');
  _processarEspec(ss, rawEspec, cabecalhoEspec);
  
  Logger.log('5. Mapeando TIRAGEM (com cruzamento)...');
  _processarTiragem(ss, rawTiragem, dicEspec, cabecalhoTiragem);
  
  Logger.log('6. Mapeando ÁRVORE (Resolvendo sub-objetos aninhados no JSON)...');
  _processarArvore(ss, rawArvore, cabecalhoArvore);
  
  var segundos = Math.round((new Date().getTime() - inicio) / 1000);
  Logger.log('✅ SUCESSO! Abas ajustadas criadas em ' + segundos + 's.');
}

// ================== MOTORES DE MAPEAMENTO ==================

function _processarEspec(ss, rawEspec, cabecalhos) {
  var linhas = [];
  linhas.push(cabecalhos); 

  for (var i = 0; i < rawEspec.length; i++) {
    var r = rawEspec[i];
    var linha = [];
    
    // Dimensões Concatenadas
    var dimBxAAberto = (r['capBaseMmAberto'] && r['capAlturaMmAberto']) ? r['capBaseMmAberto'] + 'x' + r['capAlturaMmAberto'] : '';
    var dimBxAFechado = (r['baseMm'] && r['alturaMm']) ? r['baseMm'] + 'x' + r['alturaMm'] : '';
    var dimAcabada = (dimBxAFechado && r['espessuraMm']) ? dimBxAFechado + 'x' + String(r['espessuraMm']).replace('.', ',') : '';

    // Desaninhamento Extremo (Arrays) baseados no Mapa
    var m1 = _getArrItem(r['miolos'], 0); // Miolo 1 (Aluno)
    var m2 = _getArrItem(r['miolos'], 1); // Miolo 2 (Professor)
    var enc = _getArrItem(r['encartes'], 0); // Encarte
    var ad = _getArrItem(r['adesivos'], 0);  // Adesivo

    for (var c = 0; c < cabecalhos.length; c++) {
      var col = String(cabecalhos[c]).trim();
      var val = '';
      
      // Identificação
      if (col === 'Status SKU 2026') val = r['status'];
      else if (col === '1. ID_Código SKU') val = r['skuAtual'];
      else if (col === '1. ID_Descrição') val = r['descricao'];
      else if (col === '1. ID_ISBN') val = r['isbn'];
      else if (col === '1. ID_Código SKU anterior') val = r['skuReferencia'];
      // Informações Gerais
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
      else if (col === '2. INF_Cliente personalizado') val = (String(r['clientePersonalizado']).toLowerCase() === 'true' || r['clientePersonalizado'] === true) ? 'Sim' : r['clientePersonalizado'];
      // Dimensões
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
      // Capa
      else if (col === '4. CAP_Tipo de capa') val = r['capTipoCapa'];
      else if (col === '4.Cap_orelha' || col === '4. CAP_Orelha') val = r['capLarguraOrelha'] || (r['capTemOrelhas'] ? 'Sim' : '');
      else if (col === '4. CAP_Papel da capa') val = r['capPapel'];
      else if (col === '4. CAP_Cor capa') val = r['capCor'];
      else if (col === '4. CAP_Papel revestimento') val = r['capRevestimentoPapel'];
      else if (col === '4. CAP_Papel forro') val = r['capForroPapel'];
      else if (col === '4. CAP_Acabamento da capa') val = r['capRevestimentoTipo'];
      else if (col === '4. CAP_Acabamento interno capa') val = r['capAcabamentoObs'] || r['capRevestimentoAplicacao'];
      else if (col === '4. CAP_Aproveitamento chapa capa (Cor)') val = r['capCorChapa'];
      else if (col === '4. CAP_Corte e vinco') val = _boolPT(r['capTemCorteVinco']);
      // Miolo 1 (Aluno)
      else if (col === '5. MIO_PRIN_Quantidade de páginas') val = m1['paginacao'];
      else if (col === '5. MIO_PRIN_Papel') val = m1['papel'];
      else if (col === '5. MIO_PRIN_Cor') val = m1['cor'];
      else if (col === '5. MIO_PRIN_Aproveitamento chapa miolo (Cor)') val = m1['corChapa'];
      else if (col === '5. MIO_PRIN_Corte e vinco') val = _boolPT(m1['corteVinco']);
      else if (col === '5. MIO_PRIN_Serrilha') val = _boolPT(m1['serrilha']);
      else if (col === '5. MIO_PRIN_Obs miolo') val = m1['observacoes'];
      // Miolo 2 (Professor)
      else if (col === '6. MIO_PROF_Quantidade de páginas') val = m2['paginacao'];
      else if (col === '6. MIO_PROF_Papel') val = m2['papel'];
      else if (col === '6. MIO_PROF_Cor') val = m2['cor'];
      else if (col === '6. MIO_PROF_Aproveitamento chapa miolo (Cor)') val = m2['corChapa'];
      else if (col === '6. MIO_PROF_Corte e vinco') val = _boolPT(m2['corteVinco']);
      else if (col === '6. MIO_PROF_Serrilha') val = _boolPT(m2['serrilha']);
      else if (col === '6. MIO_PROF_Obs miolo') val = m2['observacoes'];
      // Encarte
      else if (col === '7. MIO_ENCA_Quantidade de páginas') val = enc['paginacao'];
      else if (col === '7. MIO_ENCA_Papel') val = enc['papel'];
      else if (col === '7. MIO_ENCA_Cor') val = enc['cor'];
      else if (col === '7. MIO_ENCA_Aproveitamento chapa encarte (Cor)') val = enc['corChapa'];
      else if (col === '7. MIO_ENCA_Corte e vinco') val = _boolPT(enc['corteVinco']);
      else if (col === '7. MIO_ENCA_Serrilha') val = _boolPT(enc['serrilha']);
      else if (col === '7. MIO_ENCA_Obs encarte') val = enc['observacoes'];
      // Adesivo
      else if (col === '8. MIO_ADES_Quantidade de páginas') val = ad['paginacao'];
      else if (col === '8. MIO_ADES_Papel') val = ad['papel'];
      else if (col === '8. MIO_ADES_Cor') val = ad['cor'];
      else if (col === '8. MIO_ADES_Aproveitamento chapa adesivo (Cor)') val = ad['corChapa'];
      else if (col === '8. MIO_ADES_Corte e vinco') val = _boolPT(ad['corteVinco']);
      else if (col === '8. MIO_ADES_Obs adesivo') val = ad['observacoes'];
      // Acabamento / Encadernação
      else if (col === '9. ENC_Tipo acabamento') val = r['capAcabamento'];
      else if (col === '9. ENC_Posição') val = r['capPosicaoAcabamento'];
      else if (col === '9.ENC_Bitola do espiral' || col === '9. ENC_Bitola do espiral') val = r['encBitolaEspiral'];
      else if (col === '9. ENC_Cor do espiral') val = r['encCorEspiral'];
      else if (col === '9. ENC_Obs de encadernação') val = r['obsEncadernacao'];
      else if (col === '10. OBS_PROD_Observação para produção gráfica') val = r['obsNotas'];
      else if (col === '10. Orientação de montagem do livro') val = r['obsMontagem'];
      // Referências
      else if (col === '11. REF_Referência troca de chapa') val = r['capSkuReferenciaChapa'] || r['skuReferenciaChapa'];
      else if (col === '11. REF_Referência personalização') val = r['skuMaterialPadrao'];
      
      linha.push(val === undefined || val === null ? '' : val);
    }
    linhas.push(linha);
  }
  _gravarAba(ss, CONFIG_LAYOUT.abaEspecOut, linhas);
}

function _processarTiragem(ss, rawTiragem, dicEspec, cabecalhos) {
  var linhas = [];
  linhas.push(cabecalhos);

  for (var i = 0; i < rawTiragem.length; i++) {
    var r = rawTiragem[i];
    var sku = String(r['sku']).trim();
    var espec = dicEspec[sku] || {}; 
    var linha = [];

    for (var c = 0; c < cabecalhos.length; c++) {
      var col = String(cabecalhos[c]).trim();
      var val = '';

      if (col === '1. ID_Código Kit') val = r['kitCode'];
      else if (col === '1. ID_Descrição Kit') val = r['kitDescricao'];
      else if (col === '1. ID_Código somente KIT') val = r['kitCode'];
      else if (col === '1. ID_Código SKU') val = sku;
      else if (col === '1. ID_ISBN') val = espec['isbn'];
      else if (col === '1. ID_Descrição') val = r['skuDescricao'];
      else if (col === '2. INF_Marca') val = r['marca'];
      else if (col === '2. INF_Grupo da marca') val = espec['grupoMarca'];
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
      else if (col === 'MARCA') val = r['marca'];
      else if (col === 'GRAFICA') val = r['grafica1Atual'];

      linha.push(val === undefined || val === null ? '' : val);
    }
    linhas.push(linha);
  }
  _gravarAba(ss, CONFIG_LAYOUT.abaTiragemOut, linhas);
}

function _processarArvore(ss, rawArvore, cabecalhos) {
  var linhas = [];
  linhas.push(cabecalhos);

  for (var i = 0; i < rawArvore.length; i++) {
    var r = rawArvore[i];
    
    // Desaninhamento dos Objetos JSON e Sub-objetos
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

    // ===================================
    // LINHA PAI
    // ===================================
    var linhaPai = [];
    for (var c = 0; c < cabecalhos.length; c++) {
      var col = String(cabecalhos[c]).trim();
      var val = '';

      if (col === '1. ID_Código Kit') val = r['kitCode'];
      else if (col === '1. ID_Descrição Kit') val = r['description'];
      else if (col === '1. ID_Código somente KIT') val = r['kitCode'];
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

      linhaPai.push(val === undefined || val === null ? '' : val);
    }
    linhas.push(linhaPai);

    // ===================================
    // LINHAS FILHAS (SKUs desaninhados)
    // ===================================
    for (var j = 0; j < itens.length; j++) {
      var filho = itens[j];
      var prodFilho = filho['product'] || {}; // Algumas infos ficam ocultas em sub-objeto 'product'
      var linhaFilho = [];

      for (var c = 0; c < cabecalhos.length; c++) {
        var col = String(cabecalhos[c]).trim();
        var val = '';

        if (col === '1. ID_Código Kit') val = r['kitCode'];
        else if (col === '1. ID_Descrição Kit') val = r['description'];
        else if (col === '1. ID_Código somente KIT') val = r['kitCode'];
        // O SKU final está em itemCode ou dentro do sub-objeto product
        else if (col === '1. ID_Código SKU') val = filho['itemCode'] || filho['skuAtual'] || prodFilho['skuAtual'];
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

        linhaFilho.push(val === undefined || val === null ? '' : val);
      }
      linhas.push(linhaFilho);
    }
  }
  _gravarAba(ss, CONFIG_LAYOUT.abaArvoreOut, linhas);
}

// ================== HELPERS DE DADOS ==================

/** Retorna 'Sim' ou 'Não' se for booleano, ou o próprio valor */
function _boolPT(val) {
  if (val === true || String(val).toLowerCase() === 'true') return 'Sim';
  if (val === false || String(val).toLowerCase() === 'false') return 'Não';
  return val;
}

/** Extrai chaves específicas (index) de um array JSON aninhado (Ex: extrai Miolo 1 ou Miolo 2) */
function _getArrItem(jsonString, index) {
  try {
    var arr = JSON.parse(jsonString);
    if (Array.isArray(arr) && arr.length > index) {
      return arr[index] || {};
    }
  } catch(e) {}
  return {};
}

/** Transforma string em Objeto JSON com Fallback */
function _parse(jsonString, fallback) {
  try { return JSON.parse(jsonString) || fallback; } 
  catch(e) { return fallback; }
}

function _lerCabecalhoBase(ss, nomeAba) {
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) return [];
  var dados = aba.getRange(1, 1, 1, aba.getMaxColumns()).getValues()[0];
  var cols = [];
  for (var i = 0; i < dados.length; i++) {
    if (dados[i] !== "") cols.push(dados[i]); 
  }
  return cols;
}

function _lerAbaComoObjetos(ss, nomeAba) {
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) return null;
  var dados = aba.getDataRange().getValues();
  if (dados.length < 2) return []; 
  var cabecalho = dados[0];
  var objetos = [];
  for (var i = 1; i < dados.length; i++) {
    var linha = dados[i];
    var obj = {};
    for (var c = 0; c < cabecalho.length; c++) {
      obj[cabecalho[c]] = linha[c];
    }
    objetos.push(obj);
  }
  return objetos;
}

function _gravarAba(ss, nomeAba, dadosMatriz) {
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) aba = ss.insertSheet(nomeAba);
  else aba.clear(); 
  
  if (dadosMatriz.length > 0) {
    aba.getRange(1, 1, dadosMatriz.length, dadosMatriz[0].length).setValues(dadosMatriz);
    aba.getRange(1, 1, 1, dadosMatriz[0].length).setFontWeight("bold");
  }
  SpreadsheetApp.flush();
}