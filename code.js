// CEP de Origem do Vendedor (CEP correto: 20541-280)
const ORIGIN_CEP = "20541280";

// Configurações do Produto e Embalagem (Envelope ~ 13 x 18 cm)
const PRODUCT_WIDTH_CM = 13;
const PRODUCT_LENGTH_CM = 18;

function doPost(e) {
  // Intercepta chamadas de Webhook do Superfrete / Seu Rastreio via query parameter
  if (e.parameter && (e.parameter.source === "superfrete" || e.parameter.source === "seurastreio")) {
    return processarWebhookSuperFrete(e);
  }

  try {
    let requestData;
    if (e.postData && e.postData.contents) {
      requestData = JSON.parse(e.postData.contents);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Nenhum dado recebido no postData" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const action = requestData.action;

    // Intercepta chamadas de Webhook do Superfrete / Seu Rastreio enviadas de forma direta
    if (!action && requestData.event && requestData.order) {
      return processarWebhookSuperFreteDirect(requestData);
    }

    if (action === "calcular_frete") {
      const response = calcularFrete(requestData.cep_destino, requestData.quantidade, requestData);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    } 
    
    if (action === "salvar_pedido") {
      const response = salvarPedido(requestData.pedido);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "get_pedido") {
      const response = getPedido(requestData.orderId);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "get_recent_orders") {
      const response = getRecentOrders();
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "get_all_orders") {
      const response = getAllOrders();
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "atualizar_status_pedido") {
      const response = atualizarStatusPedido(requestData.orderId, requestData.status, requestData.data_pagamento);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "conciliar_recibos") {
      const response = conciliarRecibos();
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "emitir_etiqueta_pedido") {
      const response = emitirEtiquetaPedido(requestData.orderId);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "atualizar_rastreamento_pedido") {
      const response = atualizarRastreamentoPedido(requestData.orderId);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "sincronizar_todos_rastreamentos") {
      const response = sincronizarTodosRastreamentos();
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "configurar_ambiente") {
      PropertiesService.getScriptProperties().setProperty("SUPERFRETE_ENVIRONMENT", requestData.ambiente);
      return ContentService.createTextOutput(JSON.stringify({ success: true, ambiente: requestData.ambiente }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "reordenar_colunas") {
      reordenarColunasPlanilha();
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "set_property") {
      PropertiesService.getScriptProperties().setProperty(requestData.name, requestData.value);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "delete_property") {
      PropertiesService.getScriptProperties().deleteProperty(requestData.name);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "get_properties") {
      const props = PropertiesService.getScriptProperties().getProperties();
      return ContentService.createTextOutput(JSON.stringify({ success: true, properties: props }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "validar_cupom") {
      const response = validarCupom(requestData.cupom);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "test_telegram") {
      try {
        const response = enviarNotificacaoTelegram(requestData.titulo, requestData.mensagem);
        return ContentService.createTextOutput(JSON.stringify({ success: true, response: response }))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    if (action === "registrar_webhook") {
      try {
        const response = registrarWebhookNoSuperFrete();
        return ContentService.createTextOutput(JSON.stringify({ success: true, response: response }))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    if (action === "configurar_telegram") {
      try {
        const response = configurarTelegramChatId();
        return ContentService.createTextOutput(JSON.stringify(response))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    if (action === "reset_order") {
      try {
        const orderId = requestData.orderId;
        const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
        let ss = null;
        try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
        if (!ss) ss = SpreadsheetApp.openById(spreadsheetId);
        const sheet = ss.getSheetByName("Pedidos");
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        
        const statusIndex = headers.indexOf("Status");
        const packageIdIndex = headers.indexOf("Superfrete_Package_ID");
        const sfStatusIndex = headers.indexOf("Superfrete_Status");
        const trackingIndex = headers.indexOf("Codigo_Rastreio");
        const pdfIndex = headers.indexOf("Etiqueta_PDF_URL");
        const dataEmissaoIndex = headers.indexOf("Data_Emissao_Etiqueta");
        
        let rowIdx = -1;
        for (let i = 1; i < data.length; i++) {
          if (parseInt(data[i][0]) === parseInt(orderId)) {
            rowIdx = i + 1;
            break;
          }
        }
        if (rowIdx !== -1) {
          sheet.getRange(rowIdx, statusIndex + 1).setValue("Pago");
          if (packageIdIndex !== -1) sheet.getRange(rowIdx, packageIdIndex + 1).setValue("");
          if (sfStatusIndex !== -1) sheet.getRange(rowIdx, sfStatusIndex + 1).setValue("");
          if (trackingIndex !== -1) sheet.getRange(rowIdx, trackingIndex + 1).setValue("");
          if (pdfIndex !== -1) sheet.getRange(rowIdx, pdfIndex + 1).setValue("");
          if (dataEmissaoIndex !== -1) sheet.getRange(rowIdx, dataEmissaoIndex + 1).setValue("");
          return ContentService.createTextOutput(JSON.stringify({ success: true }))
            .setMimeType(ContentService.MimeType.JSON);
        }
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Pedido não encontrado" }))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    if (action === "get_cupons") {
      const response = getCupons();
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "salvar_cupom") {
      const response = salvarCupom(requestData.cupom);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "excluir_cupom") {
      const response = excluirCupom(requestData.cupom);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "alternar_status_cupom") {
      const response = alternarStatusCupom(requestData.cupom);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "get_promocoes") {
      const response = getPromocoes();
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "salvar_promocao") {
      const response = salvarPromocao(requestData.promocao);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "excluir_promocao") {
      const response = excluirPromocao(requestData.nome);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "alternar_status_promocao") {
      const response = alternarStatusPromocao(requestData.nome);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "get_custos_data") {
      const response = getCustosData();
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "salvar_insumo") {
      const response = salvarInsumo(requestData.insumo);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "salvar_compra_insumo") {
      const response = salvarCompraInsumo(requestData.compra);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "salvar_ficha_tecnica") {
      const response = salvarFichaTecnica(requestData.ficha);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "excluir_item_ficha") {
      const response = excluirItemFichaTecnica(requestData.id_produto, requestData.id_insumo);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "salvar_caixa") {
      const response = salvarCaixa(requestData.caixa);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "excluir_caixa") {
      const response = excluirCaixa(requestData.id_caixa);
      return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "recalcular_tudo") {
      const ss = SpreadsheetApp.openById("1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0");
      const sheetInsumos = obterPlanilhaInsumos(ss);
      const insumosData = sheetInsumos.getDataRange().getValues();
      const headers = insumosData[0];
      const idIdx = headers.indexOf("ID_Insumo");
      const results = {};
      for (let i = 1; i < insumosData.length; i++) {
        const id = String(insumosData[i][idIdx]).trim();
        if (id) {
          const val = recalcularCustoMedioInsumo(ss, id);
          results[id] = val;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, results: results }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Ação desconhecida" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function calcularFrete(cepDestino, quantidade, extraParams = {}) {
  const cleanCep = cepDestino.replace(/\D/g, "");
  if (cleanCep.length !== 8) {
    throw new Error("CEP de destino inválido. Deve conter 8 dígitos.");
  }

  const qty = parseInt(quantidade) || 1;
  
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!ss) ss = SpreadsheetApp.openById(spreadsheetId);
  
  const caixa = selecionarCaixaLogistica(ss, qty);
  
  // Parâmetros customizados ou cotação dinâmica com base na caixa recomendada
  const finalFormat = extraParams.custom_format !== undefined ? extraParams.custom_format : 1; // 1 = Caixa/Pacote
  const finalWeight = extraParams.custom_weight !== undefined ? extraParams.custom_weight : ((qty * 0.03567) + caixa.peso);
  
  const finalHeight = extraParams.custom_height !== undefined ? extraParams.custom_height : caixa.altura;
  const finalWidth = extraParams.custom_width !== undefined ? extraParams.custom_width : caixa.largura;
  const finalLength = extraParams.custom_length !== undefined ? extraParams.custom_length : caixa.comprimento;

  const finalServices = extraParams.custom_services !== undefined ? extraParams.custom_services : "1,2,3,4,17,31";

  // Busca o Token de forma segura de acordo com o ambiente (sandbox ou produção)
  const sfConfig = obterConfiguracaoSuperFrete();
  const token = sfConfig.token;
  if (!token) {
    throw new Error("Token da SuperFrete não configurado nas propriedades do script.");
  }

  const url = sfConfig.baseUrl + "/api/v0/calculator";
  
  const payload = {
    "from": {
      "postal_code": ORIGIN_CEP.replace(/\D/g, "")
    },
    "to": {
      "postal_code": cleanCep
    },
    "services": finalServices,
    "package": {
      "format": finalFormat,
      "weight": finalWeight,
      "width": finalWidth,
      "length": finalLength
    },
    "options": {
      "own_hand": false,
      "receipt": false,
      "insurance_value": 0,
      "use_insurance_value": false
    }
  };

  payload.package.height = finalHeight;

  const options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {
      "Authorization": "Bearer " + token,
      "Accept": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseText = response.getContentText();
  const responseData = JSON.parse(responseText);

  if (response.getResponseCode() !== 200) {
    throw new Error("Erro na API da SuperFrete: " + (responseData.message || responseText));
  }

  // Filtrar serviços bem-sucedidos
  let valorProdutos = parseFloat(extraParams.valor_produtos);
  if (isNaN(valorProdutos) || valorProdutos <= 0) {
    const kits = Math.floor(qty / 4);
    const resto = qty % 4;
    valorProdutos = kits * 45;
    if (resto === 3) valorProdutos += 40;
    else if (resto === 2) valorProdutos += 28;
    else if (resto === 1) valorProdutos += 15;
  }

  const custos = obterCustosDynamic(ss, caixa, qty);
  const lucroBruto = valorProdutos - (qty * custos.custoCadernoUnitario) - custos.custoCaixa - custos.custoExtras;

  const services = [];
  if (Array.isArray(responseData)) {
    responseData.forEach(service => {
      if (!service.error) {
        const finalPrice = parseFloat(service.price);
        const discountAmount = parseFloat(service.discount || 0);
        const isEligibleFree = lucroBruto > 0 && finalPrice <= (0.30 * lucroBruto);
        
        services.push({
          id: service.id,
          name: isEligibleFree ? (service.name + " (Grátis)") : service.name,
          price: isEligibleFree ? 0 : (finalPrice + discountAmount),
          discountPrice: isEligibleFree ? 0 : finalPrice,
          deliveryTime: service.delivery_time,
          error: null
        });
      }
    });
  } else if (responseData.error) {
    throw new Error(responseData.error);
  }

  return { 
    success: true, 
    services: services,
    embalagem: {
      nome: caixa.id,
      comprimento: finalLength,
      largura: finalWidth,
      altura: finalHeight,
      peso: finalWeight
    }
  };
}

function obterCustosDynamic(ss, caixa, qty) {
  const sheetInsumos = obterPlanilhaInsumos(ss);
  const insumosData = readSheetAsJSON(sheetInsumos);
  const insumosMap = {};
  insumosData.forEach(i => {
    insumosMap[i.ID_Insumo] = {
      custo: parseFloat(i.Custo_Unitario_Atual) || 0,
      unidade: String(i.Unidade_Medida).trim()
    };
  });
  
  // 1. Custo do caderno (Ficha Técnica)
  const sheetFicha = obterPlanilhaFichaTecnica(ss);
  const fichaData = readSheetAsJSON(sheetFicha);
  let custoCadernoUnitario = 0;
  fichaData.forEach(item => {
    if (item.ID_Produto === "caderno_14x9") {
      const insumo = insumosMap[item.ID_Insumo];
      const insumoCost = insumo ? insumo.custo : 0;
      custoCadernoUnitario += (parseFloat(item.Quantidade_Consumida) || 0) * insumoCost;
    }
  });
  if (custoCadernoUnitario === 0) {
    custoCadernoUnitario = 1.32; // Fallback de segurança
  }
  
  // 2. Custo da caixa
  let custoCaixa = 0;
  if (caixa && caixa.idInsumo) {
    const insumo = insumosMap[caixa.idInsumo];
    custoCaixa = insumo ? insumo.custo : 0;
  }
  
  // 3. Custos de fita e etiqueta (insumos extras)
  let custoExtras = 0;
  
  // Fita kraft
  const fitaKey = Object.keys(insumosMap).find(k => k.toLowerCase().indexOf("fita") !== -1);
  if (fitaKey && caixa) {
    const fitaObj = insumosMap[fitaKey];
    if (fitaObj.unidade.toLowerCase() === "metro" || fitaObj.unidade.toLowerCase() === "m" || fitaObj.unidade.toLowerCase() === "metros") {
      // Perímetro horizontal: 2 * (comprimento + largura) + 10 cm de sobra
      const metrosFita = (2 * (parseFloat(caixa.comprimento) + parseFloat(caixa.largura)) + 10) / 100;
      custoExtras += metrosFita * fitaObj.custo;
    } else {
      custoExtras += fitaObj.custo; // Custo fixo unitário
    }
  }
  
  // Etiqueta adesiva
  const etiquetaKey = Object.keys(insumosMap).find(k => k.toLowerCase().indexOf("adesivo") !== -1 || k.toLowerCase().indexOf("etiqueta") !== -1);
  if (etiquetaKey) {
    const etiquetaObj = insumosMap[etiquetaKey];
    custoExtras += etiquetaObj.custo;
  }
  
  return {
    custoCadernoUnitario: custoCadernoUnitario,
    custoCaixa: custoCaixa,
    custoExtras: custoExtras
  };
}

function selecionarCaixaLogistica(ss, qty) {
  const cacheKey = "GRADE_CAIXAS_CACHE";
  const cache = CacheService.getScriptCache();
  let caixasJson = null;
  try {
    caixasJson = cache.get(cacheKey);
  } catch (e) {
    Logger.log("Erro ao acessar o cache: " + e.message);
  }

  let caixas = [];
  if (caixasJson) {
    try {
      caixas = JSON.parse(caixasJson);
    } catch (e) {
      Logger.log("Erro ao fazer parse do cache de caixas: " + e.message);
    }
  }

  if (caixas.length === 0) {
    const sheetGrade = obterPlanilhaGradeCaixas(ss);
    const data = sheetGrade.getDataRange().getValues();
    const headers = data[0];
    
    const idIdx = headers.indexOf("ID_Caixa");
    const insumoIdx = headers.indexOf("ID_Insumo");
    const capIdx = headers.indexOf("Capacidade_Max_Cadernos");
    const compIdx = headers.indexOf("Comprimento_CM");
    const largIdx = headers.indexOf("Largura_CM");
    const altIdx = headers.indexOf("Altura_CM");
    const pesoIdx = headers.indexOf("Peso_Vazio_KG");
    
    for (let i = 1; i < data.length; i++) {
      const capacidadeRaw = parseInt(data[i][capIdx], 10);
      caixas.push({
        id: data[i][idIdx],
        idInsumo: insumoIdx !== -1 ? data[i][insumoIdx] : null,
        capacidade: isNaN(capacidadeRaw) ? 1 : capacidadeRaw,
        comprimento: parseFloat(data[i][compIdx]) || 0,
        largura: parseFloat(data[i][largIdx]) || 0,
        altura: parseFloat(data[i][altIdx]) || 0,
        peso: parseFloat(data[i][pesoIdx]) || 0
      });
    }
    
    // Salva no cache por 2 minutos (120 segundos) para minimizar acessos futuros à planilha mas refletir alterações rápidas
    try {
      cache.put(cacheKey, JSON.stringify(caixas), 120);
    } catch (cacheErr) {
      Logger.log("Erro ao salvar caixas no cache: " + cacheErr.message);
    }
  }
  
  // Ordena por capacidade crescente
  caixas.sort((a, b) => a.capacidade - b.capacidade);
  
  // Acha a menor que suporta a qty
  let selecionada = null;
  for (let c of caixas) {
    if (c.capacidade >= qty) {
      selecionada = c;
      break;
    }
  }
  
  // Se a qty for maior que a maior caixa, pega a maior de todas
  if (!selecionada && caixas.length > 0) {
    selecionada = caixas[caixas.length - 1];
  }
  
  // Fallback se a grade de caixas estiver vazia
  if (!selecionada) {
    selecionada = {
      id: "caixa_p_fallback",
      capacidade: 5,
      comprimento: 16,
      largura: 11,
      altura: 3,
      peso: 0.05
    };
  }
  
  return selecionada;
}

function salvarPedido(pedido) {
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  
  let sheet = ss.getSheetByName("Pedidos");
  if (!sheet) {
    sheet = ss.insertSheet("Pedidos");
  }

  const defaultHeaders = [
    "ID_Pedido", "Status", "Data_Pedido", "Data_Pagamento", "Nome", "WhatsApp", "CPF",
    "Entrega_Rua", "Entrega_Numero", "Entrega_Complemento", "Entrega_Bairro", "Entrega_Cidade", "Entrega_Estado", "Entrega_CEP",
    "Itens", "Qtd_Total", "Valor_Produtos", "Cupom_Codigo", "Cupom_Desconto_Valor", "Frete_Metodo", "Valor_Frete", "Total_Geral", "Observacoes", "Pix_Payload",
    "Superfrete_Package_ID", "Superfrete_Status", "Codigo_Rastreio", "Etiqueta_PDF_URL", "Data_Emissao_Etiqueta", "Valor_Frete_Pago", "Desconto_Frete",
    "Superfrete_Custo_Estimado", "Superfrete_Desconto_Estimado",
    "Superfrete_Prazo_Entrega", "Data_Postagem", "Data_Entrega",
    "Superfrete_Formato", "Superfrete_Peso", "Superfrete_Altura", "Superfrete_Largura", "Superfrete_Comprimento",
    "Superfrete_Mao_Propria", "Superfrete_Aviso_Recebimento", "Superfrete_Valor_Declarado"
  ];

  // Se a planilha estiver vazia, cria os cabeçalhos automaticamente
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(defaultHeaders);
  }

  // Lemos os cabeçalhos atuais da primeira linha
  let currentHeaders = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  
  // Se a coluna antiga "Entrega_Endereco" existir, ou a nova "Entrega_Rua" estiver faltando, ou "Data_Pedido" estiver faltando, ou os campos de cupom estiverem ausentes, ou as novas colunas estimadas estiverem ausentes, aciona a migração/reordenação
  if (currentHeaders.indexOf("Entrega_Endereco") !== -1 || currentHeaders.indexOf("Entrega_Rua") === -1 || currentHeaders.indexOf("Data_Pedido") === -1 || currentHeaders.indexOf("Cupom_Codigo") === -1 || currentHeaders.indexOf("Superfrete_Custo_Estimado") === -1) {
    reordenarColunasPlanilha();
    currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  // Determinar próximo ID de pedido (Coluna A)
  const lastRow = sheet.getLastRow();
  let nextId = 1001;
  if (lastRow > 1) {
    const lastId = sheet.getRange(lastRow, 1).getValue();
    if (lastId === "ID_Pedido") {
      nextId = 1001;
    } else {
      nextId = (parseInt(lastId) || 1000) + 1;
    }
  }

  const dataHora = new Date();
  
  // Mapeia os dados dinamicamente com base no nome do cabeçalho da coluna
  const rowValues = currentHeaders.map(header => {
    switch (header) {
      case "ID_Pedido": return nextId;
      case "Data_Pedido": return dataHora;
      case "Nome": return pedido.nome || "";
      case "WhatsApp": return pedido.whatsapp || "";
      case "CPF": return pedido.cpf || "";
      case "Entrega_Rua": return pedido.entrega_rua || "";
      case "Entrega_Numero": return pedido.entrega_numero || "";
      case "Entrega_Complemento": return pedido.entrega_complemento || "";
      case "Entrega_Bairro": return pedido.entrega_bairro || "";
      case "Entrega_Cidade": return pedido.entrega_cidade || "";
      case "Entrega_Estado": return pedido.entrega_estado || "";
      case "Entrega_CEP": return pedido.cep || "";
      case "Itens": return pedido.itens || "";
      case "Qtd_Total": return pedido.qtd_total || 0;
      case "Valor_Produtos": return pedido.valor_produtos || 0;
      case "Frete_Metodo": return pedido.frete_metodo || "";
      case "Valor_Frete": return pedido.valor_frete || 0;
      case "Total_Geral": return pedido.total_geral || 0;
      case "Observacoes": return pedido.observacoes || "";
      case "Status": return "Pendente de Pagamento";
      case "Pix_Payload": return pedido.pix_payload || "";
      case "Cupom_Codigo": return pedido.cupom_codigo || "";
      case "Cupom_Desconto_Valor": return pedido.cupom_desconto_valor !== undefined ? pedido.cupom_desconto_valor : (pedido.cupom_desconto || 0);
      default: return ""; // Se houver alguma coluna customizada criada pelo usuário
    }
  });

  const nextRow = lastRow + 1;
  sheet.getRange(nextRow, 1, 1, rowValues.length).setValues([rowValues]);

  // Se houver cupom aplicado, marcamos como usado caso seja de uso único
  if (pedido.cupom_codigo) {
    try {
      let sheetCupons = ss.getSheetByName("Cupons");
      if (sheetCupons) {
        const dataCupons = sheetCupons.getDataRange().getValues();
        const headersCupons = dataCupons[0];
        const codIdx = headersCupons.indexOf("Codigo");
        const usoUnicoIdx = headersCupons.indexOf("Uso_Unico");
        const usadoIdx = headersCupons.indexOf("Usado");
        const ativoIdx = headersCupons.indexOf("Ativo");
        
        if (codIdx !== -1 && usoUnicoIdx !== -1 && usadoIdx !== -1) {
          const cupomProcurado = String(pedido.cupom_codigo).trim().toLowerCase();
          for (let i = 1; i < dataCupons.length; i++) {
            const codCupom = String(dataCupons[i][codIdx]).trim().toLowerCase();
            if (codCupom === cupomProcurado) {
              const usoUnico = String(dataCupons[i][usoUnicoIdx]).trim();
              if (usoUnico === "Sim") {
                sheetCupons.getRange(i + 1, usadoIdx + 1).setValue("Sim");
                if (ativoIdx !== -1) {
                  sheetCupons.getRange(i + 1, ativoIdx + 1).setValue("Não");
                }
              }
              break;
            }
          }
        }
      }
    } catch (e) {
      Logger.log("Erro ao marcar cupom como usado: " + e.message);
    }
  }

  // Envia notificação pelo Telegram
  try {
    const titulo = "🛍️ Novo Pedido Recebido! (Pedido #" + nextId + ")";
    const mensagem = "Um novo pedido de cadernos foi registrado na planilha.\n\n" +
                     "• Cliente: " + (pedido.nome || "Não informado") + "\n" +
                     "• Itens: " + (pedido.itens || "Não especificado") + "\n" +
                     "• Total: R$ " + (pedido.total_geral || 0) + "\n" +
                     "• Método Frete: " + (pedido.frete_metodo || "Não escolhido");
    enviarNotificacaoTelegram(titulo, mensagem);
  } catch (errTelegram) {
    Logger.log("Erro ao enviar notificacao de novo pedido para Telegram: " + errTelegram.message);
  }

  return { success: true, orderId: nextId };
}

function getPedido(orderId) {
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  let sheet = ss.getSheetByName("Pedidos");
  if (!sheet) {
    return { success: false, error: "Planilha não encontrada" };
  }
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (parseInt(data[i][0]) === parseInt(orderId)) {
      const pedidoObj = {};
      headers.forEach((header, index) => {
        pedidoObj[header] = data[i][index];
      });
      return { success: true, pedido: pedidoObj };
    }
  }
  return { success: false, error: "Pedido não encontrado" };
}

function getRecentOrders() {
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  let sheet = ss.getSheetByName("Pedidos");
  if (!sheet) {
    return { success: false, error: "Planilha não encontrada" };
  }
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const lastRows = data.slice(Math.max(1, data.length - 10)); // Get up to last 10 rows
  
  const orders = [];
  lastRows.forEach(row => {
    const pedidoObj = {};
    headers.forEach((header, index) => {
      pedidoObj[header] = row[index];
    });
    orders.push(pedidoObj);
  });
  
  return { success: true, orders: orders };
}

function getAllOrders() {
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  let sheet = ss.getSheetByName("Pedidos");
  if (!sheet) {
    return { success: false, error: "Planilha não encontrada" };
  }
  
  let headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];

  // Se a coluna antiga "Entrega_Endereco" existir, ou a nova "Entrega_Rua" estiver faltando, ou "Data_Pedido" estiver faltando, ou os campos de cupom estiverem ausentes, aciona a migração/reordenação
  if (headers.indexOf("Entrega_Endereco") !== -1 || headers.indexOf("Entrega_Rua") === -1 || headers.indexOf("Data_Pedido") === -1 || headers.indexOf("Cupom_Codigo") === -1) {
    reordenarColunasPlanilha();
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  const data = sheet.getDataRange().getValues();
  const orders = [];
  for (let i = 1; i < data.length; i++) {
    const pedidoObj = {};
    headers.forEach((header, index) => {
      pedidoObj[header] = data[i][index] !== undefined ? data[i][index] : "";
    });
    orders.push(pedidoObj);
  }
  return { success: true, orders: orders };
}

function atualizarStatusPedido(orderId, novoStatus, dataPagamento) {
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  let sheet = ss.getSheetByName("Pedidos");
  if (!sheet) {
    return { success: false, error: "Planilha não encontrada" };
  }
  
  let headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  
  // Garante que a coluna Data_Pagamento exista
  let payDateIndex = headers.indexOf("Data_Pagamento");
  if (payDateIndex === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue("Data_Pagamento");
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    payDateIndex = headers.indexOf("Data_Pagamento");
  }
  
  const statusIndex = headers.indexOf("Status");
  if (statusIndex === -1) {
    return { success: false, error: "Coluna de Status não encontrada na planilha" };
  }
  
  const data = sheet.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (parseInt(data[i][0]) === parseInt(orderId)) {
      rowIdx = i + 1; // 1-based index
      break;
    }
  }
  
  if (rowIdx === -1) {
    return { success: false, error: "Pedido não encontrado" };
  }

  const pedido = {};
  headers.forEach((header, index) => {
    pedido[header] = data[rowIdx - 1][index];
  });

  // Validação: se o status logístico atual for "pending" (Aguardando Emissão/Sem Saldo),
  // impede a mudança para qualquer status diferente de "Cancelado (Desistência)" ou do próprio "Etiqueta Gerada"
  const currentSfStatus = String(pedido.Superfrete_Status || "").toLowerCase().trim();
  const hasPdf = !!pedido.Etiqueta_PDF_URL;
  const isSemSaldo = (currentSfStatus === "pending" && !hasPdf);

  if (isSemSaldo && novoStatus !== "Cancelado (Desistência)" && novoStatus !== "Etiqueta Gerada") {
    return { success: false, error: "Este pedido está com etiqueta pendente de saldo no SuperFrete e não pode ser alterado para outros status, exceto 'Cancelado (Desistência)'." };
  }

  // Se o novo status for "Cancelado (Desistência)"
  if (novoStatus === "Cancelado (Desistência)") {
    let packageId = pedido.Superfrete_Package_ID;
    
    // Tenta decodificar da URL da etiqueta caso não tenha o packageId na planilha
    if (!packageId && pedido.Etiqueta_PDF_URL) {
      try {
        const pdfUrl = String(pedido.Etiqueta_PDF_URL);
        const match = pdfUrl.match(/_etiqueta\/pdf\/([^?]+)/);
        if (match) {
          const decoded = Utilities.newBlob(Utilities.base64Decode(match[1])).getDataAsString();
          const parsed = JSON.parse(decoded);
          if (parsed && parsed.order_id) {
            packageId = parsed.order_id;
          }
        }
      } catch (e) {
        Logger.log("Erro ao decodificar ID para cancelamento: " + e.message);
      }
    }

    if (packageId) {
      // 1. Consultar a API da SuperFrete para obter informações atuais do pacote
      const sfConfig = obterConfiguracaoSuperFrete();
      const token = sfConfig.token;
      const baseUrl = sfConfig.baseUrl;
      let sfStatus = "";
      
      if (token) {
        try {
          const apiHeaders = {
            "Authorization": "Bearer " + token,
            "Accept": "application/json",
            "User-Agent": "FigTree (andre.figueira@gmail.com)"
          };
          const optInfo = {
            "method": "get",
            "headers": apiHeaders,
            "muteHttpExceptions": true
          };
          const resInfo = UrlFetchApp.fetch(baseUrl + "/api/v0/order/info/" + packageId, optInfo);
          const resCode = resInfo.getResponseCode();
          if (resCode === 200) {
            const dataInfo = JSON.parse(resInfo.getContentText());
            if (dataInfo.status) {
              sfStatus = String(dataInfo.status).toLowerCase().trim();
            }
          } else if (resCode === 404 || resCode === 400) {
            sfStatus = "not_found";
          }
        } catch (err) {
          Logger.log("Erro ao consultar status da etiqueta antes de cancelar: " + err.message);
        }
      }

      // Se o status da etiqueta na SuperFrete for postado ou entregue, não podemos cancelar de forma alguma!
      if (sfStatus === "posted" || sfStatus === "delivered") {
        return { success: false, error: "A etiqueta já foi postada ou entregue nos Correios e não pode ser cancelada." };
      }

      // Verificar se a data de emissão da etiqueta está dentro do prazo de 10 dias civis
      let noPrazo = false;
      let dataRef = null;
      if (pedido.Data_Emissao_Etiqueta) {
        dataRef = new Date(pedido.Data_Emissao_Etiqueta);
      } else if (pedido.Data_Pedido) {
        // Fallback: se não há data de emissão gravada, usamos a data de criação do pedido
        dataRef = new Date(pedido.Data_Pedido);
      }
      
      if (dataRef && !isNaN(dataRef.getTime())) {
        const dRefZerada = new Date(dataRef.getFullYear(), dataRef.getMonth(), dataRef.getDate());
        const dHojeZerada = new Date();
        dHojeZerada.setHours(0, 0, 0, 0);
        const diffMs = dHojeZerada.getTime() - dRefZerada.getTime();
        const diasPassados = Math.round(diffMs / (1000 * 60 * 60 * 24));
        if (diasPassados <= 10) {
          noPrazo = true;
        }
      } else {
        // Se não houver nenhuma data de referência, assumimos true para tentar o cancelamento
        noPrazo = true;
      }

      // Só tenta cancelar na API se o status logístico atual for cancelável (released, printed ou pending)
      const statusCancelaveis = ["released", "printed", "pending"];
      const deveCancelarNaAPI = statusCancelaveis.includes(sfStatus) && noPrazo;

      if (sfStatus === "pending" && !deveCancelarNaAPI) {
        // Envia notificação informando o cancelamento do rascunho pendente
        try {
          const titulo = "❌ Etiqueta Cancelada (Pedido #" + orderId + ")";
          const mensagem = "A etiqueta pendente de saldo do pedido #" + orderId + " foi cancelada e removida.";
          enviarNotificacaoTelegram(titulo, mensagem);
        } catch (errTelegram) {
          Logger.log("Erro ao enviar notificacao Telegram de cancelamento: " + errTelegram.message);
        }
      }

      if (deveCancelarNaAPI) {
        const cancelRes = cancelarEtiquetaSuperfrete(packageId);
        if (!cancelRes.success) {
          const errMsg = String(cancelRes.error).toLowerCase();
          const isAcceptableError = errMsg.includes("não pode mais ser cancelado") || 
                                    errMsg.includes("não encontrado") || 
                                    errMsg.includes("inexistente") ||
                                    errMsg.includes("conflito") ||
                                    errMsg.includes("status") ||
                                    errMsg.includes("not found") ||
                                    errMsg.includes("conflict") ||
                                    errMsg.includes("unauthorized") ||
                                    errMsg.includes("bad request") ||
                                    errMsg.includes("cannot be canceled") ||
                                    errMsg.includes("erro desconhecido");
          
          if (!isAcceptableError) {
            // Se o cancelamento na API falhar por algum outro motivo, bloqueia a alteração
            return { success: false, error: "Falha ao cancelar etiqueta na SuperFrete: " + cancelRes.error };
          }
        }
        
        // Envia notificação informando o cancelamento
        try {
          const titulo = "❌ Etiqueta Cancelada (Pedido #" + orderId + ")";
          const mensagem = "A etiqueta do pedido #" + orderId + " foi cancelada com sucesso na SuperFrete.";
          enviarNotificacaoTelegram(titulo, mensagem);
        } catch (errTelegram) {
          Logger.log("Erro ao enviar notificacao Telegram de cancelamento: " + errTelegram.message);
        }
      }

      // Se cancelado com sucesso ou se não era necessário cancelar (porque já expirou ou já estava cancelada na SuperFrete), limpa os campos de frete
      const trackingIndex = headers.indexOf("Codigo_Rastreio");
      const pdfIndex = headers.indexOf("Etiqueta_PDF_URL");
      const dataEmissaoIndex = headers.indexOf("Data_Emissao_Etiqueta");
      const packageIdIndex = headers.indexOf("Superfrete_Package_ID");
      const prazoIndex = headers.indexOf("Superfrete_Prazo_Entrega");
      const fretePagoIndex = headers.indexOf("Valor_Frete_Pago");
      const descontoFreteIndex = headers.indexOf("Desconto_Frete");
      const custoEstimadoIndex = headers.indexOf("Superfrete_Custo_Estimado");
      const descontoEstimadoIndex = headers.indexOf("Superfrete_Desconto_Estimado");
      const superfreteStatusIndex = headers.indexOf("Superfrete_Status");
      const postagemIndex = headers.indexOf("Data_Postagem");
      const entregaIndex = headers.indexOf("Data_Entrega");

      if (trackingIndex !== -1) sheet.getRange(rowIdx, trackingIndex + 1).setValue("");
      if (pdfIndex !== -1) sheet.getRange(rowIdx, pdfIndex + 1).setValue("");
      if (dataEmissaoIndex !== -1) sheet.getRange(rowIdx, dataEmissaoIndex + 1).setValue("");
      if (packageIdIndex !== -1) sheet.getRange(rowIdx, packageIdIndex + 1).setValue("");
      if (prazoIndex !== -1) sheet.getRange(rowIdx, prazoIndex + 1).setValue("");
      if (fretePagoIndex !== -1) sheet.getRange(rowIdx, fretePagoIndex + 1).setValue("");
      if (descontoFreteIndex !== -1) sheet.getRange(rowIdx, descontoFreteIndex + 1).setValue("");
      if (custoEstimadoIndex !== -1) sheet.getRange(rowIdx, custoEstimadoIndex + 1).setValue("");
      if (descontoEstimadoIndex !== -1) sheet.getRange(rowIdx, descontoEstimadoIndex + 1).setValue("");
      if (superfreteStatusIndex !== -1) sheet.getRange(rowIdx, superfreteStatusIndex + 1).setValue("");
      if (postagemIndex !== -1) sheet.getRange(rowIdx, postagemIndex + 1).setValue("");
      if (entregaIndex !== -1) sheet.getRange(rowIdx, entregaIndex + 1).setValue("");
    }
  }

  // Grava o novo status do pedido na planilha
  sheet.getRange(rowIdx, statusIndex + 1).setValue(novoStatus);
  if (dataPagamento !== undefined && dataPagamento !== null) {
    sheet.getRange(rowIdx, payDateIndex + 1).setValue(dataPagamento);
  }
  return { success: true };
}

function cancelarEtiquetaSuperfrete(packageId) {
  const sfConfig = obterConfiguracaoSuperFrete();
  const token = sfConfig.token;
  const baseUrl = sfConfig.baseUrl;
  
  if (!token) {
    return { success: false, error: "Token da SuperFrete não configurado nas propriedades do script." };
  }

  const apiHeaders = {
    "Authorization": "Bearer " + token,
    "Content-Type": "application/json",
    "Accept": "application/json",
    "User-Agent": "FigTree (andre.figueira@gmail.com)"
  };

  const payload = {
    "order": {
      "id": String(packageId),
      "description": "Meu cliente desistiu da compra"
    }
  };

  const options = {
    "method": "post",
    "headers": apiHeaders,
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(baseUrl + "/api/v0/order/cancel", options);
    const responseText = response.getContentText();
    const responseCode = response.getResponseCode();
    
    let data;
    try {
      data = JSON.parse(responseText);
    } catch (e) {
      data = {};
    }

    const isCanceled = (data[packageId] && data[packageId].canceled) || data.success || data.canceled;
    if (responseCode === 200 && (isCanceled || data.success === undefined)) {
      return { success: true };
    } else {
      let details = "";
      if (data.errors) {
        details = " Detalhes: " + JSON.stringify(data.errors);
      }
      const errorMsg = data.message || data.error || "Erro desconhecido";
      return { success: false, error: errorMsg + details };
    }
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// DADOS FIXOS DO REMETENTE (VENDEDOR) PARA A SUPERFRETE
const SELLER_NAME = "Andre Figueira";
const SELLER_CEP = "20541280"; // CEP correto da loja
const SELLER_ADDRESS = "Rua Ferreira Pontes";
const SELLER_LOCATION_NUMBER = "286";
const SELLER_COMPLEMENT = "Apt. 908 BL. A";
const SELLER_DISTRICT = "Andaraí";
const SELLER_CITY = "Rio de Janeiro";
const SELLER_STATE = "RJ";

// Mapeia o nome do metodo de frete na planilha para o ID correspondente da SuperFrete
function getServiceId(metodo) {
  const m = String(metodo).toLowerCase();
  if (m.includes("sedex")) return 2;
  if (m.includes("mini") || m.includes("envio")) return 17;
  if (m.includes("pac")) return 1;
  if (m.includes("jadlog")) return 3;
  if (m.includes("loggi")) return 31;
  return 1; // Fallback padrao: PAC
}

// Analisa e quebra o endereço livre digitado nas partes exigidas pela SuperFrete
function parseEndereco(enderecoStr, cep) {
  const parts = String(enderecoStr).split(/\s*-\s*/);
  
  let logradouroNum = parts[0].trim();
  let complemento = "";
  let bairro = "Centro";
  let cidade = "Cidade";
  let uf = "SP";

  const ufList = ["AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS","MG","PA","PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC","SP","SE","TO"];
  
  const lastPart = parts[parts.length - 1].trim();
  const isLastPartUf = ufList.includes(lastPart.toUpperCase());

  if (isLastPartUf) {
    uf = lastPart.toUpperCase();
    if (parts.length >= 3) {
      cidade = parts[parts.length - 2].trim();
      if (parts.length >= 4) {
        bairro = parts[parts.length - 3].trim();
        if (parts.length > 4) {
          complemento = parts.slice(1, parts.length - 3).map(p => p.trim()).join(" - ");
        }
      } else {
        bairro = "Centro";
      }
    } else if (parts.length === 2) {
      cidade = "Cidade";
      bairro = "Centro";
    }
  } else {
    // Caso em que a última parte contém Cidade/UF ou semelhante
    const ufParts = lastPart.split(/\s*[\/\-]\s*/);
    if (ufParts.length >= 2) {
      cidade = ufParts[0].trim();
      uf = ufParts[1].trim().toUpperCase().slice(0, 2);
    } else {
      cidade = lastPart;
      uf = "SP"; // fallback
    }

    if (parts.length >= 3) {
      bairro = parts[parts.length - 2].trim();
      if (parts.length > 3) {
        complemento = parts.slice(1, parts.length - 2).map(p => p.trim()).join(" - ");
      }
    } else {
      bairro = "Centro";
    }
  }

  // Parsear logradouroNum para separar rua, numero e extraComplemento
  let rua = logradouroNum;
  let numero = "S/N";
  let extraComplemento = "";
  
  const commaParts = logradouroNum.split(/\s*,\s*/);
  if (commaParts.length >= 2) {
    rua = commaParts[0].trim();
    numero = commaParts[1].trim();
    if (commaParts.length >= 3) {
      extraComplemento = commaParts.slice(2).join(", ").trim();
    }
  } else {
    // Regex para extrair opcionalmente "Nº", "N.", "No", etc. seguido do número no final
    const match = logradouroNum.match(/(.*?)\s*(?:N[º°º\.]*|N\.?o\.?|Número|Numero)?\s*(\d+)$/i);
    if (match) {
      rua = match[1].trim();
      numero = match[2].trim();
    }
  }
  
  // Limpar campo de número: se tiver SN ou S/N ou Sem número, mantém "S/N"
  let cleanNumero = String(numero).trim();
  if (cleanNumero.toLowerCase().includes("s/n") || 
      cleanNumero.toLowerCase().includes("sn") || 
      cleanNumero.toLowerCase().includes("sem") || 
      cleanNumero === "") {
    cleanNumero = "S/N";
  } else {
    // Remove prefixos numéricos e deixa só dígitos
    cleanNumero = cleanNumero.replace(/(?:N[º°º\.]*|N\.?o\.?|Número|Numero|#)/gi, "").trim();
    const digitMatch = cleanNumero.match(/\d+/);
    if (digitMatch) {
      cleanNumero = digitMatch[0];
    } else {
      cleanNumero = "S/N";
    }
  }
  
  // Combinar complementos do split por hifens e do split por vírgula
  let finalComplemento = complemento;
  if (extraComplemento) {
    finalComplemento = finalComplemento ? finalComplemento + " - " + extraComplemento : extraComplemento;
  }
  
  return {
    address: rua,
    location_number: cleanNumero,
    complement: finalComplemento || undefined,
    district: bairro,
    city: cidade,
    state_abbr: uf,
    postal_code: String(cep).replace(/\D/g, "")
  };
}

function emitirEtiquetaPedido(orderId) {
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  let sheet = ss.getSheetByName("Pedidos");
  if (!sheet) {
    return { success: false, error: "Planilha não encontrada" };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (parseInt(data[i][0]) === parseInt(orderId)) {
      rowIdx = i + 1; // 1-based index do getRange
      break;
    }
  }

  if (rowIdx === -1) {
    return { success: false, error: "Pedido não encontrado" };
  }

  // Mapeia os dados da linha para um objeto
  const pedido = {};
  headers.forEach((header, index) => {
    pedido[header] = data[rowIdx - 1][index];
  });

  const qty = parseInt(pedido.Qtd_Total) || 1;
  const caixa = selecionarCaixaLogistica(ss, qty);

  const rastreioIndex = headers.indexOf("Codigo_Rastreio");
  const pdfIndex = headers.indexOf("Etiqueta_PDF_URL");
  const statusIndex = headers.indexOf("Status");
  const packageIdIndex = headers.indexOf("Superfrete_Package_ID");
  const fretePagoIndex = headers.indexOf("Valor_Frete_Pago");
  const descontoFreteIndex = headers.indexOf("Desconto_Frete");
  const dataEmissaoIndex = headers.indexOf("Data_Emissao_Etiqueta");

  // Obter e limpar o número de entrega
  let cleanNumero = String(pedido.Entrega_Numero || "S/N").trim();
  if (cleanNumero.toLowerCase().includes("s/n") || 
      cleanNumero.toLowerCase().includes("sn") || 
      cleanNumero.toLowerCase().includes("sem") || 
      cleanNumero === "") {
    cleanNumero = "S/N";
  } else {
    cleanNumero = cleanNumero.replace(/(?:N[º°º\.]*|N\.?o\.?|Número|Numero|#)/gi, "").trim();
    const digitMatch = cleanNumero.match(/\d+/);
    if (digitMatch) {
      cleanNumero = digitMatch[0];
    } else {
      cleanNumero = "S/N";
    }
  }

  // Prepara as informações de entrega consolidadas para a API da SuperFrete
  const dest = {
    address: pedido.Entrega_Rua || "",
    location_number: cleanNumero,
    district: pedido.Entrega_Bairro || "",
    city: pedido.Entrega_Cidade || "",
    state_abbr: pedido.Entrega_Estado || "",
    postal_code: String(pedido.Entrega_CEP || "").replace(/\D/g, "")
  };

  if (pedido.Entrega_Complemento) {
    dest.complement = String(pedido.Entrega_Complemento).trim();
  }

  // Obtém as credenciais e URL base de acordo com o ambiente (sandbox ou produção)
  const sfConfig = obterConfiguracaoSuperFrete();
  const token = sfConfig.token;
  const baseUrl = sfConfig.baseUrl;
  
  if (!token) {
    return { success: false, error: "Token de Produção da SuperFrete não configurado nas propriedades do script." };
  }

  const apiHeaders = {
    "Authorization": "Bearer " + token,
    "Content-Type": "application/json",
    "Accept": "application/json",
    "User-Agent": "FigTree (andre.figueira@gmail.com)"
  };

  // PASSO 1: Enviar frete para o carrinho da SuperFrete
  const payloadCart = {
    "from": {
      "name": SELLER_NAME,
      "address": SELLER_ADDRESS,
      "number": SELLER_LOCATION_NUMBER,
      "complement": SELLER_COMPLEMENT,
      "district": SELLER_DISTRICT,
      "state_abbr": SELLER_STATE,
      "postal_code": SELLER_CEP.replace(/\D/g, ""),
      "city": SELLER_CITY
    },
    "to": {
      "name": pedido.Nome,
      "address": dest.address,
      "number": dest.location_number,
      "district": dest.district,
      "city": dest.city,
      "state_abbr": dest.state_abbr,
      "postal_code": dest.postal_code,
      "document": String(pedido.CPF).replace(/\D/g, "")
    },
    "service": getServiceId(pedido.Frete_Metodo),
    "products": [
      {
        "name": "Memo Book 14x9",
        "quantity": parseInt(pedido.Qtd_Total) || 1,
        "unitary_value": (parseFloat(pedido.Valor_Produtos) || 50.00) / (parseInt(pedido.Qtd_Total) || 1)
      }
    ],
    "volumes": {
      "format": 1, // 1 = Caixa/Pacote (matching calculations from Grade_Caixas)
      "height": caixa.altura,
      "width": caixa.largura,
      "length": caixa.comprimento,
      "weight": (qty * 0.03567) + caixa.peso
    },
    "platform": "MemoBook14x9"
  };

  if (dest.complement) {
    payloadCart.to.complement = dest.complement;
  }

  const optCart = {
    "method": "post",
    "headers": apiHeaders,
    "payload": JSON.stringify(payloadCart),
    "muteHttpExceptions": true
  };

  const resCart = UrlFetchApp.fetch(baseUrl + "/api/v0/cart", optCart);
  const txtCart = resCart.getContentText();
  const dataCart = JSON.parse(txtCart);

  if (resCart.getResponseCode() !== 200 || !dataCart.id) {
    let details = "";
    if (dataCart.errors) {
      details = " Detalhes: " + JSON.stringify(dataCart.errors);
    }
    return { success: false, error: "Erro ao adicionar ao carrinho da SuperFrete: " + (dataCart.message || "") + details + " (Resposta: " + txtCart + ")" };
  }
  const packageId = dataCart.id;

  // PASSO 2: Finalizar pedido e gerar etiqueta (Checkout)
  const payloadCheckout = {
    "orders": [packageId]
  };

  const optCheckout = {
    "method": "post",
    "headers": apiHeaders,
    "payload": JSON.stringify(payloadCheckout),
    "muteHttpExceptions": true
  };

  const resCheckout = UrlFetchApp.fetch(baseUrl + "/api/v0/checkout", optCheckout);
  const txtCheckout = resCheckout.getContentText();
  const dataCheckout = JSON.parse(txtCheckout);

  if (resCheckout.getResponseCode() !== 200 || !dataCheckout.success) {
    let details = "";
    if (dataCheckout.errors) {
      details = " Detalhes: " + JSON.stringify(dataCheckout.errors);
    }
    
    // VERIFICAR SE O ERRO É FALTA DE SALDO
    const isSemSaldo = txtCheckout.indexOf("Sem saldo na carteira") !== -1;
    if (isSemSaldo) {
      const dataHoraEmissao = new Date();
      
      // Grava o status do pedido como "Etiqueta Gerada"
      sheet.getRange(rowIdx, statusIndex + 1).setValue("Etiqueta Gerada");
      
      // Grava o packageId e a data de emissão na planilha
      if (packageIdIndex !== -1) sheet.getRange(rowIdx, packageIdIndex + 1).setValue(packageId);
      if (dataEmissaoIndex !== -1) sheet.getRange(rowIdx, dataEmissaoIndex + 1).setValue(dataHoraEmissao);
      
      // Grava o status logístico da SuperFrete como "pending"
      const sfStatusIndex = headers.indexOf("Superfrete_Status");
      if (sfStatusIndex !== -1) sheet.getRange(rowIdx, sfStatusIndex + 1).setValue("pending");
      
      // O PDF deve ficar vazio de acordo com o pedido do usuário
      if (pdfIndex !== -1) sheet.getRange(rowIdx, pdfIndex + 1).setValue("");
      
      // Envia notificação informando a falta de saldo
      try {
        const titulo = "⚠️ Etiqueta Pendente de Saldo (Pedido #" + orderId + ")";
        const mensagem = "Tentativa de gerar etiqueta para o pedido #" + orderId + " falhou por falta de saldo.\n\n" +
                         "O pedido foi agendado como PENDENTE no carrinho da SuperFrete.\n\n" +
                         "• Adicione saldo no painel da SuperFrete para emitir.";
        enviarNotificacaoTelegram(titulo, mensagem);
      } catch (errTelegram) {
        Logger.log("Erro ao enviar notificacao Telegram de falta de saldo: " + errTelegram.message);
      }

      // Aguarda 1.5 segundo e atualiza o rastreamento do pedido para preencher as novas colunas de dimensão e prazos
      Utilities.sleep(1500);
      try {
        atualizarRastreamentoPedido(orderId);
      } catch (err) {
        Logger.log("Erro na auto-sincronização pós-emissão sem saldo: " + err.message);
      }

      return { 
        success: true, 
        isPendingPayment: true, 
        packageId: packageId,
        labelPdfUrl: "",
        trackingCode: "", 
        message: "Sem saldo na carteira da SuperFrete." 
      };
    }
    
    return { success: false, error: "Erro no Checkout da SuperFrete: " + (dataCheckout.message || "") + details + " (Resposta: " + txtCheckout + ")" };
  }

  const purchase = dataCheckout.purchase;
  const orderResult = purchase.orders[0];
  let trackingCode = orderResult.tracking || "";

  // Se o rastreio não veio na resposta do checkout, consultamos o endpoint de info do pedido com polling
  if (!trackingCode) {
    let attempts = 0;
    while (!trackingCode && attempts < 3) {
      Utilities.sleep(2500); // Aguarda 2.5s para processamento
      attempts++;
      try {
        const optInfo = {
          "method": "get",
          "headers": apiHeaders,
          "muteHttpExceptions": true
        };
        const resInfo = UrlFetchApp.fetch(baseUrl + "/api/v0/order/info/" + packageId, optInfo);
        if (resInfo.getResponseCode() === 200) {
          const dataInfo = JSON.parse(resInfo.getContentText());
          if (dataInfo.tracking) {
            trackingCode = dataInfo.tracking;
            break;
          }
        }
      } catch (err) {
        Logger.log("Erro na tentativa " + attempts + " de obter rastreio: " + err.message);
      }
    }
  }

  // PASSO 3: Gerar link de impressão
  const payloadPrint = {
    "orders": [packageId]
  };

  const optPrint = {
    "method": "post",
    "headers": apiHeaders,
    "payload": JSON.stringify(payloadPrint),
    "muteHttpExceptions": true
  };

  const resPrint = UrlFetchApp.fetch(baseUrl + "/api/v0/tag/print", optPrint);
  const txtPrint = resPrint.getContentText();
  const dataPrint = JSON.parse(txtPrint);

  if (resPrint.getResponseCode() !== 200 || !dataPrint.url) {
    let details = "";
    if (dataPrint.errors) {
      details = " Detalhes: " + JSON.stringify(dataPrint.errors);
    }
    return { success: false, error: "Erro ao gerar link da etiqueta da SuperFrete: " + (dataPrint.message || "") + details + " (Resposta: " + txtPrint + ")" };
  }

  // Construímos a URL final estável em Base64 usando o packageId da ordem,
  // pois a URL curta temporária da SuperFrete costuma dar erro de processamento imediato (Internal Server Error)
  const base64Id = Utilities.base64Encode(JSON.stringify({ order_id: packageId }), Utilities.Charset.UTF_8);
  const labelPdfUrl = "https://etiqueta.superfrete.com/_etiqueta/pdf/" + base64Id + "?format=A6";

  // Prepara os valores adicionais de retorno da SuperFrete
  const fretePagoReal = orderResult.price !== undefined ? parseFloat(orderResult.price) : 0;
  const descontoFreteReal = orderResult.discount !== undefined ? parseFloat(orderResult.discount) : 0;
  const dataHoraEmissao = new Date();

  // Grava os dados na planilha
  sheet.getRange(rowIdx, rastreioIndex + 1).setValue(trackingCode);
  sheet.getRange(rowIdx, pdfIndex + 1).setValue(labelPdfUrl);
  sheet.getRange(rowIdx, statusIndex + 1).setValue("Etiqueta Gerada");
  
  if (packageIdIndex !== -1) sheet.getRange(rowIdx, packageIdIndex + 1).setValue(packageId);
  if (fretePagoIndex !== -1) sheet.getRange(rowIdx, fretePagoIndex + 1).setValue(fretePagoReal);
  if (descontoFreteIndex !== -1) sheet.getRange(rowIdx, descontoFreteIndex + 1).setValue(descontoFreteReal);
  if (dataEmissaoIndex !== -1) sheet.getRange(rowIdx, dataEmissaoIndex + 1).setValue(dataHoraEmissao);

  // Envia notificação Telegram de sucesso
  try {
    const titulo = "🎫 Etiqueta Gerada (Pedido #" + orderId + ")";
    const mensagem = "Etiqueta gerada com sucesso para o pedido #" + orderId + "!\n\n" +
                     "• Método Frete: " + (pedido.Frete_Metodo || "Não informado") + "\n" +
                     "• Rastreamento: " + trackingCode + "\n" +
                     "• PDF da Etiqueta: " + labelPdfUrl;
    enviarNotificacaoTelegram(titulo, mensagem);
  } catch (errTelegram) {
    Logger.log("Erro ao enviar notificacao Telegram de etiqueta gerada: " + errTelegram.message);
  }

  // Aguarda 3 segundos e atualiza o rastreamento do pedido para preencher as novas colunas de dimensão e prazos
  Utilities.sleep(3000);
  try {
    atualizarRastreamentoPedido(orderId);
  } catch (err) {
    Logger.log("Erro na auto-sincronização pós-emissão: " + err.message);
  }

  return { 
    success: true, 
    trackingCode: trackingCode, 
    labelPdfUrl: labelPdfUrl 
  };
}

function conciliarRecibos() {
  const folderId = "1RSL54bRCxiveF9cdx6cjmlsMNg2pg8Ej";
  const folder = DriveApp.getFolderById(folderId);

  const files = folder.getFiles();
  const resultados = [];

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    
    // Ignoramos arquivos que não sejam PDF
    if (!fileName.toLowerCase().endsWith(".pdf")) {
      continue;
    }

    // Extrair o ID do pedido do nome do arquivo (ex.: "1022.pdf" -> 1022)
    const matchId = fileName.match(/^(\d+)\.pdf$/i);
    if (!matchId) {
      resultados.push({ file: fileName, success: false, reason: "Nome do arquivo não contém apenas números (ex. 1022.pdf)" });
      continue;
    }

    const orderId = parseInt(matchId[1]);

    // Verificar se o pedido existe e qual o status atual dele
    const pedidoStatus = getPedidoStatusLocal(orderId);
    if (!pedidoStatus.existe) {
      resultados.push({ file: fileName, success: false, reason: "Pedido #" + orderId + " não encontrado na planilha" });
      continue;
    }

    // LOGICA DE ATUALIZACAO DE STATUS VIA WEBHOOK
    if (sfStatus === "posted") {
      // Se postado nos correios, o status do painel deve ser "Enviado"
      const statusFinais = ["enviado", "em transito", "em trânsito", "recebido (finalizado)", "recebido"];
      if (!statusFinais.includes(String(currentStatus).toLowerCase())) {
        newPanelStatus = "Enviado";
      }
    } else if (sfStatus === "delivered") {
      // Se entregue, o status do painel deve ser "Recebido (Finalizado)"
      const statusFinais = ["recebido (finalizado)", "recebido"];
      if (!statusFinais.includes(String(currentStatus).toLowerCase())) {
        newPanelStatus = "Recebido (Finalizado)";
      }
    } else if (sfStatus === "released" || sfStatus === "printed") {
      // Se liberada ou impressa na API, o status do painel deve ser "Etiqueta Gerada"
      const statusFinais = ["etiqueta gerada", "enviado", "em transito", "em trânsito", "recebido (finalizado)", "recebido"];
      if (!statusFinais.includes(String(currentStatus).toLowerCase())) {
        newPanelStatus = "Etiqueta Gerada";
      }
    } else if (sfStatus === "canceled" || sfStatus === "cancelled") {
      // Se cancelado na API, e o status do painel for logístico ativo, resetamos
      const statusAtivos = ["pronto para envio", "etiqueta gerada", "enviado", "em transito", "em trânsito"];
      if (statusAtivos.includes(String(currentStatus).toLowerCase())) {
        newPanelStatus = "Pago"; // Volta para "Pago" para permitir gerar nova etiqueta
      }
    }

    const statusLower = String(pedidoStatus.status).toLowerCase();
    const statusPagos = ["pago", "em produção", "em producao", "pronto para envio", "etiqueta gerada", "enviado", "em transito", "em trânsito", "recebido (finalizado)", "recebido"];
    if (statusPagos.includes(statusLower)) {
      resultados.push({ file: fileName, success: false, reason: "Pedido #" + orderId + " já está marcado como pago ou enviado" });
      try {
        moverArquivoParaProcessados(folder, file);
      } catch (err) {}
      continue;
    }

    // Se está pendente, rodamos o OCR
    try {
      const text = extrairTextoPDFDriveAPI(file.getId());
      
      // Buscar a data e hora de pagamento usando a Regex
      const matchData = text.match(/(\d{2})\/([a-zA-Z]{3}|\d{2})\/(\d{4})(?:\s*-\s*(\d{2}:\d{2}))?/);
      let dataPagamento = null;
      if (matchData) {
        const dia = matchData[1];
        let mes = matchData[2];
        if (isNaN(parseInt(mes))) {
          mes = converterMesBrParaNumero(mes);
        }
        const ano = matchData[3];
        const hora = matchData[4] || "";
        dataPagamento = `${dia}/${mes}/${ano}` + (hora ? ` ${hora}` : "");
      } else {
        const hoje = new Date();
        const dia = String(hoje.getDate()).padStart(2, '0');
        const mes = String(hoje.getMonth() + 1).padStart(2, '0');
        const ano = hoje.getFullYear();
        dataPagamento = `${dia}/${mes}/${ano}`;
      }

      // Atualizar a planilha
      const resUpdate = atualizarStatusPedido(orderId, "Pago", dataPagamento);
      if (resUpdate.success) {
        moverArquivoParaProcessados(folder, file);
        resultados.push({ file: fileName, success: true, orderId: orderId, dataPagamento: dataPagamento });
      } else {
        resultados.push({ file: fileName, success: false, reason: "Erro ao salvar na planilha: " + resUpdate.error });
      }

    } catch (err) {
      resultados.push({ file: fileName, success: false, reason: "Erro no processamento de OCR: " + err.message });
    }
  }

  return { success: true, resultados: resultados };
}

function getPedidoStatusLocal(orderId) {
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  const sheet = ss.getSheetByName("Pedidos");
  if (!sheet) return { existe: false };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const statusIndex = headers.indexOf("Status");

  for (let i = 1; i < data.length; i++) {
    if (parseInt(data[i][0]) === parseInt(orderId)) {
      return { existe: true, status: data[i][statusIndex] || "" };
    }
  }
  return { existe: false };
}

function converterMesBrParaNumero(mesStr) {
  const meses = {
    "jan": "01", "feb": "02", "fev": "02", "mar": "03", "apr": "04", "abr": "04",
    "may": "05", "mai": "05", "jun": "06", "jul": "07", "aug": "08", "ago": "08",
    "sep": "09", "set": "09", "oct": "10", "out": "10", "nov": "11", "dec": "12", "dez": "12"
  };
  const clean = String(mesStr).toLowerCase().substring(0, 3);
  return meses[clean] || "01";
}

function extrairTextoPDFDriveAPI(fileId) {
  const token = ScriptApp.getOAuthToken();
  const copyUrl = "https://www.googleapis.com/drive/v2/files/" + fileId + "/copy?ocr=true";
  const copyParams = {
    method: "POST",
    headers: {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify({
      title: "Temp_OCR_" + fileId,
      mimeType: "application/vnd.google-apps.document"
    }),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(copyUrl, copyParams);
  const responseData = JSON.parse(response.getContentText());
  
  if (response.getResponseCode() !== 200) {
    throw new Error("Falha no OCR do Drive API: " + response.getContentText());
  }
  
  const tempDocId = responseData.id;
  let text = "";
  try {
    const doc = DocumentApp.openById(tempDocId);
    text = doc.getBody().getText();
  } catch (e) {
    Logger.log("Erro ao ler o documento temporário: " + e.message);
  }
  
  try {
    const deleteUrl = "https://www.googleapis.com/drive/v2/files/" + tempDocId;
    const deleteParams = {
      method: "DELETE",
      headers: {
        "Authorization": "Bearer " + token
      },
      muteHttpExceptions: true
    };
    UrlFetchApp.fetch(deleteUrl, deleteParams);
  } catch (e) {
    Logger.log("Erro ao deletar o documento temporário: " + e.message);
  }
  
  return text;
}

function moverArquivoParaProcessados(folder, file) {
  let processadosFolder;
  const folders = folder.getFoldersByName("Processados");
  if (folders.hasNext()) {
    processadosFolder = folders.next();
  } else {
    processadosFolder = folder.createFolder("Processados");
  }
  file.moveTo(processadosFolder);
}

function doGet(e) {
  // Força o Google Apps Script a acionar a tela de consentimento de DriveApp de forma nativa na aba do navegador
  const folder = DriveApp.getFolderById("1RSL54bRCxiveF9cdx6cjmlsMNg2pg8Ej");
  return HtmlService.createHtmlOutput("<h1>Autorizado com Sucesso!</h1><p>O seu painel já tem acesso ao Google Drive para a conciliação de recibos Pix. Pode fechar esta aba e testar o painel administrativo.</p>");
}

function reordenarColunasPlanilha() {
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  let sheet = ss.getSheetByName("Pedidos");
  if (!sheet || sheet.getLastRow() === 0) return { success: false, error: "Planilha vazia ou inexistente" };

  const novaOrdem = [
    "ID_Pedido", "Status", "Data_Pedido", "Data_Pagamento", "Nome", "WhatsApp", "CPF",
    "Entrega_Rua", "Entrega_Numero", "Entrega_Complemento", "Entrega_Bairro", "Entrega_Cidade", "Entrega_Estado", "Entrega_CEP",
    "Itens", "Qtd_Total", "Valor_Produtos", "Cupom_Codigo", "Cupom_Desconto_Valor", "Frete_Metodo", "Valor_Frete", "Total_Geral", "Observacoes", "Pix_Payload",
    "Superfrete_Package_ID", "Superfrete_Status", "Codigo_Rastreio", "Etiqueta_PDF_URL", "Data_Emissao_Etiqueta", "Valor_Frete_Pago", "Desconto_Frete",
    "Superfrete_Custo_Estimado", "Superfrete_Desconto_Estimado",
    "Superfrete_Prazo_Entrega", "Data_Postagem", "Data_Entrega",
    "Superfrete_Formato", "Superfrete_Peso", "Superfrete_Altura", "Superfrete_Largura", "Superfrete_Comprimento",
    "Superfrete_Mao_Propria", "Superfrete_Aviso_Recebimento", "Superfrete_Valor_Declarado"
  ];

  const rawData = sheet.getDataRange().getValues();
  const headersAtuais = rawData[0];
  const dadosAtuais = rawData.slice(1);

  const idxEnderecoAntigo = headersAtuais.indexOf("Entrega_Endereco");
  const idxCep = headersAtuais.indexOf("Entrega_CEP");
  const idxDataAntigo = headersAtuais.indexOf("Data");

  const novasLinhas = [novaOrdem];

  dadosAtuais.forEach(linha => {
    // Objeto temporário contendo dados mapeados por nome de coluna
    const valoresMapeados = {};
    
    headersAtuais.forEach((header, idx) => {
      if (header) {
        valoresMapeados[header] = linha[idx];
      }
    });

    // Renomear a coluna Data
    if (idxDataAntigo !== -1 && valoresMapeados["Data"] !== undefined) {
      valoresMapeados["Data_Pedido"] = valoresMapeados["Data"];
    }

    // Se a coluna antiga de endereço existia e as novas ainda não estão preenchidas
    if (idxEnderecoAntigo !== -1 && valoresMapeados["Entrega_Endereco"]) {
      const cepVal = idxCep !== -1 ? valoresMapeados["Entrega_CEP"] : "";
      try {
        const destParsed = parseEndereco(valoresMapeados["Entrega_Endereco"], cepVal);
        if (!valoresMapeados["Entrega_Rua"]) valoresMapeados["Entrega_Rua"] = destParsed.address || "";
        if (!valoresMapeados["Entrega_Numero"]) valoresMapeados["Entrega_Numero"] = destParsed.location_number || "";
        if (!valoresMapeados["Entrega_Complemento"]) valoresMapeados["Entrega_Complemento"] = destParsed.complement || "";
        if (!valoresMapeados["Entrega_Bairro"]) valoresMapeados["Entrega_Bairro"] = destParsed.district || "";
        if (!valoresMapeados["Entrega_Cidade"]) valoresMapeados["Entrega_Cidade"] = destParsed.city || "";
        if (!valoresMapeados["Entrega_Estado"]) valoresMapeados["Entrega_Estado"] = destParsed.state_abbr || "";
      } catch (e) {
        Logger.log("Erro no parsing do endereço antigo para migração: " + e.message);
      }
    }

    // Construir a linha com base rigidamente no novo array de cabeçalhos
    const novaLinha = novaOrdem.map(colName => {
      const val = valoresMapeados[colName];
      return val !== undefined ? val : "";
    });
    
    novasLinhas.push(novaLinha);
  });

  // Limpa a planilha e grava a tabela reordenada
  sheet.clearContents();
  sheet.getRange(1, 1, novasLinhas.length, novaOrdem.length).setValues(novasLinhas);

  return { success: true };
}

function atualizarRastreamentoPedido(orderId) {
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  let sheet = ss.getSheetByName("Pedidos");
  if (!sheet) {
    return { success: false, error: "Planilha não encontrada" };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (parseInt(data[i][0]) === parseInt(orderId)) {
      rowIdx = i + 1;
      break;
    }
  }

  if (rowIdx === -1) {
    return { success: false, error: "Pedido não encontrado" };
  }

  // Mapeia os dados da linha para um objeto
  const pedido = {};
  headers.forEach((header, index) => {
    pedido[header] = data[rowIdx - 1][index];
  });

  let packageId = pedido.Superfrete_Package_ID;
  
  // Fallback: tentar extrair o packageId de Etiqueta_PDF_URL
  if (!packageId && pedido.Etiqueta_PDF_URL) {
    try {
      const pdfUrl = String(pedido.Etiqueta_PDF_URL);
      const match = pdfUrl.match(/_etiqueta\/pdf\/([^?]+)/);
      if (match) {
        const decoded = Utilities.newBlob(Utilities.base64Decode(match[1])).getDataAsString();
        const parsed = JSON.parse(decoded);
        if (parsed && parsed.order_id) {
          packageId = parsed.order_id;
        }
      }
    } catch (e) {
      Logger.log("Erro ao decodificar ID da URL PDF: " + e.message);
    }
  }

  if (!packageId) {
    return { success: false, error: "Este pedido não possui ID de pacote (Superfrete_Package_ID) nem URL de etiqueta decodificável." };
  }

  // Obtém as credenciais e URL base de acordo com o ambiente (sandbox ou produção)
  const sfConfig = obterConfiguracaoSuperFrete();
  const token = sfConfig.token;
  const baseUrl = sfConfig.baseUrl;

  if (!token) {
    return { success: false, error: "Token de Produção da SuperFrete não configurado." };
  }

  const apiHeaders = {
    "Authorization": "Bearer " + token,
    "Accept": "application/json",
    "User-Agent": "FigTree (andre.figueira@gmail.com)"
  };

  const options = {
    "method": "get",
    "headers": apiHeaders,
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(baseUrl + "/api/v0/order/info/" + packageId, options);
    const responseText = response.getContentText();
    const responseCode = response.getResponseCode();
    
    let resData;
    if (responseCode === 200) {
      resData = JSON.parse(responseText);
    } else if (responseCode === 404 || responseCode === 400) {
      // Se o pacote não existe ou foi removido da SuperFrete, tratamos pacificamente como cancelado
      resData = {
        status: "canceled",
        tracking: "",
        delivery: "",
        posted_at: "",
        delivered_at: ""
      };
    } else {
      return { success: false, error: "Erro na API da SuperFrete (" + responseCode + "): " + responseText };
    }
    
    // Obter os índices das colunas no Sheets
    const trackingIndex = headers.indexOf("Codigo_Rastreio");
    const statusIndex = headers.indexOf("Status");
    const packageIdIndex = headers.indexOf("Superfrete_Package_ID");
    const pdfIndex = headers.indexOf("Etiqueta_PDF_URL");
    const dataEmissaoIndex = headers.indexOf("Data_Emissao_Etiqueta");
    const fretePagoIndex = headers.indexOf("Valor_Frete_Pago");
    const descontoFreteIndex = headers.indexOf("Desconto_Frete");
    const custoEstimadoIndex = headers.indexOf("Superfrete_Custo_Estimado");
    const descontoEstimadoIndex = headers.indexOf("Superfrete_Desconto_Estimado");
    const superfreteStatusIndex = headers.indexOf("Superfrete_Status");
    const prazoIndex = headers.indexOf("Superfrete_Prazo_Entrega");
    const postagemIndex = headers.indexOf("Data_Postagem");
    const entregaIndex = headers.indexOf("Data_Entrega");

    // Novos índices para dimensões físicas e opcionais
    const formatoIndex = headers.indexOf("Superfrete_Formato");
    const pesoIndex = headers.indexOf("Superfrete_Peso");
    const alturaIndex = headers.indexOf("Superfrete_Altura");
    const larguraIndex = headers.indexOf("Superfrete_Largura");
    const comprimentoIndex = headers.indexOf("Superfrete_Comprimento");
    const maoPropriaIndex = headers.indexOf("Superfrete_Mao_Propria");
    const avisoRecebimentoIndex = headers.indexOf("Superfrete_Aviso_Recebimento");
    const valorDeclaradoIndex = headers.indexOf("Superfrete_Valor_Declarado");

    // Criamos uma cópia dos valores originais da linha para atualizar tudo em lote (otimização de performance)
    const rowValues = data[rowIdx - 1].slice();

    // Salvar o packageId na planilha caso ele tenha sido obtido via fallback e não estivesse gravado
    if (packageIdIndex !== -1 && !pedido.Superfrete_Package_ID) {
      rowValues[packageIdIndex] = packageId;
    }

    // Atualizar Código de Rastreamento se veio da API e estava vazio
    if (trackingIndex !== -1 && resData.tracking) {
      rowValues[trackingIndex] = resData.tracking;
    }

    // Gravar os novos campos no Sheets
    const sfStatus = resData.status ? String(resData.status).toLowerCase().trim() : "";
    if (superfreteStatusIndex !== -1 && resData.status) {
      rowValues[superfreteStatusIndex] = resData.status;
    }
    if (prazoIndex !== -1 && resData.delivery !== undefined) {
      rowValues[prazoIndex] = resData.delivery;
    }

    // Sempre atualizar os valores estimados e descontos se recebidos da API
    if (custoEstimadoIndex !== -1 && resData.price !== undefined) {
      rowValues[custoEstimadoIndex] = parseFloat(resData.price);
    }
    if (descontoEstimadoIndex !== -1 && resData.discount !== undefined) {
      rowValues[descontoEstimadoIndex] = parseFloat(resData.discount);
    }

    // Se a etiqueta foi emitida e paga (status is released, printed, posted, delivered)
    const statusEmitidos = ["released", "printed", "posted", "delivered"];
    if (statusEmitidos.includes(sfStatus)) {
      // Atualizar PDF da Etiqueta se estiver vazio
      if (pdfIndex !== -1 && !rowValues[pdfIndex]) {
        let pdfUrl = "";
        if (resData.print && resData.print.url) {
          pdfUrl = resData.print.url;
          // Converter para link estável Base64
          try {
            const base64Id = Utilities.base64Encode(JSON.stringify({ order_id: packageId }), Utilities.Charset.UTF_8);
            pdfUrl = "https://etiqueta.superfrete.com/_etiqueta/pdf/" + base64Id + "?format=A6";
          } catch (e) {}
        }
        if (pdfUrl) {
          rowValues[pdfIndex] = pdfUrl;
        }
      }
      
      // Atualizar Data de Emissao se estiver vazia
      if (dataEmissaoIndex !== -1 && !rowValues[dataEmissaoIndex]) {
        rowValues[dataEmissaoIndex] = resData.generated_at || resData.created_at || new Date();
      }
      
      // Atualizar Valor Pago e Desconto se estiverem vazios ou zero
      if (fretePagoIndex !== -1 && (!rowValues[fretePagoIndex] || parseFloat(rowValues[fretePagoIndex]) === 0)) {
        rowValues[fretePagoIndex] = resData.price !== undefined ? parseFloat(resData.price) : 0;
      }
      if (descontoFreteIndex !== -1 && (!rowValues[descontoFreteIndex] || parseFloat(rowValues[descontoFreteIndex]) === 0)) {
        rowValues[descontoFreteIndex] = resData.discount !== undefined ? parseFloat(resData.discount) : 0;
      }
    }
    
    // Gravar ou limpar datas de Postagem e Entrega de acordo com o status atual da SuperFrete
    if (postagemIndex !== -1) {
      if (resData.posted_at) {
        rowValues[postagemIndex] = resData.posted_at;
      } else if ((sfStatus === "posted" || sfStatus === "delivered") && !rowValues[postagemIndex]) {
        rowValues[postagemIndex] = resData.updated_at || new Date();
      } else if (sfStatus === "released" || sfStatus === "printed" || sfStatus === "canceled" || sfStatus === "cancelled" || sfStatus === "") {
        rowValues[postagemIndex] = "";
      }
    }
    if (entregaIndex !== -1) {
      if (resData.delivered_at) {
        rowValues[entregaIndex] = resData.delivered_at;
      } else if (sfStatus === "delivered" && !rowValues[entregaIndex]) {
        rowValues[entregaIndex] = resData.updated_at || new Date();
      } else if (sfStatus !== "delivered") {
        rowValues[entregaIndex] = "";
      }
    }

    // Gravar dimensões físicas e opcionais recebidas da API
    if (formatoIndex !== -1 && resData.format !== undefined) {
      rowValues[formatoIndex] = resData.format;
    }
    if (pesoIndex !== -1 && resData.weight !== undefined) {
      rowValues[pesoIndex] = parseFloat(resData.weight);
    }
    if (alturaIndex !== -1 && resData.height !== undefined) {
      rowValues[alturaIndex] = parseFloat(resData.height);
    }
    if (larguraIndex !== -1 && resData.width !== undefined) {
      rowValues[larguraIndex] = parseFloat(resData.width);
    }
    if (comprimentoIndex !== -1 && resData.length !== undefined) {
      rowValues[comprimentoIndex] = parseFloat(resData.length);
    }
    if (maoPropriaIndex !== -1 && resData.own_hand !== undefined) {
      rowValues[maoPropriaIndex] = resData.own_hand;
    }
    if (avisoRecebimentoIndex !== -1 && resData.receipt !== undefined) {
      rowValues[avisoRecebimentoIndex] = resData.receipt;
    }
    if (valorDeclaradoIndex !== -1 && resData.insurance_value !== undefined) {
      rowValues[valorDeclaradoIndex] = parseFloat(resData.insurance_value);
    }

    // Regras de atualização de status do painel
    let newPanelStatus = null;
    const currentStatus = pedido.Status;

    if (sfStatus === "posted") {
      // Se postado nos correios, o status do painel deve ser "Enviado"
      const statusFinais = ["enviado", "em transito", "em trânsito", "recebido (finalizado)", "recebido"];
      if (!statusFinais.includes(String(currentStatus).toLowerCase())) {
        newPanelStatus = "Enviado";
      }
    } else if (sfStatus === "delivered") {
      // Se entregue, o status do painel deve ser "Recebido (Finalizado)"
      const statusFinais = ["recebido (finalizado)", "recebido"];
      if (!statusFinais.includes(String(currentStatus).toLowerCase())) {
        newPanelStatus = "Recebido (Finalizado)";
      }
    } else if (sfStatus === "canceled" || sfStatus === "cancelled") {
      // Se a etiqueta foi cancelada na SuperFrete, o status do pedido volta para "Pronto para Envio"
      // a menos que ele já esteja como "Cancelado (Desistência)"
      if (pedido.Status !== "Cancelado (Desistência)") {
        newPanelStatus = "Pronto para Envio";
      }
      
      // Apaga da planilha os valores das colunas de etiqueta e rastreamento
      if (trackingIndex !== -1) {
        rowValues[trackingIndex] = "";
        resData.tracking = "";
      }
      if (pdfIndex !== -1) rowValues[pdfIndex] = "";
      if (dataEmissaoIndex !== -1) rowValues[dataEmissaoIndex] = "";
      
      // Limpa também o ID do pacote para permitir gerar uma nova etiqueta futuramente
      if (packageIdIndex !== -1) rowValues[packageIdIndex] = "";
      
      // Limpa dados de custos e prazos associados
      if (prazoIndex !== -1) {
        rowValues[prazoIndex] = "";
        resData.delivery = "";
      }
      if (fretePagoIndex !== -1) rowValues[fretePagoIndex] = "";
      if (descontoFreteIndex !== -1) rowValues[descontoFreteIndex] = "";
      if (custoEstimadoIndex !== -1) rowValues[custoEstimadoIndex] = "";
      if (descontoEstimadoIndex !== -1) rowValues[descontoEstimadoIndex] = "";
    }

    if (newPanelStatus && statusIndex !== -1) {
      rowValues[statusIndex] = newPanelStatus;
    }

    // Salva todas as alterações na planilha de uma única vez em lote
    sheet.getRange(rowIdx, 1, 1, rowValues.length).setValues([rowValues]);

    return { 
      success: true, 
      orderId: orderId,
      tracking: resData.tracking || pedido.Codigo_Rastreio,
      pdfUrl: pdfIndex !== -1 ? rowValues[pdfIndex] : (resData.print ? resData.print.url : ""),
      superfreteStatus: resData.status,
      deliveryTime: resData.delivery,
      postedAt: resData.posted_at,
      deliveredAt: resData.delivered_at,
      updatedPanelStatus: newPanelStatus || currentStatus,
      hasPdf: !!(resData.tracking || pedido.Etiqueta_PDF_URL)
    };

  } catch (err) {
    return { success: false, error: "Exceção ao consultar API: " + err.message };
  }
}

function sincronizarTodosRastreamentos() {
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  let sheet = ss.getSheetByName("Pedidos");
  if (!sheet) {
    return { success: false, error: "Planilha não encontrada" };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const packageIdIndex = headers.indexOf("Superfrete_Package_ID");
  const pdfUrlIndex = headers.indexOf("Etiqueta_PDF_URL");
  const statusIndex = headers.indexOf("Status");
  
  if (statusIndex === -1) {
    return { success: false, error: "Coluna de Status não encontrada na planilha" };
  }

  const resultados = [];

  for (let i = 1; i < data.length; i++) {
    const orderId = data[i][0];
    const statusVal = String(data[i][statusIndex]).toLowerCase();
    
    // Ignorar se já estiver finalizado (Recebido) ou se for pendente de pagamento ou cancelado
    if (statusVal === "recebido (finalizado)" || statusVal === "recebido" || 
        statusVal.includes("cancelado") || statusVal.includes("pendente")) {
      continue;
    }

    let packageId = packageIdIndex !== -1 ? data[i][packageIdIndex] : "";
    let pdfUrl = pdfUrlIndex !== -1 ? data[i][pdfUrlIndex] : "";

    // Se tem packageId ou se tem uma etiqueta pdfUrl que possamos decodificar, sincroniza!
    let hasPackageId = !!packageId;
    if (!hasPackageId && pdfUrl && pdfUrl.includes("_etiqueta/pdf/")) {
      hasPackageId = true;
    }

    if (hasPackageId) {
      const res = atualizarRastreamentoPedido(orderId);
      resultados.push(res);
    }
  }

  return { success: true, resultados: resultados };
}

function obterPlanilhaCupons(ss) {
  let sheet = ss.getSheetByName("Cupons");
  if (!sheet) {
    sheet = ss.insertSheet("Cupons");
    const headers = ["Codigo", "Tipo_Desconto", "Tipo_Valor", "Valor", "Uso_Unico", "Usado", "Data_Expiracao", "Produtos_Especificos", "Ativo"];
    sheet.appendRow(headers);
    // Adiciona um cupom de exemplo inicial
    sheet.appendRow(["FIGTREE10", "carrinho", "percentual", 10, "Não", "Não", "", "Qualquer", "Sim"]);
  }
  
  // Auto-migração para garantir que Produtos_Especificos e Ativo existam
  let headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  if (headers.indexOf("Produtos_Especificos") === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue("Produtos_Especificos");
    headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  }
  
  if (headers.indexOf("Ativo") === -1) {
    const colAtivo = sheet.getLastColumn() + 1;
    sheet.getRange(1, colAtivo).setValue("Ativo");
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const range = sheet.getRange(2, colAtivo, lastRow - 1, 1);
      const values = [];
      for (let r = 2; r <= lastRow; r++) {
        values.push(["Sim"]);
      }
      range.setValues(values);
    }
  }
  
  return sheet;
}

function validarCupom(cupomCodigo) {
  if (!cupomCodigo) {
    return { success: false, error: "Código do cupom não fornecido." };
  }
  
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  
  const sheet = obterPlanilhaCupons(ss);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const codIdx = headers.indexOf("Codigo");
  const tipoDescIdx = headers.indexOf("Tipo_Desconto");
  const tipoValIdx = headers.indexOf("Tipo_Valor");
  const valorIdx = headers.indexOf("Valor");
  const usoUnicoIdx = headers.indexOf("Uso_Unico");
  const usadoIdx = headers.indexOf("Usado");
  const expIdx = headers.indexOf("Data_Expiracao");
  const prodIdx = headers.indexOf("Produtos_Especificos");
  const ativoIdx = headers.indexOf("Ativo");
  
  const query = String(cupomCodigo).trim().toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    const cod = String(data[i][codIdx]).trim().toLowerCase();
    if (cod === query) {
      // Validar se está ativo
      const ativo = ativoIdx !== -1 ? String(data[i][ativoIdx]).trim() : "Sim";
      if (ativo === "Não") {
        return { success: false, error: "Este cupom de desconto foi desativado." };
      }

      // Validar uso único
      const usoUnico = String(data[i][usoUnicoIdx]).trim();
      const usado = String(data[i][usadoIdx]).trim();
      if (usoUnico === "Sim" && usado === "Sim") {
        return { success: false, error: "Este cupom de uso único já foi utilizado." };
      }
      
      // Validar data de expiração
      const expStr = data[i][expIdx];
      if (expStr) {
        const expDate = new Date(expStr);
        if (!isNaN(expDate.getTime())) {
          const expDateZerada = new Date(expDate.getFullYear(), expDate.getMonth(), expDate.getDate());
          const hojeZerada = new Date();
          hojeZerada.setHours(0, 0, 0, 0);
          if (hojeZerada.getTime() > expDateZerada.getTime()) {
            if (ativoIdx !== -1) {
              sheet.getRange(i + 1, ativoIdx + 1).setValue("Não");
            }
            return { success: false, error: "Este cupom expirou em " + expDate.toLocaleDateString("pt-BR") + "." };
          }
        }
      }
      
      return {
        success: true,
        cupom: {
          codigo: data[i][codIdx],
          tipo_desconto: data[i][tipoDescIdx],
          tipo_valor: data[i][tipoValIdx],
          valor: parseFloat(data[i][valorIdx]) || 0,
          uso_unico: usoUnico,
          usado: usado,
          data_expiracao: expStr ? new Date(expStr).toISOString().split('T')[0] : "",
          produtos_especificos: data[i][prodIdx] || "",
          ativo: ativoIdx !== -1 ? String(data[i][ativoIdx]).trim() : "Sim"
        }
      };
    }
  }
  
  return { success: false, error: "Cupom inválido ou não encontrado." };
}

function getCupons() {
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  
  const sheet = obterPlanilhaCupons(ss);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const codIdx = headers.indexOf("Codigo");
  const expIdx = headers.indexOf("Data_Expiracao");
  const ativoIdx = headers.indexOf("Ativo");
  
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    let ativoVal = ativoIdx !== -1 ? String(row[ativoIdx]).trim() : "Sim";
    const expStr = expIdx !== -1 ? row[expIdx] : null;
    
    if (ativoVal !== "Não" && expStr) {
      try {
        const expDate = new Date(expStr);
        if (!isNaN(expDate.getTime())) {
          const expDateZerada = new Date(expDate.getFullYear(), expDate.getMonth(), expDate.getDate());
          const hojeZerada = new Date();
          hojeZerada.setHours(0, 0, 0, 0);
          if (hojeZerada.getTime() > expDateZerada.getTime()) {
            ativoVal = "Não";
            if (ativoIdx !== -1) {
              sheet.getRange(i + 1, ativoIdx + 1).setValue("Não");
            }
          }
        }
      } catch (e) {
        Logger.log("Erro ao inativar cupom expirado no getCupons: " + e.message);
      }
    }
    
    const cupomObj = {};
    headers.forEach((header, index) => {
      if (header === "Ativo") {
        cupomObj[header] = ativoVal;
      } else if (header === "Data_Expiracao" && row[index]) {
        try {
          const d = new Date(row[index]);
          if (!isNaN(d.getTime())) {
            cupomObj[header] = d.toISOString().split('T')[0];
          } else {
            cupomObj[header] = row[index];
          }
        } catch (e) {
          cupomObj[header] = row[index];
        }
      } else {
        cupomObj[header] = row[index];
      }
    });
    list.push(cupomObj);
  }
  
  return { success: true, cupons: list };
}

function salvarCupom(cupomData) {
  if (!cupomData || !cupomData.codigo) {
    return { success: false, error: "Código do cupom inválido." };
  }
  
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  
  const sheet = obterPlanilhaCupons(ss);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const codIdx = headers.indexOf("Codigo");
  const query = String(cupomData.codigo).trim().toUpperCase();
  
  const rowValues = headers.map(header => {
    switch (header) {
      case "Codigo": return query;
      case "Tipo_Desconto": return cupomData.tipo_desconto || "carrinho";
      case "Tipo_Valor": return cupomData.tipo_valor || "percentual";
      case "Valor": return parseFloat(cupomData.valor) || 0;
      case "Uso_Unico": return cupomData.uso_unico || "Não";
      case "Usado": return cupomData.usado || "Não";
      case "Data_Expiracao": 
        if (cupomData.data_expiracao) {
          const parts = cupomData.data_expiracao.split("-");
          if (parts.length === 3) {
            return new Date(parts[0], parts[1] - 1, parts[2]);
          }
          return new Date(cupomData.data_expiracao);
        }
        return "";
      case "Produtos_Especificos": return cupomData.produtos_especificos || "Qualquer";
      case "Ativo": return cupomData.ativo || "Sim";
      default: return "";
    }
  });
  
  let foundRowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    const cod = String(data[i][codIdx]).trim().toUpperCase();
    if (cod === query) {
      foundRowIdx = i + 1;
      break;
    }
  }
  
  if (foundRowIdx !== -1) {
    sheet.getRange(foundRowIdx, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  
  return { success: true };
}

function excluirCupom(cupomCodigo) {
  if (!cupomCodigo) {
    return { success: false, error: "Código do cupom inválido." };
  }
  
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  
  const sheet = obterPlanilhaCupons(ss);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const codIdx = headers.indexOf("Codigo");
  const query = String(cupomCodigo).trim().toUpperCase();
  
  for (let i = 1; i < data.length; i++) {
    const cod = String(data[i][codIdx]).trim().toUpperCase();
    if (cod === query) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  
  return { success: false, error: "Cupom não encontrado." };
}

function alternarStatusCupom(cupomCodigo) {
  if (!cupomCodigo) {
    return { success: false, error: "Código do cupom inválido." };
  }
  
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  
  const sheet = obterPlanilhaCupons(ss);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const codIdx = headers.indexOf("Codigo");
  const ativoIdx = headers.indexOf("Ativo");
  
  if (codIdx === -1 || ativoIdx === -1) {
    return { success: false, error: "Colunas da planilha de cupons não configuradas." };
  }
  
  const query = String(cupomCodigo).trim().toUpperCase();
  for (let i = 1; i < data.length; i++) {
    const cod = String(data[i][codIdx]).trim().toUpperCase();
    if (cod === query) {
      const statusAtual = String(data[i][ativoIdx]).trim();
      const novoStatus = (statusAtual === "Não") ? "Sim" : "Não";
      sheet.getRange(i + 1, ativoIdx + 1).setValue(novoStatus);
      return { success: true, novoStatus: novoStatus };
    }
  }
  
  return { success: false, error: "Cupom não encontrado para alteração de status." };
}

function obterPlanilhaPromocoes(ss) {
  let sheet = ss.getSheetByName("Promocoes");
  if (!sheet) {
    sheet = ss.insertSheet("Promocoes");
    const headers = ["Nome", "Capa", "Miolo", "Tipo_Valor", "Valor", "Ativa"];
    sheet.appendRow(headers);
    // Adiciona uma promoção de exemplo inativa
    sheet.appendRow(["Mês das Mães", "Capa 2", "Qualquer", "percentual", 20, "Não"]);
  }
  return sheet;
}

function getPromocoes() {
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  
  const sheet = obterPlanilhaPromocoes(ss);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const promoObj = {};
    headers.forEach((header, index) => {
      promoObj[header] = row[index];
    });
    list.push(promoObj);
  }
  
  return { success: true, promocoes: list };
}

function salvarPromocao(promoData) {
  if (!promoData || !promoData.nome) {
    return { success: false, error: "Nome da promoção inválido." };
  }
  
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  
  const sheet = obterPlanilhaPromocoes(ss);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const nomeIdx = headers.indexOf("Nome");
  const capaIdx = headers.indexOf("Capa");
  const mioloIdx = headers.indexOf("Miolo");
  
  const queryNome = String(promoData.nome).trim();
  const queryCapa = String(promoData.capa || "Qualquer").trim();
  const queryMiolo = String(promoData.miolo || "Qualquer").trim();
  
  const rowValues = headers.map(header => {
    switch (header) {
      case "Nome": return queryNome;
      case "Capa": return queryCapa;
      case "Miolo": return queryMiolo;
      case "Tipo_Valor": return promoData.tipo_valor || "percentual";
      case "Valor": return parseFloat(promoData.valor) || 0;
      case "Ativa": return promoData.ativa || "Sim";
      default: return "";
    }
  });
  
  let foundRowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    const nome = String(data[i][nomeIdx]).trim();
    const capa = String(data[i][capaIdx]).trim();
    const miolo = String(data[i][mioloIdx]).trim();
    
    if (nome.toLowerCase() === queryNome.toLowerCase()) {
      foundRowIdx = i + 1;
      break;
    }
    if (capa.toLowerCase() === queryCapa.toLowerCase() && miolo.toLowerCase() === queryMiolo.toLowerCase()) {
      foundRowIdx = i + 1;
      break;
    }
  }
  
  if (foundRowIdx !== -1) {
    sheet.getRange(foundRowIdx, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  
  return { success: true };
}

function excluirPromocao(promoNome) {
  if (!promoNome) {
    return { success: false, error: "Nome da promoção inválido." };
  }
  
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  
  const sheet = obterPlanilhaPromocoes(ss);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const nomeIdx = headers.indexOf("Nome");
  const query = String(promoNome).trim().toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    const nome = String(data[i][nomeIdx]).trim().toLowerCase();
    if (nome === query) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  
  return { success: false, error: "Promoção não encontrada." };
}

function alternarStatusPromocao(promoNome) {
  if (!promoNome) {
    return { success: false, error: "Nome da promoção inválido." };
  }
  
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  
  const sheet = obterPlanilhaPromocoes(ss);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const nomeIdx = headers.indexOf("Nome");
  const ativaIdx = headers.indexOf("Ativa");
  
  if (nomeIdx === -1 || ativaIdx === -1) {
    return { success: false, error: "Colunas da planilha de promoções não configuradas." };
  }
  
  const query = String(promoNome).trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    const nome = String(data[i][nomeIdx]).trim().toLowerCase();
    if (nome === query) {
      const statusAtual = String(data[i][ativaIdx]).trim();
      const novoStatus = (statusAtual === "Não") ? "Sim" : "Não";
      sheet.getRange(i + 1, ativaIdx + 1).setValue(novoStatus);
      return { success: true, novoStatus: novoStatus };
    }
  }
  
  return { success: false, error: "Promoção não encontrada para alteração de status." };
}

function obterPlanilhaInsumos(ss) {
  let sheet = ss.getSheetByName("Insumos");
  if (!sheet) {
    sheet = ss.insertSheet("Insumos");
    const headers = ["ID_Insumo", "Nome", "Unidade_Medida", "Custo_Unitario_Atual", "Estoque_Atual"];
    sheet.appendRow(headers);
    // Insere insumos iniciais padrão
    const defaultInsumos = [
      ["papel_a4_marfim", "Papel A4 75g/m² Marfim", "Folha", 0.0864, 0],
      ["papel_kraft_240g", "Papel Kraft 240g/m²", "Folha", 1.00, 0],
      ["grampos_26_8", "Grampos 26/8", "Unidade", 0.00316, 0],
      ["tinta", "Tinta / Impressão", "Unidade", 0.17, 0],
      ["caixa_p", "Caixa de Papelão Pardo P (Até 5 cadernos)", "Unidade", 1.618, 0],
      ["caixa_m", "Caixa de Papelão Pardo M (Até 10 cadernos)", "Unidade", 1.20, 0]
    ];
    defaultInsumos.forEach(row => sheet.appendRow(row));
  }
  return sheet;
}

function obterPlanilhaComprasInsumos(ss) {
  let sheet = ss.getSheetByName("Compras_Insumos");
  if (!sheet) {
    sheet = ss.insertSheet("Compras_Insumos");
    const headers = ["Data_Compra", "ID_Insumo", "Quantidade", "Valor_Total", "Custo_Unitario", "Fornecedor"];
    sheet.appendRow(headers);
    
    // Insere compras iniciais estimadas
    const dataHoje = new Date();
    const defaultCompras = [
      [dataHoje, "papel_a4_marfim", 500, 43.20, 0.0864, "Fornecedor Padrão"],
      [dataHoje, "papel_kraft_240g", 25, 25.00, 1.00, "Fornecedor Padrão"],
      [dataHoje, "grampos_26_8", 5000, 15.80, 0.00316, "Fornecedor Padrão"],
      [dataHoje, "caixa_p", 50, 80.90, 1.618, "Fornecedor Embalagens"],
      [dataHoje, "caixa_m", 50, 60.00, 1.20, "Fornecedor Embalagens"]
    ];
    defaultCompras.forEach(row => sheet.appendRow(row));
  }
  return sheet;
}

function obterPlanilhaFichaTecnica(ss) {
  let sheet = ss.getSheetByName("Ficha_Tecnica");
  if (!sheet) {
    sheet = ss.insertSheet("Ficha_Tecnica");
    const headers = ["ID_Produto", "ID_Insumo", "Quantidade_Consumida"];
    sheet.appendRow(headers);
    
    // Insere consumo padrão para o caderno_14x9
    const defaultFicha = [
      ["caderno_14x9", "papel_a4_marfim", 7.5],
      ["caderno_14x9", "papel_kraft_240g", 0.5],
      ["caderno_14x9", "grampos_26_8", 2.0],
      ["caderno_14x9", "tinta", 1.0]
    ];
    defaultFicha.forEach(row => sheet.appendRow(row));
  }
  return sheet;
}

function obterPlanilhaGradeCaixas(ss) {
  let sheet = ss.getSheetByName("Grade_Caixas");
  if (!sheet) {
    sheet = ss.insertSheet("Grade_Caixas");
    const headers = ["ID_Caixa", "ID_Insumo", "Capacidade_Max_Cadernos", "Comprimento_CM", "Largura_CM", "Altura_CM", "Peso_Vazio_KG"];
    sheet.appendRow(headers);
    
    // Insere as caixas P e M do usuário
    const defaultGrade = [
      ["caixa_p", "caixa_p", 5, 16, 11, 3, 0.05],
      ["caixa_m", "caixa_m", 10, 17, 12, 5, 0.08]
    ];
    defaultGrade.forEach(row => sheet.appendRow(row));
  }
  return sheet;
}

function recalcularCustoMedioInsumo(ss, idInsumo) {
  const sheetCompras = obterPlanilhaComprasInsumos(ss);
  const dataCompras = sheetCompras.getDataRange().getValues();
  const headersCompras = dataCompras[0];
  
  const idIdx = headersCompras.indexOf("ID_Insumo");
  const qtyIdx = headersCompras.indexOf("Quantidade");
  const totalIdx = headersCompras.indexOf("Valor_Total");
  
  let totalQty = 0;
  let totalValor = 0;
  
  for (let i = 1; i < dataCompras.length; i++) {
    const rowId = String(dataCompras[i][idIdx]).trim();
    if (rowId === idInsumo) {
      const q = parseFloat(dataCompras[i][qtyIdx]) || 0;
      const v = parseFloat(dataCompras[i][totalIdx]) || 0;
      totalQty += q;
      totalValor += v;
    }
  }
  
  if (totalQty === 0) {
    // Se não houver compras registradas para este insumo, não altera o custo atual
    return 0;
  }
  let custoMedio = totalValor / totalQty;
  
  // Atualiza na planilha Insumos
  const sheetInsumos = obterPlanilhaInsumos(ss);
  const dataInsumos = sheetInsumos.getDataRange().getValues();
  const headersInsumos = dataInsumos[0];
  const idInsumosIdx = headersInsumos.indexOf("ID_Insumo");
  const custoIdx = headersInsumos.indexOf("Custo_Unitario_Atual");
  
  for (let i = 1; i < dataInsumos.length; i++) {
    const insumoId = String(dataInsumos[i][idInsumosIdx]).trim();
    if (insumoId === idInsumo) {
      sheetInsumos.getRange(i + 1, custoIdx + 1).setValue(custoMedio);
      break;
    }
  }
  
  return custoMedio;
}

function getCustosData() {
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  
  const sheetInsumos = obterPlanilhaInsumos(ss);
  const sheetCompras = obterPlanilhaComprasInsumos(ss);
  const sheetFicha = obterPlanilhaFichaTecnica(ss);
  const sheetGrade = obterPlanilhaGradeCaixas(ss);
  
  const insumos = readSheetAsJSON(sheetInsumos);
  const compras = readSheetAsJSON(sheetCompras);
  const ficha = readSheetAsJSON(sheetFicha);
  const grade = readSheetAsJSON(sheetGrade);
  
  // Limpa o cache da grade de caixas ao consultar dados de custos para forçar a atualização se houver edições diretas na planilha
  try {
    CacheService.getScriptCache().remove("GRADE_CAIXAS_CACHE");
  } catch (e) {
    Logger.log("Erro ao limpar cache GRADE_CAIXAS_CACHE: " + e.message);
  }

  return {
    success: true,
    insumos: insumos,
    compras: compras,
    ficha: ficha,
    grade: grade
  };
}

function readSheetAsJSON(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const obj = {};
    headers.forEach((header, index) => {
      if (row[index] instanceof Date) {
        obj[header] = row[index].toISOString().split('T')[0];
      } else {
        obj[header] = row[index];
      }
    });
    list.push(obj);
  }
  return list;
}

function salvarInsumo(insumoData) {
  if (!insumoData || !insumoData.id_insumo) {
    return { success: false, error: "ID do insumo inválido." };
  }
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!ss) ss = SpreadsheetApp.openById(spreadsheetId);
  
  const sheet = obterPlanilhaInsumos(ss);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const idIdx = headers.indexOf("ID_Insumo");
  const query = String(insumoData.id_insumo).trim();
  
  const rowValues = headers.map(header => {
    switch (header) {
      case "ID_Insumo": return query;
      case "Nome": return insumoData.nome || "";
      case "Unidade_Medida": return insumoData.unidade_medida || "Unidade";
      case "Custo_Unitario_Atual": return parseFloat(insumoData.custo_unitario_atual) || 0;
      case "Estoque_Atual": return parseFloat(insumoData.estoque_atual) || 0;
      default: return "";
    }
  });
  
  let foundRowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]).trim() === query) {
      foundRowIdx = i + 1;
      break;
    }
  }
  
  if (foundRowIdx !== -1) {
    sheet.getRange(foundRowIdx, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  return { success: true };
}

function salvarCompraInsumo(compraData) {
  if (!compraData || !compraData.id_insumo) {
    return { success: false, error: "ID do insumo inválido para a compra." };
  }
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!ss) ss = SpreadsheetApp.openById(spreadsheetId);
  
  const sheet = obterPlanilhaComprasInsumos(ss);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  let dataVal = new Date();
  if (compraData.data_compra) {
    const parts = compraData.data_compra.split("-");
    if (parts.length === 3) {
      dataVal = new Date(parts[0], parts[1] - 1, parts[2]);
    } else {
      dataVal = new Date(compraData.data_compra);
    }
  }
  
  const total = parseFloat(compraData.valor_total) || 0;
  const qty = parseFloat(compraData.quantidade) || 1;
  const custoUnit = qty > 0 ? total / qty : 0;
  
  const rowValues = headers.map(header => {
    switch (header) {
      case "Data_Compra": return dataVal;
      case "ID_Insumo": return String(compraData.id_insumo).trim();
      case "Quantidade": return qty;
      case "Valor_Total": return total;
      case "Custo_Unitario": return custoUnit;
      case "Fornecedor": return compraData.fornecedor || "";
      default: return "";
    }
  });
  
  sheet.appendRow(rowValues);
  
  // Recalcula o custo unitário atual
  recalcularCustoMedioInsumo(ss, String(compraData.id_insumo).trim());
  
  return { success: true };
}

function salvarFichaTecnica(fichaData) {
  if (!fichaData || !fichaData.id_insumo) {
    return { success: false, error: "Dados da ficha técnica inválidos." };
  }
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!ss) ss = SpreadsheetApp.openById(spreadsheetId);
  
  const sheet = obterPlanilhaFichaTecnica(ss);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const prodIdx = headers.indexOf("ID_Produto");
  const insIdx = headers.indexOf("ID_Insumo");
  
  const queryProd = String(fichaData.id_produto || "caderno_14x9").trim();
  const queryIns = String(fichaData.id_insumo).trim();
  
  const rowValues = headers.map(header => {
    switch (header) {
      case "ID_Produto": return queryProd;
      case "ID_Insumo": return queryIns;
      case "Quantidade_Consumida": return parseFloat(fichaData.quantidade_consumida) || 0;
      default: return "";
    }
  });
  
  let foundRowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][prodIdx]).trim() === queryProd && String(data[i][insIdx]).trim() === queryIns) {
      foundRowIdx = i + 1;
      break;
    }
  }
  
  if (foundRowIdx !== -1) {
    sheet.getRange(foundRowIdx, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  return { success: true };
}

function excluirItemFichaTecnica(prodId, insumoId) {
  if (!insumoId) {
    return { success: false, error: "Insumo inválido." };
  }
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!ss) ss = SpreadsheetApp.openById(spreadsheetId);
  
  const sheet = obterPlanilhaFichaTecnica(ss);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const prodIdx = headers.indexOf("ID_Produto");
  const insIdx = headers.indexOf("ID_Insumo");
  
  const queryProd = String(prodId || "caderno_14x9").trim();
  const queryIns = String(insumoId).trim();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][prodIdx]).trim() === queryProd && String(data[i][insIdx]).trim() === queryIns) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: "Item de Ficha Técnica não encontrado." };
}

function salvarCaixa(caixaData) {
  if (!caixaData || !caixaData.id_caixa) {
    return { success: false, error: "ID da caixa inválido." };
  }
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!ss) ss = SpreadsheetApp.openById(spreadsheetId);
  
  const sheet = obterPlanilhaGradeCaixas(ss);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const idIdx = headers.indexOf("ID_Caixa");
  const query = String(caixaData.id_caixa).trim();
  
  const rowValues = headers.map(header => {
    switch (header) {
      case "ID_Caixa": return query;
      case "ID_Insumo": return String(caixaData.id_insumo || query).trim();
      case "Capacidade_Max_Cadernos": return parseInt(caixaData.capacidade_max_cadernos) || 1;
      case "Comprimento_CM": return parseFloat(caixaData.comprimento_cm) || 0;
      case "Largura_CM": return parseFloat(caixaData.largura_cm) || 0;
      case "Altura_CM": return parseFloat(caixaData.altura_cm) || 0;
      case "Peso_Vazio_KG": return parseFloat(caixaData.peso_vazio_kg) || 0;
      default: return "";
    }
  });
  
  let foundRowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]).trim() === query) {
      foundRowIdx = i + 1;
      break;
    }
  }
  
  if (foundRowIdx !== -1) {
    sheet.getRange(foundRowIdx, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  
  // Invalida o cache para que a alteração seja lida na próxima cotação/emissão
  try {
    CacheService.getScriptCache().remove("GRADE_CAIXAS_CACHE");
  } catch (e) {
    Logger.log("Erro ao limpar cache GRADE_CAIXAS_CACHE: " + e.message);
  }

  return { success: true };
}

function excluirCaixa(caixaId) {
  if (!caixaId) {
    return { success: false, error: "ID de caixa inválido." };
  }
  const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!ss) ss = SpreadsheetApp.openById(spreadsheetId);
  
  const sheet = obterPlanilhaGradeCaixas(ss);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const idIdx = headers.indexOf("ID_Caixa");
  const query = String(caixaId).trim();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]).trim() === query) {
      sheet.deleteRow(i + 1);
      
      // Invalida o cache para que a alteração seja lida na próxima cotação/emissão
      try {
        CacheService.getScriptCache().remove("GRADE_CAIXAS_CACHE");
      } catch (e) {
        Logger.log("Erro ao limpar cache GRADE_CAIXAS_CACHE: " + e.message);
      }

      return { success: true };
    }
  }
  return { success: false, error: "Caixa não encontrada na grade." };
}

function processarWebhookSuperFrete(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    return processarWebhookSuperFreteDirect(payload);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Erro ao parsear JSON: " + error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function processarWebhookSuperFreteDirect(payload) {
  try {
    const dataObj = payload.data || {};
    const packageId = dataObj.id ? String(dataObj.id).trim() : "";
    const trackingCode = dataObj.tracking ? String(dataObj.tracking).trim() : "";
    
    // Tenta obter também um orderId de fallback (caso o payload mude ou em testes rápidos)
    let orderId = payload.order ? payload.order.orderId : null;
    
    if (!packageId && !trackingCode && !orderId) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Nenhum ID de pacote, codigo de rastreio ou ID de pedido encontrado no payload" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const spreadsheetId = "1Bm7cx-uDRJiJaRo9k8jdJfsxNUs-LhIxSkBwZ98yI-0";
    let ss = null;
    try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
    if (!ss) ss = SpreadsheetApp.openById(spreadsheetId);
    
    const sheet = ss.getSheetByName("Pedidos");
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Planilha nao encontrada" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const colIdPedido = headers.indexOf("ID_Pedido");
    const colPackageId = headers.indexOf("Superfrete_Package_ID");
    const colTracking = headers.indexOf("Codigo_Rastreio");
    
    if (colIdPedido === -1) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Coluna ID_Pedido nao encontrada" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (!orderId) {
      // Busca na planilha pelo Package ID ou pelo Código de Rastreio
      for (let i = 1; i < data.length; i++) {
        const rowPackageId = colPackageId !== -1 ? String(data[i][colPackageId]).trim() : "";
        const rowTracking = colTracking !== -1 ? String(data[i][colTracking]).trim() : "";
        
        const matchPackage = packageId && rowPackageId && (rowPackageId === packageId);
        const matchTracking = trackingCode && rowTracking && (rowTracking === trackingCode);
        
        if (matchPackage || matchTracking) {
          orderId = data[i][colIdPedido];
          break;
        }
      }
    }
    
    if (!orderId) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Nenhum pedido correspondente encontrado na planilha para o pacote " + packageId + " / " + trackingCode }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Executa a sincronização oficial de rastreamento usando a função existente
    const res = atualizarRastreamentoPedido(orderId);
    
    if (res && res.success) {
      try {
        const sfStatus = String(res.superfreteStatus).toLowerCase().trim();
        let titulo = "📦 Atualização de Frete (Pedido #" + orderId + ")";
        
        let customMensagem = "";
        
        if (sfStatus === "canceled" || sfStatus === "cancelled") {
          titulo = "❌ Etiqueta Cancelada (Pedido #" + orderId + ")";
        } else if (sfStatus === "pending") {
          titulo = "⚠️ Etiqueta Pendente de Saldo (Pedido #" + orderId + ")";
        } else if (sfStatus === "released" || sfStatus === "printed") {
          titulo = "🎫 Etiqueta Gerada / Paga (Pedido #" + orderId + ")";
          customMensagem = "A etiqueta do pedido #" + orderId + " foi emitida e paga com sucesso!\n\n" +
                           "• Status Interno: " + res.updatedPanelStatus + "\n" +
                           "• Status SuperFrete: " + res.superfreteStatus + "\n" +
                           "• Rastreamento: " + (res.tracking || "Pendente") + "\n" +
                           "• PDF da Etiqueta: " + (res.pdfUrl || "Disponível no painel");
        } else if (sfStatus === "posted") {
          titulo = "🚚 Pedido Postado nos Correios (Pedido #" + orderId + ")";
        } else if (sfStatus === "delivered") {
          titulo = "🎉 Pedido Entregue (Pedido #" + orderId + ")";
        }
        
        const mensagem = customMensagem || ("O status do pedido #" + orderId + " foi atualizado.\n\n" +
                         "• Status Interno: " + res.updatedPanelStatus + "\n" +
                         "• Status SuperFrete: " + res.superfreteStatus + "\n" +
                         "• Rastreamento: " + (res.tracking || "Pendente"));
                         
        enviarNotificacaoTelegram(titulo, mensagem);
      } catch (errTelegram) {
        Logger.log("Erro ao processar envio de notificacao Telegram: " + errTelegram.message);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Cadastra o webhook automaticamente na API oficial do SuperFrete.
 * Pode ser executada diretamente do editor do Google Apps Script.
 */
function registrarWebhookNoSuperFrete() {
  const sfConfig = obterConfiguracaoSuperFrete();
  const token = sfConfig.token;
  if (!token) {
    throw new Error("Token da SuperFrete nao configurado nas propriedades do script.");
  }

  let webAppUrl = "";
  try {
    webAppUrl = ScriptApp.getService().getUrl();
    if (webAppUrl) {
      webAppUrl = webAppUrl.replace("/dev", "/exec");
    }
  } catch (e) {
    Logger.log("Nao foi possivel detectar a URL do Web App automaticamente: " + e.message);
  }

  if (!webAppUrl || webAppUrl.indexOf("/dev") !== -1) {
    webAppUrl = "https://script.google.com/macros/s/AKfycbycjYy2UiWKMJAQcrFEs9IFJ42LTKtSJMgUufCYErjlX9U80m6V7M1pHvRb0gbH5TSF/exec";
  }

  // Garante que a URL termine com o parametro source
  if (webAppUrl.indexOf("?") === -1) {
    webAppUrl += "?source=superfrete";
  } else if (webAppUrl.indexOf("source=superfrete") === -1) {
    webAppUrl += "&source=superfrete";
  }

  const url = sfConfig.baseUrl + "/api/v0/webhook";
  
  // Primeiro, lista e remove webhooks antigos para evitar duplicidade ou URLs desatualizadas (como /dev)
  try {
    const listOptions = {
      method: "get",
      headers: {
        "Authorization": "Bearer " + token
      },
      muteHttpExceptions: true
    };
    const listResponse = UrlFetchApp.fetch(url, listOptions);
    if (listResponse.getResponseCode() === 200) {
      const existingWebhooks = JSON.parse(listResponse.getContentText());
      const webhookList = Array.isArray(existingWebhooks) ? existingWebhooks : (existingWebhooks.results || existingWebhooks.data || []);
      for (const webhook of webhookList) {
        if (webhook.id) {
          Logger.log("Removendo webhook antigo/invalido: " + webhook.id + " (" + webhook.url + ")");
          const deleteUrl = url + "/" + webhook.id;
          const deleteOptions = {
            method: "delete",
            headers: {
              "Authorization": "Bearer " + token
            },
            muteHttpExceptions: true
          };
          UrlFetchApp.fetch(deleteUrl, deleteOptions);
        }
      }
    }
  } catch (e) {
    Logger.log("Nao foi possivel limpar webhooks antigos: " + e.message);
  }

  const payload = {
    name: "Notificacao Automatica FigTree",
    url: webAppUrl,
    events: [
      "order.posted",
      "order.delivered",
      "order.cancelled",
      "order.released",
      "order.printed",
      "order.paid",
      "order.pending"
    ]
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": "Bearer " + token
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const responseText = response.getContentText();
  const responseCode = response.getResponseCode();
  
  Logger.log("Response Code: " + responseCode);
  Logger.log("Response Text: " + responseText);
  
  return {
    statusCode: responseCode,
    response: responseText
  };
}

/**
 * Envia uma notificação para o Telegram
 */
function enviarNotificacaoTelegram(titulo, mensagem) {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    let token = scriptProps.getProperty("TELEGRAM_BOT_TOKEN");
    if (!token) {
      token = "8782927815:AAEgY0cK1rC-d5dZOAYbA1xkDNySYoL81fY";
    }
    const chatId = scriptProps.getProperty("TELEGRAM_CHAT_ID");
    
    if (!chatId) {
      Logger.log("Telegram Chat ID não configurado nas propriedades do script.");
      return { success: false, error: "Telegram Chat ID não configurado" };
    }
    
    const url = "https://api.telegram.org/bot" + token + "/sendMessage";
    const formattedText = "<b>" + titulo + "</b>\n\n" + mensagem;
    
    const payload = {
      chat_id: chatId,
      text: formattedText,
      parse_mode: "HTML"
    };
    
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log("Telegram response code: " + responseCode);
    
    return {
      success: responseCode === 200,
      statusCode: responseCode,
      response: responseText
    };
  } catch (err) {
    Logger.log("Erro ao enviar notificação para o Telegram: " + err.message);
    throw err;
  }
}

/**
 * Configura o TELEGRAM_CHAT_ID a partir da última interação com o bot no /getUpdates
 */
function configurarTelegramChatId() {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    let token = scriptProps.getProperty("TELEGRAM_BOT_TOKEN");
    if (!token) {
      token = "8782927815:AAEgY0cK1rC-d5dZOAYbA1xkDNySYoL81fY";
      scriptProps.setProperty("TELEGRAM_BOT_TOKEN", token);
    }
    
    // Limpar propriedades antigas do NTFY para não deixar "lixo"
    const oldProps = ["NTFY_TOPIC", "NTFY_TOPIC_SALES", "NTFY_TOPIC_SHIPPING", "NTFY_SERVER_URL", "NTFY_ACCESS_TOKEN"];
    oldProps.forEach(function(prop) {
      try {
        scriptProps.deleteProperty(prop);
      } catch (e) {
        Logger.log("Erro ao deletar propriedade " + prop + ": " + e.message);
      }
    });
    
    const url = "https://api.telegram.org/bot" + token + "/getUpdates";
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      return { success: false, error: "Erro ao chamar getUpdates: " + responseText };
    }
    
    const data = JSON.parse(responseText);
    if (!data.ok) {
      return { success: false, error: "getUpdates retornou ok: false" };
    }
    
    const results = data.result;
    if (!results || results.length === 0) {
      return { 
        success: false, 
        error: "Nenhuma mensagem recente encontrada. Por favor, envie uma mensagem para o bot no Telegram (ex: /start ou qualquer mensagem) e tente configurar novamente." 
      };
    }
    
    const lastUpdate = results[results.length - 1];
    let chatId = null;
    let username = "";
    let firstName = "";
    
    if (lastUpdate.message) {
      chatId = lastUpdate.message.chat.id;
      username = lastUpdate.message.chat.username || "";
      firstName = lastUpdate.message.chat.first_name || "";
    } else if (lastUpdate.my_chat_member) {
      chatId = lastUpdate.my_chat_member.chat.id;
      username = lastUpdate.my_chat_member.chat.username || "";
      firstName = lastUpdate.my_chat_member.chat.first_name || "";
    }
    
    if (!chatId) {
      return { success: false, error: "Não foi possível obter o Chat ID do último update." };
    }
    
    scriptProps.setProperty("TELEGRAM_CHAT_ID", chatId.toString());
    
    // Envia mensagem de confirmação
    const msgTest = "✅ Integração do Telegram configurada com sucesso!\nEste chat receberá as notificações de vendas e fretes do FigTree.";
    UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/sendMessage", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        chat_id: chatId,
        text: msgTest
      }),
      muteHttpExceptions: true
    });
    
    return {
      success: true,
      chatId: chatId,
      username: username,
      firstName: firstName,
      message: "Telegram Chat ID configurado com sucesso!"
    };
  } catch (err) {
    Logger.log("Erro ao configurar Telegram Chat ID: " + err.message);
    return { success: false, error: err.message };
  }
}

/**
 * Retorna as credenciais e URL base corretas para a API da SuperFrete de acordo com o ambiente configurado (sandbox ou production).
 */
function obterConfiguracaoSuperFrete() {
  const scriptProps = PropertiesService.getScriptProperties();
  const env = String(scriptProps.getProperty("SUPERFRETE_ENVIRONMENT") || "production").toLowerCase().trim();
  
  if (env === "sandbox") {
    return {
      token: scriptProps.getProperty("SUPERFRETE_SANDBOX_TOKEN"),
      baseUrl: "https://sandbox.superfrete.com",
      isSandbox: true
    };
  } else {
    return {
      token: scriptProps.getProperty("SUPERFRETE_TOKEN"),
      baseUrl: "https://api.superfrete.com",
      isSandbox: false
    };
  }
}

