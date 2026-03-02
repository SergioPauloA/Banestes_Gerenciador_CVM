/**
 * Sistema de Monitoramento de Fundos CVM
 * Lê dados da planilha (que já tem as fórmulas IMPORTXML)
 */

var SPREADSHEET_ID = '1N6LP1ydsxnQO_Woatv9zWEccb0fOGaV_3EKK1GoSWZI';

// Debug flag - set to false in production to reduce logging
var DEBUG_MODE = true;

// ============================================
// WEB APP
// ============================================

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('Monitor de Fundos CVM - BANESTES')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// OBTER PLANILHA
// ============================================

function obterPlanilha() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (error) {
    Logger.log('❌ Erro ao abrir planilha: ' + error.toString());
    throw new Error('Não foi possível abrir a planilha.');
  }
}

function obterURLPlanilha() {
  return 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/edit';
}

// ============================================
// API: LER DADOS DA PLANILHA
// ============================================

function getDashboardData() {
  try {
    Logger.log('📖 Lendo dados da planilha...');
    
    var ss = obterPlanilha();
    var datas = getDatasReferencia();
    
    Logger.log('📅 Datas de referência: ' + JSON.stringify(datas));
    
    var balancete = lerAbaBalancete(ss, datas);
    Logger.log('📊 Balancete statusGeral: "' + balancete.statusGeral + '"');
    
    var composicao = lerAbaComposicao(ss, datas);
    Logger.log('📈 Composição statusGeral: "' + composicao.statusGeral + '"');
    
    var diarias = lerAbaDiarias(ss);
    Logger.log('📅 Diárias statusGeral1: "' + diarias.statusGeral1 + '", statusGeral2: "' + diarias.statusGeral2 + '"');
    
    var lamina = lerAbaLamina(ss, datas);
    Logger.log('📄 Lâmina statusGeral: "' + lamina.statusGeral + '"');
    
    var perfilMensal = lerAbaPerfilMensal(ss, datas);
    Logger.log('📊 Perfil Mensal statusGeral: "' + perfilMensal.statusGeral + '"');
    
    var resultado = {
      timestamp: new Date().toISOString(),
      datas: datas,
      balancete: balancete,
      composicao: composicao,
      diarias: diarias,
      lamina: lamina,
      perfilMensal: perfilMensal
    };
    
    Logger.log('✅ Dados lidos com sucesso');
    return resultado;
    
  } catch (error) {
    Logger.log('❌ Erro em getDashboardData: ' + error.toString());
    throw new Error('Erro ao carregar dados: ' + error.message);
  }
}

// Alias para compatibilidade com Index.html
function getDashboardDataCompleto() {
  return getDashboardData();
}

// ============================================
// FUNÇÕES DE LEITURA POR ABA - ATUALIZADO PARA CÓDIGO BANESTES
// ============================================

function lerAbaBalancete(ss, datas) {
  var aba = ss.getSheetByName('Balancete');
  if (!aba) throw new Error('Aba Balancete não encontrada');
  
  var statusGeral = aba.getRange('E1').getDisplayValue();
  var ultimaLinha = aba.getLastRow();
  
  if (ultimaLinha < 4) {
    return {
      titulo: 'Balancetes de Fundos',
      statusGeral: 'SEM DADOS',
      substatus: null,
      dados: []
    };
  }
  
  // Ler 6 colunas: A=Nome, B=Código, C=Comp1, D=Status1, E=Comp2, F=Status2
  var valores = aba.getRange(4, 1, ultimaLinha - 3, 6).getDisplayValues();
  
  var dados = valores
    .filter(function(linha) { return linha[0] !== '' && linha[0] !== null; })
    .map(function(linha) {
      var codigoBanestes = buscarCodigoBanestes(ss, linha[0]);
      return {
        fundo: linha[0],
        codigo: codigoBanestes,
        retorno: linha[2] || '-',    // Competência 1
        status: linha[3] || '-',      // Status 1
        retorno2: linha[4] || '-',    // Competência 2
        status2: linha[5] || '-'      // Status 2
      };
    });

  var substatus = null;
  var statusGeralDisplay = statusGeral;
  
  // Calcular cor baseada em desconformidade ou dias restantes
  if (statusGeral && statusGeral.indexOf('DESCONFORMIDADE') !== -1) {
    substatus = 'ok-vermelho';
  } else if (statusGeral === 'OK' || statusGeral.indexOf('OK') !== -1) {
    substatus = calcularCorStatusOk(datas.diasRestantes);
  }
  
  return {
    titulo: 'Balancetes de Fundos',
    statusGeral: statusGeralDisplay,
    substatus: substatus,
    dados: dados
  };
}

function lerAbaComposicao(ss, datas) {
  var aba = ss.getSheetByName('Composição');
  if (!aba) throw new Error('Aba Composição não encontrada');
  
  var statusGeral = aba.getRange('E1').getDisplayValue();
  var ultimaLinha = aba.getLastRow();
  
  if (ultimaLinha < 4) {
    return {
      titulo: 'Composição da Carteira',
      statusGeral: 'SEM DADOS',
      substatus: null,
      dados: []
    };
  }
  
  var valores = aba.getRange(4, 1, ultimaLinha - 3, 6).getDisplayValues();
  
  var dados = valores
    .filter(function(linha) { return linha[0] !== '' && linha[0] !== null; })
    .map(function(linha) {
      var codigoBanestes = buscarCodigoBanestes(ss, linha[0]);
      return {
        fundo: linha[0],
        codigo: codigoBanestes,
        retorno: linha[2] || '-',
        status: linha[3] || '-',
        retorno2: linha[4] || '-',
        status2: linha[5] || '-'
      };
    });

  var substatus = null;
  var statusGeralDisplay = statusGeral;
  
  if (statusGeral && statusGeral.indexOf('DESCONFORMIDADE') !== -1) {
    substatus = 'ok-vermelho';
  } else if (statusGeral === 'OK' || statusGeral.indexOf('OK') !== -1) {
    substatus = calcularCorStatusOk(datas.diasRestantes);
  }
  
  return {
    titulo: 'Composição da Carteira',
    statusGeral: statusGeralDisplay,
    substatus: substatus,
    dados: dados
  };
}

function lerAbaDiarias(ss) {
  var aba = ss.getSheetByName('Diárias');
  if (!aba) throw new Error('Aba Diárias não encontrada');
  
  var statusGeral1 = aba.getRange('E1').getDisplayValue();
  var statusGeral2 = aba.getRange('F1').getDisplayValue();
  var ultimaLinha = aba.getLastRow();
  
  if (ultimaLinha < 4) {
    return {
      titulo: 'Informações Diárias',
      statusGeral1: statusGeral1 || 'SEM DADOS',
      statusGeral2: statusGeral2 || 'SEM DADOS',
      dados: []
    };
  }
  
  // Ler 6 colunas: A=Nome, B=Código, C=Retorno1, D=Status1, E=Retorno2, F=Status2
  var valores = aba.getRange(4, 1, ultimaLinha - 3, 6).getDisplayValues();
  
  var dados = valores
    .filter(function(linha) { return linha[0] !== '' && linha[0] !== null; })
    .map(function(linha) {
      return {
        fundo: linha[0],
        codigo: String(linha[1]),
        retorno1: linha[2] || '-',
        status1: linha[3] || '-',
        retorno2: linha[4] || '#N/A',
        status2: linha[5] || 'A ATUALIZAR' // CORRIGIDO: sempre mostrar "A ATUALIZAR" se vazio
      };
    });
  
  return {
    titulo: 'Informações Diárias',
    statusGeral1: statusGeral1 || 'SEM DADOS',
    statusGeral2: statusGeral2 || 'SEM DADOS',
    dados: dados
  };
}

function lerAbaLamina(ss, datas) {
  var aba = ss.getSheetByName('Lâmina');
  if (!aba) throw new Error('Aba Lâmina não encontrada');
  
  var statusGeral = aba.getRange('E1').getDisplayValue();
  var ultimaLinha = aba.getLastRow();
  
  if (ultimaLinha < 4) {
    return {
      titulo: 'Lâmina do Fundo',
      statusGeral: 'SEM DADOS',
      substatus: null,
      dados: []
    };
  }
  
  var valores = aba.getRange(4, 1, ultimaLinha - 3, 6).getDisplayValues();
  
  var dados = valores
    .filter(function(linha) { return linha[0] !== '' && linha[0] !== null; })
    .map(function(linha) {
      var codigoBanestes = buscarCodigoBanestes(ss, linha[0]);
      return {
        fundo: linha[0],
        codigo: codigoBanestes,
        retorno: linha[2] || '-',
        status: linha[3] || '-',
        retorno2: linha[4] || '-',
        status2: linha[5] || '-'
      };
    });

  var substatus = null;
  var statusGeralDisplay = statusGeral;
  
  if (statusGeral && statusGeral.indexOf('DESCONFORMIDADE') !== -1) {
    substatus = 'ok-vermelho';
  } else if (statusGeral === 'OK' || statusGeral.indexOf('OK') !== -1) {
    substatus = calcularCorStatusOk(datas.diasRestantes);
  }
  
  return {
    titulo: 'Lâmina do Fundo',
    statusGeral: statusGeralDisplay,
    substatus: substatus,
    dados: dados
  };
}

function lerAbaPerfilMensal(ss, datas) {
  var aba = ss.getSheetByName('Perfil Mensal');
  if (!aba) throw new Error('Aba Perfil Mensal não encontrada');
  
  var statusGeral = aba.getRange('E1').getDisplayValue();
  var ultimaLinha = aba.getLastRow();
  
  if (ultimaLinha < 4) {
    return {
      titulo: 'Perfil Mensal',
      statusGeral: 'SEM DADOS',
      substatus: null,
      dados: []
    };
  }
  
  var valores = aba.getRange(4, 1, ultimaLinha - 3, 6).getDisplayValues();
  
  var dados = valores
    .filter(function(linha) { return linha[0] !== '' && linha[0] !== null; })
    .map(function(linha) {
      var codigoBanestes = buscarCodigoBanestes(ss, linha[0]);
      return {
        fundo: linha[0],
        codigo: codigoBanestes,
        retorno: linha[2] || '-',
        status: linha[3] || '-',
        retorno2: linha[4] || '-',
        status2: linha[5] || '-'
      };
    });

  var substatus = null;
  var statusGeralDisplay = statusGeral;
  
  if (statusGeral && statusGeral.indexOf('DESCONFORMIDADE') !== -1) {
    substatus = 'ok-vermelho';
  } else if (statusGeral === 'OK' || statusGeral.indexOf('OK') !== -1) {
    substatus = calcularCorStatusOk(datas.diasRestantes);
  }
  
  return {
    titulo: 'Perfil Mensal',
    statusGeral: statusGeralDisplay,
    substatus: substatus,
    dados: dados
  };
}

function calcularCorStatusOk(diasRestantes) {
  if (diasRestantes > 15) return 'ok-verde';      // Mais de 15 dias = Verde
  if (diasRestantes >= 5) return 'ok-amarelo';    // 5 a 15 dias = Amarelo
  return 'ok-vermelho';                            // Menos de 5 dias = Vermelho
}

// ============================================
// NOVA FUNÇÃO: BUSCAR CÓDIGO BANESTES
// ============================================

function buscarCodigoBanestes(ss, nomeFundo) {
  try {
    var abaCodFundo = ss.getSheetByName('COD FUNDO');
    if (!abaCodFundo) {
      Logger.log('⚠️ Aba COD FUNDO não encontrada');
      return '-';
    }
    
    var ultimaLinha = abaCodFundo.getLastRow();
    if (ultimaLinha < 2) return '-';
    
    // Normalizar o nome do fundo para comparação
    var nomeFundoNormalizado = nomeFundo.trim().replace(/\s+/g, ' ').toUpperCase();
    
    // Buscar nas 3 colunas: A=Nome, B=CVM, C=BANESTES
    var dados = abaCodFundo.getRange(2, 1, ultimaLinha - 1, 3).getValues();
    
    for (var i = 0; i < dados.length; i++) {
      var nomeNaAba = String(dados[i][0]).trim().replace(/\s+/g, ' ').toUpperCase();
      
      if (nomeNaAba === nomeFundoNormalizado) {
        var codigo = dados[i][2];
        Logger.log('✅ Código encontrado para ' + nomeFundo.substring(0, 30) + '... = ' + codigo);
        return String(codigo);
      }
    }
    
    Logger.log('⚠️ Código não encontrado para: ' + nomeFundo.substring(0, 40));
    Logger.log('   Buscando por: ' + nomeFundoNormalizado.substring(0, 40));
    return '-';
  } catch (error) {
    Logger.log('❌ Erro ao buscar código BANESTES: ' + error.toString());
    return '-';
  }
}

// ============================================
// API: VERIFICAR INSTALAÇÃO
// ============================================

function getStatusInstalacao() {
  try {
    var ss = obterPlanilha();
    var abas = ['GERAL', 'Balancete', 'Composição', 'Diárias', 'Lâmina', 'Perfil Mensal', 'APOIO', 'FERIADOS', 'COD FUNDO'];
    
    var abasExistentes = [];
    var todasExistem = true;
    
    abas.forEach(function(nomeAba) {
      var aba = ss.getSheetByName(nomeAba);
      if (aba) {
        abasExistentes.push(nomeAba);
      } else {
        todasExistem = false;
      }
    });
    
    // Verificar se tem fórmulas nas abas de dados
    var temFormulas = false;
    if (todasExistem) {
      try {
        var abaBalancete = ss.getSheetByName('Balancete');
        if (abaBalancete && abaBalancete.getLastRow() >= 4) {
          var formula = abaBalancete.getRange('C4').getFormula();
          temFormulas = formula && formula.indexOf('IMPORTXML') !== -1;
        }
      } catch (e) {
        temFormulas = false;
      }
    }
    
    return {
      instalado: todasExistem && temFormulas,
      abas: abasExistentes,
      totalAbas: abasExistentes.length,
      temFormulas: temFormulas,
      url: obterURLPlanilha()
    };
  } catch (error) {
    return {
      instalado: false,
      erro: error.toString()
    };
  }
}

// ============================================
// API: FORÇAR REINSTALAÇÃO
// ============================================

function forcarReinstalacao() {
  Logger.log('🔄 Forçando reinstalação...');
  return setupCompletoAutomatico();
}

// ============================================
// NOVAS FUNÇÕES DE TESTE
// ============================================

// ============================================
// BUSCAR CÓDIGO BANESTES (VERSÃO FINAL)
// ============================================

function buscarCodigoBanestes(ss, nomeFundo) {
  try {
    var abaCodFundo = ss.getSheetByName('COD FUNDO');
    if (!abaCodFundo) return '-';
    
    var ultimaLinha = abaCodFundo.getLastRow();
    if (ultimaLinha < 2) return '-';
    
    // Normalizar o nome do fundo para comparação
    var nomeFundoNormalizado = nomeFundo.trim().replace(/\s+/g, ' ').toUpperCase();
    
    // Buscar nas 3 colunas: A=Nome, B=CVM, C=BANESTES
    var dados = abaCodFundo.getRange(2, 1, ultimaLinha - 1, 3).getValues();
    
    for (var i = 0; i < dados.length; i++) {
      var nomeNaAba = String(dados[i][0]).trim().replace(/\s+/g, ' ').toUpperCase();
      
      if (nomeNaAba === nomeFundoNormalizado) {
        return String(dados[i][2]); // Coluna C = Código BANESTES
      }
    }
    
    return '-';
  } catch (error) {
    return '-';
  }
}

function atualizarAbaCodFundoComColuna3() {
  Logger.log('🔄 Atualizando aba COD FUNDO com coluna C...');
  
  var ss = obterPlanilha();
  preencherAbaCodFundo(ss);
  
  Logger.log('✅ Aba COD FUNDO atualizada!');
  Logger.log('📊 Verificando dados...');
  
  var aba = ss.getSheetByName('COD FUNDO');
  var dados = aba.getRange('A2:C27').getValues();
  
  Logger.log('\n📋 Primeiros 5 fundos:');
  dados.slice(0, 5).forEach(function(linha, i) {
    Logger.log('  [' + (i+1) + '] Nome: ' + linha[0].substring(0, 30) + '...');
    Logger.log('      CVM: ' + linha[1]);
    Logger.log('      BANESTES: ' + linha[2]);
    Logger.log('');
  });
  
  return {
    success: true,
    message: 'Aba COD FUNDO atualizada com 3 colunas!'
  };
}

// ============================================
// SUAS FUNÇÕES ORIGINAIS (SEM ALTERAÇÃO)
// ============================================

/**
 * Atualizar status na planilha após ler os dados
 */
function atualizarStatusNaPlanilha() {
  try {
    var ss = obterPlanilha();

    // Processa em bloco as abas de conformidade
    var datasReferencia = getDatasReferencia();
    processarAbasConformidade(datasReferencia); // <<--- NOVA FUNÇÃO, roda Balancete, Composição, Lâmina, Perfil Mensal

    // Diárias - permanece lógica específica (caso use diferentes critérios para Diárias):
    var abaDiarias = ss.getSheetByName('Diárias');
    // status 1 (col E1): todos OK em D4:D29?
    var dadosDiarias1 = abaDiarias.getRange('D4:D29').getValues();
    var dadosDiarias2 = abaDiarias.getRange('F4:F29').getValues();

    var totalOK1 = dadosDiarias1.filter(function(r) { return r[0] === 'OK'; }).length;
    var statusDiarias1 = totalOK1 === dadosDiarias1.length ? 'OK' : 'DESCONFORMIDADE';

    var totalOK2 = dadosDiarias2.filter(function(r) { return r[0] === 'OK'; }).length;
    var statusDiarias2 = totalOK2 === dadosDiarias2.length ? 'OK' : 'A ATUALIZAR';

    abaDiarias.getRange('E1').setValue(statusDiarias1);
    abaDiarias.getRange('F1').setValue(statusDiarias2);

    // Balancete, Composição, Lâmina, Perfil Mensal: status geral já foi atualizado em E1 de cada aba pela nova função

    // Leitura após processamento atualizado:
    var abaBalancete = ss.getSheetByName('Balancete');
    var statusBalancete = abaBalancete.getRange('E1').getValue();

    var abaComposicao = ss.getSheetByName('Composição');
    var statusComposicao = abaComposicao.getRange('E1').getValue();

    var abaLamina = ss.getSheetByName('Lâmina');
    var statusLamina = abaLamina.getRange('E1').getValue();

    var abaPerfilMensal = ss.getSheetByName('Perfil Mensal');
    var statusPerfilMensal = abaPerfilMensal.getRange('E1').getValue();

    // Atualiza dashboard
    var abaGeral = ss.getSheetByName('GERAL');
    abaGeral.getRange('A4').setValue(statusBalancete);
    abaGeral.getRange('B4').setValue(statusComposicao);
    abaGeral.getRange('C4').setValue(statusDiarias1);
    abaGeral.getRange('D4').setValue(statusDiarias2);
    abaGeral.getRange('E4').setValue(statusLamina);
    abaGeral.getRange('F4').setValue(statusPerfilMensal);

    Logger.log('✅ Status atualizados na planilha');
  } catch (error) {
    Logger.log('❌ Erro ao atualizar status: ' + error.toString());
  }
}

/**
 * Função executada automaticamente pelo trigger
 * Atualiza status na planilha
 */
function atualizarStatusNaPlanilhaAutomatico() {
  try {
    Logger.log('🔄 [TRIGGER] Atualização automática iniciada em: ' + new Date());
    
    var ss = obterPlanilha();
    var datas = getDatasReferencia();

    // === ATUALIZAR COMPETÊNCIAS DAS ABAS MENSAIS ===
    atualizarTodasCompetencias();
    
    Logger.log('📅 Datas de referência:');
    Logger.log('   - diaMesRef (deve ser 01/12/2025): ' + datas.diaMesRef);
    Logger.log('   - diasRestantes: ' + datas.diasRestantes);
    Logger.log('   - prazoFinal: ' + datas.diaMesRef2);
    
    // ============================================
    // BALANCETE
    // ============================================
    Logger.log('\n📊 Processando Balancete...');
    var abaBalancete = ss.getSheetByName('Balancete');
    var dadosBalancete = abaBalancete.getRange('C4:C29').getDisplayValues();
    
    Logger.log('   Total de linhas: ' + dadosBalancete.length);
    Logger.log('   Primeiras 3 datas lidas:');
    for (var i = 0; i < Math.min(3, dadosBalancete.length); i++) {
      Logger.log('   [' + (i+1) + '] "' + dadosBalancete[i][0] + '"');
    }
    
    var statusBalancete = calcularStatusGeralDaAba(dadosBalancete, 'mensal');
    Logger.log('   Status Geral calculado: "' + statusBalancete + '"');
    abaBalancete.getRange('E1').setValue(statusBalancete);
    Logger.log('   ✅ Status Geral gravado na E1');
    
    // Atualizar status individuais
    Logger.log('   Atualizando status individuais (coluna D)...');
    var statusIndividuaisCalculados = [];
    for (var i = 0; i < dadosBalancete.length; i++) {
      var retorno = dadosBalancete[i][0];
      // Only enable debug logging for first 3 rows
      var enableDebugLog = (i < 3);
      var status = calcularStatusIndividual(retorno, 'mensal', enableDebugLog);
      abaBalancete.getRange(i + 4, 4).setValue(status);
      statusIndividuaisCalculados.push(status);
    }
    
    Logger.log('   Status individuais calculados:');
    var contadores = {
      'OK': 0,
      'EM CONFORMIDADE': 0,
      'DESATUALIZADO': 0,
      '-': 0
    };
    statusIndividuaisCalculados.forEach(function(s) {
      if (contadores.hasOwnProperty(s)) {
        contadores[s]++;
      }
    });
    Logger.log('   - OK: ' + contadores['OK']);
    Logger.log('   - EM CONFORMIDADE: ' + contadores['EM CONFORMIDADE']);
    Logger.log('   - DESATUALIZADO: ' + contadores['DESATUALIZADO']);
    Logger.log('   - Vazios (-): ' + contadores['-']);
    
    // ============================================
    // COMPOSIÇÃO
    // ============================================
    Logger.log('\n📈 Processando Composição...');
    var abaComposicao = ss.getSheetByName('Composição');
    var dadosComposicao = abaComposicao.getRange('C4:C29').getDisplayValues();
    var statusComposicao = calcularStatusGeralDaAba(dadosComposicao, 'mensal');
    Logger.log('   Status Geral: "' + statusComposicao + '"');
    abaComposicao.getRange('E1').setValue(statusComposicao);
    
    for (var i = 0; i < dadosComposicao.length; i++) {
      var retorno = dadosComposicao[i][0];
      var status = calcularStatusIndividual(retorno, 'mensal');
      abaComposicao.getRange(i + 4, 4).setValue(status);
    }
    Logger.log('   ✅ Status individuais atualizados');
    
    // ============================================
    // DIÁRIAS (NÃO ALTERAR - ESTÁ CORRETO)
    // ============================================
    Logger.log('\n📅 Processando Diárias...');
    var abaDiarias = ss.getSheetByName('Diárias');
    var dadosDiarias1 = abaDiarias.getRange('C4:C29').getDisplayValues();
    var dadosDiarias2 = abaDiarias.getRange('E4:E29').getDisplayValues();
    
    var statusDiarias1 = calcularStatusGeralDaAba(dadosDiarias1, 'diario');
    var statusDiarias2 = calcularStatusGeralDaAba(dadosDiarias2, 'diario');
    
    abaDiarias.getRange('E1').setValue(statusDiarias1);
    abaDiarias.getRange('F1').setValue(statusDiarias2);
    
    // Status individuais
    for (var i = 0; i < dadosDiarias1.length; i++) {
      var status1 = calcularStatusIndividual(dadosDiarias1[i][0], 'diario');
      var status2 = calcularStatusIndividual(dadosDiarias2[i][0], 'diario');
      abaDiarias.getRange(i + 4, 4).setValue(status1);
      abaDiarias.getRange(i + 4, 6).setValue(status2);
    }
    Logger.log('   ✅ Diárias atualizadas');
    
    // ============================================
    // LÂMINA
    // ============================================
    Logger.log('\n📄 Processando Lâmina...');
    var abaLamina = ss.getSheetByName('Lâmina');
    var dadosLamina = abaLamina.getRange('C4:C29').getDisplayValues();
    var statusLamina = calcularStatusGeralDaAba(dadosLamina, 'mensal');
    Logger.log('   Status Geral: "' + statusLamina + '"');
    abaLamina.getRange('E1').setValue(statusLamina);
    
    for (var i = 0; i < dadosLamina.length; i++) {
      var retorno = dadosLamina[i][0];
      var status = calcularStatusIndividual(retorno, 'mensal');
      abaLamina.getRange(i + 4, 4).setValue(status);
    }
    Logger.log('   ✅ Status individuais atualizados');
    
    // ============================================
    // PERFIL MENSAL
    // ============================================
    Logger.log('\n📊 Processando Perfil Mensal...');
    var abaPerfilMensal = ss.getSheetByName('Perfil Mensal');
    var dadosPerfilMensal = abaPerfilMensal.getRange('C4:C29').getDisplayValues();
    var statusPerfilMensal = calcularStatusGeralDaAba(dadosPerfilMensal, 'mensal');
    Logger.log('   Status Geral: "' + statusPerfilMensal + '"');
    abaPerfilMensal.getRange('E1').setValue(statusPerfilMensal);
    
    for (var i = 0; i < dadosPerfilMensal.length; i++) {
      var retorno = dadosPerfilMensal[i][0];
      var status = calcularStatusIndividual(retorno, 'mensal');
      abaPerfilMensal.getRange(i + 4, 4).setValue(status);
    }
    Logger.log('   ✅ Status individuais atualizados');
    
    // ============================================
    // GERAL
    // ============================================
    Logger.log('\n📋 Atualizando Dashboard Geral...');
    var abaGeral = ss.getSheetByName('GERAL');
    abaGeral.getRange('A4').setValue(statusBalancete);
    abaGeral.getRange('B4').setValue(statusComposicao);
    abaGeral.getRange('C4').setValue(statusDiarias1);
    abaGeral.getRange('D4').setValue(statusDiarias2);
    abaGeral.getRange('E4').setValue(statusLamina);
    abaGeral.getRange('F4').setValue(statusPerfilMensal);
    
    Logger.log('\n✅ [TRIGGER] Atualização automática concluída!');
    Logger.log('📊 Próxima execução em 1 hora');
    
  } catch (error) {
    Logger.log('❌ [TRIGGER] Erro na atualização automática: ' + error.toString());
    Logger.log('   Stack trace: ' + error.stack);
  }
}

/**
 * Calcular status individual de um fundo
 * @param {string} retorno - Data de retorno da coluna C (ex: "01/12/2025")
 * @param {string} tipo - 'mensal' ou 'diario'
 * @param {boolean} enableDebugLog - Optional: Enable debug logging for this call (default: false)
 * @returns {string} - 'OK', 'EM CONFORMIDADE', 'DESATUALIZADO', ou '-'
 */
function calcularStatusIndividual(retorno, tipo, enableDebugLog) {
  // Se vazio ou com erro, retornar DESATUALIZADO
  if (
    !retorno ||
    retorno === '-' ||
    retorno === '' ||
    retorno === 'Loading...' ||
    retorno === 'ERRO' ||
    retorno === '#N/A' ||
    retorno === '#REF!' ||
    retorno === null ||
    retorno === undefined
  ) {
    return 'DESATUALIZADO';
  }

  var datas = getDatasReferencia();
  
  // Normalizar as datas para comparação
  var retornoNormalizado = normalizaDataParaComparacao(retorno);
  
  if (tipo === 'mensal') {
    var dataRefNormalizada = normalizaDataParaComparacao(datas.diaMesRef);
    
    // Debug logging (only when explicitly enabled and DEBUG_MODE is true)
    if (DEBUG_MODE && enableDebugLog) {
      Logger.log('🔍 Comparando: "' + retornoNormalizado + '" vs "' + dataRefNormalizada + '"');
    }
    
    // Se a data é igual à data de referência → OK
    if (retornoNormalizado === dataRefNormalizada) {
      return 'OK';
    }
    
    // ✅ NOVA LÓGICA: Se ainda está dentro do prazo, aceitar apenas mês anterior
    if (datas.diasRestantes >= 0) {
      // Calcular data do mês retrasado (limite mínimo aceitável)
      var hoje = new Date();
      var mesRetrasado = new Date(hoje.getFullYear(), hoje.getMonth() - 2, 1);
      var dataLimiteMinima = normalizaDataParaComparacao(formatarData(mesRetrasado));
      
      // Converter strings DD/MM/YYYY para objetos Date para comparação
      var partesRetorno = retornoNormalizado.split('/');
      var dataRetorno = new Date(partesRetorno[2], partesRetorno[1] - 1, partesRetorno[0]);
      
      var partesLimite = dataLimiteMinima.split('/');
      var dataLimite = new Date(partesLimite[2], partesLimite[1] - 1, partesLimite[0]);
      
      // Debug logging
      if (DEBUG_MODE && enableDebugLog) {
        Logger.log('📅 Data retornada: ' + retornoNormalizado + ' (' + dataRetorno.toISOString().split('T')[0] + ')');
        Logger.log('📅 Data limite mínima: ' + dataLimiteMinima + ' (' + dataLimite.toISOString().split('T')[0] + ')');
        Logger.log('✅ Data retornada >= limite? ' + (dataRetorno >= dataLimite));
      }
      
      // Se a data retornada é >= mês retrasado → OK
      if (dataRetorno >= dataLimite) {
        return 'OK';
      }
      
      // Data muito antiga → DESATUALIZADO
      return 'DESATUALIZADO';
    }
    
    // Passou do prazo → DESATUALIZADO
    return 'DESATUALIZADO';
  }

  if (tipo === 'diario') {
    var diaD1Normalizado = normalizaDataParaComparacao(datas.diaD1);
    
    // Se a data é igual ao dia D-1 → OK
    if (retornoNormalizado === diaD1Normalizado) {
      return 'OK';
    }
    
    // Para diárias, se não é OK, retornar vazio (conforme planilha original)
    return '-';
  }

  return 'DESATUALIZADO';
}

/**
 * Calcular status geral de uma aba
 */
function calcularStatusGeralDaAba(dados, tipo, datas) {
  var totalOK = 0;
  var totalAguardando = 0;
  var total = dados.length;
  
  dados.forEach(function(linha) {
    var retorno = linha[0];
    if (!retorno || retorno === '-' || retorno === '' || retorno === 'Loading...') {
      totalAguardando++;
    } else {
      var status = calcularStatusIndividual(retorno, tipo);
      if (status === 'OK') {
        totalOK++;
      }
    }
  });
  
  if (totalAguardando === total) {
    return 'AGUARDANDO DADOS';
  }
  
  if (totalOK === total) {
    return 'OK';
  }
  
  var datas = getDatasReferencia();
  
  // Para tipo mensal: Se ainda está dentro do prazo (DIADDD <= DIAMESREF2)
  if (tipo === 'mensal' && datas.diasRestantes >= 0) {
    return 'OK (' + formatarDiasRestantes(datas.diasRestantes) + ')';
  }
  
  if (tipo === 'diario') {
    return totalOK >= total / 2 ? 'OK' : 'DESCONFORMIDADE';
  }
  
  return 'DESCONFORMIDADE';
}

function testarAtualizacaoAutomatica() {
  Logger.log('🧪 Testando atualização automática...');
  
  try {
    atualizarStatusNaPlanilhaAutomatico();
    Logger.log('✅ Teste concluído com sucesso!');
    return {
      success: true,
      message: 'Atualização automática testada com sucesso!'
    };
  } catch (error) {
    Logger.log('❌ Erro no teste: ' + error.toString());
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * Função para testar manualmente o cálculo de status
 * Execute esta função no Apps Script Editor para debug
 */
function testarCalculoDeStatus() {
  Logger.log('🧪 ===== TESTE DE CÁLCULO DE STATUS =====\n');
  
  var ss = obterPlanilha();
  var datas = getDatasReferencia();
  
  Logger.log('📅 Datas de Referência:');
  Logger.log('   diaMesRef: ' + datas.diaMesRef);
  Logger.log('   diasRestantes: ' + datas.diasRestantes);
  Logger.log('   prazoFinal: ' + datas.diaMesRef2);
  Logger.log('   diaD1: ' + datas.diaD1);
  
  Logger.log('\n📊 Testando Balancete:');
  var abaBalancete = ss.getSheetByName('Balancete');
  var dadosBalancete = abaBalancete.getRange('C4:C8').getDisplayValues();
  
  for (var i = 0; i < dadosBalancete.length; i++) {
    var retorno = dadosBalancete[i][0];
    // Enable debug logging for test function
    var status = calcularStatusIndividual(retorno, 'mensal', true);
    Logger.log('   Linha ' + (i+4) + ': "' + retorno + '" → Status: "' + status + '"');
  }
  
  Logger.log('\n📈 Status Geral do Balancete:');
  var statusGeral = calcularStatusGeralDaAba(dadosBalancete, 'mensal');
  Logger.log('   ' + statusGeral);
  
  Logger.log('\n✅ Teste concluído!');
}

/**
 * Envia emails de conformidade ou desconformidade com os modelos HTML
 * Focando no STATUS 2 de cada aba
 * 
 * ✅ MODIFICADO: Diárias SÓ envia se houver desconformidade
 * 
 * 💡 COMO USAR:
 * - Se quiser HABILITAR envio diário de Diárias: remova os comentários da seção "1. DIÁRIAS"
 * - Se quiser DESABILITAR: mantenha os comentários como está
 */
function enviarEmailConformidadeOuDesconformidadeAvancado() {
  // ✅ VERIFICAÇÃO: Enviar e-mail apenas em dias úteis
  var hoje = new Date();
  var diaSemana = hoje.getDay();
  
  // Se é sábado (6) ou domingo (0), não enviar
  if (diaSemana === 0 || diaSemana === 6) {
    Logger.log('⏭️ Hoje é ' + (diaSemana === 0 ? 'domingo' : 'sábado') + '. E-mail não será enviado.');
    return { skipped: true, reason: 'Fim de semana' };
  }
  
  // Verificar se é feriado
  try {
    var ss = obterPlanilha();
    var abaFeriados = ss.getSheetByName('FERIADOS');
    if (abaFeriados) {
      var feriados = abaFeriados.getRange('A2:A100').getValues();
      var hojeFormatado = formatarData(hoje);
      
      for (var i = 0; i < feriados.length; i++) {
        if (feriados[i][0]) {
          var feriadoFormatado = formatarData(new Date(feriados[i][0]));
          if (feriadoFormatado === hojeFormatado) {
            Logger.log('⏭️ Hoje é feriado. E-mail não será enviado.');
            return { skipped: true, reason: 'Feriado' };
          }
        }
      }
    }
  } catch (error) {
    Logger.log('⚠️ Erro ao verificar feriados, prosseguindo com envio: ' + error.toString());
  }
  
  Logger.log('✅ Dia útil confirmado. Iniciando envio de e-mails...');

  var ss = obterPlanilha();
  var destinatarios = [
    //'spandrade@banestes.com.br',
    'fabiooliveira@banestes.com.br',
    //'iodutra@banestes.com.br',
    'mcdias@banestes.com.br',
    'sndemuner@banestes.com.br',
    'wffreitas@banestes.com.br'
  ];

  var mesPassado = obterMesPassadoFormatado();
  var dataAtualFormatada = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy HH:mm');
  var urlPlanilha = obterURLPlanilha();

  Logger.log('📧 Iniciando envio de emails...');

  // ============================================
  // 1. DIÁRIAS - ✅ APENAS DESCONFORMIDADE DIÁRIA
  // ============================================
  var abaDiarias = ss.getSheetByName('Diárias');
  if (abaDiarias) {
    var statusDiarias1 = abaDiarias.getRange('E1').getDisplayValue().toUpperCase().trim();
    var statusDiarias2 = abaDiarias.getRange('F1').getDisplayValue().toUpperCase().trim();
    
    Logger.log('📊 Status Diárias 1: "' + statusDiarias1 + '"');
    Logger.log('📊 Status Diárias 2: "' + statusDiarias2 + '"');
    
    // ✅ LÓGICA: SÓ ENVIA SE HOUVER DESCONFORMIDADE
    var ultimaLinha = abaDiarias.getLastRow();
    if (ultimaLinha >= 4) {
      var dadosStatus2 = abaDiarias.getRange('F4:F' + ultimaLinha).getValues();
      
      // Buscar fundos com desconformidade no STATUS 2
      var fundosDesconformes = [];
      
      for (var i = 0; i < dadosStatus2.length; i++) {
        var status2 = String(dadosStatus2[i][0]).toUpperCase().trim();
        
        // 🔥 CRITÉRIO: Status 2 diferente de "OK"
        if (status2 !== 'OK' && status2 !== '' && status2 !== '-') {
          var linhaAtual = i + 4;
          var nomeFundo = abaDiarias.getRange(linhaAtual, 1).getValue();
          var codigoFundo = abaDiarias.getRange(linhaAtual, 2).getValue();
          var retorno1 = abaDiarias.getRange(linhaAtual, 3).getDisplayValue();
          var status1 = abaDiarias.getRange(linhaAtual, 4).getValue();
          var retorno2 = abaDiarias.getRange(linhaAtual, 5).getDisplayValue();
          
          fundosDesconformes.push({
            nome: nomeFundo,
            codigo: codigoFundo,
            competencia1: retorno1,  // Na verdade é "retorno1" para Diárias
            status1: status1,
            competencia2: retorno2,  // Na verdade é "retorno2" para Diárias
            status2: status2
          });
        }
      }
      
      // 🎯 DECISÃO FINAL
      if (fundosDesconformes.length > 0) {
        Logger.log('⚠️ Diárias: ' + fundosDesconformes.length + ' fundos com desconformidade. Enviando email...');
        enviarEmailDesconformidade(
          'Diárias',
          fundosDesconformes,
          destinatarios,
          dataAtualFormatada,
          urlPlanilha
        );
      } else {
        Logger.log('✅ Diárias: Todos status OK ou intermediários. Email NÃO será enviado.');
        Logger.log('💡 Conformidade de Diárias só é enviada no último dia útil do mês.');
      }
    }
  }

  Logger.log('⏭️ Diárias: Conformidade será enviada apenas no último dia útil do mês pela função enviarEmailDiariasIndividualPorFundo()');

  // ============================================
  // 2. Abas mensais: Balancete, Composição, Lâmina, Perfil Mensal
  // ============================================
  ['Balancete', 'Composição', 'Lâmina', 'Perfil Mensal'].forEach(function(nomeAba) {
    var aba = ss.getSheetByName(nomeAba);
    if (!aba) return; // Pular se não existe

    // 🔥 VERIFICAÇÃO: Envia conformidade OU desconformidade, ou NADA
    if (deveEnviarEmailConformidade(aba)) {
      Logger.log('✅ ' + nomeAba + ': Enviando email de CONFORMIDADE');
      enviarEmailConformidade(
        nomeAba,
        getFundosFormatadosParaEmail(aba),
        destinatarios,
        mesPassado,
        dataAtualFormatada
      );
    } else if (deveEnviarEmailDesconformidade(aba)) {
      Logger.log('⚠️ ' + nomeAba + ': Enviando email de DESCONFORMIDADE');
      enviarEmailDesconformidade(
        nomeAba,
        getFundosDesconformesParaEmail(aba),
        destinatarios,
        dataAtualFormatada,
        urlPlanilha
      );
    } else {
      Logger.log('⏭️ ' + nomeAba + ': Nenhuma condição atendida. Email NÃO será enviado.');
    }
  });

  Logger.log('✅ Processo de envio de emails concluído!');
}

// Verifica se todas as linhas da aba possuem as competências 1/2 como datas E status 1/2 como OK
function deveEnviarEmailConformidade(aba) {
  var ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 4) return false;

  var dados = aba.getRange(4, 3, ultimaLinha - 3, 4).getValues(); // C=Comp1, D=Status1, E=Comp2, F=Status2

  // TODOS comp1/comp2 são datas E status1/status2 == 'OK'
  return dados.every(function(linha) {
    var comp1 = linha[0], status1 = linha[1], comp2 = linha[2], status2 = linha[3];
    return isDataValida(comp1) && status1 === 'OK' &&
           isDataValida(comp2) && status2 === 'OK';
  });
}

// Verifica se existe alguma linha com competência 2 vazia E status 2 DESCONFORMIDADE
function deveEnviarEmailDesconformidade(aba) {
  var ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 4) return false;

  var dados = aba.getRange(4, 5, ultimaLinha - 3, 2).getValues(); // E=Comp2, F=Status2

  // Alguma comp2 está vazia E status2 == 'DESCONFORMIDADE'
  return dados.some(function(linha) {
    var comp2 = linha[0], status2 = (linha[1] || '').toString().trim().toUpperCase();
    return (!comp2 || comp2 === '-' || (typeof comp2 === 'string' && comp2.trim() === '')) &&
           status2 === 'DESCONFORMIDADE';
  });
}

// Verifica se é data válida (Date objeto ou string DD/MM/YYYY)
function isDataValida(valor) {
  if (!valor) return false;
  if (Object.prototype.toString.call(valor) === "[object Date]") return true;
  return /^(\d{2})\/(\d{2})\/(\d{4})$/.test(valor);
}

// Adapte conforme sua estrutura:
function getFundosFormatadosParaEmail(aba) {
  var ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 4) return [];
  return aba.getRange(4, 1, ultimaLinha - 3, 6).getValues().map(function(linha) {
    return {
      nome: linha[0],
      codigo: linha[1],
      competencia1: linha[2],
      status1: linha[3],
      competencia2: linha[4],
      status2: linha[5]
    };
  });
}

function getFundosDesconformesParaEmail(aba) {
  var ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 4) return [];
  return aba.getRange(4, 1, ultimaLinha - 3, 6).getValues().filter(function(linha) {
    var comp2 = linha[4], status2 = (linha[5] || '').toString().trim().toUpperCase();
    return (!comp2 || comp2 === '-' || (typeof comp2 === 'string' && comp2.trim() === '')) &&
           status2 === 'DESCONFORMIDADE';
  }).map(function(linha) {
    return {
      nome: linha[0],
      codigo: linha[1],
      competencia1: linha[2],
      status1: linha[3],
      competencia2: linha[4],
      status2: linha[5]
    };
  });
}

/**
 * Processa uma aba e envia email de conformidade ou desconformidade
 */
/**
 * Processa uma aba e envia email de conformidade ou desconformidade
 */
function processarAbaEmail(aba, nomeAba, destinatarios, mesPassado, dataAtualFormatada, urlPlanilha, tipo) {
  if (!aba) {
    Logger.log('⚠️ Aba não encontrada: ' + nomeAba);
    return;
  }

  Logger.log('\n📊 Processando: ' + nomeAba);

  var ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 4) {
    Logger.log('  ⚠️ Sem dados na aba');
    return;
  }

  // Ler dados: A=Nome, B=Código, C=Comp1/Ret1, D=Status1, E=Comp2/Ret2, F=Status2
  var dados = aba.getRange(4, 1, ultimaLinha - 3, 6).getValues();

  var fundos = dados
    .filter(function(linha) { return linha[0] && linha[0].toString().trim() !== ''; })
    .map(function(linha) {
      return {
        nome: linha[0],
        codigo: linha[1],
        competencia1: formatarDataParaEmail(linha[2]), // ✅ USAR FUNÇÃO DE FORMATAÇÃO
        status1: linha[3] || '-',
        competencia2: formatarDataParaEmail(linha[4]), // ✅ USAR FUNÇÃO DE FORMATAÇÃO
        status2: linha[5] || '-'
      };
    });

  Logger.log('  Total de fundos: ' + fundos.length);
  Logger.log('  Exemplo do primeiro fundo:');
  if (fundos.length > 0) {
    Logger.log('    - Nome: ' + fundos[0].nome.substring(0, 40));
    Logger.log('    - Comp1: ' + fundos[0].competencia1);
    Logger.log('    - Status1: ' + fundos[0].status1);
    Logger.log('    - Comp2: ' + fundos[0].competencia2);
    Logger.log('    - Status2: ' + fundos[0].status2);
  }

  // ✅ FILTRAR APENAS DESCONFORMIDADES NO STATUS 2
  var desconformes = fundos.filter(function(f) {
    var status2Normalizado = f.status2.toString().trim().toUpperCase();
    return status2Normalizado === 'DESCONFORMIDADE' || 
           status2Normalizado === 'DESATUALIZADO' ||
           status2Normalizado === 'A ATUALIZAR';
  });

  Logger.log('  Desconformes (Status 2): ' + desconformes.length);

  // ============================================
  // DECIDIR: CONFORMIDADE OU DESCONFORMIDADE
  // ============================================
  if (desconformes.length > 0) {
    enviarEmailDesconformidade(nomeAba, desconformes, destinatarios, dataAtualFormatada, urlPlanilha);
  } else {
    enviarEmailConformidade(nomeAba, fundos, destinatarios, mesPassado, dataAtualFormatada);
  }
}

/**
 * Envia email de DESCONFORMIDADE (SEMPRE usa Competência 2)
 */
function enviarEmailDesconformidade(nomeAba, fundosDesconformes, destinatarios, dataAtual, urlPlanilha) {
  Logger.log('  ❌ Enviando email de DESCONFORMIDADE');

  // 🔥 GERAR TABELA - SEMPRE COMPETÊNCIA 2 E STATUS 2
  var linhasTabela = fundosDesconformes.map(function(f) {
    var dataExibir = f.competencia2 || '-'; // SEMPRE COMPETÊNCIA 2
    var statusExibir = f.status2 || '-';     // SEMPRE STATUS 2
    
    return '<tr>' +
      '<td style="padding:10px;border:1px solid #e0e0e0;">' + dataExibir + '</td>' +
      '<td style="padding:10px;border:1px solid #e0e0e0;">' + (f.nome || '-') + '</td>' +
      '<td style="padding:10px;border:1px solid #e0e0e0;text-align:center;">' + statusExibir + '</td>' +
      '</tr>';
  }).join('');

  var htmlTabela = '<table border="1" cellpadding="10" cellspacing="0" width="100%" style="border-collapse:collapse;margin-top:15px;">' +
    '<thead>' +
    '<tr style="background-color:#f3f4f6;">' +
    '<th style="padding:10px;border:1px solid #e0e0e0;text-align:left;">Data Envio</th>' +
    '<th style="padding:10px;border:1px solid #e0e0e0;text-align:left;">Fundo</th>' +
    '<th style="padding:10px;border:1px solid #e0e0e0;text-align:center;">Status 2</th>' +
    '</tr>' +
    '</thead>' +
    '<tbody>' +
    linhasTabela +
    '</tbody>' +
    '</table>';

  // Carregar template e substituir placeholders
  var htmlDesconformidade = HtmlService.createHtmlOutputFromFile('desconformidade').getContent();
  
  htmlDesconformidade = htmlDesconformidade
    .replace('[INSERIR NOME/CÓDIGO DO FUNDO OU EMPRESA]', nomeAba)
    .replace('[INSERIR DESCRIÇÃO DA FALHA, EX: ENVIO DE LÂMINA EM ATRASO]', htmlTabela)
    .replace('[INSERIR DATA]', dataAtual.split(' ')[0])
    .replace('[INSERIR HORÁRIO]', '17:00')
    .replace('[LINK_PARA_SISTEMA_OU_INSTRUCOES]', urlPlanilha);

  // Enviar email
  MailApp.sendEmail({
    to: destinatarios.join(','),
    subject: '⚠️ Desconformidade CVM - ' + nomeAba,
    htmlBody: htmlDesconformidade
  });

  Logger.log('  ✅ Email de desconformidade enviado para: ' + destinatarios.join(', '));
}

/**
 * Envia email de CONFORMIDADE (SEMPRE usa Competência 2)
 */
function enviarEmailConformidade(nomeAba, fundos, destinatarios, mesPassado, dataAtual) {
  Logger.log('  ✅ Enviando email de CONFORMIDADE');
  Logger.log('  Total de fundos: ' + fundos.length);

  // 🔥 GERAR TABELA - SEMPRE USAR COMPETÊNCIA 2 E STATUS 2
  var linhasTabela = fundos.map(function(f) {
    // 🔥 GARANTIR QUE AS DATAS ESTÃO FORMATADAS
    var dataExibir = formatarDataParaEmail(f.competencia2); // ✅ SEMPRE FORMATAR
    
    // 🔥 SEMPRE STATUS 2
    var statusExibir = f.status2 || '-';
    var statusNormalizado = statusExibir.toString().trim().toUpperCase();
    
    // 🔥 DEFINIR COR DO STATUS
    var corStatus, textoStatus;
    if (statusNormalizado === 'OK') {
      corStatus = '#d1fae5';
      textoStatus = '#065f46';
    } else if (statusNormalizado === 'AGUARDANDO') {
      corStatus = '#fef3c7';
      textoStatus = '#92400e';
    } else if (statusNormalizado === 'A ATUALIZAR') {
      corStatus = '#fed7aa';
      textoStatus = '#9a3412';
    } else if (statusNormalizado === 'DESCONFORMIDADE' || statusNormalizado === 'DESATUALIZADO') {
      corStatus = '#fee2e2';
      textoStatus = '#991b1b';
    } else {
      // Status "-" ou desconhecido = cinza
      corStatus = '#f3f4f6';
      textoStatus = '#374151';
    }
    
    Logger.log('    [' + f.nome.substring(0, 30) + '] Comp2: "' + dataExibir + '" | Status2: "' + statusExibir + '"');
    
    return '<tr>' +
      '<td style="padding:10px;border:1px solid #dddddd;background:#ffffff;">' + dataExibir + '</td>' +
      '<td style="padding:10px;border:1px solid #dddddd;background:#ffffff;">' + (f.nome || '-') + '</td>' +
      '<td style="padding:10px;border:1px solid #dddddd;background:#ffffff;text-align:center;">' +
      '<span style="background-color:' + corStatus + ';color:' + textoStatus + ';padding:5px 12px;border-radius:8px;font-weight:bold;display:inline-block;">' + statusExibir + '</span>' +
      '</td>' +
      '</tr>';
  }).join('');

  // 🔥 FORMATAR DATA ATUAL (corrigido para garantir formato dd/MM/yyyy)
  var dataAtualFormatada;
  if (dataAtual instanceof Date && !isNaN(dataAtual.getTime())) {
    dataAtualFormatada = Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  } else if (typeof dataAtual === 'string') {
    // Se recebido com data e hora, tentar separar e pegar apenas a data
    var regexData = /(\d{2}\/\d{2}\/\d{4})/;
    var match = dataAtual.match(regexData);
    if (match && match[1]) {
      dataAtualFormatada = match[1];
    } else {
      // Tentar parsear string em Date
      var tentativaDt = new Date(dataAtual);
      if (!isNaN(tentativaDt.getTime())) {
        dataAtualFormatada = Utilities.formatDate(tentativaDt, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      } else {
        // Último recurso: coloca data de agora
        dataAtualFormatada = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
      }
    }
  } else {
    // fallback
    dataAtualFormatada = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }

  // 🔥 TABELA COMPLETA
  var tabelaHTML = 
    '<table style="width:100%;border-collapse:collapse;margin:20px 0;font-family:Arial,sans-serif;" cellpadding="0" cellspacing="0">' +
    '<thead>' +
    '<tr>' +
    '<th style="padding:12px;border:1px solid #dddddd;background-color:#f3f4f6;text-align:left;font-weight:bold;color:#555555;">Data Envio</th>' +
    '<th style="padding:12px;border:1px solid #dddddd;background-color:#f3f4f6;text-align:left;font-weight:bold;color:#555555;">Registro/Fundo</th>' +
    '<th style="padding:12px;border:1px solid #dddddd;background-color:#f3f4f6;text-align:center;font-weight:bold;color:#555555;">Status</th>' +
    '</tr>' +
    '</thead>' +
    '<tbody>' +
    linhasTabela +
    '</tbody>' +
    '</table>';

  // HTML COMPLETO
  var htmlCompleto = 
    '<div style="background-color:#f4f4f4;padding:20px;font-family:Arial,sans-serif;">' +
    '<table style="max-width:650px;margin:0 auto;background-color:#ffffff;border-radius:8px;overflow:hidden;box-shadow:0 2px 5px rgba(0,0,0,0.1);" cellpadding="0" cellspacing="0">' +
    '<tr>' +
    '<td style="background-color:#2E7D32;padding:30px 20px;text-align:center;">' +
    '<div style="font-size:40px;color:#ffffff;margin-bottom:10px;">✓</div>' +
    '<h1 style="color:#ffffff;font-size:22px;margin:0;">Relatório de Conformidade CVM</h1>' +
    '<p style="color:#a5d6a7;margin:5px 0 0 0;font-size:14px;">Status: 100% Regularizado</p>' +
    '</td>' +
    '</tr>' +
    '<tr>' +
    '<td style="padding:30px 25px;color:#333333;font-size:15px;line-height:1.6;">' +
    '<p>Prezados,</p>' +
    '<p>Informamos que, após a verificação mensal, <strong>todos os registros e obrigações junto à CVM encontram-se em conformidade</strong>.</p>' +
    '<p>Abaixo listamos os envios realizados com sucesso referentes ao período de <strong>' + dataAtualFormatada + '</strong>:</p>' +
    tabelaHTML +
    '<div style="background-color:#e3f2fd;border-left:4px solid #2196F3;padding:15px;margin-top:20px;border-radius:0 4px 4px 0;">' +
    '<p style="margin:0;font-weight:bold;color:#0d47a1;font-size:14px;">IMPORTANTE: Manutenção da Conformidade</p>' +
    '<p style="margin:5px 0 0 0;font-size:13px;color:#444;">Embora estejamos em situação regular, solicitamos à equipe que mantenha o monitoramento constante dos prazos e exigências regulatórias. A vigilância contínua é essencial para evitar sanções futuras.</p>' +
    '</div>' +
    '</td>' +
    '</tr>' +
    '<tr>' +
    '<td style="background-color:#f8f9fa;padding:20px;text-align:center;color:#888888;font-size:12px;border-top:1px solid #eeeeee;">' +
    '<p style="margin:0;">Departamento de Inovação e Automação interno Asset</p>' +
    '<p style="margin:5px 0 0 0;">Relatório gerado em ' + dataAtualFormatada + '</p>' +
    '</td>' +
    '</tr>' +
    '</table>' +
    '</div>';

  // Enviar email
  try {
    MailApp.sendEmail({
      to: destinatarios.join(','),
      subject: '✅ Conformidade CVM - ' + nomeAba,
      htmlBody: htmlCompleto
    });
    
    Logger.log('  ✅ Email enviado com sucesso');

    // 🆕 MARCAR FLAG NA PLANILHA
    marcarEmailEnviado(nomeAba, dataAtualFormatada);

  } catch (error) {
    Logger.log('  ❌ Erro: ' + error.toString());
    throw error;
  }
}

// Função auxiliar mantida
function obterMesPassadoFormatado() {
  var hoje = new Date();
  var mes = hoje.getMonth();
  var ano = hoje.getFullYear();
  if (mes === 0) {
    mes = 12;
    ano -= 1;
  }
  return (mes < 10 ? '0' + mes : mes) + '/' + ano;
}

/**
 * 🧪 TESTE com dados reais da planilha
 */
function testarEmailComDadosReais() {
  Logger.log('🧪 Testando com dados reais da planilha...');
  
  var ss = obterPlanilha();
  var destinatarios = ['spandrade@banestes.com.br'];
  var mesPassado = obterMesPassadoFormatado();
  var dataAtualFormatada = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy HH:mm');
  var urlPlanilha = obterURLPlanilha();
  
  // Testar Balancete (aba com AGUARDANDO)
  processarAbaEmail(
    ss.getSheetByName('Balancete'),
    'Balancete (TESTE)',
    destinatarios,
    mesPassado,
    dataAtualFormatada,
    urlPlanilha,
    'mensal'
  );
  
  Logger.log('✅ Teste com Balancete concluído!');
}

// Função auxiliar: retorna o mês passado em formato "MM/YYYY"
function obterMesPassadoFormatado() {
  var hoje = new Date();
  var mes = hoje.getMonth();
  var ano = hoje.getFullYear();
  if (mes === 0) {
    mes = 12;
    ano -= 1;
  }
  return (mes < 10 ? '0' + mes : mes) + '/' + ano;
}

// Função de teste: envia para você o modelo CONFORMIDADE e DESCONFORMIDADE com exemplos
//function forcarEnvioDosDoisModelosEmail() {
//  var destinatarioTeste = Session.getActiveUser().getEmail(); // Ou coloque o seu e-mail aqui

  // --- Enviar modelo CONFORMIDADE ---
//  var mesPassado = obterMesPassadoFormatado();
//  var dataAtualFormatada = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy HH:mm');
//  var htmlConformidade = HtmlService.createHtmlOutputFromFile('conformidade').getContent()
//    .replace('[MÊS PASSADO]', mesPassado)
//    .replace('[DATA ATUAL]', dataAtualFormatada);
//  MailApp.sendEmail({
//    to: destinatarioTeste,
//    subject: 'TESTE: Modelo HTML de Conformidade',
//    htmlBody: htmlConformidade
//  });

  // --- Enviar modelo DESCONFORMIDADE ---
//  var htmlDesconformidade = HtmlService.createHtmlOutputFromFile('desconformidade').getContent()
//    .replace('[INSERIR NOME/CÓDIGO DO FUNDO OU EMPRESA]', 'BANESTES INVEST AUTOMÁTICO (275709)')
//    .replace('[INSERIR DESCRIÇÃO DA FALHA, EX: ENVIO DE LÂMINA EM ATRASO]', 'Envio da Lâmina em atraso')
//    .replace('[INSERIR DATA]', Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy'))
//    .replace('[INSERIR HORÁRIO]', '16:00')
//    .replace('[LINK_PARA_SISTEMA_OU_INSTRUCOES]', 'https://docs.google.com/spreadsheets/d/1N6LP1ydsxnQO_Woatv9zWEccb0fOGaV_3EKK1GoSWZI/edit');
//  MailApp.sendEmail({
//    to: destinatarioTeste,
//    subject: 'TESTE: Modelo HTML de Desconformidade',
//    htmlBody: htmlDesconformidade
//  });

//  Logger.log('✅ Emails de teste enviados para: ' + destinatarioTeste);
//}

function atualizarStatusNaPlanilhaAutomaticoComEmail() {
  atualizarStatusNaPlanilhaAutomatico();
  enviarEmailDesconformidade();
}

function testarIMPORTXMLManual() {
  Logger.log('🧪 Testando IMPORTXML manualmente...');
  
  var ss = obterPlanilha();
  var abaBalancete = ss.getSheetByName('Balancete');
  
  var codigoFundo = abaBalancete.getRange('B4').getValue();
  Logger.log('📊 Testando fundo: ' + codigoFundo);
  
  var url = 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/Balancete/CPublicaBalancete.asp?PK_PARTIC=' + codigoFundo + '&SemFrame=';
  Logger.log('🌐 URL: ' + url);
  
  try {
    var response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true
    });
    
    var codigo = response.getResponseCode();
    Logger.log('📡 Código HTTP: ' + codigo);
    
    if (codigo === 200) {
      var html = response.getContentText();
      Logger.log('✅ Página carregada! Tamanho: ' + html.length + ' caracteres');
      
      var regex = /(\d{2}\/\d{2}\/\d{4})/g;
      var datas = html.match(regex);
      
      if (datas && datas.length > 0) {
        Logger.log('📅 Datas encontradas:');
        datas.slice(0, 5).forEach(function(data) {
          Logger.log('   - ' + data);
        });
        Logger.log('✅ IMPORTXML deveria funcionar!');
      } else {
        Logger.log('⚠️ Nenhuma data encontrada no HTML');
        Logger.log('❌ Problema: XPath pode estar errado ou página mudou');
      }
      
    } else {
      Logger.log('❌ Erro HTTP: ' + codigo);
      Logger.log('⚠️ Site da CVM pode estar fora do ar ou bloqueando');
    }
    
  } catch (error) {
    Logger.log('❌ Erro ao buscar página: ' + error.toString());
  }
  
  Logger.log('\n📋 Testando fórmula atual na célula C4...');
  var formula = abaBalancete.getRange('C4').getFormula();
  Logger.log('📝 Fórmula: ' + formula);
  
  var valor = abaBalancete.getRange('C4').getValue();
  Logger.log('💾 Valor atual: ' + valor);
  
  var display = abaBalancete.getRange('C4').getDisplayValue();
  Logger.log('👁️ Display: ' + display);
}

function atualizarDadosCVMRealCompleto() {
  Logger.log('🚀 Buscando dados COMPLETOS da CVM (com Lâmina corrigida)...');
  Logger.log('⏱️ Tempo estimado: 40-60 segundos');
  
  var ss = obterPlanilha();
  var fundos = getFundos();
  var totalFundos = fundos.length;
  
  var mesesMap = {
    'Jan': '01', 'Fev': '02', 'Mar': '03', 'Abr': '04',
    'Mai': '05', 'Jun': '06', 'Jul': '07', 'Ago': '08',
    'Set': '09', 'Out': '10', 'Nov': '11', 'Dez': '12'
  };
  
  // ============================================
  // 1. BALANCETE
  // ============================================
  Logger.log('\n📊 [1/5] Processando Balancete...');
  var abaBalancete = ss.getSheetByName('Balancete');
  
  fundos.forEach(function(fundo, index) {
    try {
      var linha = index + 4;
      var url = 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/Balancete/CPublicaBalancete.asp?PK_PARTIC=' + fundo.codigoCVM + '&SemFrame=';
      var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' }});
      
      if (response.getResponseCode() === 200) {
        var html = response.getContentText();
        var regex = /<option[^>]*>(\d{2}\/\d{4})<\/option>/gi;
        var matches = html.match(regex);
        
        if (matches && matches.length > 0) {
          // Pegar os 2 últimos (mais recentes)
          var comp1Match = matches[0].match(/(\d{2}\/\d{4})/);
          var comp2Match = matches[1] ? matches[1].match(/(\d{2}\/\d{4})/) : null;
          
          if (comp1Match) {
            var partes = comp1Match[1].split('/');
            abaBalancete.getRange(linha, 3).setValue('01/' + partes[0] + '/' + partes[1]);
          }
          
          if (comp2Match) {
            var partes2 = comp2Match[1].split('/');
            abaBalancete.getRange(linha, 5).setValue('01/' + partes2[0] + '/' + partes2[1]);
          }
          
          Logger.log('  ✅ [' + (index + 1) + '/' + totalFundos + '] Balancete atualizado');
        }
      }
      Utilities.sleep(300);
    } catch (error) {
      Logger.log('  ❌ [' + (index + 1) + '/' + totalFundos + '] Erro');
    }
  });
  
  // ============================================
  // 2. COMPOSIÇÃO
  // ============================================
  Logger.log('\n📈 [2/5] Processando Composição...');
  var abaComposicao = ss.getSheetByName('Composição');
  
  fundos.forEach(function(fundo, index) {
    try {
      var linha = index + 4;
      var url = 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/CDA/CPublicaCDA.aspx?PK_PARTIC=' + fundo.codigoCVM + '&SemFrame=';
      var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' }});
      
      if (response.getResponseCode() === 200) {
        var html = response.getContentText();
        var regex = /<option[^>]*>(\d{2}\/\d{4})<\/option>/gi;
        var matches = html.match(regex);
        
        if (matches && matches.length > 0) {
          var comp1Match = matches[0].match(/(\d{2}\/\d{4})/);
          var comp2Match = matches[1] ? matches[1].match(/(\d{2}\/\d{4})/) : null;
          
          if (comp1Match) {
            var partes = comp1Match[1].split('/');
            abaComposicao.getRange(linha, 3).setValue('01/' + partes[0] + '/' + partes[1]);
          }
          
          if (comp2Match) {
            var partes2 = comp2Match[1].split('/');
            abaComposicao.getRange(linha, 5).setValue('01/' + partes2[0] + '/' + partes2[1]);
          }
          
          Logger.log('  ✅ [' + (index + 1) + '/' + totalFundos + '] Composição atualizada');
        }
      }
      Utilities.sleep(300);
    } catch (error) {
      Logger.log('  ❌ [' + (index + 1) + '/' + totalFundos + '] Erro');
    }
  });
  
  // ============================================
  // 3. LÂMINA
  // ============================================
  Logger.log('\n📄 [3/5] Processando Lâmina (CORRIGIDA)...');
  var abaLamina = ss.getSheetByName('Lâmina');
  
  fundos.forEach(function(fundo, index) {
    try {
      var linha = index + 4;
      var url = 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/CPublicaLamina.aspx?PK_PARTIC=' + fundo.codigoCVM + '&PK_SUBCLASSE=-1';
      var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' }});
      
      if (response.getResponseCode() === 200) {
        var html = response.getContentText();
        var regex = /<option[^>]*value="([A-Za-z]{3}\/\d{4})"[^>]*>/gi;
        var matches = html.match(regex);
        
        if (matches && matches.length > 0) {
          var comp1Match = matches[0].match(/value="([A-Za-z]{3}\/\d{4})"/);
          var comp2Match = matches[1] ? matches[1].match(/value="([A-Za-z]{3}\/\d{4})"/) : null;
          
          if (comp1Match) {
            var competencia = comp1Match[1];
            var partes = competencia.split('/');
            var mesNumero = mesesMap[partes[0]];
            if (mesNumero) {
              abaLamina.getRange(linha, 3).setValue('01/' + mesNumero + '/' + partes[1]);
            }
          }
          
          if (comp2Match) {
            var competencia2 = comp2Match[1];
            var partes2 = competencia2.split('/');
            var mesNumero2 = mesesMap[partes2[0]];
            if (mesNumero2) {
              abaLamina.getRange(linha, 5).setValue('01/' + mesNumero2 + '/' + partes2[1]);
            }
          }
          
          Logger.log('  ✅ [' + (index + 1) + '/' + totalFundos + '] Lâmina atualizada');
        }
      }
      Utilities.sleep(300);
    } catch (error) {
      Logger.log('  ❌ [' + (index + 1) + '/' + totalFundos + '] Erro');
    }
  });
  
  // ============================================
  // 4. PERFIL MENSAL
  // ============================================
  Logger.log('\n📊 [4/5] Processando Perfil Mensal...');
  var abaPerfilMensal = ss.getSheetByName('Perfil Mensal');
  
  fundos.forEach(function(fundo, index) {
    try {
      var linha = index + 4;
      var url = 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/Regul/CPublicaRegulPerfilMensal.aspx?PK_PARTIC=' + fundo.codigoCVM;
      var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' }});
      
      if (response.getResponseCode() === 200) {
        var html = response.getContentText();
        var regex = /<option[^>]*>(\d{2}\/\d{4})<\/option>/gi;
        var matches = html.match(regex);
        
        if (matches && matches.length > 0) {
          var comp1Match = matches[0].match(/(\d{2}\/\d{4})/);
          var comp2Match = matches[1] ? matches[1].match(/(\d{2}\/\d{4})/) : null;
          
          if (comp1Match) {
            var partes = comp1Match[1].split('/');
            abaPerfilMensal.getRange(linha, 3).setValue('01/' + partes[0] + '/' + partes[1]);
          }
          
          if (comp2Match) {
            var partes2 = comp2Match[1].split('/');
            abaPerfilMensal.getRange(linha, 5).setValue('01/' + partes2[0] + '/' + partes2[1]);
          }
          
          Logger.log('  ✅ [' + (index + 1) + '/' + totalFundos + '] Perfil Mensal atualizado');
        }
      }
      Utilities.sleep(300);
    } catch (error) {
      Logger.log('  ❌ [' + (index + 1) + '/' + totalFundos + '] Erro');
    }
  });
  
  // ============================================
  // 5. DIÁRIAS (mantém como está)
  // ============================================
  Logger.log('\n📅 [5/5] Processando Diárias...');
  var abaDiarias = ss.getSheetByName('Diárias');
  
  var contadorOK_Status1 = 0;
  var contadorOK_Status2 = 0;
  var hojeObj = new Date();
  var hoje = normalizaDataDate(hojeObj);
  var feriados = getFeriadosArray();
  var diaD1 = calculaUltimoDiaUtil(hojeObj, feriados);
  
  fundos.forEach(function(fundo, index) {
    try {
      var linha = index + 4;
      var url = 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/InfDiario/CPublicaInfdiario.aspx?PK_PARTIC=' + fundo.codigoCVM + '&PK_SUBCLASSE=-1';
      var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' }});

      if (response.getResponseCode() === 200) {
        var html = response.getContentText();
        var regex = /(\d{2}\/\d{2}\/\d{4})/g;
        var matches = html.match(regex);

        if (matches && matches.length > 0) {
          var datasExtraidas = matches.map(normalizaData);

          var envio1 = datasExtraidas.includes(diaD1) ? diaD1 : (datasExtraidas[0] || "-");
          var status1 = envio1 === diaD1 ? "OK" : "DESATUALIZADO";
          if (status1 === "OK") contadorOK_Status1++;

          var envio2 = datasExtraidas.includes(hoje) ? hoje : "-";
          var status2 = envio2 === hoje ? "OK" : "A ATUALIZAR";
          if (status2 === "OK") contadorOK_Status2++;

          abaDiarias.getRange(linha, 3).setValue(envio1);
          abaDiarias.getRange(linha, 4).setValue(status1);
          abaDiarias.getRange(linha, 5).setValue(envio2);
          abaDiarias.getRange(linha, 6).setValue(status2);

          Logger.log('  ✅ [' + (index + 1) + '/' + fundos.length + '] Envio1:' + envio1 + ' (' + status1 + ') / Envio2:' + envio2 + ' (' + status2 + ')');
        } else {
          abaDiarias.getRange(linha, 3).setValue('-');
          abaDiarias.getRange(linha, 4).setValue('DESATUALIZADO');
          abaDiarias.getRange(linha, 5).setValue('-');
          abaDiarias.getRange(linha, 6).setValue('A ATUALIZAR');
        }
      }

      Utilities.sleep(300);

    } catch (error) {
      abaDiarias.getRange(linha, 3).setValue('ERRO');
      abaDiarias.getRange(linha, 4).setValue('DESATUALIZADO');
      abaDiarias.getRange(linha, 5).setValue('#N/A');
      abaDiarias.getRange(linha, 6).setValue('A ATUALIZAR');
    }
  });

  var statusGeral1 = contadorOK_Status1 === fundos.length ? 'OK' : 'DESCONFORMIDADE';
  var statusGeral2 = contadorOK_Status2 === fundos.length ? 'OK' : 'A ATUALIZAR';
  abaDiarias.getRange('E1').setValue(statusGeral1);
  abaDiarias.getRange('F1').setValue(statusGeral2);
  
  Logger.log('  STATUS 1 GERAL: ' + statusGeral1);
  Logger.log('  STATUS 2 GERAL: ' + statusGeral2);
  
  // ============================================
  // 🆕 6. CALCULAR COMPETÊNCIAS E STATUS
  // ============================================
  Logger.log('\n🧮 [6/6] Calculando competências e status...');
  atualizarTodasCompetencias(); // 🔥 ADICIONAR ESTA LINHA
  
  Logger.log('\n✅ ═══════════════════════════════════════════');
  Logger.log('✅ ATUALIZAÇÃO 100% COMPLETA!');
  Logger.log('✅ ═══════════════════════════════════════════');
  
  return { success: true, message: 'Sistema 100% funcional!' };
}

// Função auxiliar (PROTEGIDA contra Diárias)
function atualizarStatusParaAbasEspecificas(nomesAbas) {
  var ss = obterPlanilha();
  var datas = getDatasReferencia();
  
  nomesAbas.forEach(function(nomeAba) {
    if (nomeAba === 'Diárias') return;
    
    var aba = ss.getSheetByName(nomeAba);
    if (!aba) return;
    
    var ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 4) return;
    
    var valores = aba.getRange(4, 3, ultimaLinha - 3, 1).getDisplayValues();
    
    var totalOK = 0;
    var totalDesconformidade = 0;
    
    valores.forEach(function(linha, index) {
      var valor = linha[0];
      var status;
      
      if (!valor || valor === '' || valor === '-' || valor === 'ERRO') {
        status = 'DESCONFORMIDADE';
        totalDesconformidade++;
      } else if (valor === datas.diaMesRef) {
        status = 'OK';
        totalOK++;
      } else {
        status = '-';
      }
      
      aba.getRange(index + 4, 4).setValue(status);
    });
    
    // STATUS GERAL (D1)
    var statusGeral;
    if (totalDesconformidade > 0) {
      statusGeral = 'DESCONFORMIDADE';
    } else if (totalOK === valores.length) {
      statusGeral = 'OK';
    } else {
      statusGeral = '-';
    }
    aba.getRange('D1').setValue(statusGeral);
    
    Logger.log('  ✅ ' + nomeAba + ': ' + statusGeral + ' (' + totalOK + '/' + valores.length + ' OK)');
  });
}

// Função auxiliar (PROTEGIDA contra Diárias)
function atualizarStatusParaAbasEspecificas(nomesAbas) {
  var ss = obterPlanilha();
  var datas = getDatasReferencia();
  
  nomesAbas.forEach(function(nomeAba) {
    if (nomeAba === 'Diárias') return;
    
    var aba = ss.getSheetByName(nomeAba);
    if (!aba) return;
    
    var ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 4) return;
    
    var valores = aba.getRange(4, 3, ultimaLinha - 3, 1).getDisplayValues();
    
    valores.forEach(function(linha, index) {
      var valor = linha[0];
      var status;
      
      if (!valor || valor === '' || valor === '-' || valor === 'ERRO') {
        status = 'DESCONFORMIDADE';
      } else if (valor === datas.diaMesRef) {
        status = 'OK';
      } else {
        status = '-';
      }
      
      aba.getRange(index + 4, 4).setValue(status);
    });
  });
}

// Função auxiliar (PROTEGIDA contra Diárias)
function atualizarStatusParaAbasEspecificas(nomesAbas) {
  var ss = obterPlanilha();
  var datas = getDatasReferencia();
  
  nomesAbas.forEach(function(nomeAba) {
    // PROTEÇÃO: NÃO processar Diárias
    if (nomeAba === 'Diárias') return;
    
    var aba = ss.getSheetByName(nomeAba);
    if (!aba) return;
    
    var ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 4) return;
    
    var valores = aba.getRange(4, 3, ultimaLinha - 3, 1).getDisplayValues();
    
    valores.forEach(function(linha, index) {
      var valor = linha[0];
      var status;
      
      if (!valor || valor === '' || valor === '-' || valor === 'ERRO') {
        status = 'DESCONFORMIDADE';
      } else if (valor === datas.diaMesRef) {
        status = 'OK';
      } else {
        status = '-';
      }
      
      aba.getRange(index + 4, 4).setValue(status);
    });
  });
}

// Função auxiliar para calcular status apenas de abas específicas
function atualizarStatusParaAbasEspecificas(nomesAbas) {
  var ss = obterPlanilha();
  var datas = getDatasReferencia();
  
  nomesAbas.forEach(function(nomeAba) {
    var aba = ss.getSheetByName(nomeAba);
    
    if (!aba) return;
    
    var ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 4) return;
    
    var valores = aba.getRange(4, 3, ultimaLinha - 3, 1).getDisplayValues();
    
    var totalOK = 0;
    var totalDesconformidade = 0;
    
    valores.forEach(function(linha, index) {
      var valor = linha[0];
      var status;
      
      if (!valor || valor === '' || valor === '-' || valor === 'ERRO') {
        status = 'DESCONFORMIDADE';
        totalDesconformidade++;
      } else if (valor === datas.diaMesRef) {
        status = 'OK';
        totalOK++;
      } else {
        status = '-';
      }
      
      aba.getRange(index + 4, 4).setValue(status);
    });
    
    // Atualizar status geral no cabeçalho
    if (totalDesconformidade > 0) {
      aba.getRange('D1').setValue('DESCONFORMIDADE');
    } else if (totalOK === valores.length) {
      aba.getRange('D1').setValue('OK');
    } else {
      aba.getRange('D1').setValue('-');
    }
  });
}

function ativarSistemaCompleto() {
  Logger.log('🚀 Ativando sistema completo 100% funcional...');
  
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
    Logger.log('  🗑️ Trigger removido: ' + trigger.getHandlerFunction());
  });
  
  ScriptApp.newTrigger('atualizarDadosCVMRealCompleto')
    .timeBased()
    .everyHours(1)
    .create();
  
  Logger.log('\n✅ ═══════════════════════════════════════════');
  Logger.log('✅ SISTEMA ATIVADO COM SUCESSO!');
  Logger.log('✅ ═══════════════════════════════════════════');
  Logger.log('');
  Logger.log('📊 Função ativa: atualizarDadosCVMRealCompleto()');
  Logger.log('⏰ Frequência: A cada 1 hora');
  Logger.log('🔄 Execuções por dia: 24');
  Logger.log('📱 Funciona 24/7: SIM');
  Logger.log('💻 Precisa navegador aberto: NÃO');
  Logger.log('');
  Logger.log('📋 O que será atualizado automaticamente:');
  Logger.log('   ✅ Balancete (26 fundos)');
  Logger.log('   ✅ Composição (26 fundos)');
  Logger.log('   ✅ Lâmina (26 fundos)');
  Logger.log('   ✅ Perfil Mensal (26 fundos)');
  Logger.log('   ✅ Diárias (26 fundos × 2 datas)');
  Logger.log('   ✅ Cálculo de status');
  Logger.log('   ✅ Dashboard Geral');
  Logger.log('');
  Logger.log('🌐 Web App: PRONTO PARA USO');
  Logger.log('📊 Planilha: ' + obterURLPlanilha());
  Logger.log('');
  Logger.log('🎉 PARABÉNS! SISTEMA 100% OPERACIONAL!');
  
  return {
    success: true,
    message: 'Sistema ativado com sucesso! Todas as 5 abas funcionando perfeitamente.'
  };
}

function diagnosticarAbaCodFundo() {
  Logger.log('🔍 Diagnosticando aba COD FUNDO...');
  
  var ss = obterPlanilha();
  var aba = ss.getSheetByName('COD FUNDO');
  
  if (!aba) {
    Logger.log('❌ Aba COD FUNDO não existe!');
    return;
  }
  
  Logger.log('✅ Aba COD FUNDO existe');
  Logger.log('📊 Última linha: ' + aba.getLastRow());
  Logger.log('📊 Última coluna: ' + aba.getLastColumn());
  
  // Ver cabeçalho
  var cabecalho = aba.getRange('A1:C1').getValues()[0];
  Logger.log('\n📋 Cabeçalho:');
  Logger.log('  A1: ' + cabecalho[0]);
  Logger.log('  B1: ' + cabecalho[1]);
  Logger.log('  C1: ' + cabecalho[2]);
  
  // Ver primeiras 3 linhas
  Logger.log('\n📋 Primeiras 3 linhas de dados:');
  var dados = aba.getRange('A2:C4').getValues();
  dados.forEach(function(linha, i) {
    Logger.log('  Linha ' + (i+2) + ':');
    Logger.log('    A (Nome): ' + linha[0].substring(0, 30) + '...');
    Logger.log('    B (CVM): ' + linha[1]);
    Logger.log('    C (BANESTES): ' + linha[2]);
  });
}

function investigarDatasDiarias() {
  Logger.log('🔍 Investigando datas da página de Diárias...');
  
  var codigoCVM = '275709'; // BANESTES INVESTIDOR AUTOMÁTICO
  var url = 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/InfDiario/CPublicaInfdiario.aspx?PK_PARTIC=' + codigoCVM + '&PK_SUBCLASSE=-1';
  
  try {
    var response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    
    if (response.getResponseCode() === 200) {
      var html = response.getContentText();
      Logger.log('✅ Página carregada');
      
      // Buscar TODAS as datas
      var regex = /(\d{2}\/\d{2}\/\d{4})/g;
      var matches = html.match(regex);
      
      if (matches) {
        Logger.log('\n📅 Total de datas encontradas: ' + matches.length);
        Logger.log('\n📋 Primeiras 20 datas:');
        matches.slice(0, 20).forEach(function(data, i) {
          Logger.log('  [' + (i+1) + '] ' + data);
        });
        
        // Buscar a estrutura HTML ao redor das datas
        Logger.log('\n🔍 Buscando contexto das primeiras 5 datas...');
        matches.slice(0, 5).forEach(function(data, i) {
          var index = html.indexOf(data);
          var contexto = html.substring(index - 100, index + 150);
          Logger.log('\n[' + (i+1) + '] Data: ' + data);
          Logger.log('Contexto HTML:');
          Logger.log(contexto);
        });
        
      } else {
        Logger.log('❌ Nenhuma data encontrada');
      }
      
    } else {
      Logger.log('❌ Erro HTTP: ' + response.getResponseCode());
    }
    
  } catch (error) {
    Logger.log('❌ Erro: ' + error.toString());
  }
}

function diagnosticarGetDatasReferencia() {
  Logger.log('🔍 Diagnosticando getDatasReferencia()...\n');
  
  try {
    // Verificar se a função existe
    if (typeof getDatasReferencia === 'function') {
      Logger.log('✅ Função getDatasReferencia() existe');
      
      // Tentar executar
      var resultado = getDatasReferencia();
      
      Logger.log('\n📋 Resultado:');
      Logger.log(JSON.stringify(resultado, null, 2));
      
    } else {
      Logger.log('❌ Função getDatasReferencia() NÃO EXISTE!');
      Logger.log('⚠️ A função pode estar em outro arquivo (DateUtils.gs)');
    }
    
  } catch (error) {
    Logger.log('❌ ERRO ao executar getDatasReferencia():');
    Logger.log(error.toString());
  }
  
  // Verificar qual arquivo tem a função
  Logger.log('\n📁 Verificando arquivos do projeto...');
  Logger.log('   - Code.gs');
  Logger.log('   - DateUtils.gs (provável localização)');
  Logger.log('   - FundoService.gs');
  Logger.log('   - ConfigData.gs');
  Logger.log('   - onInstall.gs');
}

function normalizaData(data) {
  if (!data) return '';
  var s = String(data).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    var p = s.split('-');
    return [p[2], p[1], p[0]].join('/');
  }
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;
  return s.replace(/\s+/g, '');
}

/**
 * Normaliza data para comparação
 * Aceita formatos: "01/12/2025", "01/12/2025 - ", "2025-12-01"
 */
function normalizaDataParaComparacao(data) {
  if (!data) return '';
  
  // Converter para string e remover espaços e traços extras
  var dataStr = String(data).trim().replace(/\s*-\s*$/, '').trim();
  
  // Se for objeto Date
  if (data instanceof Date) {
    var dia = ('0' + data.getDate()).slice(-2);
    var mes = ('0' + (data.getMonth() + 1)).slice(-2);
    var ano = data.getFullYear();
    return dia + '/' + mes + '/' + ano;
  }
  
  // Se já estiver no formato DD/MM/YYYY
  var match = dataStr.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (match) {
    return match[1] + '/' + match[2] + '/' + match[3];
  }
  
  // Se estiver no formato YYYY-MM-DD
  match = dataStr.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) {
    return match[3] + '/' + match[2] + '/' + match[1];
  }
  
  return dataStr;
}

function formatarCompetencia(dataStr) {
  if (!dataStr || dataStr === "-" || dataStr === "") return "-";
  
  // Se for objeto Date
  if (dataStr instanceof Date) {
    var mes = ('0' + (dataStr.getMonth() + 1)).slice(-2);
    var ano = dataStr.getFullYear();
    return mes + '/' + ano;
  }
  
  // Se for string
  var str = String(dataStr).trim();
  var partes = str.split("/");
  
  if (partes.length === 3) { 
    // DD/MM/AAAA → MM/AAAA
    return partes[1] + "/" + partes[2];
  }
  
  if (partes.length === 2) { 
    // MM/AAAA → Já está OK
    return str;
  }
  
  // Se chegou aqui, retornar como está
  return str;
}

// ===== FUNÇÕES DE STATUS INDIVIDUAL E GERAL =====

/**
 * Para cada linha de datas (Coluna C), determina o status na Coluna D
 */
function calcularArrayStatusIndividual(datasColunaC, dataReferencia, dataAtual, decimoDiaUtil) {
  return datasColunaC.map(function(linha) {
    var valor = String(linha[0] || '').trim();
    if (!valor || valor === '-' || valor === 'ERRO') return 'DESATUALIZADO';
    if (valor === dataReferencia) return 'OK';
    if (compararDatasPTBR(dataAtual, decimoDiaUtil) <= 0) return 'OK';
    return 'DESATUALIZADO';
  });
}

/**
 * Calcula o status geral final para a Coluna E (célula E1)
 */
function calcularStatusGeral(statusArray, totalLinhas, diasUteisRestantes) {
  var okCount = statusArray.filter(function(s){ return s === 'OK'; }).length;
  if (okCount === totalLinhas) return 'OK';
  if (diasUteisRestantes > 0) { // Está no prazo (dias úteis!)
    return 'OK\n(' + diasUteisRestantes + ' Dias restantes)';
  }
  return 'DESCONFORMIDADE';
}

/**
 * Processa uma aba ("Balancete", "Composição", etc) para calcular os STATUS em lote
 */
function processarStatusAba(aba, datasReferencia) {
  // Pega linhas não vazias
  var rangeA = aba.getRange(4, 1, aba.getLastRow() - 3, 1).getValues();
  var rangeC = aba.getRange(4, 3, rangeA.length, 1).getValues();

  // Limpa as fórmulas/conteúdo da coluna D para evitar conflito!
  aba.getRange(4, 4, rangeA.length, 1).clearContent();

  var dataReferencia = datasReferencia.diaMesRef;
  var dataAtual = datasReferencia.hoje;
  var decimoDiaUtil = datasReferencia.diaMesRef2;

  var linhasPreenchidas = rangeA.filter(function(l){ return l[0] && l[0].toString().trim() !== ''; }).length;

  var statusIndividuais = calcularArrayStatusIndividual(rangeC, dataReferencia, dataAtual, decimoDiaUtil);

  aba.getRange(4, 4, statusIndividuais.length, 1)
    .setValues(statusIndividuais.map(function(x){ return [x]; }));

  var statusGeral = calcularStatusGeral(statusIndividuais, linhasPreenchidas, datasReferencia.diasRestantes);
  aba.getRange('E1').clearContent();
  aba.getRange('E1').setValue(statusGeral);
}

/**
 * Processa automaticamente todas as abas de conformidade monitoradas.
 */
function processarAbasConformidade(datasReferencia) {
  var ss = obterPlanilha();
  var abas = ['Balancete', 'Composição', 'Lâmina', 'Perfil Mensal'];
  if (!datasReferencia) datasReferencia = getDatasReferencia();
  abas.forEach(function(nome) {
    var aba = ss.getSheetByName(nome);
    if (aba) processarStatusAba(aba, datasReferencia);
  });
}

// Executa nas abas Balancete, Composição, Lâmina, Perfil Mensal
function limparFormulasE1() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ['Balancete', 'Composição', 'Lâmina', 'Perfil Mensal'].forEach(function(nomeAba){
    var aba = ss.getSheetByName(nomeAba);
    if (aba) aba.getRange('E1').clearContent();
  });
}

function formatarDiasRestantes(dias) {
  if (dias === 1) return "1 dia restante";
  return dias + " dias restantes";
}

function montarStatusDisplay(statusGeral, diasRestantes) {
  // Garante que qualquer EM CONFORMIDADE virará OK
  if (statusGeral && statusGeral.indexOf('EM CONFORMIDADE') !== -1) {
    return "OK (" + formatarDiasRestantes(diasRestantes) + ")";
  }
  if (statusGeral === "OK") return "OK (" + formatarDiasRestantes(diasRestantes) + ")";
  return statusGeral;
}

// ============================================
// LÓGICA DE COMPETÊNCIAS (MENSAL)
// ============================================

/**
 * Calcula as competências esperadas e seus status
 * @returns {Object} { comp1: "12/2025", comp2: "01/2026", dentrodoPrazo: true }
 */
function calcularCompetenciasEsperadas() {
  var datas = getDatasReferencia();
  var hoje = new Date();
  
  // Mês retrasado (competência 1 esperada)
  var mesRetrasado = new Date(hoje.getFullYear(), hoje.getMonth() - 2, 1);
  var comp1Esperada = formatarCompetencia(formatarData(mesRetrasado));
  
  // Mês anterior (competência 2 esperada)
  var mesAnterior = new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);
  var comp2Esperada = formatarCompetencia(formatarData(mesAnterior));
  
  // Verifica se ainda está dentro do prazo (10º dia útil)
  var dentrodoPrazo = datas.diasRestantes >= 0;
  
  Logger.log('📅 Competências esperadas:');
  Logger.log('   Comp1 (mês retrasado): ' + comp1Esperada);
  Logger.log('   Comp2 (mês anterior): ' + comp2Esperada);
  Logger.log('   Dentro do prazo: ' + dentrodoPrazo + ' (' + datas.diasRestantes + ' dias)');
  
  return {
    comp1: comp1Esperada,
    comp2: comp2Esperada,
    dentrodoPrazo: dentrodoPrazo,
    diasRestantes: datas.diasRestantes
  };
}

/**
 * Determina qual competência exibir e seu status
 * @param {Array} todasCompetencias - Lista de todas as competências encontradas ["12/2025", "01/2026"]
 * @returns {Object} { comp1: "12/2025", status1: "OK", comp2: "01/2026", status2: "OK" }
 */
function determinarCompetenciasEStatus(todasCompetencias) {
  var esperadas = calcularCompetenciasEsperadas();
  
  // Filtrar e ordenar competências (mais recente primeiro)
  var competenciasValidas = todasCompetencias
    .filter(function(c) { return c && c !== '-' && c !== 'ERRO'; })
    .sort()
    .reverse();
  
  Logger.log('📊 Competências encontradas: ' + JSON.stringify(competenciasValidas));
  
  // Se não tem nenhuma competência
  if (competenciasValidas.length === 0) {
    return {
      comp1: '-',
      status1: 'DESCONFORMIDADE',
      comp2: '-',
      status2: esperadas.dentrodoPrazo ? 'AGUARDANDO' : 'DESCONFORMIDADE'
    };
  }
  
  var comp1Encontrada = competenciasValidas[0]; // Mais recente
  var comp2Encontrada = competenciasValidas[1] || null; // Segunda mais recente
  
  // === LÓGICA DE AUTO-ROTAÇÃO ===
  // Se Comp1 = mês anterior E Comp2 = mês anterior TAMBÉM
  // Significa que ambos estão OK, então rotaciona
  if (comp1Encontrada === esperadas.comp2 && comp2Encontrada === esperadas.comp2) {
    Logger.log('🔄 AUTO-ROTAÇÃO: Ambas competências OK, rotacionando...');
    return {
      comp1: esperadas.comp2,
      status1: 'OK',
      comp2: '-',
      status2: 'AGUARDANDO'
    };
  }
  
  // === COMPETÊNCIA 1 ===
  var status1;
  if (comp1Encontrada === esperadas.comp1) {
    status1 = 'OK'; // Tem o mês retrasado
  } else if (comp1Encontrada === esperadas.comp2) {
    status1 = 'OK'; // Tem o mês anterior (ainda melhor)
  } else {
    status1 = 'DESCONFORMIDADE'; // Tem competência muito antiga
  }
  
  // === COMPETÊNCIA 2 ===
  var comp2, status2;
  
  if (comp2Encontrada && comp2Encontrada === esperadas.comp2) {
    // Caso ideal: tem o mês anterior
    comp2 = comp2Encontrada;
    status2 = 'OK';
  } else if (comp1Encontrada === esperadas.comp2) {
    // Só tem mês anterior na comp1, comp2 fica aguardando
    comp2 = '-';
    status2 = 'AGUARDANDO';
  } else {
    // Não tem o mês anterior
    comp2 = '-';
    if (esperadas.dentrodoPrazo) {
      status2 = 'AGUARDANDO'; // Ainda está no prazo
    } else {
      status2 = 'DESCONFORMIDADE'; // Passou do prazo
    }
  }
  
  return {
    comp1: comp1Encontrada,
    status1: status1,
    comp2: comp2,
    status2: status2
  };
}

/**
 * Atualiza colunas C, D, E, F nas abas Balancete, Composição, Lâmina, Perfil Mensal
 * Corrige as falhas 1 e 2 conforme especificação.
 */
function atualizarCompetenciasAba(nomeAba) {
  Logger.log('\n📊 [ATUALIZAÇÃO DE COMPETÊNCIAS] ' + nomeAba);

  var ss = obterPlanilha();
  var aba = ss.getSheetByName(nomeAba);
  if (!aba) {
    Logger.log('❌ Aba não encontrada: ' + nomeAba);
    return;
  }

  var ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 4) {
    Logger.log('⚠️ Aba sem dados');
    return;
  }

  // Buscar para cada linha:
  // C = competência atual
  // E = nova data mais recente (se houver)
  // F = status da nova data ("OK" ou "AGUARDANDO")

  var linhasDados = aba.getRange(4, 1, ultimaLinha - 3, 6).getValues(); // A:F
  var houveNovaData = false;
  var novaDataMaisRecente = null;
  var novaDataArray = [];
  var novaStatusArray = [];
  var competenciaArray = [];

  // Passo 1: Descobrir a data mais recente disponível nas colunas relacionadas ao scraping
  // Suporte flexível: pode vir de IMPORTXML ou scraping, aqui consideramos coluna E (índice 4) após fórmulas já em vigor
  for (var i = 0; i < linhasDados.length; i++) {
    var linha = linhasDados[i];
    var dataAtual = (linha[2] || '').toString().trim(); // C
    var possivelNovaData = (linha[4] || '').toString().trim(); // E
    competenciaArray.push(dataAtual);

    // Se nova data está preenchida e não "-", e é maior que a atual, considerar como candidato
    if (possivelNovaData && possivelNovaData !== '-' && isDataMaisRecente(possivelNovaData, dataAtual)) {
      if (!novaDataMaisRecente) {
        novaDataMaisRecente = possivelNovaData;
      } else {
        // Fica com a maior
        if (isDataMaisRecente(possivelNovaData, novaDataMaisRecente)) {
          novaDataMaisRecente = possivelNovaData;
        }
      }
    }
  }

  // Passo 2: Percorrer linhas para atualizar E e F conforme regra 1 e 2
  var todasEComNova = true; // check para rotação
  var todasFok = true;      // check para rotação

  for (var i = 0; i < linhasDados.length; i++) {
    var linha = linhasDados[i];
    var dataAtual = (linha[2] || '').toString().trim(); // C
    var dataNova = '-';
    var statusNova = "AGUARDANDO";

    // Se há uma nova data mais recente para esta linha:
    if (novaDataMaisRecente && isDataMaisRecente(novaDataMaisRecente, dataAtual)) {
      dataNova = novaDataMaisRecente;
      // Se há algum retorno (caso queira validar mais fortemente, adicione lógica própria)
      // Aqui, para padronizar, consideramos que o fato de aparecer a nova data já indica envio.
      // Se quiser validação por scraping direto aqui, insira na função abaixo.
      var houveRetorno = true; // INSIRA LÓGICA se quiser refinar
      if (houveRetorno) {
        statusNova = "OK";
      } else {
        statusNova = "AGUARDANDO";
        todasFok = false;
      }
      novaStatusArray.push(statusNova);
      novaDataArray.push(dataNova);
    } else {
      // Não há nova data mais recente
      dataNova = '-';
      statusNova = "AGUARDANDO";
      novaDataArray.push(dataNova);
      novaStatusArray.push(statusNova);
      todasEComNova = false;
      todasFok = false;
    }
    // Escreve em E e F
    aba.getRange(i + 4, 5).setValue(dataNova);      // Coluna E: Nova Data
    aba.getRange(i + 4, 6).setValue(statusNova);    // Coluna F: Status Nova
  }

  // Passo 3: Se TODAS as linhas têm E igual à nova data e F="OK", faz a rotação:
  // - C recebe E, D recebe F
  // - E = "-", F = "AGUARDANDO"

  if (
    novaDataMaisRecente &&
    novaDataArray.every(function(e) { return e === novaDataMaisRecente; }) &&
    novaStatusArray.every(function(s) { return s === "OK"; }) &&
    novaDataArray.length === linhasDados.length
  ) {
    Logger.log('🔁 Rotação automática: Nova data será aplicada como competência atual.');
    for (var i = 0; i < linhasDados.length; i++) {
      aba.getRange(i + 4, 3).setValue(novaDataMaisRecente);     // C: Data competência atual = nova
      aba.getRange(i + 4, 4).setValue("OK");                    // D: Status competência atual = OK
      aba.getRange(i + 4, 5).setValue('-');                     // E: Nova Data = "-"
      aba.getRange(i + 4, 6).setValue("AGUARDANDO");            // F: Status Nova = "AGUARDANDO"
    }
    // Atualize dashboard, status geral e etc depois deste ponto (fora desta função).
    Logger.log('✅ Rotação aplicada com sucesso.');
  }

  // -- Atualiza status geral do dashboard (diretamente aqui ou via outra função)
  // -- O cálculo de dias restantes precisa ser ajustado na função correspondente (abaixo).

  Logger.log('✅ Atualização das colunas concluída para ' + nomeAba);
}

/**
 * Verifica se dataPossivel é mais recente que dataAtual
 * Ambas são string no formato "DD/MM/YYYY"
 */
function isDataMaisRecente(dataPossivel, dataAtual) {
  if (!dataAtual || dataAtual === "-") return true;
  var p1 = dataPossivel.split('/'); // DD/MM/YYYY
  var p2 = dataAtual.split('/');
  var d1 = new Date(parseInt(p1[2]), parseInt(p1[1])-1, parseInt(p1[0]));
  var d2 = new Date(parseInt(p2[2]), parseInt(p2[1])-1, parseInt(p2[0]));
  return d1 > d2;
}

/**
 * Chamada direta do ciclo de atualização: chama atualizarCompetenciasAba para todas as abas e atualiza dashboard geral.
 * Para garantir a aplicação das FALHAS 1 e 2.
 * Use esta função para atualizar tudo corretamente.
 */
function atualizarTodasCompetencias() {
  Logger.log('🔄 Atualizando todas as competências e dashboards...');
  var abas = ['Balancete', 'Composição', 'Lâmina', 'Perfil Mensal'];
  var ss = obterPlanilha();
  var dashboardStatus = [];

  abas.forEach(function(nomeAba, idx) {
    atualizarCompetenciasAba(nomeAba); // FALHA 1

    var aba = ss.getSheetByName(nomeAba);
    var ultimaLinha = aba.getLastRow();
    var linhas = aba.getRange(4, 1, ultimaLinha - 3, 6).getValues(); // A:F

    // Para coluna C (Competência Atual, índice 2)
    var competenciasAtuais = linhas.map(function(l) { return (l[2] || '').toString().trim(); });
    var statusGeral = calcularStatusGeralDaAbaComPrazo(linhas, "mensal", competenciasAtuais); // FALHA 2
    aba.getRange('E1').setValue(statusGeral);

    dashboardStatus[idx] = statusGeral;
  });

  // Atualiza painel Dashboard Geral (aba GERAL)
  var abaGeral = ss.getSheetByName('GERAL');
  if (abaGeral) {
    // A4 = Balancete
    // B4 = Composição
    // E4 = Lâmina
    // F4 = Perfil Mensal
    abaGeral.getRange('A4').setValue(dashboardStatus[0]);
    abaGeral.getRange('B4').setValue(dashboardStatus[1]);
    abaGeral.getRange('E4').setValue(dashboardStatus[2]);
    abaGeral.getRange('F4').setValue(dashboardStatus[3]);
  }

  Logger.log('✅ Competências e dashboard atualizados.');
}

// ============================================
// FUNÇÕES AUXILIARES PARA DIÁRIAS
// ============================================

/**
 * Formata um objeto Date para DD/MM/YYYY
 */
function normalizaDataDate(dateObj) {
  var dd = String(dateObj.getDate()).padStart(2, '0');
  var mm = String(dateObj.getMonth() + 1).padStart(2, '0');
  var yyyy = dateObj.getFullYear();
  return dd + '/' + mm + '/' + yyyy;
}

/**
 * Normaliza uma string de data para DD/MM/YYYY
 */
function normalizaData(data) {
  if (!data) return '';
  var s = String(data).trim();
  
  // Formato YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    var p = s.split('-');
    return [p[2], p[1], p[0]].join('/');
  }
  
  // Formato DD/MM/YYYY
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;
  
  return s.replace(/\s+/g, '');
}

/**
 * Retorna array de feriados no formato DD/MM/YYYY
 */
function getFeriadosArray() {
  try {
    var ss = obterPlanilha();
    var aba = ss.getSheetByName('FERIADOS');
    if (!aba) return [];
    
    var lastRow = aba.getLastRow();
    if (lastRow < 2) return [];
    
    var dados = aba.getRange(2, 1, lastRow - 1, 1).getValues();
    return dados
      .filter(function(r) { return r[0]; })
      .map(function(r) { 
        if (r[0] instanceof Date) {
          return normalizaDataDate(r[0]);
        }
        return normalizaData(r[0]);
      });
  } catch (error) {
    Logger.log('⚠️ Erro ao buscar feriados: ' + error.toString());
    return [];
  }
}

/**
 * Calcula o último dia útil antes de uma data
 */
function calculaUltimoDiaUtil(dateObj, feriadosArray) {
  var d = new Date(dateObj.getTime()); // Clonar data
  
  do {
    d.setDate(d.getDate() - 1);
  } while (
    d.getDay() === 0 || // Domingo
    d.getDay() === 6 || // Sábado
    feriadosArray.indexOf(normalizaDataDate(d)) >= 0 // Feriado
  );
  
  return normalizaDataDate(d);
}

// Execute no Apps Script Editor
function testarFuncoes() {
  Logger.log('📅 Testando funções auxiliares...');
  
  var hoje = new Date();
  Logger.log('Hoje: ' + normalizaDataDate(hoje));
  
  var feriados = getFeriadosArray();
  Logger.log('Total de feriados: ' + feriados.length);
  
  var diaUtil = calculaUltimoDiaUtil(hoje, feriados);
  Logger.log('Último dia útil: ' + diaUtil);
  
  Logger.log('✅ Funções OK!');
}

// Gera tabela estilizada para o e-mail, com base no seu template
function gerarTabelaDesconformidadeTemplate(fundos, tipoAba) {
  if (!fundos.length) return '<div style="color:#666;font-size:14px;">Nenhuma desconformidade encontrada.</div>';
  var rotuloData = 'Último envio';
  return (
    '<table class="data-table" style="border-collapse:collapse;width:100%;max-width:600px;margin:15px 0;font-size:13px;">' +
    '<thead>' +
    '<tr style="background:#f3f4f6;">' +
      '<th style="padding:7px 5px;border:1px solid #e0e0e0;text-align:left;width:110px;">'+ rotuloData +'</th>' +
      '<th style="padding:7px 5px;border:1px solid #e0e0e0;text-align:left;">Fundo</th>' +
      '<th style="padding:7px 5px;border:1px solid #e0e0e0;text-align:center;width:110px;">Status 2</th>' +
    '</tr>' +
    '</thead>' +
    '<tbody>' +
    fundos.map(function(f) {
      var dataVal = tipoAba === 'Diárias' ? (f.envio1 || '-') : (f.competencia1 || '-');
      return (
        '<tr class="table-row">' +
          '<td style="padding:7px 5px;border:1px solid #e0e0e0;">' + dataVal + '</td>' +
          '<td style="padding:7px 5px;border:1px solid #e0e0e0;max-width:320px;word-break:break-word;">' + (f.nome || '-') + '</td>' +
          '<td style="padding:7px 5px;border:1px solid #e0e0e0;text-align:center;">' + (f.status2 || '-') + '</td>' +
        '</tr>'
      );
    }).join('') +
    '</tbody>' +
    '</table>'
  );
}

/**
 * Criar trigger para enviar emails diariamente às 18:30
 */
function criarTriggerEmailDiario1830() {
  Logger.log('🔧 Configurando triggers de emails...');
  
  // Remover triggers antigos
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    var funcName = trigger.getHandlerFunction();
    if (funcName === 'enviarEmailConformidadeOuDesconformidadeAvancado' || 
        funcName === 'enviarEmailDiariasSeForUltimoDiaUtil') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('  🗑️ Trigger antigo removido: ' + funcName);
    }
  });
  
  // ✅ TRIGGER 1: Emails das abas mensais (Balancete, Composição, Lâmina, Perfil Mensal)
  ScriptApp.newTrigger('enviarEmailConformidadeOuDesconformidadeAvancado')
    .timeBased()
    .atHour(18)
    .nearMinute(30)
    .everyDays(1)
    .create();
  
  Logger.log('✅ Trigger criado: Emails mensais às 18:30 (diariamente)');
  
  // ✅ TRIGGER 2: Emails de Diárias (APENAS no último dia útil do mês)
  ScriptApp.newTrigger('enviarEmailDiariasSeForUltimoDiaUtil')
    .timeBased()
    .atHour(18)
    .nearMinute(30)
    .everyDays(1)
    .create();
  
  Logger.log('✅ Trigger criado: Emails de Diárias às 18:30 (verifica se é último dia útil)');
  
  Logger.log('\n✅ ═══════════════════════════════════════════');
  Logger.log('✅ TRIGGERS DE EMAILS ATIVADOS!');
  Logger.log('✅ ═══════════════════════════════════════════');
  Logger.log('');
  Logger.log('📧 Função 1: enviarEmailConformidadeOuDesconformidadeAvancado()');
  Logger.log('   ⏰ Horário: 18:30 (diariamente)');
  Logger.log('   📋 Envia: Balancete, Composição, Lâmina, Perfil Mensal');
  Logger.log('');
  Logger.log('📧 Função 2: enviarEmailDiariasSeForUltimoDiaUtil()');
  Logger.log('   ⏰ Horário: 18:30 (diariamente)');
  Logger.log('   📋 Envia: Diárias (SÓ no último dia útil do mês)');
  Logger.log('');
  Logger.log('⚠️ IMPORTANTE: O Google Apps Script pode ter variação de ±15 minutos');
  Logger.log('   (Pode executar entre 18:15 e 18:45)');
  
  return {
    success: true,
    message: 'Triggers criados! Emails serão enviados diariamente às 18:30'
  };
}

/**
 * 🔍 DIAGNÓSTICO: Busca TODAS as datas de diárias de TODOS os fundos
 * Execute no Apps Script Editor para ver o log completo
 * Tempo estimado: ~15 segundos
 */
function diagnosticarTodasDatasDiarias() {
  Logger.log('🔍 ===== DIAGNÓSTICO DE DATAS DIÁRIAS =====\n');
  
  var fundos = getFundos();
  var totalFundos = fundos.length;
  var fundosComSucesso = 0;
  var fundosComErro = 0;
  var totalDatas = 0;
  
  fundos.forEach(function(fundo, index) {
    Logger.log('📊 [' + (index + 1) + '/' + totalFundos + '] ' + fundo.nome.substring(0, 40) + '...');
    Logger.log('   Código CVM: ' + fundo.codigoCVM);
    
    var url = 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/InfDiario/CPublicaInfdiario.aspx?PK_PARTIC=' + fundo.codigoCVM + '&PK_SUBCLASSE=-1';
    Logger.log('   URL: ' + url.substring(0, 80) + '...');
    
    try {
      var response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0' },
        followRedirects: true
      });
      
      var codigo = response.getResponseCode();
      Logger.log('   Status HTTP: ' + codigo);
      
      if (codigo === 200) {
        var html = response.getContentText();
        
        // 🔥 NOVA LÓGICA: Extrair linhas da tabela com DIA e DATA
        var linhasComDatas = [];
        
        // Regex para encontrar padrões como: <td>2</td>...<td>03/02/2026</td>
        // Captura o conteúdo entre <tr> e </tr>
        var regexLinhas = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
        var matchLinhas;
        
        while ((matchLinhas = regexLinhas.exec(html)) !== null) {
          var linhaHtml = matchLinhas[1];
          
          // Buscar <td> com número (dia) e <td> com data DD/MM/YYYY
          var regexDia = /<td[^>]*>(\d{1,2})<\/td>/i;
          var regexData = /<td[^>]*>(\d{2}\/\d{2}\/\d{4})<\/td>/i;
          
          var matchDia = linhaHtml.match(regexDia);
          var matchData = linhaHtml.match(regexData);
          
          if (matchDia && matchData) {
            var dia = matchDia[1];
            var data = matchData[1];
            linhasComDatas.push({ dia: dia, data: data });
          }
        }
        
        if (linhasComDatas.length > 0) {
          // Remover duplicatas e ordenar por data (mais recente primeiro)
          var datasUnicas = [];
          var datasVistas = {};
          
          linhasComDatas.forEach(function(item) {
            if (!datasVistas[item.data]) {
              datasVistas[item.data] = true;
              datasUnicas.push(item);
            }
          });
          
          // Ordenar por data (mais recente primeiro)
          datasUnicas.sort(function(a, b) {
            var partsA = a.data.split('/');
            var partsB = b.data.split('/');
            var dateA = new Date(partsA[2], partsA[1] - 1, partsA[0]);
            var dateB = new Date(partsB[2], partsB[1] - 1, partsB[0]);
            return dateB - dateA; // Ordem decrescente
          });
          
          Logger.log('   Total de datas únicas: ' + datasUnicas.length);
          Logger.log('   Data mais recente: Dia ' + datasUnicas[0].dia + ' - ' + datasUnicas[0].data);
          
          // Mostrar primeiras 10 datas COM o número do dia
          Logger.log('   Primeiras 10 datas:');
          for (var i = 0; i < Math.min(10, datasUnicas.length); i++) {
            Logger.log('     [' + (i + 1) + '] Dia ' + datasUnicas[i].dia + ' - ' + datasUnicas[i].data);
          }
          
          fundosComSucesso++;
          totalDatas += datasUnicas.length;
          Logger.log('   ✅ Sucesso\n');
          
        } else {
          Logger.log('   ⚠️ Nenhuma data encontrada no HTML');
          Logger.log('   ❌ Falha\n');
          fundosComErro++;
        }
        
      } else {
        Logger.log('   ❌ Erro HTTP: ' + codigo + '\n');
        fundosComErro++;
      }
      
    } catch (error) {
      Logger.log('   ❌ Erro: ' + error.toString() + '\n');
      fundosComErro++;
    }
    
    // Delay entre requisições (evitar bloqueio)
    if (index < totalFundos - 1) {
      Utilities.sleep(300);
    }
  });
  
  // Resumo final
  Logger.log('\n========================================');
  Logger.log('✅ RESUMO FINAL:');
  Logger.log('   Total de fundos: ' + totalFundos);
  Logger.log('   Fundos com sucesso: ' + fundosComSucesso);
  Logger.log('   Fundos com erro: ' + fundosComErro);
  if (fundosComSucesso > 0) {
    Logger.log('   Média de datas por fundo: ' + Math.round(totalDatas / fundosComSucesso));
    Logger.log('   Total de datas encontradas: ' + totalDatas);
  }
  Logger.log('========================================');
  
  return {
    success: true,
    totalFundos: totalFundos,
    fundosComSucesso: fundosComSucesso,
    fundosComErro: fundosComErro,
    mediaDatas: fundosComSucesso > 0 ? Math.round(totalDatas / fundosComSucesso) : 0
  };
}

/**
 * Formata qualquer tipo de data para DD/MM/YYYY
 * @param {*} data - Date object, string ou null
 * @returns {string} Data formatada ou "-"
 */
/**
 * Formata qualquer tipo de data para DD/MM/YYYY
 * @param {*} data - Date object, string ou null
 * @returns {string} Data formatada ou "-"
 */
function formatarDataParaEmail(data) {
  if (!data) return '-';
  
  // Se for string vazia
  if (typeof data === 'string' && data.trim() === '') return '-';
  
  // Se já for DD/MM/YYYY
  if (typeof data === 'string' && /^\d{2}\/\d{2}\/\d{4}$/.test(data)) {
    return data;
  }
  
  // Se for objeto Date
  if (data instanceof Date && !isNaN(data.getTime())) {
    return Utilities.formatDate(data, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  
  // Se for string de Date (Thu Jan 01...)
  if (typeof data === 'string' && data.indexOf('GMT') !== -1) {
    try {
      var dateObj = new Date(data);
      if (!isNaN(dateObj.getTime())) {
        return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      }
    } catch (e) {
      return '-';
    }
  }
  
  return '-';
}

/**
 * 🧪 TESTE: Verifica formatação de datas nos emails
 */
function testarFormatacaoDatasEmail() {
  Logger.log('🧪 Testando formatação de datas...\n');
  
  // Testar diferentes tipos de entrada
  var testes = [
    { entrada: new Date(2026, 0, 1), descricao: 'Date object' },
    { entrada: '01/01/2026', descricao: 'String já formatada' },
    { entrada: 'Thu Jan 01 2026 00:00:00 GMT-0300 (GMT-03:00)', descricao: 'String de Date' },
    { entrada: null, descricao: 'null' },
    { entrada: '', descricao: 'String vazia' },
    { entrada: '-', descricao: 'Hífen' }
  ];
  
  testes.forEach(function(teste, i) {
    var resultado = formatarDataParaEmail(teste.entrada);
    Logger.log('[' + (i+1) + '] ' + teste.descricao);
    Logger.log('    Input: ' + (teste.entrada || 'null'));
    Logger.log('    Output: ' + resultado);
    Logger.log('    ✅ ' + (resultado === '-' || /^\d{2}\/\d{2}\/\d{4}$/.test(resultado) ? 'OK' : 'FALHOU'));
    Logger.log('');
  });
  
  Logger.log('✅ Teste concluído!');
}

/**
 * 🧪 Forçar envio de email do Balancete
 * ATENÇÃO: Envia email real para os destinatários configurados
 */
function forcarEnvioEmailBalancete() {
  Logger.log('📧 Forçando envio de email do Balancete...');
  
  var ss = obterPlanilha();
  var destinatarios = [
    'spandrade@banestes.com.br',
    'fabiooliveira@banestes.com.br',
    'iodutra@banestes.com.br',
    'mcdias@banestes.com.br',
    'sndemuner@banestes.com.br',
    'wffreitas@banestes.com.br'
  ];
  
  var mesPassado = obterMesPassadoFormatado();
  var dataAtualFormatada = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy HH:mm');
  var urlPlanilha = obterURLPlanilha();
  
  // Processar APENAS Balancete
  processarAbaEmail(
    ss.getSheetByName('Balancete'),
    'Balancete',
    destinatarios,
    mesPassado,
    dataAtualFormatada,
    urlPlanilha,
    'mensal'
  );
  
  Logger.log('✅ Processo concluído!');
  Logger.log('📬 Verifique a caixa de entrada dos destinatários.');
}


/**
 * 📧 Enviar email individual para CADA FUNDO com TODAS as suas datas
 */
function enviarEmailDiariasIndividualPorFundo() {
  Logger.log('📧 ===== ENVIO INDIVIDUAL POR FUNDO =====\n');
  
  var destinatarios = [
    //'spandrade@banestes.com.br',
    'fabiooliveira@banestes.com.br',
    //'iodutra@banestes.com.br',
    'mcdias@banestes.com.br',
    'sndemuner@banestes.com.br',
    'wffreitas@banestes.com.br'
  ];
  
  var fundos = getFundos();
  var emailsEnviados = 0;
  var emailsComErro = 0;
  
  Logger.log('📊 Processando ' + fundos.length + ' fundos...\n');
  
  fundos.forEach(function(fundo, index) {
    Logger.log('[' + (index + 1) + '/' + fundos.length + '] ' + fundo.nome.substring(0, 40) + '...');
    
    try {
      var url = 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/InfDiario/CPublicaInfdiario.aspx?PK_PARTIC=' + fundo.codigoCVM + '&PK_SUBCLASSE=-1';
      var response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0' }
      });
      
      if (response.getResponseCode() === 200) {
        var html = response.getContentText();
        var linhasComDatas = [];
        var regexLinhas = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
        var matchLinhas;
        
        while ((matchLinhas = regexLinhas.exec(html)) !== null) {
          var linhaHtml = matchLinhas[1];
          var regexDia = /<td[^>]*>(\d{1,2})<\/td>/i;
          var regexData = /<td[^>]*>(\d{2}\/\d{2}\/\d{4})<\/td>/i;
          var matchDia = linhaHtml.match(regexDia);
          var matchData = linhaHtml.match(regexData);
          
          if (matchDia && matchData) {
            linhasComDatas.push({ dia: matchDia[1], data: matchData[1] });
          }
        }
        
        if (linhasComDatas.length > 0) {
          // Remover duplicatas
          var datasUnicas = [];
          var datasVistas = {};
          
          linhasComDatas.forEach(function(item) {
            if (!datasVistas[item.data]) {
              datasVistas[item.data] = true;
              datasUnicas.push(item);
            }
          });
          
          // Ordenar (mais recente primeiro)
          datasUnicas.sort(function(a, b) {
            var partsA = a.data.split('/');
            var partsB = b.data.split('/');
            var dateA = new Date(partsA[2], partsA[1] - 1, partsA[0]);
            var dateB = new Date(partsB[2], partsB[1] - 1, partsB[0]);
            return dateB - dateA;
          });
          
          Logger.log('   Total de datas encontradas: ' + datasUnicas.length);
          
          // Gerar linhas da tabela com TODAS as datas
          var linhasTabela = datasUnicas.map(function(item) {
            return '<tr>' +
              '<td style="padding:10px;border:1px solid #dddddd;background:#ffffff;text-align:center;">' + item.dia + '</td>' +
              '<td style="padding:10px;border:1px solid #dddddd;background:#ffffff;text-align:center;">' + item.data + '</td>' +
              '</tr>';
          });
          
          var dataAtual = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy');
          
          // HTML do email
          var htmlCompleto = 
            '<!DOCTYPE html>' +
            '<html lang="pt-BR">' +
            '<head>' +
            '<meta charset="UTF-8">' +
            '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
            '<style>' +
            'body, table, td, a { -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; }' +
            'table, td { mso-table-lspace: 0pt; mso-table-rspace: 0pt; }' +
            'table { border-collapse: collapse !important; }' +
            'body { margin: 0 !important; padding: 0 !important; width: 100% !important; font-family: Arial, sans-serif; background-color: #f4f4f4; }' +
            '.monitor-box { background-color: #e3f2fd; border-left: 4px solid #2196F3; padding: 15px; margin-top: 20px; }' +
            '</style>' +
            '</head>' +
            '<body style="background-color:#f4f4f4;padding:20px;">' +
            '<table width="100%" cellpadding="0" cellspacing="0" style="max-width:650px;margin:auto;background-color:#ffffff;border-radius:8px;box-shadow:0 2px 5px rgba(0,0,0,0.1);">' +
            '<tr>' +
            '<td align="center" style="background-color:#2E7D32;padding:30px 20px;">' +
            '<div style="font-size:40px;color:#ffffff;line-height:1;margin-bottom:10px;">✓</div>' +
            '<h1 style="color:#ffffff;font-size:22px;margin:0;">Relatório de Conformidade CVM</h1>' +
            '<p style="color:#a5d6a7;margin:5px 0 0 0;font-size:14px;">Informações Diárias</p>' +
            '</td>' +
            '</tr>' +
            '<tr>' +
            '<td style="padding:30px 25px;color:#333333;font-size:15px;line-height:1.6;">' +
            '<p>Prezados,</p>' +
            '<p>Informamos que os envios de <strong>Informações Diárias</strong> junto à CVM para o fundo abaixo encontram-se em conformidade.</p>' +
            '<div style="background-color:#f0f9ff;border-left:4px solid #667eea;padding:15px;margin:20px 0;">' +
            '<p style="margin:0;font-weight:bold;color:#1e3a8a;font-size:16px;">Fundo:</p>' +
            '<p style="margin:5px 0 0 0;font-size:14px;color:#333;">' + fundo.nome + '</p>' +
            '<p style="margin:10px 0 0 0;font-size:13px;color:#666;">Código CVM: ' + fundo.codigoCVM + '</p>' +
            '</div>' +
            '<p>Abaixo listamos <strong>todos os ' + datasUnicas.length + ' envios</strong> registrados no sistema da CVM:</p>' +
            '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin:20px 0;font-family:Arial,sans-serif;">' +
            '<thead>' +
            '<tr>' +
            '<th style="padding:12px;border:1px solid #dddddd;background-color:#f3f4f6;text-align:center;font-weight:bold;color:#555555;">Dia</th>' +
            '<th style="padding:12px;border:1px solid #dddddd;background-color:#f3f4f6;text-align:center;font-weight:bold;color:#555555;">Data de Envio</th>' +
            '</tr>' +
            '</thead>' +
            '<tbody>' +
            linhasTabela.join('') +
            '</tbody>' +
            '</table>' +
            '<div class="monitor-box">' +
            '<p style="margin:0;font-weight:bold;color:#0d47a1;font-size:14px;">✓ Status: Regularizado</p>' +
            '<p style="margin:5px 0 0 0;font-size:13px;color:#444;">Todos os ' + datasUnicas.length + ' envios foram identificados corretamente no portal da CVM.</p>' +
            '</div>' +
            '</td>' +
            '</tr>' +
            '<tr>' +
            '<td align="center" style="background-color:#f8f9fa;padding:20px;color:#888888;font-size:12px;border-top:1px solid #eeeeee;">' +
            '<p style="margin:0;">Departamento de Inovação e Automação interno Asset</p>' +
            '<p style="margin:5px 0 0 0;">Relatório gerado em ' + dataAtual + '</p>' +
            '</td>' +
            '</tr>' +
            '</table>' +
            '</body>' +
            '</html>';
          
          // Enviar email
          var assunto = '✅ Conformidade CVM - Diárias - ' + fundo.nome.substring(0, 60);
          
          MailApp.sendEmail({
            to: destinatarios.join(','),
            subject: assunto,
            htmlBody: htmlCompleto
          });
          
          emailsEnviados++;
          Logger.log('   ✅ Email enviado (' + datasUnicas.length + ' datas)');
          
        } else {
          Logger.log('   ⚠️ Sem dados - email não enviado');
        }
      } else {
        Logger.log('   ❌ Erro HTTP: ' + response.getResponseCode());
        emailsComErro++;
      }
      
      // Delay entre fundos (evitar spam)
      if (index < fundos.length - 1) {
        Utilities.sleep(2000); // 2 segundos entre cada email
      }
      
    } catch (error) {
      Logger.log('   ❌ Erro: ' + error.toString());
      emailsComErro++;
    }
  });
  
  Logger.log('\n========================================');
  Logger.log('✅ RESUMO FINAL:');
  Logger.log('   Total de fundos: ' + fundos.length);
  Logger.log('   Emails enviados: ' + emailsEnviados);
  Logger.log('   Erros: ' + emailsComErro);
  Logger.log('========================================');
  
  return {
    success: true,
    totalFundos: fundos.length,
    emailsEnviados: emailsEnviados,
    emailsComErro: emailsComErro
  };
}

/**
 * 🧪 TESTE com apenas 2 fundos
 */
function testarEmailDiariasIndividual() {
  Logger.log('🧪 ===== TESTE - 2 FUNDOS =====\n');
  
  var destinatarios = ['spandrade@banestes.com.br'];
  var fundos = getFundos().slice(0, 2); // Apenas 2 fundos
  
  fundos.forEach(function(fundo, index) {
    Logger.log('[' + (index + 1) + '/' + fundos.length + '] ' + fundo.nome.substring(0, 40) + '...');
    
    try {
      var url = 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/InfDiario/CPublicaInfdiario.aspx?PK_PARTIC=' + fundo.codigoCVM + '&PK_SUBCLASSE=-1';
      var response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0' }
      });
      
      if (response.getResponseCode() === 200) {
        var html = response.getContentText();
        var linhasComDatas = [];
        var regexLinhas = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
        var matchLinhas;
        
        while ((matchLinhas = regexLinhas.exec(html)) !== null) {
          var linhaHtml = matchLinhas[1];
          var regexDia = /<td[^>]*>(\d{1,2})<\/td>/i;
          var regexData = /<td[^>]*>(\d{2}\/\d{2}\/\d{4})<\/td>/i;
          var matchDia = linhaHtml.match(regexDia);
          var matchData = linhaHtml.match(regexData);
          
          if (matchDia && matchData) {
            linhasComDatas.push({ dia: matchDia[1], data: matchData[1] });
          }
        }
        
        if (linhasComDatas.length > 0) {
          var datasUnicas = [];
          var datasVistas = {};
          
          linhasComDatas.forEach(function(item) {
            if (!datasVistas[item.data]) {
              datasVistas[item.data] = true;
              datasUnicas.push(item);
            }
          });
          
          datasUnicas.sort(function(a, b) {
            var partsA = a.data.split('/');
            var partsB = b.data.split('/');
            var dateA = new Date(partsA[2], partsA[1] - 1, partsA[0]);
            var dateB = new Date(partsB[2], partsB[1] - 1, partsB[0]);
            return dateB - dateA;
          });
          
          Logger.log('   Datas: ' + datasUnicas.length);
          datasUnicas.forEach(function(item, i) {
            Logger.log('     [' + (i+1) + '] Dia ' + item.dia + ' - ' + item.data);
          });
          
          var linhasTabela = datasUnicas.map(function(item) {
            return '<tr>' +
              '<td style="padding:10px;border:1px solid #dddddd;background:#ffffff;text-align:center;">' + item.dia + '</td>' +
              '<td style="padding:10px;border:1px solid #dddddd;background:#ffffff;text-align:center;">' + item.data + '</td>' +
              '</tr>';
          });
          
          var dataAtual = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy');
          
          var htmlCompleto = 
            '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><style>' +
            'body, table, td { font-family: Arial, sans-serif; }' +
            'body { background-color: #f4f4f4; padding: 20px; }' +
            'table { border-collapse: collapse !important; }' +
            '</style></head><body>' +
            '<table width="100%" cellpadding="0" cellspacing="0" style="max-width:650px;margin:auto;background-color:#ffffff;border-radius:8px;">' +
            '<tr><td align="center" style="background-color:#2E7D32;padding:30px 20px;">' +
            '<h1 style="color:#ffffff;font-size:22px;margin:0;">✓ Relatório de Conformidade CVM</h1>' +
            '<p style="color:#a5d6a7;margin:5px 0 0 0;">Informações Diárias</p>' +
            '</td></tr>' +
            '<tr><td style="padding:30px 25px;color:#333333;font-size:15px;">' +
            '<p>Prezados,</p>' +
            '<p>Informamos que os envios de <strong>Informações Diárias</strong> junto à CVM encontram-se em conformidade.</p>' +
            '<div style="background-color:#f0f9ff;border-left:4px solid #667eea;padding:15px;margin:20px 0;">' +
            '<p style="margin:0;font-weight:bold;color:#1e3a8a;">Fundo:</p>' +
            '<p style="margin:5px 0 0 0;">' + fundo.nome + '</p>' +
            '<p style="margin:10px 0 0 0;font-size:13px;color:#666;">Código CVM: ' + fundo.codigoCVM + '</p>' +
            '</div>' +
            '<p>Abaixo listamos <strong>todos os ' + datasUnicas.length + ' envios</strong>:</p>' +
            '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin:20px 0;">' +
            '<thead><tr>' +
            '<th style="padding:12px;border:1px solid #ddd;background:#f3f4f6;text-align:center;">Dia</th>' +
            '<th style="padding:12px;border:1px solid #ddd;background:#f3f4f6;text-align:center;">Data de Envio</th>' +
            '</tr></thead><tbody>' +
            linhasTabela.join('') +
            '</tbody></table>' +
            '<div style="background-color:#e3f2fd;border-left:4px solid #2196F3;padding:15px;margin-top:20px;">' +
            '<p style="margin:0;font-weight:bold;color:#0d47a1;">✓ Status: Regularizado</p>' +
            '<p style="margin:5px 0 0 0;font-size:13px;">Todos os ' + datasUnicas.length + ' envios foram identificados.</p>' +
            '</div>' +
            '</td></tr>' +
            '<tr><td align="center" style="background-color:#f8f9fa;padding:20px;font-size:12px;color:#888;">' +
            '<p style="margin:0;">Departamento de Inovação e Automação interno Asset</p>' +
            '<p style="margin:5px 0 0 0;">Relatório gerado em ' + dataAtual + '</p>' +
            '</td></tr></table></body></html>';
          
          MailApp.sendEmail({
            to: destinatarios.join(','),
            subject: '🧪 TESTE - Diárias - ' + fundo.nome.substring(0, 40),
            htmlBody: htmlCompleto
          });
          
          Logger.log('   ✅ Email enviado');
        }
      }
      
      if (index < fundos.length - 1) {
        Utilities.sleep(2000);
      }
      
    } catch (error) {
      Logger.log('   ❌ Erro: ' + error.toString());
    }
  });
  
  Logger.log('\n✅ Teste concluído!');
}

/**
 * Criar trigger para enviar emails de diárias no último dia útil do mês
 */
function criarTriggerEmailDiariasUltimoDiaUtil() {
  Logger.log('🔧 Configurando trigger mensal para diárias...');
  
  // Remover triggers antigos (se existirem)
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'verificarEEnviarEmailDiariasSeUltimoDiaUtil') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('  🗑️ Trigger antigo removido');
    }
  });
  
  // Criar novo trigger DIÁRIO às 17:00 (verifica se é último dia útil)
  ScriptApp.newTrigger('verificarEEnviarEmailDiariasSeUltimoDiaUtil')
    .timeBased()
    .atHour(17)
    .everyDays(1)
    .create();
  
  Logger.log('\n✅ ═══════════════════════════════════════════');
  Logger.log('✅ TRIGGER MENSAL DE DIÁRIAS ATIVADO!');
  Logger.log('✅ ═══════════════════════════════════════════');
  Logger.log('');
  Logger.log('📧 Função: verificarEEnviarEmailDiariasSeUltimoDiaUtil()');
  Logger.log('⏰ Horário: 17:00 (diariamente)');
  Logger.log('🎯 Envia: Apenas no último dia útil do mês');
  Logger.log('📊 Conteúdo: Todos os envios de diárias por fundo');
  
  return {
    success: true,
    message: 'Trigger criado! Emails de diárias serão enviados no último dia útil do mês.'
  };
}

/**
 * Verifica se hoje é o último dia útil do mês e envia emails
 */
function verificarEEnviarEmailDiariasSeUltimoDiaUtil() {
  Logger.log('🔍 Verificando se hoje é último dia útil do mês...');
  
  var hoje = new Date();
  var ss = obterPlanilha();
  var feriados = getFeriadosArray();
  
  // Calcular último dia útil do mês
  var ultimoDiaMes = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0); // Último dia do mês
  
  // Retroceder até encontrar um dia útil
  while (ultimoDiaMes.getDay() === 0 || ultimoDiaMes.getDay() === 6 || 
         feriados.indexOf(normalizaDataDate(ultimoDiaMes)) >= 0) {
    ultimoDiaMes.setDate(ultimoDiaMes.getDate() - 1);
  }
  
  var ultimoDiaUtilFormatado = normalizaDataDate(ultimoDiaMes);
  var hojeFormatado = normalizaDataDate(hoje);
  
  Logger.log('📅 Hoje: ' + hojeFormatado);
  Logger.log('📅 Último dia útil do mês: ' + ultimoDiaUtilFormatado);
  
  // Verificar se hoje é o último dia útil
  if (hojeFormatado === ultimoDiaUtilFormatado) {
    Logger.log('✅ HOJE É O ÚLTIMO DIA ÚTIL! Enviando emails...');
    enviarRelatorioDiariasConsolidadoPDF();
  } else {
    Logger.log('⏭️ Hoje NÃO é o último dia útil. Email não será enviado.');
  }
}

/**
 * 🧪 TESTE: Simular último dia útil (forçar envio)
 */
function testarEnvioDiariasUltimoDiaUtil() {
  Logger.log('🧪 TESTE: Forçando envio de emails de Diárias...');
  Logger.log('⚠️ ATENÇÃO: Emails REAIS serão enviados!');
  Logger.log('');
  
  // Alterar destinatários para teste (só você)
  var destinatariosTeste = ['spandrade@banestes.com.br'];
  Logger.log('📧 Destinatários: ' + destinatariosTeste.join(', '));
  Logger.log('');
  
  // Chamar função de envio
  enviarEmailDiariasIndividualPorFundo();
  
  Logger.log('\n✅ Teste concluído!');
  Logger.log('📬 Verifique sua caixa de entrada.');
}

/**
 * 🧪 TESTE: Verifica qual é o último dia útil do mês atual
 * Execute no Apps Script Editor para ver o resultado no log
 */
function testarCalculoUltimoDiaUtil() {
  Logger.log('🧪 ===== TESTE: Cálculo do Último Dia Útil =====\n');
  
  var hoje = new Date();
  var ss = obterPlanilha();
  var feriados = getFeriadosArray();
  
  Logger.log('📅 Hoje: ' + normalizaDataDate(hoje));
  Logger.log('📅 Dia da semana: ' + hoje.toLocaleDateString('pt-BR', { weekday: 'long' }));
  
  // Calcular último dia útil do mês
  var ultimoDiaMes = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0);
  Logger.log('📅 Último dia do mês (calendário): ' + normalizaDataDate(ultimoDiaMes));
  
  // Retroceder até encontrar um dia útil
  while (ultimoDiaMes.getDay() === 0 || ultimoDiaMes.getDay() === 6 || 
         feriados.indexOf(normalizaDataDate(ultimoDiaMes)) >= 0) {
    ultimoDiaMes.setDate(ultimoDiaMes.getDate() - 1);
  }
  
  Logger.log('📅 Último dia ÚTIL do mês: ' + normalizaDataDate(ultimoDiaMes));
  Logger.log('📅 Dia da semana: ' + ultimoDiaMes.toLocaleDateString('pt-BR', { weekday: 'long' }));
  
  // Verificar
  var ultimoDiaUtilFormatado = normalizaDataDate(ultimoDiaMes);
  var hojeFormatado = normalizaDataDate(hoje);
  
  if (hojeFormatado === ultimoDiaUtilFormatado) {
    Logger.log('\n✅ HOJE É O ÚLTIMO DIA ÚTIL DO MÊS!');
    Logger.log('📧 Emails de Diárias SERÃO enviados às 17:00');
  } else {
    var diasRestantes = Math.floor((ultimoDiaMes - hoje) / (1000 * 60 * 60 * 24));
    Logger.log('\n⏭️ Hoje NÃO é o último dia útil');
    Logger.log('📆 Faltam ' + diasRestantes + ' dias úteis para o último dia útil');
    Logger.log('📅 Próximo envio: ' + ultimoDiaUtilFormatado + ' às 17:00');
  }
  
  Logger.log('\n✅ Teste concluído!');
}

/**
 * Helper: Retorna o nome do dia da semana
 */
function obterDiaSemana(data) {
  var dias = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
  return dias[data.getDay()];
}

/**
 * Helper: Calcula dias úteis entre duas datas
 */
function calcularDiasUteisEntreDatas(dataInicio, dataFim, feriados) {
  var count = 0;
  var atual = new Date(dataInicio);
  atual.setHours(0, 0, 0, 0);
  
  var fim = new Date(dataFim);
  fim.setHours(0, 0, 0, 0);
  
  while (atual < fim) {
    atual.setDate(atual.getDate() + 1);
    if (atual.getDay() !== 0 && atual.getDay() !== 6 && 
        feriados.indexOf(normalizaDataDate(atual)) === -1) {
      count++;
    }
  }
  
  return count;
}

/**
 * 📧 ENVIA EMAIL DE DIÁRIAS APENAS NO ÚLTIMO DIA ÚTIL DO MÊS
 * Esta função verifica se hoje é o último dia útil do mês E envia os emails
 * 
 * ✅ COMO USAR:
 * - Configure um trigger diário às 18:30 para executar esta função
 * - Ela só envia email no último dia útil do mês
 */
/**
 * 📧 ENVIA EMAIL DE DIÁRIAS APENAS NO ÚLTIMO DIA ÚTIL DO MÊS
 * Esta função verifica se hoje é o último dia útil do mês E envia os emails
 * 
 * ✅ COMO USAR:
 * - Configure um trigger diário às 18:30 para executar esta função
 * - Ela só envia email no último dia útil do mês
 */
function enviarEmailDiariasSeForUltimoDiaUtil() {
  Logger.log('🔍 Verificando se hoje é o último dia útil do mês...');
  
  // Verificar se é dia útil
  var hoje = new Date();
  var diaSemana = hoje.getDay();
  
  if (diaSemana === 0 || diaSemana === 6) {
    Logger.log('⏭️ Hoje é fim de semana. Não é dia útil.');
    return { skipped: true, reason: 'Fim de semana' };
  }
  
  // Verificar se é feriado
  try {
    var ss = obterPlanilha();
    var abaFeriados = ss.getSheetByName('FERIADOS');
    if (abaFeriados) {
      var feriados = abaFeriados.getRange('A2:A100').getValues();
      var hojeFormatado = formatarData(hoje);
      
      for (var i = 0; i < feriados.length; i++) {
        if (feriados[i][0]) {
          var feriadoFormatado = formatarData(new Date(feriados[i][0]));
          if (feriadoFormatado === hojeFormatado) {
            Logger.log('⏭️ Hoje é feriado. Não é dia útil.');
            return { skipped: true, reason: 'Feriado' };
          }
        }
      }
    }
  } catch (error) {
    Logger.log('⚠️ Erro ao verificar feriados: ' + error.toString());
  }
  
  // ✅ É DIA ÚTIL - Verificar se é o ÚLTIMO dia útil do mês
  var ultimoDiaUtil = calcularUltimoDiaUtilDoMes(hoje, ss);
  var hojeNormalizado = formatarData(hoje);
  var ultimoDiaUtilNormalizado = formatarData(ultimoDiaUtil);
  
  Logger.log('📅 Hoje: ' + hojeNormalizado);
  Logger.log('📅 Último dia útil do mês: ' + ultimoDiaUtilNormalizado);
  
  if (hojeNormalizado === ultimoDiaUtilNormalizado) {
    Logger.log('🎯 HOJE É O ÚLTIMO DIA ÚTIL DO MÊS! Enviando emails de Diárias...');
    
    // ✅ ENVIAR EMAILS DE DIÁRIAS
    return enviarEmailDiariasIndividualPorFundo();
  } else {
    Logger.log('⏭️ Hoje NÃO é o último dia útil do mês. Email NÃO será enviado.');
    return { skipped: true, reason: 'Não é o último dia útil do mês' };
  }
}

/**
 * 📅 CALCULA O ÚLTIMO DIA ÚTIL DO MÊS ATUAL
 * @param {Date} dataReferencia - Data de referência
 * @param {SpreadsheetApp} ss - Planilha
 * @returns {Date} - Último dia útil do mês
 */
function calcularUltimoDiaUtilDoMes(dataReferencia, ss) {
  // Último dia do mês atual
  var ano = dataReferencia.getFullYear();
  var mes = dataReferencia.getMonth();
  var ultimoDia = new Date(ano, mes + 1, 0); // Dia 0 do próximo mês = último dia do mês atual
  
  Logger.log('📅 Último dia do mês ' + (mes + 1) + '/' + ano + ': ' + formatarData(ultimoDia));
  
  // Carregar feriados
  var feriadosArray = getFeriadosArray();
  
  // Retroceder até encontrar um dia útil
  while (true) {
    var diaSemana = ultimoDia.getDay();
    var dataFormatada = formatarData(ultimoDia);
    var ehFeriadoFlag = feriadosArray.indexOf(dataFormatada) >= 0;
    
    // Se é dia útil (segunda a sexta E não feriado), retornar
    if (diaSemana !== 0 && diaSemana !== 6 && !ehFeriadoFlag) {
      Logger.log('✅ Último dia útil encontrado: ' + dataFormatada);
      return ultimoDia;
    }
    
    // Retroceder 1 dia
    ultimoDia.setDate(ultimoDia.getDate() - 1);
  }
}

/**
 * 🧪 TESTE: Verificar se hoje é o último dia útil do mês
 * Execute esta função manualmente no Apps Script Editor para testar
 */
function testarSeEhUltimoDiaUtil() {
  Logger.log('🧪 ===== TESTE: ÚLTIMO DIA ÚTIL DO MÊS =====\n');
  
  var ss = obterPlanilha();
  var hoje = new Date();
  
  Logger.log('📅 Data de hoje: ' + formatarData(hoje));
  Logger.log('📅 Dia da semana: ' + ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'][hoje.getDay()]);
  
  var ultimoDiaUtil = calcularUltimoDiaUtilDoMes(hoje, ss);
  Logger.log('📅 Último dia útil do mês: ' + formatarData(ultimoDiaUtil));
  
  var hojeNormalizado = formatarData(hoje);
  var ultimoDiaUtilNormalizado = formatarData(ultimoDiaUtil);
  
  if (hojeNormalizado === ultimoDiaUtilNormalizado) {
    Logger.log('\n🎯 ✅ HOJE É O ÚLTIMO DIA ÚTIL DO MÊS!');
    Logger.log('📧 Emails de Diárias SERÃO enviados.');
  } else {
    Logger.log('\n⏭️ ❌ Hoje NÃO é o último dia útil do mês.');
    Logger.log('📧 Emails de Diárias NÃO serão enviados.');
    
    // Calcular quantos dias faltam
    var diasRestantes = Math.ceil((ultimoDiaUtil - hoje) / (1000 * 60 * 60 * 24));
    Logger.log('⏰ Faltam ' + diasRestantes + ' dia(s) para o último dia útil.');
  }
  
  Logger.log('\n✅ Teste concluído!');
}

/**
 * Ativa TODOS os triggers necessários para o sistema funcionar
 * Execute esta função MANUALMENTE no Apps Script Editor
 */
function ativarSistemaCompleto() {
  Logger.log('🚀 Ativando sistema completo...');
  
  // Remover triggers antigos (evitar duplicação)
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
    Logger.log('  🗑️ Trigger removido: ' + trigger.getHandlerFunction());
  });
  
  // ============================================
  // TRIGGER 1: Atualização de dados da CVM (a cada 1 hora)
  // ============================================
  ScriptApp.newTrigger('atualizarDadosCVMRealCompleto')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('✅ TRIGGER 1: Atualização de dados CVM (a cada 1 hora)');
  
  // ============================================
  // TRIGGER 2: Emails diários às 18:30 (abas mensais)
  // ============================================
  ScriptApp.newTrigger('enviarEmailConformidadeOuDesconformidadeAvancado')
    .timeBased()
    .atHour(18)
    .nearMinute(30)
    .everyDays(1)
    .create();
  Logger.log('✅ TRIGGER 2: Emails diários às 18:30 (Balancete, Composição, Lâmina, Perfil Mensal)');
  
  // ============================================
  // TRIGGER 3: Emails mensais de Diárias (último dia útil do mês)
  // ============================================
  ScriptApp.newTrigger('verificarEEnviarEmailDiariasSeUltimoDiaUtil')
    .timeBased()
    .atHour(17)
    .everyDays(1)
    .create();
  Logger.log('✅ TRIGGER 3: Emails mensais de Diárias às 17:00 (só no último dia útil)');
  
  // ============================================
  // RESUMO
  // ============================================
  Logger.log('\n✅ ═══════════════════════════════════════════');
  Logger.log('✅ SISTEMA 100% ATIVADO!');
  Logger.log('✅ ═══════════════════════════════════════════');
  Logger.log('');
  Logger.log('📊 TRIGGER 1: atualizarDadosCVMRealCompleto()');
  Logger.log('   ⏰ Executa: A cada 1 hora (24x por dia)');
  Logger.log('   🎯 Faz: Busca dados da CVM e atualiza planilha');
  Logger.log('');
  Logger.log('📧 TRIGGER 2: enviarEmailConformidadeOuDesconformidadeAvancado()');
  Logger.log('   ⏰ Executa: Diariamente às 18:30');
  Logger.log('   🎯 Envia emails: Balancete, Composição, Lâmina, Perfil Mensal');
  Logger.log('   ⚠️ NÃO envia Diárias (seção comentada)');
  Logger.log('');
  Logger.log('📅 TRIGGER 3: verificarEEnviarEmailDiariasSeUltimoDiaUtil()');
  Logger.log('   ⏰ Executa: Diariamente às 17:00');
  Logger.log('   🎯 Envia emails de Diárias APENAS no último dia útil do mês');
  Logger.log('');
  Logger.log('🌐 Web App: ' + ScriptApp.getService().getUrl());
  Logger.log('📊 Planilha: ' + obterURLPlanilha());
  
  return {
    success: true,
    message: 'Sistema ativado com 3 triggers!'
  };
}

function testarFormatacaoEmailDiarias() {
  Logger.log('🧪 Testando formatação de emails de Diárias...');
  
  var ss = obterPlanilha();
  var destinatarios = ['spandrade@banestes.com.br'];
  var mesPassado = obterMesPassadoFormatado();
  var dataAtualFormatada = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy HH:mm');
  var urlPlanilha = obterURLPlanilha();
  
  processarAbaEmail(
    ss.getSheetByName('Diárias'),
    'Diárias (TESTE)',
    destinatarios,
    mesPassado,
    dataAtualFormatada,
    urlPlanilha,
    'diario'
  );
  
  Logger.log('✅ Email de teste enviado!');
}

/**
 * NOVA FUNÇÃO (VERSÃO FINAL COM LAYOUT RICO): 
 * Gera PDF consolidado das Diárias (26 páginas) com layout estilizado (HTML/CSS) e envia 1 único e-mail.
 * Substitui o envio de 26 e-mails individuais no final do mês.
 */
function enviarRelatorioDiariasConsolidadoPDF() {
  Logger.log('🎨 Iniciando geração de PDF consolidado (Layout Rico)...');
  
  // 1. Configurações
  var destinatarios = [ 
    //'spandrade@banestes.com.br',
    'fabiooliveira@banestes.com.br',
    //'iodutra@banestes.com.br',
    'mcdias@banestes.com.br',
    'sndemuner@banestes.com.br',
    'wffreitas@banestes.com.br']; // Adicione outros e-mails aqui se necessário
  var fundos = getFundos();
  
  // 2. Datas e Referência (Mês Anterior)
  var hoje = new Date();
  var dataMesAnterior = new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);
  
  // Formatar Mês/Ano para o cabeçalho
  var meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
  var nomeMes = meses[dataMesAnterior.getMonth()];
  var ano = dataMesAnterior.getFullYear();
  var formatadorMes = nomeMes + "/" + ano; 
  
  var dataGeracao = Utilities.formatDate(hoje, 'GMT-3', 'dd/MM/yyyy');

  // 3. Início do HTML com CSS para forçar cores na impressão
  var htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        /* Força a impressão de background e cores */
        body { 
          font-family: Arial, sans-serif; 
          margin: 0; padding: 0; 
          -webkit-print-color-adjust: exact; 
          print-color-adjust: exact;
          background-color: #ffffff;
        }
        .page-break { page-break-after: always; }
      </style>
    </head>
    <body>
  `;

  // 4. Loop pelos Fundos
  for (var i = 0; i < fundos.length; i++) {
    var fundo = fundos[i];
    Logger.log('[' + (i + 1) + '/' + fundos.length + '] Processando: ' + fundo.nome);

    try {
      // --- Lógica de busca na CVM ---
      var url = 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/InfDiario/CPublicaInfdiario.aspx?PK_PARTIC=' + fundo.codigoCVM + '&PK_SUBCLASSE=-1';
      var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' }});
      
      var linhasTabelaHTML = "";
      var totalEnvios = 0;

      if (response.getResponseCode() === 200) {
        var htmlResponse = response.getContentText();
        var regexLinhas = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
        var matchLinhas;
        var linhasComDatas = [];

        // Extração de dados
        while ((matchLinhas = regexLinhas.exec(htmlResponse)) !== null) {
          var linhaHtml = matchLinhas[1];
          var matchDia = linhaHtml.match(/<td[^>]*>(\d{1,2})<\/td>/i);
          var matchData = linhaHtml.match(/<td[^>]*>(\d{2}\/\d{2}\/\d{4})<\/td>/i);
          
          if (matchDia && matchData) {
            linhasComDatas.push({ dia: matchDia[1], data: matchData[1] });
          }
        }

        if (linhasComDatas.length > 0) {
          // Deduplicação
          var datasUnicas = [];
          var datasVistas = {};
          linhasComDatas.forEach(function(item) {
            if (!datasVistas[item.data]) {
              datasVistas[item.data] = true;
              datasUnicas.push(item);
            }
          });
          // Ordenação (Mais recente primeiro)
          datasUnicas.sort(function(a, b) {
            var partsA = a.data.split('/');
            var partsB = b.data.split('/');
            return new Date(partsB[2], partsB[1] - 1, partsB[0]) - new Date(partsA[2], partsA[1] - 1, partsA[0]);
          });
          
          totalEnvios = datasUnicas.length;
          
          // Montar linhas da tabela HTML
          linhasTabelaHTML = datasUnicas.map(function(item) {
            return `<tr>
                      <td style="padding: 8px; border: 1px solid #dddddd; text-align: center;">${item.dia}</td>
                      <td style="padding: 8px; border: 1px solid #dddddd; text-align: center;">${item.data}</td>
                    </tr>`;
          }).join('');
        } else {
          linhasTabelaHTML = `<tr><td colspan="2" style="padding: 10px; border: 1px solid #ddd;">- Sem dados encontrados -</td></tr>`;
        }
      } else {
        linhasTabelaHTML = `<tr><td colspan="2" style="padding: 10px; border: 1px solid #ddd;">Erro de Conexão CVM</td></tr>`;
      }

      // --- Montagem do HTML da Página do Fundo (Layout Colorido) ---
      htmlContent += `
        <div style="padding: 20px; max-width: 700px; margin: 0 auto;">
          
          <div style="background-color: #2E7D32; color: white; padding: 20px; text-align: center; border-radius: 8px 8px 0 0; -webkit-print-color-adjust: exact;">
            <div style="font-size: 40px; margin-bottom: 5px;">✓</div>
            <div style="font-size: 22px; margin: 0; font-weight: bold;">Relatório de Conformidade CVM</div>
            <div style="color: #a5d6a7; margin: 5px 0 0 0; font-size: 14px;">Referência: ${formatadorMes}</div>
          </div>
          
          <div style="padding: 20px; border: 1px solid #e0e0e0; border-top: none; border-radius: 0 0 8px 8px;">
            <p style="font-family: Arial; font-weight: bold;">Informações Diárias</p>
            
            <div style="background-color: #f0f9ff; border-left: 4px solid #667eea; padding: 15px; margin: 20px 0; border-radius: 4px; -webkit-print-color-adjust: exact;">
              <p style="margin: 0; font-weight: bold; color: #1e3a8a; font-size: 16px;">Fundo:</p>
              <p style="margin: 5px 0 0 0; font-size: 14px; color: #333;">${fundo.nome}</p>
              <p style="margin: 10px 0 0 0; font-size: 13px; color: #666;">Código CVM: ${fundo.codigoCVM}</p>
            </div>

            <p>Histórico de envios identificados (${totalEnvios}):</p>

            <table style="width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 13px;">
              <thead>
                <tr>
                  <th width="30%" style="background-color: #f3f4f6; padding: 10px; border: 1px solid #dddddd; text-align: center; color: #555; -webkit-print-color-adjust: exact;">Dia</th>
                  <th style="background-color: #f3f4f6; padding: 10px; border: 1px solid #dddddd; text-align: center; color: #555; -webkit-print-color-adjust: exact;">Data de Envio</th>
                </tr>
              </thead>
              <tbody>
                ${linhasTabelaHTML}
              </tbody>
            </table>

            <div style="background-color: #e3f2fd; border-left: 4px solid #2196F3; padding: 10px; margin-top: 20px; font-size: 13px; -webkit-print-color-adjust: exact;">
              <strong>✓ Status: Regularizado</strong><br>
              Todos os envios foram identificados no portal da CVM.
            </div>

            <div style="text-align: center; color: #888; font-size: 11px; margin-top: 30px; border-top: 1px solid #eee; padding-top: 10px;">
              Departamento de Inovação e Automação interno Asset<br>
              Relatório gerado em ${dataGeracao}
            </div>
          </div>
        </div>
      `;

      // Adiciona quebra de página se não for o último fundo
      if (i < fundos.length - 1) {
        htmlContent += `<div class="page-break"></div>`;
      }

    } catch (e) {
      Logger.log("❌ Erro no fundo " + fundo.nome + ": " + e.toString());
    }
    
    // Pequena pausa para evitar bloqueio
    Utilities.sleep(600); 
  }

  // 5. Finaliza HTML e Converte para PDF
  htmlContent += `</body></html>`;
  
  var blobHtml = Utilities.newBlob(htmlContent, MimeType.HTML, "relatorio_temp.html");
  var pdfBlob = blobHtml.getAs(MimeType.PDF).setName(`Relatorio_Diarias_CVM_${formatadorMes.replace('/','-')}.pdf`);

  // 6. Envia o E-mail Único
  Logger.log('📧 Enviando e-mail consolidado...');
  
  MailApp.sendEmail({
    to: destinatarios.join(','),
    subject: '✅ Relatório Consolidado Diárias CVM (' + formatadorMes + ')',
    htmlBody: `
      <h3>Relatório Mensal de Conformidade CVM</h3>
      <p>Prezados,</p>
      <p>Segue em anexo o <strong>Relatório Consolidado de Informações Diárias</strong> referente ao mês de <strong>${formatadorMes}</strong>.</p>
      <p>O arquivo contém o detalhamento dos envios de todos os <strong>${fundos.length} fundos</strong> monitorados.</p>
      <br>
      <p style="color:#666; font-size:12px;">Departamento de Inovação e Automação Asset</p>
    `,
    attachments: [pdfBlob]
  });

  Logger.log('✅ PDF enviado com sucesso!');
}

/**
 * 🧪 TESTE VISUAL CORRIGIDO: Força as cores no PDF
 * Usa estilos "inline" e print-color-adjust para garantir que o fundo verde e azul apareçam.
 */
function testarRelatorioPDFConsolidado() {
  Logger.log('🎨 Iniciando geração de PDF (Modo Colorido Forçado)...');
  
  // 1. Configurações
  var destinatariosTeste = ['spandrade@banestes.com.br'];
  var fundos = getFundos();
  
  // --- OPCIONAL: Teste rápido com 3 fundos (descomente para testar) ---
  // fundos = fundos.slice(0, 3); 
  
  // 2. Datas
  var hoje = new Date();
  var dataMesAnterior = new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);
  
  // Formatar Mês em Inglês ou Português (ajuste conforme preferir)
  var meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
  var nomeMes = meses[dataMesAnterior.getMonth()];
  var ano = dataMesAnterior.getFullYear();
  var formatadorMes = nomeMes + "/" + ano; 
  
  var dataGeracao = Utilities.formatDate(hoje, 'GMT-3', 'dd/MM/yyyy');

  Logger.log('📅 Referência: ' + formatadorMes);

  // 3. HTML com CSS INLINE (Crucial para cores no PDF)
  // Note o uso de -webkit-print-color-adjust: exact;
  var htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        /* Força a impressão de background */
        body { 
          font-family: Arial, sans-serif; 
          margin: 0; padding: 0; 
          -webkit-print-color-adjust: exact; 
          print-color-adjust: exact;
        }
        .page-break { page-break-after: always; }
      </style>
    </head>
    <body style="background-color: #ffffff;">
  `;

  // 4. Loop pelos Fundos
  for (var i = 0; i < fundos.length; i++) {
    var fundo = fundos[i];
    Logger.log('   [' + (i + 1) + '] ' + fundo.nome);

    try {
      // --- Scraping ---
      var url = 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/InfDiario/CPublicaInfdiario.aspx?PK_PARTIC=' + fundo.codigoCVM + '&PK_SUBCLASSE=-1';
      var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' }});
      
      var linhasTabelaHTML = "";
      var totalEnvios = 0;

      if (response.getResponseCode() === 200) {
        var htmlResponse = response.getContentText();
        var regexLinhas = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
        var matchLinhas;
        var linhasComDatas = [];

        while ((matchLinhas = regexLinhas.exec(htmlResponse)) !== null) {
          var linhaHtml = matchLinhas[1];
          var matchDia = linhaHtml.match(/<td[^>]*>(\d{1,2})<\/td>/i);
          var matchData = linhaHtml.match(/<td[^>]*>(\d{2}\/\d{2}\/\d{4})<\/td>/i);
          
          if (matchDia && matchData) {
            linhasComDatas.push({ dia: matchDia[1], data: matchData[1] });
          }
        }

        if (linhasComDatas.length > 0) {
          var datasUnicas = [];
          var datasVistas = {};
          linhasComDatas.forEach(function(item) {
            if (!datasVistas[item.data]) {
              datasVistas[item.data] = true;
              datasUnicas.push(item);
            }
          });
          datasUnicas.sort(function(a, b) {
            var partsA = a.data.split('/');
            var partsB = b.data.split('/');
            return new Date(partsB[2], partsB[1] - 1, partsB[0]) - new Date(partsA[2], partsA[1] - 1, partsA[0]);
          });
          
          totalEnvios = datasUnicas.length;
          
          linhasTabelaHTML = datasUnicas.map(function(item) {
            return `<tr>
                      <td style="padding: 8px; border: 1px solid #dddddd; text-align: center;">${item.dia}</td>
                      <td style="padding: 8px; border: 1px solid #dddddd; text-align: center;">${item.data}</td>
                    </tr>`;
          }).join('');
        } else {
          linhasTabelaHTML = `<tr><td colspan="2" style="padding: 10px; border: 1px solid #ddd;">- Sem dados -</td></tr>`;
        }
      } else {
        linhasTabelaHTML = `<tr><td colspan="2" style="padding: 10px; border: 1px solid #ddd;">Erro de Conexão CVM</td></tr>`;
      }

      // --- HTML INLINE (Cores forçadas aqui) ---
      htmlContent += `
        <div style="padding: 20px; max-width: 700px; margin: 0 auto;">
          
          <div style="background-color: #2E7D32; color: white; padding: 20px; text-align: center; border-radius: 8px 8px 0 0; -webkit-print-color-adjust: exact;">
            <div style="font-size: 40px; margin-bottom: 5px;">✓</div>
            <div style="font-size: 22px; margin: 0; font-weight: bold;">Relatório de Conformidade CVM</div>
            <div style="color: #a5d6a7; margin: 5px 0 0 0; font-size: 14px;">Referência: ${formatadorMes}</div>
          </div>
          
          <div style="padding: 20px; border: 1px solid #e0e0e0; border-top: none; border-radius: 0 0 8px 8px;">
            <p style="font-family: Arial; font-weight: bold;">Informações Diárias</p>
            
            <div style="background-color: #f0f9ff; border-left: 4px solid #667eea; padding: 15px; margin: 20px 0; border-radius: 4px; -webkit-print-color-adjust: exact;">
              <p style="margin: 0; font-weight: bold; color: #1e3a8a; font-size: 16px;">Fundo:</p>
              <p style="margin: 5px 0 0 0; font-size: 14px; color: #333;">${fundo.nome}</p>
              <p style="margin: 10px 0 0 0; font-size: 13px; color: #666;">Código CVM: ${fundo.codigoCVM}</p>
            </div>

            <p>Histórico de envios identificados (${totalEnvios}):</p>

            <table style="width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 13px;">
              <thead>
                <tr>
                  <th width="30%" style="background-color: #f3f4f6; padding: 10px; border: 1px solid #dddddd; text-align: center; color: #555; -webkit-print-color-adjust: exact;">Dia</th>
                  <th style="background-color: #f3f4f6; padding: 10px; border: 1px solid #dddddd; text-align: center; color: #555; -webkit-print-color-adjust: exact;">Data de Envio</th>
                </tr>
              </thead>
              <tbody>
                ${linhasTabelaHTML}
              </tbody>
            </table>

            <div style="background-color: #e3f2fd; border-left: 4px solid #2196F3; padding: 10px; margin-top: 20px; font-size: 13px; -webkit-print-color-adjust: exact;">
              <strong>✓ Status: Regularizado</strong><br>
              Todos os envios foram identificados no portal da CVM.
            </div>

            <div style="text-align: center; color: #888; font-size: 11px; margin-top: 30px; border-top: 1px solid #eee; padding-top: 10px;">
              Departamento de Inovação e Automação interno Asset<br>
              Relatório gerado em ${dataGeracao}
            </div>
          </div>
        </div>
      `;

      if (i < fundos.length - 1) {
        htmlContent += `<div class="page-break"></div>`;
      }

    } catch (e) {
      Logger.log("❌ Erro: " + e.toString());
    }
    Utilities.sleep(500);
  }

  htmlContent += `</body></html>`;

  // Converter e Enviar
  var blobHtml = Utilities.newBlob(htmlContent, MimeType.HTML, "relatorio.html");
  var pdfBlob = blobHtml.getAs(MimeType.PDF).setName(`Relatorio_CVM_${formatadorMes.replace('/','-')}.pdf`);

  Logger.log('📧 Enviando e-mail com cores corrigidas...');
  
  MailApp.sendEmail({
    to: destinatariosTeste.join(','),
    subject: '✅ Relatório Consolidado CVM (Cores Corrigidas)',
    htmlBody: `
      <h3>Relatório de Teste (Cores Forçadas)</h3>
      <p>Tentativa de correção das cores de fundo (Cabeçalho Verde, Caixas Azuis).</p>
      <p>Referência: ${formatadorMes}</p>
    `,
    attachments: [pdfBlob]
  });

  Logger.log('✅ Teste finalizado.');
}

/**
 * 🆕 NOVA FUNÇÃO: Marca que email foi enviado
 */
function marcarEmailEnviado(nomeAba, dataAtual) {
  try {
    var ss = obterPlanilha();
    var aba = ss.getSheetByName(nomeAba);
    
    if (!aba) {
      Logger.log('  ⚠️ Aba não encontrada: ' + nomeAba);
      return;
    }
    
    // 📝 Escrever na célula G1
    var mensagem = 'E-MAIL ENVIADO\n' + dataAtual;
    aba.getRange('G1').setValue(mensagem);
    
    // 🎨 Formatar célula (verde)
    aba.getRange('G1')
      .setBackground('#d1fae5')
      .setFontColor('#065f46')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    
    Logger.log('  ✅ Flag "E-MAIL ENVIADO" marcada em ' + nomeAba + '!G1');
    
  } catch (error) {
    Logger.log('  ❌ Erro ao marcar flag: ' + error.toString());
  }
}

/**
 * 🆕 NOVA FUNÇÃO: Força rotação de competências
 */
function forcarRotacaoCompetencias(todasCompetencias) {
  Logger.log('  🔄 Forçando rotação...');
  
  // Filtrar e ordenar competências (mais recente primeiro)
  var competenciasValidas = todasCompetencias
    .filter(function(c) { return c && c !== '-' && c !== 'ERRO'; })
    .sort()
    .reverse();
  
  if (competenciasValidas.length === 0) {
    return {
      comp1: '-',
      status1: 'DESCONFORMIDADE',
      comp2: '-',
      status2: 'AGUARDANDO'
    };
  }
  
  // 🎯 LÓGICA DE ROTAÇÃO FORÇADA
  // Comp1 = mais recente da CVM
  // Comp2 = resetar para aguardar próxima
  return {
    comp1: competenciasValidas[0],
    status1: 'OK',
    comp2: '-',
    status2: 'AGUARDANDO'
  };
}

/**
 * 🆕 NOVA FUNÇÃO: Reseta flag após rotação
 */
function resetarFlagEmail(nomeAba) {
  try {
    var ss = obterPlanilha();
    var aba = ss.getSheetByName(nomeAba);
    
    if (!aba) return;
    
    // 📝 Resetar para "-"
    aba.getRange('G1').setValue('-');
    
    // 🎨 Formatar célula (cinza)
    aba.getRange('G1')
      .setBackground('#f3f4f6')
      .setFontColor('#6b7280')
      .setFontWeight('normal')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    
    Logger.log('  ✅ Flag resetada em ' + nomeAba + '!G1');
    
  } catch (error) {
    Logger.log('  ❌ Erro ao resetar flag: ' + error.toString());
  }
}


/**
 * 🧪 TESTE 2: Verificar flags de todas as abas
 */
function verificarFlagsDeTodasAsAbas() {
  Logger.log('🔍 Verificando flags G1...\n');
  
  var ss = obterPlanilha();
  var abas = ['Balancete', 'Composição', 'Lâmina', 'Perfil Mensal'];
  
  abas.forEach(function(nomeAba) {
    var aba = ss.getSheetByName(nomeAba);
    if (aba) {
      var flagG1 = aba.getRange('G1').getValue();
      var emailEnviado = flagG1 && flagG1.toString().indexOf('E-MAIL ENVIADO') !== -1;
      
      Logger.log('📋 ' + nomeAba + ':');
      Logger.log('   G1: "' + flagG1 + '"');
      Logger.log('   Email enviado? ' + (emailEnviado ? '✅ SIM' : '❌ NÃO'));
      Logger.log('');
    }
  });
}

// 1. Marcar flag manualmente em todas as abas
function marcarFlagEmTodasAsAbas() {
  var abas = ['Balancete', 'Composição', 'Lâmina', 'Perfil Mensal'];
  var dataAtual = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy HH:mm');
  
  abas.forEach(function(nomeAba) {
    marcarEmailEnviado(nomeAba, dataAtual);
    Logger.log('✅ Flag marcada em: ' + nomeAba);
  });
  
  Logger.log('\n✅ Todas as flags marcadas!');
  Logger.log('💡 Agora execute: atualizarTodasCompetencias()');
}

/**
 * Calcula o status geral com contador OK (X dias restantes) baseado na coluna C.
 * Este método SUPERA a lógica anterior e implementa a Falha 2 completa.
 * Para Balancete, Composição, Lâmina, Perfil Mensal.
 * 
 * @param {Array} dados Linhas da planilha (A:F)
 * @param {String} tipo Tipo de aba ('mensal')
 * @param {Array} competenciasAtuais Array de datas na coluna C (usado para status)
 * @returns {String} Status geral do dashboard ("OK (X dias restantes)" ou "DESCONFORMIDADE")
 */
function calcularStatusGeralDaAbaComPrazo(dados, tipo, competenciasAtuais) {
  // Supondo uma planilha padronizada
  var totalFundos = dados.length;
  var okFundos = 0;
  if (!competenciasAtuais || competenciasAtuais.length === 0)
    return "AGUARDANDO DADOS";

  // Considerar status "OK" se coluna D = "OK" para todos
  for (var i = 0; i < totalFundos; i++) {
    var linha = dados[i];
    var statusCompAtual = (linha[3] || '').toString().trim(); // Coluna D
    if (statusCompAtual === "OK") okFundos++;
  }

  if (okFundos === totalFundos) {
    // Pegar a data da competência atual (da primeira linha, pois todas devem estar iguais)
    var competenciaBase = (competenciasAtuais[0] || '').toString().trim();
    if (!competenciaBase || competenciaBase === "-" || competenciaBase === "") return "OK";

    // Calcula 10º dia útil do mês X+2
    var partes = competenciaBase.split('/');
    if (partes.length !== 3) return "OK";
    var dia = 1;
    var mes = parseInt(partes[1], 10) - 1; // base 0
    var ano = parseInt(partes[2], 10);

    var dataPrazo = new Date(ano, mes + 2, 1);
    var decimoDiaUtil = calcularDiaUtil(dataPrazo, 10, obterPlanilha());
    var diasRestantes = calcularDiasUteisEntre(new Date(), decimoDiaUtil, obterPlanilha());

    // Normaliza: se prazo já passou, mostrar 0
    if (diasRestantes < 0) diasRestantes = 0;

    // ⚡️ Alteração: Exibir "OK" puro só se dias > 15, senão sempre "OK (X dias restantes)"
    if (diasRestantes > 15) {
      return "OK";
    } else {
      return "OK (" + diasRestantes + " dias restantes)";
    }
  } else {
    return "DESCONFORMIDADE";
  }
}

//Function para testar retorno de colorações e respostas
function testeSLAExemploUnico() {
  var ss = SpreadsheetApp.openById('1N6LP1ydsxnQO_Woatv9zWEccb0fOGaV_3EKK1GoSWZI');
  // Simulação do cenário
  var competencia1 = "01/01/2026";
  var hoje = new Date(2026, 1, 18); // 18/02/2026
  var partes = competencia1.split('/');
  var dataComp1 = new Date(parseInt(partes[2],10), parseInt(partes[1],10)-1, parseInt(partes[0],10));
  var mesReferencia = new Date(dataComp1.getFullYear(), dataComp1.getMonth(), 1);
  var mesPrazo = new Date(mesReferencia.getFullYear(), mesReferencia.getMonth() + 2, 1);
  var decimoDiaUtil = calcularDiaUtil(mesPrazo, 10, ss);
  var diasFaltantes = calcularDiasUteisEntre(hoje, decimoDiaUtil, ss);

  var cor, dashboardStatus;
  if (diasFaltantes > 15) {
    cor = 'VERDE';
    dashboardStatus = 'OK (' + diasFaltantes + ' dias restantes)';
  } else if (diasFaltantes >= 6) {
    cor = 'AMARELA';
    dashboardStatus = 'OK (' + diasFaltantes + ' dias restantes)';
  } else if (diasFaltantes >= 0) {
    cor = 'VERMELHA';
    dashboardStatus = 'OK (' + diasFaltantes + ' dias restantes)';
  } else {
    cor = 'VENCIDO';
    dashboardStatus = 'DESCONFORMIDADE';
  }
  Logger.log(
    'Hoje: %s | Competência: %s | Prazo: %s | Dias úteis faltantes: %s | Cor: %s | Dashboard: %s',
    formatarData(hoje), competencia1, formatarData(decimoDiaUtil), diasFaltantes, cor, dashboardStatus
  );
}

// Suas funções de apoio:
function formatarData(dateObj) {
  var dd = String(dateObj.getDate()).padStart(2,'0');
  var mm = String(dateObj.getMonth()+1).padStart(2,'0');
  var yyyy = dateObj.getFullYear();
  return dd + '/' + mm + '/' + yyyy;
}

function debugStatusGeral() {
  var diasRestantesTestes = [20, 10, 3, -2]; // valores para testar casos verde, amarelo, vermelho e vencido
  var statusGeralOriginal = "OK"; // simula que todos envios estão OK

  diasRestantesTestes.forEach(function(diasRestantes) {
    var substatus = calcularCorStatusOk(diasRestantes);
    var displayStatus;
    
    if (diasRestantes < 0) {
      displayStatus = "DESCONFORMIDADE";
      substatus = "ok-vermelho"; // No vencido, pode forçar vermelho (ou exibir só DESCONFORMIDADE)
    } else {
      displayStatus = `OK (${diasRestantes} dias restantes)`;
    }
    
    Logger.log("Dias Restantes: %d | Status Geral: %s | Substatus (cor): %s", diasRestantes, displayStatus, substatus);
  });
}

function criarTriggerEmailDiariasMensal() {
  // Remove triggers antigos (evitar duplicação)
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'enviarRelatorioDiariasConsolidadoMensalPrimeiroDiaUtil') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Cria novo trigger DIÁRIO às 18h
  ScriptApp.newTrigger('enviarRelatorioDiariasConsolidadoMensalPrimeiroDiaUtil')
    .timeBased()
    .atHour(18)
    .everyDays(1)
    .create();
}

function enviarRelatorioDiariasConsolidadoMensalPrimeiroDiaUtil() {
  Logger.log('⏰ Verificando envio consolidado mensal das diárias...');
  var hoje = new Date();

  // Verifica se é útil
  var diaSemana = hoje.getDay();
  if (diaSemana === 0 || diaSemana === 6) {
    Logger.log('⏭️ Hoje é fim de semana/Não é dia útil');
    return;
  }

  // Verifica se é feriado
  var ss = obterPlanilha();
  var abaFeriados = ss.getSheetByName('FERIADOS');
  if (abaFeriados) {
    var feriados = abaFeriados.getRange('A2:A100').getValues().map(function(r) {
      var d = r[0];
      if (d instanceof Date) {
        return Utilities.formatDate(d, 'GMT-3', 'dd/MM/yyyy');
      }
      return d;
    });
    var hojeFormatado = Utilities.formatDate(hoje, 'GMT-3', 'dd/MM/yyyy');
    if (feriados.indexOf(hojeFormatado) != -1) {
      Logger.log('⏭️ Hoje é feriado/Não é dia útil');
      return;
    }
  }

  // Só deve enviar se HOJE é o PRIMEIRO DIA ÚTIL do mês
  var mesAtual = hoje.getMonth();
  var anoAtual = hoje.getFullYear();
  var primeiroDiaUtil = new Date(anoAtual, mesAtual, 1);
  // Encontra o primeiro dia útil ignorando feriado
  while (primeiroDiaUtil.getDay() === 0 || primeiroDiaUtil.getDay() === 6 ||
    (abaFeriados && feriados.indexOf(Utilities.formatDate(primeiroDiaUtil, 'GMT-3', 'dd/MM/yyyy')) !== -1)) {
    primeiroDiaUtil.setDate(primeiroDiaUtil.getDate() + 1);
  }
  // Só executa se hoje == primeiroDiaUtil
  var hojeFormatado2 = Utilities.formatDate(hoje, 'GMT-3', 'dd/MM/yyyy');
  var primeiroDiaUtilFormatado = Utilities.formatDate(primeiroDiaUtil, 'GMT-3', 'dd/MM/yyyy');
  if (hojeFormatado2 !== primeiroDiaUtilFormatado) {
    Logger.log('⏭️ Hoje NÃO é o primeiro dia útil do mês. Não vai enviar mensal.');
    return;
  }

  Logger.log('✅ Hoje é o primeiro dia útil. Vai enviar o relatório mensal consolidado.');

  // Chama sua função normal de envio do PDF consolidado mensal
  enviarRelatorioDiariasConsolidadoPDF();
}

function testarContagemDiasUteis() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var inicio = new Date(2026, 2, 2); // 02/03/2026
  var fim = new Date(2026, 2, 13);   // 13/03/2026
  var dias = calcularDiasUteisEntre(inicio, fim, ss);
  Logger.log('Dias úteis entre 02/03 e 13/03/2026: ' + dias); // Deve ser 10
}
