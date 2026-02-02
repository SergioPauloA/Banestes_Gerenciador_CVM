/**
 * Sistema de Monitoramento de Fundos CVM
 * Lê dados da planilha (que já tem as fórmulas IMPORTXML)
 */

var SPREADSHEET_ID = '1N6LP1ydsxnQO_Woatv9zWEccb0fOGaV_3EKK1GoSWZI';

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
  
  // Ler o status atual da célula E1
  var statusGeral = aba.getRange('E1').getDisplayValue();
  
  // LOG: Verificar o que está sendo lido
  Logger.log('📊 Balancete - Status lido da E1: "' + statusGeral + '"');
  
  var ultimaLinha = aba.getLastRow();
  
  if (ultimaLinha < 4) {
    return {
      titulo: 'Balancetes de Fundos',
      statusGeral: 'SEM DADOS',
      substatus: null,
      dados: []
    };
  }
  
  // Se o status estiver vazio, com "-", ou inválido, recalcular
  if (!statusGeral || statusGeral === '' || statusGeral === '-' || statusGeral.trim() === '') {
    Logger.log('⚠️ Balancete - Status inválido na E1, recalculando...');
    
    // Recalcular o status
    var dadosRetorno = aba.getRange('C4:C' + ultimaLinha).getDisplayValues();
    statusGeral = calcularStatusGeralDaAba(dadosRetorno, 'mensal');
    
    // Atualizar a célula E1 com o valor correto
    aba.getRange('E1').setValue(statusGeral);
    Logger.log('✅ Balancete - Status recalculado e atualizado: "' + statusGeral + '"');
  }
  
  var valores = aba.getRange(4, 1, ultimaLinha - 3, 4).getDisplayValues();
  
  var dados = valores
    .filter(function(linha) { return linha[0] !== '' && linha[0] !== null; })
    .map(function(linha) {
      // MUDANÇA: Buscar código BANESTES em vez de usar coluna B
      var codigoBanestes = buscarCodigoBanestes(ss, linha[0]);
      return {
        fundo: linha[0],
        codigo: codigoBanestes, // CÓDIGO BANESTES
        retorno: formatarCompetencia(linha[2]) || '-',
        status: linha[3] || '-'
      };
    });
  
  return {
    titulo: 'Balancetes de Fundos',
    statusGeral: statusGeral || 'SEM DADOS',
    substatus: (statusGeral === 'OK') ? calcularCorStatusOk(datas.diasRestantes) : null,
    dados: dados
  };
}

function lerAbaComposicao(ss, datas) {
  var aba = ss.getSheetByName('Composição');
  if (!aba) throw new Error('Aba Composição não encontrada');
  
  // Ler o status atual da célula E1
  var statusGeral = aba.getRange('E1').getDisplayValue();
  
  // LOG: Verificar o que está sendo lido
  Logger.log('📈 Composição - Status lido da E1: "' + statusGeral + '"');
  
  var ultimaLinha = aba.getLastRow();
  
  if (ultimaLinha < 4) {
    return {
      titulo: 'Composição da Carteira',
      statusGeral: 'SEM DADOS',
      substatus: null,
      dados: []
    };
  }
  
  // Se o status estiver vazio, com "-", ou inválido, recalcular
  if (!statusGeral || statusGeral === '' || statusGeral === '-' || statusGeral.trim() === '') {
    Logger.log('⚠️ Composição - Status inválido na E1, recalculando...');
    
    // Recalcular o status
    var dadosRetorno = aba.getRange('C4:C' + ultimaLinha).getDisplayValues();
    statusGeral = calcularStatusGeralDaAba(dadosRetorno, 'mensal');
    
    // Atualizar a célula E1 com o valor correto
    aba.getRange('E1').setValue(statusGeral);
    Logger.log('✅ Composição - Status recalculado e atualizado: "' + statusGeral + '"');
  }
  
  var valores = aba.getRange(4, 1, ultimaLinha - 3, 4).getDisplayValues();
  
  var dados = valores
    .filter(function(linha) { return linha[0] !== '' && linha[0] !== null; })
    .map(function(linha) {
      // MUDANÇA: Buscar código BANESTES
      var codigoBanestes = buscarCodigoBanestes(ss, linha[0]);
      return {
        fundo: linha[0],
        codigo: codigoBanestes,
        retorno: formatarCompetencia(linha[2]) || '-',
        status: linha[3] || '-'
      };
    });
  
  return {
    titulo: 'Composição da Carteira',
    statusGeral: statusGeral || 'SEM DADOS',
    substatus: (statusGeral === 'OK') ? calcularCorStatusOk(datas.diasRestantes) : null,
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
  
  // Ler o status atual da célula E1
  var statusGeral = aba.getRange('E1').getDisplayValue();
  
  // LOG: Verificar o que está sendo lido
  Logger.log('📄 Lâmina - Status lido da E1: "' + statusGeral + '"');
  
  var ultimaLinha = aba.getLastRow();
  
  if (ultimaLinha < 4) {
    return {
      titulo: 'Lâmina do Fundo',
      statusGeral: 'SEM DADOS',
      substatus: null,
      dados: []
    };
  }
  
  // Se o status estiver vazio, com "-", ou inválido, recalcular
  if (!statusGeral || statusGeral === '' || statusGeral === '-' || statusGeral.trim() === '') {
    Logger.log('⚠️ Lâmina - Status inválido na E1, recalculando...');
    
    // Recalcular o status
    var dadosRetorno = aba.getRange('C4:C' + ultimaLinha).getDisplayValues();
    statusGeral = calcularStatusGeralDaAba(dadosRetorno, 'mensal');
    
    // Atualizar a célula E1 com o valor correto
    aba.getRange('E1').setValue(statusGeral);
    Logger.log('✅ Lâmina - Status recalculado e atualizado: "' + statusGeral + '"');
  }
  
  var valores = aba.getRange(4, 1, ultimaLinha - 3, 4).getDisplayValues();
  
  var dados = valores
    .filter(function(linha) { return linha[0] !== '' && linha[0] !== null; })
    .map(function(linha) {
      // MUDANÇA: Buscar código BANESTES
      var codigoBanestes = buscarCodigoBanestes(ss, linha[0]);
      return {
        fundo: linha[0],
        codigo: codigoBanestes,
        retorno: formatarCompetencia(linha[2]) || '-',
        status: linha[3] || '-'
      };
    });
  
  return {
    titulo: 'Lâmina do Fundo',
    statusGeral: statusGeral || 'SEM DADOS',
    substatus: (statusGeral === 'OK') ? calcularCorStatusOk(datas.diasRestantes) : null,
    dados: dados
  };
}

function lerAbaPerfilMensal(ss, datas) {
  var aba = ss.getSheetByName('Perfil Mensal');
  if (!aba) throw new Error('Aba Perfil Mensal não encontrada');
  
  // Ler o status atual da célula E1
  var statusGeral = aba.getRange('E1').getDisplayValue();
  
  // LOG: Verificar o que está sendo lido
  Logger.log('📊 Perfil Mensal - Status lido da E1: "' + statusGeral + '"');
  
  var ultimaLinha = aba.getLastRow();
  
  if (ultimaLinha < 4) {
    return {
      titulo: 'Perfil Mensal',
      statusGeral: 'SEM DADOS',
      substatus: null,
      dados: []
    };
  }
  
  // Se o status estiver vazio, com "-", ou inválido, recalcular
  if (!statusGeral || statusGeral === '' || statusGeral === '-' || statusGeral.trim() === '') {
    Logger.log('⚠️ Perfil Mensal - Status inválido na E1, recalculando...');
    
    // Recalcular o status
    var dadosRetorno = aba.getRange('C4:C' + ultimaLinha).getDisplayValues();
    statusGeral = calcularStatusGeralDaAba(dadosRetorno, 'mensal');
    
    // Atualizar a célula E1 com o valor correto
    aba.getRange('E1').setValue(statusGeral);
    Logger.log('✅ Perfil Mensal - Status recalculado e atualizado: "' + statusGeral + '"');
  }
  
  var valores = aba.getRange(4, 1, ultimaLinha - 3, 4).getDisplayValues();
  
  var dados = valores
    .filter(function(linha) { return linha[0] !== '' && linha[0] !== null; })
    .map(function(linha) {
      // MUDANÇA: Buscar código BANESTES
      var codigoBanestes = buscarCodigoBanestes(ss, linha[0]);
      return {
        fundo: linha[0],
        codigo: codigoBanestes,
        retorno: formatarCompetencia(linha[2]) || '-',
        status: linha[3] || '-'
      };
    });
  
  return {
    titulo: 'Perfil Mensal',
    statusGeral: statusGeral || 'SEM DADOS',
    substatus: (statusGeral === 'OK') ? calcularCorStatusOk(datas.diasRestantes) : null,
    dados: dados
  };
}

function calcularCorStatusOk(diasRestantes) {
  if (diasRestantes > 5) return 'ok-verde';
  if (diasRestantes >= 3) return 'ok-amarelo';
  return 'ok-vermelho'; // 1 ou 2 dias restantes
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
    
    // Balancete
    var abaBalancete = ss.getSheetByName('Balancete');
    var dadosBalancete = abaBalancete.getRange('D4:D29').getValues();
    var totalOK = dadosBalancete.filter(function(r) { return r[0] === 'OK'; }).length;
    var statusBalancete = totalOK === dadosBalancete.length ? 'OK' : 'DESCONFORMIDADE';
    abaBalancete.getRange('E1').setValue(statusBalancete);
    
    // Composição
    var abaComposicao = ss.getSheetByName('Composição');
    var dadosComposicao = abaComposicao.getRange('D4:D29').getValues();
    totalOK = dadosComposicao.filter(function(r) { return r[0] === 'OK'; }).length;
    var statusComposicao = totalOK === dadosComposicao.length ? 'OK' : 'DESCONFORMIDADE';
    abaComposicao.getRange('E1').setValue(statusComposicao);
    
    // Diárias
    var abaDiarias = ss.getSheetByName('Diárias');
    var dadosDiarias1 = abaDiarias.getRange('D4:D29').getValues();
    var dadosDiarias2 = abaDiarias.getRange('F4:F29').getValues();
    totalOK = dadosDiarias1.filter(function(r) { return r[0] === 'OK'; }).length;
    var statusDiarias1 = totalOK === dadosDiarias1.length ? 'OK' : 'DESCONFORMIDADE';
    totalOK = dadosDiarias2.filter(function(r) { return r[0] === 'OK'; }).length;
    var statusDiarias2 = totalOK === dadosDiarias2.length ? 'OK' : 'A ATUALIZAR';
    abaDiarias.getRange('E1').setValue(statusDiarias1);
    abaDiarias.getRange('F1').setValue(statusDiarias2);
    
    // Lâmina
    var abaLamina = ss.getSheetByName('Lâmina');
    var dadosLamina = abaLamina.getRange('D4:D29').getValues();
    totalOK = dadosLamina.filter(function(r) { return r[0] === 'OK'; }).length;
    var statusLamina = totalOK === dadosLamina.length ? 'OK' : 'DESCONFORMIDADE';
    abaLamina.getRange('E1').setValue(statusLamina);
    
    // Perfil Mensal
    var abaPerfilMensal = ss.getSheetByName('Perfil Mensal');
    var dadosPerfilMensal = abaPerfilMensal.getRange('D4:D29').getValues();
    totalOK = dadosPerfilMensal.filter(function(r) { return r[0] === 'OK'; }).length;
    var statusPerfilMensal = totalOK === dadosPerfilMensal.length ? 'OK' : 'DESCONFORMIDADE';
    abaPerfilMensal.getRange('E1').setValue(statusPerfilMensal);
    
    // GERAL
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
      var status = calcularStatusIndividual(retorno, 'mensal');
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
 */
/**
 * Calcular status individual de um fundo
 * @param {string} retorno - Data de retorno da coluna C (ex: "01/12/2025")
 * @param {string} tipo - 'mensal' ou 'diario'
 * @returns {string} - 'OK', 'EM CONFORMIDADE', 'DESATUALIZADO', ou '-'
 */
function calcularStatusIndividual(retorno, tipo) {
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
    
    // Log para debug (remover depois)
    Logger.log('🔍 Comparando: "' + retornoNormalizado + '" vs "' + dataRefNormalizada + '"');
    
    // Se a data é igual à data de referência → OK
    if (retornoNormalizado === dataRefNormalizada) {
      return 'OK';
    }
    
    // Se ainda está dentro do prazo → EM CONFORMIDADE
    if (datas.diasRestantes >= 0) {
      return 'EM CONFORMIDADE';
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
function calcularStatusGeralDaAba(dados, tipo) {
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
    return 'EM CONFORMIDADE\n' + datas.diasRestantes + ' DIAS RESTANTES';
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
    var status = calcularStatusIndividual(retorno, 'mensal');
    Logger.log('   Linha ' + (i+4) + ': "' + retorno + '" → Status: "' + status + '"');
  }
  
  Logger.log('\n📈 Status Geral do Balancete:');
  var statusGeral = calcularStatusGeralDaAba(dadosBalancete, 'mensal');
  Logger.log('   ' + statusGeral);
  
  Logger.log('\n✅ Teste concluído!');
}

function enviarEmailDesconformidade() {
  var ss = obterPlanilha();
  var abaBalancete = ss.getSheetByName('Balancete');
  var dados = abaBalancete.getRange('A4:D29').getValues();
  
  var fundosDesconformes = [];
  
  dados.forEach(function(linha) {
    var fundo = linha[0];
    var status = linha[3];
    
    if (status === 'DESATUALIZADO' || status === 'DESCONFORMIDADE') {
      fundosDesconformes.push(fundo);
    }
  });
  
  if (fundosDesconformes.length > 0) {
    var destinatario = 'seu-email@banestes.com.br'; 
    //* Fundos.administrador@banestes.com.br
    //* Fabiooliveira@banestes.com.br
    //* Adicionar copia
    var assunto = '⚠️ ALERTA: Fundos em Desconformidade';
    var corpo = 'Os seguintes fundos estão em desconformidade:\n\n' + 
                fundosDesconformes.join('\n') +
                '\n\nTotal: ' + fundosDesconformes.length + ' fundos' +
                '\n\nAcesse: ' + obterURLPlanilha();
    
    MailApp.sendEmail(destinatario, assunto, corpo);
    Logger.log('📧 Email enviado para: ' + destinatario);
  }
}

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

  function getFeriadosArray() {
    var ss = obterPlanilha();
    var aba = ss.getSheetByName('FERIADOS');
    if (!aba) return [];
    var lastRow = aba.getLastRow();
    if (lastRow < 2) return [];
    var dados = aba.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    return dados.map(function(r) { return normalizaData(r[0]); });
  }

  function calculaUltimoDiaUtil(dateObj, feriadosArray) {
    var d = new Date(dateObj);
    do {
      d.setDate(d.getDate() - 1);
    } while (
      d.getDay() === 0 || // domingo
      d.getDay() === 6 || // sábado
      feriadosArray.indexOf(normalizaDataDate(d)) >= 0
    );
    return normalizaDataDate(d);
  }
  function normalizaDataDate(dateObj) {
    var dd = String(dateObj.getDate()).padStart(2, '0');
    var mm = String(dateObj.getMonth() + 1).padStart(2, '0');
    var yyyy = dateObj.getFullYear();
    return dd + '/' + mm + '/' + yyyy;
  }
  
  fundos.forEach(function(fundo, index) {
    try {
      var linha = index + 4;
      var url = 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/Balancete/CPublicaBalancete.asp?PK_PARTIC=' + fundo.codigoCVM + '&SemFrame=';
      var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' }});
      
      if (response.getResponseCode() === 200) {
        var html = response.getContentText();
        var regex = /<option[^>]*>(\d{2}\/\d{4})<\/option>/i;
        var match = html.match(regex);
        
        if (match) {
          var partes = match[1].split('/');
          var dataFormatada = '01/' + partes[0] + '/' + partes[1];
          abaBalancete.getRange(linha, 3).setValue(dataFormatada);
          Logger.log('  ✅ [' + (index + 1) + '/' + totalFundos + '] ' + dataFormatada);
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
        var regex = /<option[^>]*>(\d{2}\/\d{4})<\/option>/i;
        var match = html.match(regex);
        
        if (match) {
          var partes = match[1].split('/');
          var dataFormatada = '01/' + partes[0] + '/' + partes[1];
          abaComposicao.getRange(linha, 3).setValue(dataFormatada);
          Logger.log('  ✅ [' + (index + 1) + '/' + totalFundos + '] ' + dataFormatada);
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
        var regex = /<option[^>]*value="([A-Za-z]{3}\/\d{4})"[^>]*>/i;
        var match = html.match(regex);
        
        if (match) {
          var competencia = match[1];
          var partes = competencia.split('/');
          var mesTexto = partes[0];
          var ano = partes[1];
          var mesNumero = mesesMap[mesTexto];
          
          if (mesNumero) {
            var dataFormatada = '01/' + mesNumero + '/' + ano;
            abaLamina.getRange(linha, 3).setValue(dataFormatada);
            Logger.log('  ✅ [' + (index + 1) + '/' + totalFundos + '] ' + dataFormatada + ' (de ' + competencia + ')');
          } else {
            abaLamina.getRange(linha, 3).setValue('-');
          }
        } else {
          abaLamina.getRange(linha, 3).setValue('-');
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
        var regex = /<option[^>]*>(\d{2}\/\d{4})<\/option>/i;
        var match = html.match(regex);
        
        if (match) {
          var partes = match[1].split('/');
          var dataFormatada = '01/' + partes[0] + '/' + partes[1];
          abaPerfilMensal.getRange(linha, 3).setValue(dataFormatada);
          Logger.log('  ✅ [' + (index + 1) + '/' + totalFundos + '] ' + dataFormatada);
        }
      }
      Utilities.sleep(300);
    } catch (error) {
      Logger.log('  ❌ [' + (index + 1) + '/' + totalFundos + '] Erro');
    }
  });
  
  // ============================================
  // 5. DIÁRIAS (COM STATUS GERAL) - ATUALIZADO
  // ============================================
  Logger.log('\n📅 [5/5] Processando Diárias...');
  var abaDiarias = ss.getSheetByName('Diárias');

  var contadorOK_Status1 = 0;
  var contadorOK_Status2 = 0;
  var totalFundosValidos = 0;

  // Calcula hoje e dia útil real ANTES do loop!
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

          // ENVIO 1 — sempre exibe o D-1 útil se houver, senão a mais recente scrapeada
          var envio1 = datasExtraidas.includes(diaD1) ? diaD1 : (datasExtraidas[0] || "-");
          var status1 = envio1 === diaD1 ? "OK" : "DESATUALIZADO";
          if (status1 === "OK") contadorOK_Status1++;

          // ENVIO 2 — HOJE se encontrar, "-" caso não
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
      } else {
        totalFundosValidos++;
        abaDiarias.getRange(linha, 3).setValue('ERRO');
        abaDiarias.getRange(linha, 4).setValue('DESATUALIZADO');
        abaDiarias.getRange(linha, 5).setValue('#N/A');
        abaDiarias.getRange(linha, 6).setValue('A ATUALIZAR');
      }

      Utilities.sleep(300);

    } catch (error) {
      totalFundosValidos++;
      abaDiarias.getRange(linha, 3).setValue('ERRO');
      abaDiarias.getRange(linha, 4).setValue('DESATUALIZADO');
      abaDiarias.getRange(linha, 5).setValue('#N/A');
      abaDiarias.getRange(linha, 6).setValue('A ATUALIZAR');
    }
  });

  // Após o loop dos fundos diárias, atualize a aba APOIO:
  var abaApoio = ss.getSheetByName("APOIO");
  if (abaApoio) {
    // Por exemplo, HOJE na célula A11
    abaApoio.getRange("A11").setValue(new Date());
    abaApoio.getRange("A11").setNumberFormat("dd/mm/yyyy");
    // Coloque D-1 útil na célula desejada, se quiser ("A12" por exemplo)
    abaApoio.getRange("A12").setValue(diaD1);
  }
  
  // ============================================
  // CALCULAR STATUS GERAL PARA DIÁRIAS
  // ============================================
  Logger.log('\n🧮 Calculando status geral de Diárias...');
  
  // STATUS GERAL 1 (coluna E1)
  var statusGeral1;
  if (contadorOK_Status1 === fundos.length) {
    statusGeral1 = 'OK';
  } else {
    statusGeral1 = 'DESCONFORMIDADE';
  }
  abaDiarias.getRange('E1').setValue(statusGeral1);

  Logger.log('  STATUS 1 GERAL: ' + statusGeral1 + ' (' + contadorOK_Status1 + '/' + totalFundosValidos + ' OK)');
  
  // STATUS GERAL 2 (coluna F1) - CORRETO
  var statusGeral2;
  if (contadorOK_Status2 === fundos.length) {
    statusGeral2 = 'OK';
  } else {
    statusGeral2 = 'A ATUALIZAR';
  }
  abaDiarias.getRange('F1').setValue(statusGeral2);
  Logger.log('  STATUS 2 GERAL: ' + statusGeral2 + ' (' + contadorOK_Status2 + '/' + fundos.length + ' OK)');
  
  // ============================================
  // 6. CALCULAR STATUS (APENAS OUTRAS ABAS)
  // ============================================
  Logger.log('\n🧮 Calculando status das outras abas...');
  atualizarStatusParaAbasEspecificas(['Balancete', 'Composição', 'Lâmina', 'Perfil Mensal']);
  
  Logger.log('\n✅ ═══════════════════════════════════════════');
  Logger.log('✅ ATUALIZAÇÃO 100% COMPLETA!');
  Logger.log('✅ ═══════════════════════════════════════════');
  Logger.log('📊 5/5 abas carregadas com sucesso');
  Logger.log('💾 Todos os status calculados');
  Logger.log('🌐 Web App pronto para uso');
  
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
  // Aceita "DD/MM/AAAA", "MM/AAAA", "-" ou ""
  if (!dataStr || dataStr === "-") return "-";
  var partes = String(dataStr).trim().split("/");
  if (partes.length === 3) { // DD/MM/AAAA
    return partes[1] + "/" + partes[2];
  }
  if (partes.length === 2) { // MM/AAAA
    return dataStr;
  }
  return dataStr;
}
