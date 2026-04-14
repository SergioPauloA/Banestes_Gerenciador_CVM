/**
 * DateUtils.gs - Funções de cálculo de datas
 * 
 * CORREÇÃO FINAL: Ignora aba APOIO se a data for fim de semana
 */

// Execution-level caches (reset on each new script invocation)
var _datasReferenciaCache = null;
var _feriadosCache = null;

function getDatasReferencia() {
  if (_datasReferenciaCache !== null) {
    return _datasReferenciaCache;
  }

  Logger.log('📅 getDatasReferencia: calculando datas...');
  
  var ss = SpreadsheetApp.openById('1N6LP1ydsxnQO_Woatv9zWEccb0fOGaV_3EKK1GoSWZI');
  var hoje = new Date();
  
  // 🔥 Se hoje é fim de semana, ajustar para o próximo dia útil
  var diaParaCalculo = new Date(hoje);
  while (diaParaCalculo.getDay() === 0 || diaParaCalculo.getDay() === 6) {
    diaParaCalculo.setDate(diaParaCalculo.getDate() + 1);
  }
  
  Logger.log('  Data real: ' + formatarData(hoje));
  Logger.log('  Data de trabalho: ' + formatarData(diaParaCalculo));
  
  // DIADREF1 (D-2 em dias ÚTEIS)
  var dataD1 = calcularDiaUtil(diaParaCalculo, -2, ss);
  var diaD1 = formatarData(dataD1);
  
  // DIADREF2 (D-1 em dias ÚTEIS)
  var dataD2 = calcularDiaUtil(diaParaCalculo, -1, ss);
  var diaD2 = formatarData(dataD2);
  
  // Supondo que o mês de referência é "mesAnterior" (competência X):
  var mesReferencia = new Date(diaParaCalculo.getFullYear(), diaParaCalculo.getMonth() - 1, 1); // X
  var diaMesRef = formatarData(mesReferencia);

  // PRAZO: 10º dia útil do mês ATUAL
  var mesAtual = new Date(diaParaCalculo.getFullYear(), diaParaCalculo.getMonth(), 1);
  var decimoDiaUtil = calcularDiaUtil(mesAtual, 10, ss);
  var diaMesRef2 = formatarData(decimoDiaUtil);
  
  // Calcular dias restantes até o prazo
  var diasRestantes = calcularDiasUteisEntre(diaParaCalculo, decimoDiaUtil, ss);
  
  // 🔥 Se o prazo do mês atual já passou (deadline estritamente antes de hoje), avançar para o próximo ciclo:
  // Nota: calcularDiasUteisEntre pode retornar 0 quando o prazo caiu num fim de semana/feriado
  // imediatamente antes de hoje, portanto usamos comparação de datas diretamente.
  // - nova competência aguardada = 1º dia do mês atual
  // - decimoDiaUtil = 10º dia útil do próximo mês (novo prazo)
  var hojeNorm = new Date(diaParaCalculo.getFullYear(), diaParaCalculo.getMonth(), diaParaCalculo.getDate()).getTime();
  var prazoNorm = new Date(decimoDiaUtil.getFullYear(), decimoDiaUtil.getMonth(), decimoDiaUtil.getDate()).getTime();
  if (prazoNorm < hojeNorm) {
    var novoMesReferencia = new Date(diaParaCalculo.getFullYear(), diaParaCalculo.getMonth(), 1);
    diaMesRef = formatarData(novoMesReferencia);
    var proximoMes = new Date(diaParaCalculo.getFullYear(), diaParaCalculo.getMonth() + 1, 1);
    decimoDiaUtil = calcularDiaUtil(proximoMes, 10, ss);
    diaMesRef2 = formatarData(decimoDiaUtil);
    diasRestantes = calcularDiasUteisEntre(diaParaCalculo, decimoDiaUtil, ss);
    Logger.log('  ⚡ Prazo do mês atual expirou. Avançando para o próximo ciclo.');
  }
  
  Logger.log('  1º dia mês referência: ' + diaMesRef);
  Logger.log('  10º dia útil (prazo): ' + diaMesRef2);
  Logger.log('  🔥 Dias restantes: ' + diasRestantes);
  
  _datasReferenciaCache = {
    hoje: formatarData(diaParaCalculo),
    diaMesRef: diaMesRef,
    diaMesRef2: diaMesRef2,
    diaDD: formatarData(diaParaCalculo),
    diaD1: diaD1,
    diaD2: diaD2,
    diasRestantes: diasRestantes
  };
  return _datasReferenciaCache;
}


function calcularDatasManualmente() {
  var ss = SpreadsheetApp.openById('1N6LP1ydsxnQO_Woatv9zWEccb0fOGaV_3EKK1GoSWZI');
  var hoje = new Date();
  
  // 🔥 CORREÇÃO CRÍTICA: Se hoje é fim de semana, ajustar para o próximo dia útil
  var diaParaCalculo = new Date(hoje);
  while (diaParaCalculo.getDay() === 0 || diaParaCalculo.getDay() === 6) {
    diaParaCalculo.setDate(diaParaCalculo.getDate() + 1); // AVANÇAR para segunda
  }
  
  Logger.log('📅 calcularDatasManualmente:');
  Logger.log('  Data real (hoje): ' + formatarData(hoje));
  Logger.log('  Data ajustada (dia útil): ' + formatarData(diaParaCalculo));
  
  // DIADREF1 (D-2 em dias ÚTEIS)
  var dataD1 = calcularDiaUtil(diaParaCalculo, -2, ss);
  var diaD1 = formatarData(dataD1);
  
  // DIADREF2 (D-1 em dias ÚTEIS)
  var dataD2 = calcularDiaUtil(diaParaCalculo, -1, ss);
  var diaD2 = formatarData(dataD2);
  
  // DIAMESREF (1º dia do mês anterior)
  var mesAnterior = new Date(diaParaCalculo.getFullYear(), diaParaCalculo.getMonth() - 1, 1);
  var diaMesRef = formatarData(mesAnterior);
  
  // DIAMESREF2 (10º dia útil do mês atual)
  var mesAtual = new Date(diaParaCalculo.getFullYear(), diaParaCalculo.getMonth(), 1);
  var decimoDiaUtil = calcularDiaUtil(mesAtual, 10, ss);
  var diaMesRef2 = formatarData(decimoDiaUtil);
  
  // Calcular dias restantes até o prazo
  var diasRestantes = calcularDiasUteisEntre(diaParaCalculo, decimoDiaUtil, ss);
  
  // 🔥 Se o prazo do mês atual já passou (deadline estritamente antes de hoje), avançar para o próximo ciclo:
  // Nota: calcularDiasUteisEntre pode retornar 0 quando o prazo caiu num fim de semana/feriado
  // imediatamente antes de hoje, portanto usamos comparação de datas diretamente.
  // - competência aguardada passa a ser o 1º dia do mês atual
  // - decimoDiaUtil = 10º dia útil do próximo mês (novo prazo)
  var hojeNorm = new Date(diaParaCalculo.getFullYear(), diaParaCalculo.getMonth(), diaParaCalculo.getDate()).getTime();
  var prazoNorm = new Date(decimoDiaUtil.getFullYear(), decimoDiaUtil.getMonth(), decimoDiaUtil.getDate()).getTime();
  if (prazoNorm < hojeNorm) {
    var mesReferenciaAtualizado = new Date(diaParaCalculo.getFullYear(), diaParaCalculo.getMonth(), 1);
    diaMesRef = formatarData(mesReferenciaAtualizado);
    var mesProximoPrazo = new Date(diaParaCalculo.getFullYear(), diaParaCalculo.getMonth() + 1, 1);
    decimoDiaUtil = calcularDiaUtil(mesProximoPrazo, 10, ss);
    diaMesRef2 = formatarData(decimoDiaUtil);
    diasRestantes = calcularDiasUteisEntre(diaParaCalculo, decimoDiaUtil, ss);
    Logger.log('  ⚡ Prazo do mês atual expirou. Avançando para o próximo ciclo.');
  }
  
  Logger.log('  10º dia útil (prazo): ' + diaMesRef2);
  Logger.log('  🔥 Dias restantes: ' + diasRestantes);
  
  return {
    hoje: formatarData(diaParaCalculo), // 🔥 USAR DATA AJUSTADA
    diaMesRef: diaMesRef,
    diaMesRef2: diaMesRef2,
    diaDD: formatarData(diaParaCalculo),
    diaD1: diaD1,
    diaD2: diaD2,
    diasRestantes: diasRestantes
  };
}

function calcularDiaUtil(dataInicial, diasUteis, ss) {
  var resultado = new Date(dataInicial);
  if (diasUteis === 0) return resultado;

  var diasAdicionados = 0;

  // Para datas futuras (ex.: 10º dia útil do mês), iniciar a contagem no próprio
  // dia inicial e somar apenas quando ele for dia útil.
  if (diasUteis > 0) {
    while (true) {
      var diaSemana = resultado.getDay();
      if (diaSemana !== 0 && diaSemana !== 6 && !ehFeriado(resultado, ss)) {
        diasAdicionados++;
      }
      if (diasAdicionados >= diasUteis) break;
      resultado.setDate(resultado.getDate() + 1);
    }
    return resultado;
  }

  // Para datas passadas (D-1, D-2), manter regra original: não contar o dia inicial.
  var diasRestantes = Math.abs(diasUteis);
  while (diasAdicionados < diasRestantes) {
    resultado.setDate(resultado.getDate() - 1);
    var diaSemana = resultado.getDay();
    if (diaSemana !== 0 && diaSemana !== 6 && !ehFeriado(resultado, ss)) {
      diasAdicionados++;
    }
  }

  return resultado;
}

/**
 * Calcula dias úteis RESTANTES entre duas datas
 * 
 * REGRA CORRIGIDA:
 * - NÃO conta o dia de HOJE (já estamos nele)
 * - NÃO conta o dia do PRAZO (é o deadline)
 * - NÃO conta fins de semana
 * - NÃO conta feriados
 * 
 * Exemplo: Hoje 03/02/2026 até prazo 13/02/2026
 * Conta: 04, 05, 06, 07, 10, 11, 12 = 7 dias úteis
 */
function calcularDiasUteisEntre(dataInicio, dataFim, ss) {
  var diasUteis = 0;
  var dataAtual = new Date(dataInicio);
  
  // Normalizar datas para meia-noite
  dataAtual.setHours(0, 0, 0, 0);
  var dataFimNormalizada = new Date(dataFim);
  dataFimNormalizada.setHours(0, 0, 0, 0);
  
  // Se o prazo já passou
  if (dataFimNormalizada <= dataAtual) {
    var temp = new Date(dataAtual);
    temp.setDate(temp.getDate() - 1);
    
    while (temp > dataFimNormalizada) {
      var diaSemana = temp.getDay();
      if (diaSemana !== 0 && diaSemana !== 6) {
        if (!ehFeriado(temp, ss)) {
          diasUteis--;
        }
      }
      temp.setDate(temp.getDate() - 1);
    }
    return diasUteis;
  }
  
  // 🔥 CONTAR DE AMANHÃ ATÉ ANTES DO PRAZO
  var temp = new Date(dataAtual);
  temp.setDate(temp.getDate() + 1); // Pular HOJE
  
  while (temp < dataFimNormalizada) { // Parar ANTES do prazo
    var diaSemana = temp.getDay();
    
    if (diaSemana !== 0 && diaSemana !== 6) {
      if (!ehFeriado(temp, ss)) {
        diasUteis++;
      }
    }
    
    temp.setDate(temp.getDate() + 1);
  }
  
  return diasUteis;
}

function ehFeriado(data, ss) {
  try {
    if (_feriadosCache === null) {
      if (!ss) {
        ss = SpreadsheetApp.openById('1N6LP1ydsxnQO_Woatv9zWEccb0fOGaV_3EKK1GoSWZI');
      }
      var abaFeriados = ss.getSheetByName('FERIADOS');
      _feriadosCache = [];
      if (abaFeriados) {
        var feriadosData = abaFeriados.getRange('A2:A100').getValues();
        for (var j = 0; j < feriadosData.length; j++) {
          if (feriadosData[j][0]) {
            _feriadosCache.push(formatarData(new Date(feriadosData[j][0])));
          }
        }
      }
    }
    
    var dataFormatada = formatarData(data);
    return _feriadosCache.indexOf(dataFormatada) >= 0;
  } catch (error) {
    return false;
  }
}

function formatarData(data) {
  var dia = String(data.getDate()).padStart(2, '0');
  var mes = String(data.getMonth() + 1).padStart(2, '0');
  var ano = data.getFullYear();
  return dia + '/' + mes + '/' + ano;
}

// ============================================
// FUNÇÕES DE TESTE
// ============================================

function testarContagemDiasCompleta() {
  Logger.log('🧪 ===== TESTE DE CONTAGEM DE DIAS =====\n');
  
  var ss = SpreadsheetApp.openById('1N6LP1ydsxnQO_Woatv9zWEccb0fOGaV_3EKK1GoSWZI');
  
  // Teste 1: 03/02/2026 até 13/02/2026
  Logger.log('📅 TESTE 1: 03/02/2026 até 13/02/2026');
  var hoje1 = new Date(2026, 1, 3);
  var prazo1 = new Date(2026, 1, 13);
  var resultado1 = calcularDiasUteisEntre(hoje1, prazo1, ss);
  Logger.log('Resultado: ' + resultado1 + ' dias');
  Logger.log('Esperado: 7 dias');
  Logger.log(resultado1 === 7 ? '✅ PASSOU\n' : '❌ FALHOU\n');
  
  // Teste 2: Usando getDatasReferencia (data real)
  Logger.log('📅 TESTE 2: Usando getDatasReferencia()');
  var datas = getDatasReferencia();
  Logger.log('Dias restantes: ' + datas.diasRestantes);
  Logger.log('Hoje: ' + datas.hoje);
  Logger.log('Prazo: ' + datas.diaMesRef2);
  Logger.log('\n✅ Teste concluído!');
}

function testarAtualizacaoApoio() {
  Logger.log('🧪 Testando atualização da aba APOIO...\n');
  atualizarAbaApoioComDatas();
}

function verificarAbaApoio() {
  var ss = SpreadsheetApp.openById('1N6LP1ydsxnQO_Woatv9zWEccb0fOGaV_3EKK1GoSWZI');
  var abaApoio = ss.getSheetByName('APOIO');
  
  if (!abaApoio) {
    Logger.log('❌ Aba APOIO não existe!');
    return;
  }
  
  Logger.log('✅ Aba APOIO encontrada!');
  Logger.log('\n📋 Verificando células importantes:');
  
  // Linha 17
  var a17 = abaApoio.getRange('A17').getValue();
  var b17 = abaApoio.getRange('B17').getValue();
  var c17 = abaApoio.getRange('C17').getValue();
  
  Logger.log('\n🔹 LINHA 17:');
  Logger.log('  A17 (HOJE): ' + a17);
  Logger.log('  B17 (D-2): ' + b17);
  Logger.log('  C17 (LINHAD): ' + c17);
  
  // Linha 18
  var a18 = abaApoio.getRange('A18').getValue();
  var b18 = abaApoio.getRange('B18').getValue();
  var c18 = abaApoio.getRange('C18').getValue();
  
  Logger.log('\n🔹 LINHA 18:');
  Logger.log('  A18: ' + a18);
  Logger.log('  B18 (D-1): ' + b18);
  Logger.log('  C18: ' + c18);
  
  // Testar getDisplayValue
  Logger.log('\n🔹 USANDO getDisplayValue():');
  Logger.log('  B17 (display): ' + abaApoio.getRange('B17').getDisplayValue());
  Logger.log('  B18 (display): ' + abaApoio.getRange('B18').getDisplayValue());
  
  // Linha 1 (DATA MENSAL REFERENCIA)
  var d1 = abaApoio.getRange('D1').getValue();
  var e1 = abaApoio.getRange('E1').getValue();
  
  Logger.log('\n🔹 LINHA 1:');
  Logger.log('  D1 (1º mês anterior): ' + d1);
  Logger.log('  E1 (1º mês atual): ' + e1);
}

/**
 * Helper function para criar named ranges de datas
 */
function criarNamedRangesDatas(ss, abaApoio) {
  ss.setNamedRange('DIAMESREF', abaApoio.getRange('D1'));
  ss.setNamedRange('DIAMESREF2', abaApoio.getRange('F1'));
  ss.setNamedRange('DIADDD', abaApoio.getRange('A17'));
}

function criarAbaApoioComValores() {
  var ss = SpreadsheetApp.openById('1N6LP1ydsxnQO_Woatv9zWEccb0fOGaV_3EKK1GoSWZI');
  var abaApoio = ss.getSheetByName('APOIO');
  
  if (!abaApoio) {
    Logger.log('❌ Aba APOIO não existe. Criando...');
    abaApoio = ss.insertSheet('APOIO');
  } else {
    Logger.log('✅ Aba APOIO encontrada. Limpando...');
    abaApoio.clear();
  }
  
  // Verificar/criar aba FERIADOS
  var abaFeriados = ss.getSheetByName('FERIADOS');
  if (!abaFeriados) {
    Logger.log('⚠️ Criando aba FERIADOS...');
    abaFeriados = ss.insertSheet('FERIADOS');
    abaFeriados.getRange('A1').setValue('DATA');
    abaFeriados.getRange('A2').setValue(new Date(2026, 0, 1));  // Ano Novo
    abaFeriados.getRange('A3').setValue(new Date(2026, 3, 21)); // Tiradentes
    abaFeriados.getRange('A4').setValue(new Date(2026, 4, 1));  // Dia do Trabalho
    abaFeriados.getRange('A5').setValue(new Date(2026, 8, 7));  // Independência
    abaFeriados.getRange('A6').setValue(new Date(2026, 9, 12)); // N. Sra. Aparecida
    abaFeriados.getRange('A7').setValue(new Date(2026, 10, 2)); // Finados
    abaFeriados.getRange('A8').setValue(new Date(2026, 10, 15)); // Proclamação
    abaFeriados.getRange('A9').setValue(new Date(2026, 11, 25)); // Natal
  }
  
  // ============================================
  // CALCULAR DATAS VIA CÓDIGO
  // ============================================
  var hoje = new Date();
  
  // Se hoje é fim de semana, recuar para sexta-feira
  var diaParaCalculo = new Date(hoje);
  while (diaParaCalculo.getDay() === 0 || diaParaCalculo.getDay() === 6) {
    diaParaCalculo.setDate(diaParaCalculo.getDate() - 1);
  }
  
  // D-2 (2 dias úteis atrás)
  var dataD2Uteis = calcularDiaUtil(diaParaCalculo, -2, ss);
  
  // D-1 (1 dia útil atrás)
  var dataD1Util = calcularDiaUtil(diaParaCalculo, -1, ss);
  
  // 1º dia do mês anterior
  var mesAnterior = new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);
  
  // 1º dia do mês atual
  var mesAtual = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  
  // 10º dia útil do mês atual
  var decimoDiaUtil = calcularDiaUtil(mesAtual, 10, ss);
  
  Logger.log('📅 Datas calculadas:');
  Logger.log('  Hoje: ' + formatarData(hoje));
  Logger.log('  D-2 úteis: ' + formatarData(dataD2Uteis));
  Logger.log('  D-1 útil: ' + formatarData(dataD1Util));
  Logger.log('  1º mês anterior: ' + formatarData(mesAnterior));
  Logger.log('  1º mês atual: ' + formatarData(mesAtual));
  Logger.log('  10º dia útil: ' + formatarData(decimoDiaUtil));
  
  // ============================================
  // LINHA 1: DATA MENSAL REFERENCIA
  // ============================================
  abaApoio.getRange('C1').setValue('DATA MENSAL REFERENCIA');
  abaApoio.getRange('D1').setValue(formatarData(mesAnterior));
  abaApoio.getRange('E1').setValue(formatarData(mesAtual));
  abaApoio.getRange('F1').setValue(formatarData(decimoDiaUtil));
  
  // ============================================
  // LINHA 2-6: URLs
  // ============================================
  abaApoio.getRange('C2').setValue('BALANCETE');
  abaApoio.getRange('D2').setValue('https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/Balancete/CPublicaBalancete.asp?PK_PARTIC=');
  abaApoio.getRange('E2').setValue('&SemFrame=');
  abaApoio.getRange('F2').setValue('/html/body/form/table/tbody/tr[1]/td/select');
  
  abaApoio.getRange('C3').setValue('COMPOSIÇÃO');
  abaApoio.getRange('D3').setValue('https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/CDA/CPublicaCDA.aspx?PK_PARTIC=');
  abaApoio.getRange('E3').setValue('&SemFrame=');
  abaApoio.getRange('F3').setValue('/html/body/form/table/tbody/tr[1]/td/select');
  
  abaApoio.getRange('C4').setValue('DIÁRIAS');
  abaApoio.getRange('D4').setValue('https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/InfDiario/CPublicaInfdiario.aspx?PK_PARTIC=');
  abaApoio.getRange('E4').setValue('&PK_SUBCLASSE=-1');
  abaApoio.getRange('F4').setValue('/html/body/form/table[2]/tbody/tr[');
  abaApoio.getRange('G4').setValue(']/td[8]');
  
  abaApoio.getRange('C5').setValue('LÂMINA');
  abaApoio.getRange('D5').setValue('https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/CPublicaLamina.aspx?PK_PARTIC=');
  abaApoio.getRange('E5').setValue('&PK_SUBCLASSE=-1');
  abaApoio.getRange('F5').setValue('/html/body/form/table[1]/tbody/tr[1]/td/select');
  
  abaApoio.getRange('C6').setValue('PERFIL MENSAL');
  abaApoio.getRange('D6').setValue('https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/Regul/CPublicaRegulPerfilMensal.aspx?PK_PARTIC=');
  abaApoio.getRange('F6').setValue('/html/body/form/table[1]/tbody/tr[3]/td[2]/select');
  
  // ============================================
  // LINHA 16-18: DATAS CALCULADAS
  // ============================================
  abaApoio.getRange('A16').setValue('HOJE');
  abaApoio.getRange('B16').setValue('DATAD');
  abaApoio.getRange('C16').setValue('LINHAD');
  
  // Linha 17: D-2
  abaApoio.getRange('A17').setValue(formatarData(hoje));
  abaApoio.getRange('B17').setValue(formatarData(calcularDiaUtil(hoje, -1, ss)));
  abaApoio.getRange('C17').setValue(dataD2Uteis.getDate() + 1);
  abaApoio.getRange('D17').setValue(dataD2Uteis.getDate() + 1);
  abaApoio.getRange('E17').setValue(formatarData(new Date(dataD2Uteis.getTime() + 86400000)));
  
  // Linha 18: D-1
  abaApoio.getRange('A18').setValue(formatarData(dataD1Util));
  abaApoio.getRange('B18').setValue(formatarData(dataD1Util));
  abaApoio.getRange('B18').setValue(formatarData(calcularDiaUtil(hoje, -2, ss)));
  abaApoio.getRange('C18').setValue(dataD1Util.getDate() + 1);
  
  // ============================================
  // LINHA 20-21: XPATH
  // ============================================
  abaApoio.getRange('A20').setValue('HTMLDP1');
  abaApoio.getRange('B20').setValue('HTMLDP2');
  abaApoio.getRange('A21').setValue('/html/body/form/table[2]/tbody/tr[');
  abaApoio.getRange('B21').setValue(']/td[8]');
  
  SpreadsheetApp.flush();
  
  // ============================================
  // CRIAR NOMES PARA AS DATAS (para uso em fórmulas)
  // ============================================
  try {
    // Remover nomes existentes primeiro para evitar conflitos
    var nomesExistentes = ss.getNamedRanges();
    nomesExistentes.forEach(function(nr) {
      var nome = nr.getName();
      if (nome === 'DIAMESREF' || nome === 'DIAMESREF2' || nome === 'DIADDD') {
        nr.remove();
        Logger.log('  🗑️ Nome existente removido: ' + nome);
      }
    });
    
    // Criar os named ranges
    criarNamedRangesDatas(ss, abaApoio);
    Logger.log('  ✅ Named ranges criados com sucesso:');
    Logger.log('     - DIAMESREF: APOIO!D1');
    Logger.log('     - DIAMESREF2: APOIO!F1');
    Logger.log('     - DIADDD: APOIO!A17');
  } catch (e) {
    Logger.log('❌ Erro ao criar named ranges: ' + e.toString());
    throw new Error('Falha ao criar named ranges necessários para as fórmulas: ' + e.toString());
  }
  
  Logger.log('\n✅ Aba APOIO preenchida com VALORES calculados!');
  Logger.log('✅ Agora execute: verificarAbaApoio()');
}

// ===== NOVAS FUNÇÕES UTILITÁRIAS DE DATA PARA CÁLCULO DE STATUS =====

/**
 * Calcula o 10º dia útil do mês atual (opcionalmente, recebe array de feriados se quiser)
 */
function calcularDecimoDiaUtil(referencia, feriados) {
  var hoje = referencia ? new Date(referencia) : new Date();
  var ano = hoje.getFullYear();
  var mes = hoje.getMonth();
  var date = new Date(ano, mes, 1); // 1º dia do mês
  var uteis = 0;
  feriados = feriados || [];
  while (uteis < 10) {
    var diaSemana = date.getDay();
    var ehFeriado = feriados.some(function(f){
      return (
        f.getDate() === date.getDate() &&
        f.getMonth() === date.getMonth() &&
        f.getFullYear() === date.getFullYear()
      );
    });
    if (diaSemana !== 0 && diaSemana !== 6 && !ehFeriado) {
      uteis++;
    }
    if (uteis < 10) date.setDate(date.getDate() + 1);
  }
  return date;
}

/**
 * Compara duas datas "DD/MM/YYYY", retorna -1 (d1<d2), 0 (iguais) ou 1 (d1>d2)
 */
function compararDatasPTBR(d1, d2) {
  var d1a = d1.split('/').reverse().join('-');
  var d2a = d2.split('/').reverse().join('-');
  return d1a < d2a ? -1 : (d1a === d2a ? 0 : 1);
}

/**
 * Dias restantes entre datas "DD/MM/YYYY"
 */
function calcularDiasRestantesPTBR(data1, data2) {
  var d1 = new Date(data1.split('/').reverse().join('-'));
  var d2 = new Date(data2.split('/').reverse().join('-'));
  var delta = Math.ceil((d2 - d1) / (1000*60*60*24));
  return delta >= 0 ? delta : 0;
}
