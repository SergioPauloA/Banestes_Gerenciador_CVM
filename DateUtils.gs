/**
 * DateUtils.gs - Funções de cálculo de datas
 */

function getDatasReferencia() {
  try {
    var ss = SpreadsheetApp.openById('1N6LP1ydsxnQO_Woatv9zWEccb0fOGaV_3EKK1GoSWZI');
    var abaApoio = ss.getSheetByName('APOIO');
    
    if (!abaApoio) {
      Logger.log('⚠️ Aba APOIO não encontrada. Calculando manualmente.');
      return calcularDatasManualmente();
    }
    
    // Tentar ler da linha 17 (onde estão os cálculos)
    var diaAtual = abaApoio.getRange('A17').getValue(); // HOJE
    var diaD1 = abaApoio.getRange('B17').getDisplayValue(); // D-1 útil (OK agora!)
    var diaD2 = abaApoio.getRange('B18').getDisplayValue(); // D-2 útil
    
    // Se algum valor está vazio, calcular manualmente
    if (!diaD1 || !diaD2 || diaD1 === '' || diaD2 === '') {
      Logger.log('⚠️ Aba APOIO com dados incompletos. Calculando manualmente.');
      return calcularDatasManualmente();
    }
    
    // 1º dia do mês anterior
    var mesAnterior = new Date(diaAtual);
    mesAnterior.setMonth(mesAnterior.getMonth() - 1);
    mesAnterior.setDate(1);
    var diaMesRef = formatarData(mesAnterior);
    
    // Calcular 10º dia útil do mês atual
    var mesAtual = new Date(diaAtual.getFullYear(), diaAtual.getMonth(), 1);
    var decimoDiaUtil = calcularDiaUtil(mesAtual, 10, ss);
    var diaMesRef2 = formatarData(decimoDiaUtil);
    
    // Calcular dias restantes até o prazo (10º dia útil)
    var hoje = new Date(diaAtual);
    var diasRestantes = calcularDiasUteisEntre(hoje, decimoDiaUtil, ss);
    
    Logger.log('📅 Datas obtidas da aba APOIO:');
    Logger.log('  Hoje: ' + formatarData(diaAtual));
    Logger.log('  D-2 (DIADREF1): ' + diaD1);
    Logger.log('  D-1 (DIADREF2): ' + diaD2);
    Logger.log('  1º mês anterior: ' + diaMesRef);
    Logger.log('  10º dia útil (prazo): ' + diaMesRef2);
    Logger.log('  Dias restantes: ' + diasRestantes);
    
    return {
      hoje: formatarData(diaAtual),
      diaMesRef: diaMesRef,
      diaMesRef2: diaMesRef2,
      diaDD: formatarData(diaAtual),
      diaD1: diaD1, // D-2 (dias úteis)
      diaD2: diaD2, // D-1 (dias úteis)
      diasRestantes: diasRestantes
    };
    
  } catch (error) {
    Logger.log('❌ Erro ao ler aba APOIO: ' + error.toString());
    Logger.log('⚠️ Usando datas calculadas manualmente.');
    return calcularDatasManualmente();
  }
}

function calcularDatasManualmente() {
  var ss = SpreadsheetApp.openById('1N6LP1ydsxnQO_Woatv9zWEccb0fOGaV_3EKK1GoSWZI');
  var hoje = new Date();
  
  // DIADDD (Hoje)
  var diaDD = formatarData(hoje);
  
  // IMPORTANTE: Se hoje é fim de semana, recuar para a última sexta-feira
  var diaParaCalculo = new Date(hoje);
  while (diaParaCalculo.getDay() === 0 || diaParaCalculo.getDay() === 6) {
    diaParaCalculo.setDate(diaParaCalculo.getDate() - 1);
  }
  
  // DIADREF1 (D-2 em dias ÚTEIS)
  var dataD1 = calcularDiaUtil(diaParaCalculo, -2, ss);
  var diaD1 = formatarData(dataD1);
  
  // DIADREF2 (D-1 em dias ÚTEIS)
  var dataD2 = calcularDiaUtil(diaParaCalculo, -1, ss);
  var diaD2 = formatarData(dataD2);
  
  // DIAMESREF (1º dia do mês anterior)
  var mesAnterior = new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);
  var diaMesRef = formatarData(mesAnterior);
  
  // DIAMESREF2 (10º dia útil do mês atual)
  var mesAtual = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  var decimoDiaUtil = calcularDiaUtil(mesAtual, 10, ss);
  var diaMesRef2 = formatarData(decimoDiaUtil);
  
  // Calcular dias restantes até o prazo (10º dia útil)
  var diasRestantes = calcularDiasUteisEntre(hoje, decimoDiaUtil, ss);
  
  Logger.log('📅 Datas calculadas manualmente:');
  Logger.log('  Hoje: ' + diaDD);
  Logger.log('  D-2 (DIADREF1 - úteis): ' + diaD1);
  Logger.log('  D-1 (DIADREF2 - úteis): ' + diaD2);
  Logger.log('  1º mês anterior: ' + diaMesRef);
  Logger.log('  10º dia útil (prazo): ' + diaMesRef2);
  Logger.log('  Dias restantes: ' + diasRestantes);
  
  return {
    hoje: diaDD,
    diaMesRef: diaMesRef,
    diaMesRef2: diaMesRef2,
    diaDD: diaDD,
    diaD1: diaD1,
    diaD2: diaD2,
    diasRestantes: diasRestantes
  };
}

function calcularDiaUtil(dataInicial, diasUteis, ss) {
  var resultado = new Date(dataInicial);
  var diasAdicionados = 0;
  var direcao = diasUteis > 0 ? 1 : -1;
  var diasRestantes = Math.abs(diasUteis);
  
  while (diasAdicionados < diasRestantes) {
    resultado.setDate(resultado.getDate() + direcao);
    
    var diaSemana = resultado.getDay();
    // Se não é sábado (6) nem domingo (0)
    if (diaSemana !== 0 && diaSemana !== 6) {
      // Se não é feriado
      if (!ehFeriado(resultado, ss)) {
        diasAdicionados++;
      }
    }
  }
  
  return resultado;
}

function calcularDiasUteisEntre(dataInicio, dataFim, ss) {
  var diasUteis = 0;
  var dataAtual = new Date(dataInicio);
  
  // Normalizar datas para meia-noite para comparação correta
  dataAtual.setHours(0, 0, 0, 0);
  var dataFimNormalizada = new Date(dataFim);
  dataFimNormalizada.setHours(0, 0, 0, 0);
  
  // Se a data de fim já passou, calcular dias negativos (prazo expirado)
  if (dataFimNormalizada < dataAtual) {
    // Inverter e retornar negativo
    var temp = new Date(dataAtual);
    while (dataFimNormalizada < temp) {
      temp.setDate(temp.getDate() - 1);
      var diaSemana = temp.getDay();
      if (diaSemana !== 0 && diaSemana !== 6) {
        if (!ehFeriado(temp, ss)) {
          diasUteis--;
        }
      }
    }
    return diasUteis;
  }
  
  // Contar dias úteis até a data fim (NÃO incluindo o dia de hoje, MAS incluindo o prazo)
  dataAtual.setDate(dataAtual.getDate() + 1); // Começar do dia seguinte
  while (dataAtual <= dataFimNormalizada) {
    var diaSemana = dataAtual.getDay();
    // Se não é sábado (6) nem domingo (0)
    if (diaSemana !== 0 && diaSemana !== 6) {
      // Se não é feriado
      if (!ehFeriado(dataAtual, ss)) {
        diasUteis++;
      }
    }
    dataAtual.setDate(dataAtual.getDate() + 1);
  }
  
  return diasUteis;
}

function ehFeriado(data, ss) {
  try {
    if (!ss) {
      ss = SpreadsheetApp.openById('1N6LP1ydsxnQO_Woatv9zWEccb0fOGaV_3EKK1GoSWZI');
    }
    
    var abaFeriados = ss.getSheetByName('FERIADOS');
    
    if (!abaFeriados) return false;
    
    var feriados = abaFeriados.getRange('A2:A100').getValues();
    var dataFormatada = formatarData(data);
    
    for (var i = 0; i < feriados.length; i++) {
      if (feriados[i][0]) {
        var feriadoFormatado = formatarData(new Date(feriados[i][0]));
        if (feriadoFormatado === dataFormatada) {
          return true;
        }
      }
    }
    
    return false;
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

// Função de teste
function testarGetDatasReferencia() {
  Logger.log('🧪 Testando getDatasReferencia()...\n');
  
  var datas = getDatasReferencia();
  
  Logger.log('📅 RESULTADO:');
  Logger.log('  Hoje: ' + datas.hoje);
  Logger.log('  1º mês anterior: ' + datas.diaMesRef);
  Logger.log('  Prazo (10º dia útil): ' + datas.diaMesRef2);
  Logger.log('  D-2 (DIADREF1 - úteis): ' + datas.diaD1);
  Logger.log('  D-1 (DIADREF2 - úteis): ' + datas.diaD2);
  Logger.log('  Dias restantes: ' + datas.diasRestantes);
  
  Logger.log('\n✅ Teste concluído!');
  
  return datas;
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
  
  Logger.log('\n✅ Aba APOIO preenchida com VALORES calculados!');
  Logger.log('✅ Agora execute: verificarAbaApoio()');
}
