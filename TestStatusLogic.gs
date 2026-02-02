/**
 * TestStatusLogic.gs - Testes manuais para a lógica de status
 * 
 * Execute estas funções no editor do Google Apps Script para testar:
 * 1. testarCalculoDiasRestantes() - Testa o cálculo de dias restantes
 * 2. testarStatusIndividual() - Testa o status individual
 * 3. testarStatusGeral() - Testa o status geral
 */

function testarCalculoDiasRestantes() {
  Logger.log('🧪 TESTE: Cálculo de Dias Restantes\n');
  Logger.log('='.repeat(50));
  
  var datas = getDatasReferencia();
  
  Logger.log('\n📅 DATAS OBTIDAS:');
  Logger.log('  Hoje: ' + datas.hoje);
  Logger.log('  1º do mês anterior (DIAMESREF): ' + datas.diaMesRef);
  Logger.log('  10º dia útil do mês (DIAMESREF2): ' + datas.diaMesRef2);
  Logger.log('  D-2 úteis (DIADREF1): ' + datas.diaD1);
  Logger.log('  D-1 útil (DIADREF2): ' + datas.diaD2);
  Logger.log('  Dias restantes: ' + datas.diasRestantes);
  
  Logger.log('\n✅ VALIDAÇÕES:');
  
  if (datas.diasRestantes >= 0) {
    Logger.log('  ✅ Ainda dentro do prazo');
    Logger.log('  ✅ Status esperado: "EM CONFORMIDADE ' + datas.diasRestantes + ' DIAS RESTANTES"');
  } else {
    Logger.log('  ❌ Prazo expirado');
    Logger.log('  ❌ Status esperado: "DESCONFORMIDADE"');
  }
  
  Logger.log('\n' + '='.repeat(50));
  Logger.log('✅ Teste concluído!\n');
  
  return datas;
}

function testarStatusIndividual() {
  Logger.log('🧪 TESTE: Status Individual\n');
  Logger.log('='.repeat(50));
  
  var datas = getDatasReferencia();
  
  // Teste 1: Data igual ao mês de referência
  var retorno1 = datas.diaMesRef;
  var status1 = calcularStatusIndividual(retorno1, 'mensal');
  Logger.log('\n📋 TESTE 1: Data = DIAMESREF');
  Logger.log('  Retorno: ' + retorno1);
  Logger.log('  Status: ' + status1);
  Logger.log('  Esperado: OK');
  Logger.log('  Resultado: ' + (status1 === 'OK' ? '✅ PASSOU' : '❌ FALHOU'));
  
  // Teste 2: Data diferente, mas dentro do prazo
  var retorno2 = '15/01/2026'; // Data aleatória
  var status2 = calcularStatusIndividual(retorno2, 'mensal');
  Logger.log('\n📋 TESTE 2: Data diferente, dentro do prazo');
  Logger.log('  Retorno: ' + retorno2);
  Logger.log('  Status: ' + status2);
  Logger.log('  Esperado: ' + (datas.diasRestantes >= 0 ? 'EM CONFORMIDADE' : 'DESATUALIZADO'));
  Logger.log('  Resultado: ' + (status2 === (datas.diasRestantes >= 0 ? 'EM CONFORMIDADE' : 'DESATUALIZADO') ? '✅ PASSOU' : '❌ FALHOU'));
  
  // Teste 3: Valor vazio
  var retorno3 = '';
  var status3 = calcularStatusIndividual(retorno3, 'mensal');
  Logger.log('\n📋 TESTE 3: Valor vazio');
  Logger.log('  Retorno: (vazio)');
  Logger.log('  Status: ' + status3);
  Logger.log('  Esperado: DESATUALIZADO');
  Logger.log('  Resultado: ' + (status3 === 'DESATUALIZADO' ? '✅ PASSOU' : '❌ FALHOU'));
  
  // Teste 4: Valor de erro
  var retorno4 = '#N/A';
  var status4 = calcularStatusIndividual(retorno4, 'mensal');
  Logger.log('\n📋 TESTE 4: Erro #N/A');
  Logger.log('  Retorno: #N/A');
  Logger.log('  Status: ' + status4);
  Logger.log('  Esperado: DESATUALIZADO');
  Logger.log('  Resultado: ' + (status4 === 'DESATUALIZADO' ? '✅ PASSOU' : '❌ FALHOU'));
  
  Logger.log('\n' + '='.repeat(50));
  Logger.log('✅ Teste concluído!\n');
}

function testarStatusGeral() {
  Logger.log('🧪 TESTE: Status Geral\n');
  Logger.log('='.repeat(50));
  
  var datas = getDatasReferencia();
  
  // Teste 1: Todos OK
  var dados1 = [
    [datas.diaMesRef],
    [datas.diaMesRef],
    [datas.diaMesRef]
  ];
  var status1 = calcularStatusGeralDaAba(dados1, 'mensal');
  Logger.log('\n📋 TESTE 1: Todos com data OK');
  Logger.log('  Status: ' + status1);
  Logger.log('  Esperado: OK');
  Logger.log('  Resultado: ' + (status1 === 'OK' ? '✅ PASSOU' : '❌ FALHOU'));
  
  // Teste 2: Alguns OK, alguns não, mas dentro do prazo
  var dados2 = [
    [datas.diaMesRef],
    ['15/01/2026'],
    ['16/01/2026']
  ];
  var status2 = calcularStatusGeralDaAba(dados2, 'mensal');
  Logger.log('\n📋 TESTE 2: Alguns OK, dentro do prazo');
  Logger.log('  Status: ' + status2);
  var esperado2 = datas.diasRestantes >= 0 
    ? 'EM CONFORMIDADE\n' + datas.diasRestantes + ' DIAS RESTANTES' 
    : 'DESCONFORMIDADE';
  Logger.log('  Esperado: ' + esperado2);
  Logger.log('  Resultado: ' + (status2 === esperado2 ? '✅ PASSOU' : '❌ FALHOU'));
  
  // Teste 3: Todos aguardando
  var dados3 = [
    [''],
    ['Loading...'],
    ['-']
  ];
  var status3 = calcularStatusGeralDaAba(dados3, 'mensal');
  Logger.log('\n📋 TESTE 3: Todos aguardando dados');
  Logger.log('  Status: ' + status3);
  Logger.log('  Esperado: AGUARDANDO DADOS');
  Logger.log('  Resultado: ' + (status3 === 'AGUARDANDO DADOS' ? '✅ PASSOU' : '❌ FALHOU'));
  
  Logger.log('\n' + '='.repeat(50));
  Logger.log('✅ Teste concluído!\n');
}

function executarTodosTestes() {
  Logger.log('🧪🧪🧪 EXECUTANDO TODOS OS TESTES 🧪🧪🧪\n\n');
  
  try {
    testarCalculoDiasRestantes();
    Logger.log('\n\n');
    testarStatusIndividual();
    Logger.log('\n\n');
    testarStatusGeral();
    
    Logger.log('\n\n' + '='.repeat(50));
    Logger.log('✅✅✅ TODOS OS TESTES CONCLUÍDOS ✅✅✅');
    Logger.log('='.repeat(50) + '\n');
    
  } catch (error) {
    Logger.log('\n\n❌❌❌ ERRO DURANTE OS TESTES ❌❌❌');
    Logger.log('Erro: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
  }
}
