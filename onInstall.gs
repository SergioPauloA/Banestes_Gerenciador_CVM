/**
 * Instalação Automática do Sistema
 */

function criarTriggerAutomaticoHorario() {
  Logger.log('⚙️ Criando trigger de atualização automática...');
  
  // Remover triggers antigos do mesmo tipo (evitar duplicação)
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'atualizarStatusNaPlanilhaAutomatico') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('  🗑️ Trigger antigo removido');
    }
  });
  
  // Criar novo trigger (executa a cada 1 hora)
  ScriptApp.newTrigger('atualizarStatusNaPlanilhaAutomatico')
    .timeBased()
    .everyHours(1)
    .create();
  
  Logger.log('✅ Trigger criado com sucesso!');
  Logger.log('⏰ A função atualizarStatusNaPlanilhaAutomatico() será executada:');
  Logger.log('   - A cada 1 hora');
  Logger.log('   - Automaticamente, sem precisar fazer nada');
  Logger.log('   - Mesmo se você fechar o navegador');
  
  return {
    success: true,
    message: 'Trigger criado! Status serão atualizados automaticamente a cada 1 hora.'
  };
}

function reverterParaEstadoFuncional() {
  Logger.log('🔄 Revertendo para estado funcional anterior...');
  
  var ss = obterPlanilha();
  
  // ============================================
  // ABA BALANCETE
  // ============================================
  Logger.log('📊 Revertendo Balancete...');
  var abaBalancete = ss.getSheetByName('Balancete');
  
  // Status Geral E1 - Texto estático
  abaBalancete.getRange('E1').clearContent();
  abaBalancete.getRange('E1').setValue('AGUARDANDO DADOS');
  abaBalancete.getRange('E1').setBackground('#fef3c7').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
  
  // Limpar coluna D (STATUS) - deixar vazio para o Apps Script preencher
  abaBalancete.getRange('D4:D29').clearContent();
  abaBalancete.getRange('D4:D29').setValue('Aguardando...');
  
  Logger.log('  ✅ Balancete revertido');
  
  // ============================================
  // ABA COMPOSIÇÃO
  // ============================================
  Logger.log('📈 Revertendo Composição...');
  var abaComposicao = ss.getSheetByName('Composição');
  
  abaComposicao.getRange('E1').clearContent();
  abaComposicao.getRange('E1').setValue('AGUARDANDO DADOS');
  abaComposicao.getRange('E1').setBackground('#fef3c7').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
  
  abaComposicao.getRange('D4:D29').clearContent();
  abaComposicao.getRange('D4:D29').setValue('Aguardando...');
  
  Logger.log('  ✅ Composição revertida');
  
  // ============================================
  // ABA DIÁRIAS
  // ============================================
  Logger.log('📅 Revertendo Diárias...');
  var abaDiarias = ss.getSheetByName('Diárias');
  
  abaDiarias.getRange('E1').clearContent();
  abaDiarias.getRange('E1').setValue('AGUARDANDO DADOS');
  abaDiarias.getRange('E1').setBackground('#fef3c7').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
  
  abaDiarias.getRange('F1').clearContent();
  abaDiarias.getRange('F1').setValue('AGUARDANDO DADOS');
  abaDiarias.getRange('F1').setBackground('#fef3c7').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
  
  abaDiarias.getRange('D4:D29').clearContent();
  abaDiarias.getRange('D4:D29').setValue('Aguardando...');
  
  abaDiarias.getRange('F4:F29').clearContent();
  abaDiarias.getRange('F4:F29').setValue('Aguardando...');
  
  Logger.log('  ✅ Diárias revertida');
  
  // ============================================
  // ABA LÂMINA
  // ============================================
  Logger.log('📄 Revertendo Lâmina...');
  var abaLamina = ss.getSheetByName('Lâmina');
  
  abaLamina.getRange('E1').clearContent();
  abaLamina.getRange('E1').setValue('AGUARDANDO DADOS');
  abaLamina.getRange('E1').setBackground('#fef3c7').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
  
  abaLamina.getRange('D4:D29').clearContent();
  abaLamina.getRange('D4:D29').setValue('Aguardando...');
  
  Logger.log('  ✅ Lâmina revertida');
  
  // ============================================
  // ABA PERFIL MENSAL
  // ============================================
  Logger.log('📊 Revertendo Perfil Mensal...');
  var abaPerfilMensal = ss.getSheetByName('Perfil Mensal');
  
  abaPerfilMensal.getRange('E1').clearContent();
  abaPerfilMensal.getRange('E1').setValue('AGUARDANDO DADOS');
  abaPerfilMensal.getRange('E1').setBackground('#fef3c7').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
  
  abaPerfilMensal.getRange('D4:D29').clearContent();
  abaPerfilMensal.getRange('D4:D29').setValue('Aguardando...');
  
  Logger.log('  ✅ Perfil Mensal revertido');
  
  // ============================================
  // ABA GERAL
  // ============================================
  Logger.log('📋 Revertendo GERAL...');
  var abaGeral = ss.getSheetByName('GERAL');
  
  abaGeral.getRange('A3').setValue('Balancetes de Fundos');
  abaGeral.getRange('B3').setValue('Composição da Carteira');
  abaGeral.getRange('C3:D3').merge().setValue('Informações Diárias');
  abaGeral.getRange('E3').setValue('Lâmina do Fundo');
  abaGeral.getRange('F3').setValue('Perfil Mensal');
  
  abaGeral.getRange('A4').setValue('AGUARDANDO DADOS');
  abaGeral.getRange('B4').setValue('AGUARDANDO DADOS');
  abaGeral.getRange('C4').setValue('AGUARDANDO DADOS');
  abaGeral.getRange('D4').setValue('AGUARDANDO DADOS');
  abaGeral.getRange('E4').setValue('AGUARDANDO DADOS');
  abaGeral.getRange('F4').setValue('AGUARDANDO DADOS');
  
  abaGeral.getRange('A3:F4').setHorizontalAlignment('center').setVerticalAlignment('middle');
  
  Logger.log('  ✅ GERAL revertida');
  
  Logger.log('✅ REVERSÃO COMPLETA!');
  Logger.log('📊 Sistema voltou ao estado funcional');
  Logger.log('⏳ IMPORTXML vai continuar buscando dados da CVM');
  Logger.log('💡 Use atualizarStatusNaPlanilha() para calcular status depois que IMPORTXML carregar');
  
  return {
    success: true,
    message: 'Sistema revertido com sucesso! Estado funcional restaurado.'
  };
}

function corrigirFormulasSemErro() {
  Logger.log('🔧 Corrigindo com fórmulas universais...');
  
  var ss = SpreadsheetApp.openById('1N6LP1ydsxnQO_Woatv9zWEccb0fOGaV_3EKK1GoSWZI');
  
  // 1. BALANCETE - Status Geral (SEM FÓRMULA COMPLEXA)
  Logger.log('📊 Balancete...');
  var abaBalancete = ss.getSheetByName('Balancete');
  abaBalancete.getRange('E1').setValue('AGUARDANDO DADOS');
  abaBalancete.getRange('E1').setBackground('#fef3c7').setHorizontalAlignment('center').setFontWeight('bold');
  
  // 2. COMPOSIÇÃO
  Logger.log('📈 Composição...');
  var abaComposicao = ss.getSheetByName('Composição');
  abaComposicao.getRange('E1').setValue('AGUARDANDO DADOS');
  abaComposicao.getRange('E1').setBackground('#fef3c7').setHorizontalAlignment('center').setFontWeight('bold');
  
  // 3. DIÁRIAS
  Logger.log('📅 Diárias...');
  var abaDiarias = ss.getSheetByName('Diárias');
  abaDiarias.getRange('E1').setValue('AGUARDANDO DADOS');
  abaDiarias.getRange('F1').setValue('AGUARDANDO DADOS');
  abaDiarias.getRange('E1:F1').setBackground('#fef3c7').setHorizontalAlignment('center').setFontWeight('bold');
  
  // 4. LÂMINA
  Logger.log('📄 Lâmina...');
  var abaLamina = ss.getSheetByName('Lâmina');
  abaLamina.getRange('E1').setValue('AGUARDANDO DADOS');
  abaLamina.getRange('E1').setBackground('#fef3c7').setHorizontalAlignment('center').setFontWeight('bold');
  
  // 5. PERFIL MENSAL
  Logger.log('📊 Perfil Mensal...');
  var abaPerfilMensal = ss.getSheetByName('Perfil Mensal');
  abaPerfilMensal.getRange('E1').setValue('AGUARDANDO DADOS');
  abaPerfilMensal.getRange('E1').setBackground('#fef3c7').setHorizontalAlignment('center').setFontWeight('bold');
  
  // 6. GERAL - Referências diretas
  Logger.log('📋 GERAL...');
  var abaGeral = ss.getSheetByName('GERAL');
  
  // Em vez de fórmulas, vamos usar valores fixos por enquanto
  abaGeral.getRange('A3').setValue('Balancetes de Fundos');
  abaGeral.getRange('B3').setValue('Composição da Carteira');
  abaGeral.getRange('C3:D3').merge().setValue('Informações Diárias');
  abaGeral.getRange('E3').setValue('Lâmina do Fundo');
  abaGeral.getRange('F3').setValue('Perfil Mensal');
  
  abaGeral.getRange('A4').setValue('AGUARDANDO DADOS');
  abaGeral.getRange('B4').setValue('AGUARDANDO DADOS');
  abaGeral.getRange('C4').setValue('AGUARDANDO DADOS');
  abaGeral.getRange('D4').setValue('AGUARDANDO DADOS');
  abaGeral.getRange('E4').setValue('AGUARDANDO DADOS');
  abaGeral.getRange('F4').setValue('AGUARDANDO DADOS');
  
  abaGeral.getRange('A3:F4').setHorizontalAlignment('center').setVerticalAlignment('middle');
  
  // 7. APOIO
  Logger.log('⚙️ APOIO...');
  var abaApoio = ss.getSheetByName('APOIO');
  
  // A11 - Data de hoje (valor, não fórmula)
  abaApoio.getRange('A11').setValue(new Date());
  abaApoio.getRange('A11').setNumberFormat('dd/mm/yyyy');
  
  // F4 e G4 - XPath fixo
  abaApoio.getRange('F4').setValue('/html/body/form/table[2]/tbody/tr[2]/td[8]');
  abaApoio.getRange('G4').setValue('/html/body/form/table[2]/tbody/tr[3]/td[8]');
  
  Logger.log('✅ TODAS AS FÓRMULAS CORRIGIDAS (valores estáticos)!');
  Logger.log('⏳ Aguarde para o IMPORTXML carregar');
  
  return {
    success: true,
    message: 'Fórmulas corrigidas! Status será calculado pelo Apps Script.'
  };
}

function onInstall(e) {
  Logger.log('🚀 Instalação iniciada...');
  var resultado = setupCompletoAutomatico();
  Logger.log(JSON.stringify(resultado));
  return resultado;
}

function setupCompletoAutomatico() {
  try {
    Logger.log('📦 Etapa 1: Obtendo planilha...');
    var ss = obterPlanilha();
    
    Logger.log('📦 Etapa 2: Criando estrutura...');
    criarEstruturaPlanilhaCompleta(ss);
    
    Logger.log('📦 Etapa 3: Preenchendo COD FUNDO...');
    preencherAbaCodFundo(ss);
    
    Logger.log('📦 Etapa 4: Preenchendo FERIADOS...');
    preencherAbaFeriados(ss);
    
    Logger.log('📦 Etapa 5: Preenchendo APOIO...');
    //preencherAbaApoio(ss);
    criarAbaApoioComValores();
    
    Logger.log('📦 Etapa 6: Definindo nomes...');
    definirNomesApoio(ss);
    
    Logger.log('📦 Etapa 7: Criando fórmulas...');
    criarFormulasAbas(ss);
    
    Logger.log('📦 Etapa 8: Configurando GERAL...');
    configurarAbaGeral(ss);
    
    Logger.log('✅ INSTALAÇÃO CONCLUÍDA!');
    
    return {
      success: true,
      message: 'Sistema instalado! Aguarde alguns segundos para as fórmulas IMPORTXML carregarem.',
      url: obterURLPlanilha()
    };
    
  } catch (error) {
    Logger.log('❌ ERRO: ' + error.toString());
    return {
      success: false,
      message: error.toString()
    };
  }
}

// ============================================
// CRIAR ESTRUTURA
// ============================================

function criarEstruturaPlanilhaCompleta(ss) {
  var abas = ['GERAL', 'Balancete', 'Composição', 'Diárias', 'Lâmina', 'Perfil Mensal', 'APOIO', 'FERIADOS', 'COD FUNDO'];
  
  abas.forEach(function(nomeAba) {
    var aba = ss.getSheetByName(nomeAba);
    if (!aba) {
      ss.insertSheet(nomeAba);
      Logger.log('  ✅ Criada: ' + nomeAba);
    }
  });
}

// ============================================
// CRIAR FÓRMULAS
// ============================================

// Fórmula para status geral (coluna E)
var FORMULA_STATUS_GERAL = '=SE(CONT.SE(D:D;"OK")=CONT.SE(A4:A;"<>"&"");"OK";SE(DIADDD<=DIAMESREF2;"EM CONFORMIDADE " & CARACT(10) & DATADIF(DIADDD;DIAMESREF2;"D") & " DIAS RESTANTES";"DESCONFORMIDADE"))';

// Função helper para criar fórmula de status individual (coluna D)
function criarFormulaStatusIndividual(linha) {
  return '=SE(C' + linha + '=DIAMESREF;"OK";SE(DIADDD<=DIAMESREF2;"EM CONFORMIDADE";"DESATUALIZADO"))';
}

function criarFormulasAbas(ss) {
  criarFormulasBalancete(ss);
  criarFormulasComposicao(ss);
  criarFormulasDiarias(ss);
  criarFormulasLamina(ss);
  criarFormulasPerfilMensal(ss);
}

function criarFormulasBalancete(ss) {
  var aba = ss.getSheetByName('Balancete');
  aba.clear();
  
  // Título A1:D2
  aba.getRange('A1:D2').merge().setValue('Balancetes de Fundos')
    .setBackground('#667eea').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  
  // Status Geral E1:E2
  aba.getRange('E1:E2').merge();
  aba.getRange('E1').setValue('AGUARDANDO DADOS')
    .setBackground('#fef3c7').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
  
  // Cabeçalhos linha 3
  aba.getRange('A3:F3').setValues([['FUNDO', 'COD', 'COMPETÊNCIA 1', 'STATUS 1', 'COMPETÊNCIA 2', 'STATUS 2']])
    .setBackground('#667eea').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  
  var fundos = getFundos();
  for (var i = 0; i < fundos.length; i++) {
    var linha = i + 4;
    var linhaCodFundo = i + 2;
    
    aba.getRange(linha, 1).setFormula("='COD FUNDO'!A" + linhaCodFundo);
    aba.getRange(linha, 2).setFormula("='COD FUNDO'!C" + linhaCodFundo); // Código BANESTES
    aba.getRange(linha, 3).setFormula('=IFERROR(ÍNDICE(IMPORTXML(CONCATENAR(LINKB1;\'COD FUNDO\'!B' + linhaCodFundo + ';LINKB2);HTMLB);1);"-")');
    aba.getRange(linha, 4).setValue('Aguardando...');
    aba.getRange(linha, 5).setFormula('=IFERROR(ÍNDICE(IMPORTXML(CONCATENAR(LINKB1;\'COD FUNDO\'!B' + linhaCodFundo + ';LINKB2);HTMLB);2);"-")');
    aba.getRange(linha, 6).setValue('Aguardando...');
  }
  
  aba.setColumnWidth(1, 400);
  aba.setColumnWidth(2, 80);
  aba.setColumnWidth(3, 130);
  aba.setColumnWidth(4, 150);
  aba.setColumnWidth(5, 130);
  aba.setColumnWidth(6, 150);
  aba.setFrozenRows(3);

  // 🆕 ADICIONAR CÉLULA DE CONTROLE
  criarCelulaControleEmail(aba);
  
  Logger.log('  ✅ Balancete criado (6 colunas com competências)');
}

function criarFormulasComposicao(ss) {
  var aba = ss.getSheetByName('Composição');
  aba.clear();
  
  aba.getRange('A1:D2').merge().setValue('Composição da Carteira')
    .setBackground('#667eea').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  
  aba.getRange('E1:E2').merge();
  aba.getRange('E1').setValue('AGUARDANDO DADOS')
    .setBackground('#fef3c7').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
  
  aba.getRange('A3:F3').setValues([['FUNDO', 'COD', 'COMPETÊNCIA 1', 'STATUS 1', 'COMPETÊNCIA 2', 'STATUS 2']])
    .setBackground('#667eea').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  
  var fundos = getFundos();
  for (var i = 0; i < fundos.length; i++) {
    var linha = i + 4;
    var linhaCodFundo = i + 2;
    
    aba.getRange(linha, 1).setFormula("='COD FUNDO'!A" + linhaCodFundo);
    aba.getRange(linha, 2).setFormula("='COD FUNDO'!C" + linhaCodFundo);
    aba.getRange(linha, 3).setFormula('=IFERROR(ÍNDICE(IMPORTXML(CONCATENAR(LINKC1;\'COD FUNDO\'!B' + linhaCodFundo + ';LINKC2);HTMLC);1);"-")');
    aba.getRange(linha, 4).setValue('Aguardando...');
    aba.getRange(linha, 5).setFormula('=IFERROR(ÍNDICE(IMPORTXML(CONCATENAR(LINKC1;\'COD FUNDO\'!B' + linhaCodFundo + ';LINKC2);HTMLC);2);"-")');
    aba.getRange(linha, 6).setValue('Aguardando...');
  }
  
  aba.setColumnWidth(1, 400);
  aba.setColumnWidth(2, 80);
  aba.setColumnWidth(3, 130);
  aba.setColumnWidth(4, 150);
  aba.setColumnWidth(5, 130);
  aba.setColumnWidth(6, 150);
  aba.setFrozenRows(3);
  
  Logger.log('  ✅ Composição criada (6 colunas com competências)');
}

function criarFormulasDiarias(ss) {
  var aba = ss.getSheetByName('Diárias');
  aba.clear();
  
  aba.getRange('A1:D2').merge().setValue('Informações Diárias')
    .setBackground('#667eea').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  
  aba.getRange('E1:E2').merge();
  aba.getRange('E1').setFormula('=SE(CONT.SE(D:D;"OK")=CONT.SE(A4:A;"<>"&"");"OK";"DESCONFORMIDADE")')
    .setBackground('#fef3c7').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
  
  aba.getRange('F1:F2').merge();
  aba.getRange('F1').setFormula('=SE(CONT.SE(F:F;"OK")=CONT.SE(A4:A;"<>"&"");"OK";"A ATUALIZAR")')
    .setBackground('#fef3c7').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
  
  aba.getRange('A3:F3').setValues([['FUNDO', 'COD', 'RETORNO 1', 'STATUS 1', 'RETORNO 2', 'STATUS 2']])
    .setBackground('#667eea').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  
  var fundos = getFundos();
  for (var i = 0; i < fundos.length; i++) {
    var linha = i + 4;
    var linhaCodFundo = i + 2;
    
    aba.getRange(linha, 1).setFormula("='COD FUNDO'!A" + linhaCodFundo);
    aba.getRange(linha, 2).setFormula("='COD FUNDO'!B" + linhaCodFundo);
    aba.getRange(linha, 3).setFormula('=IFERROR(ÍNDICE(IMPORTXML(CONCATENAR(LINKD1;\'COD FUNDO\'!B' + linhaCodFundo + ';LINKD2);HTMLD1);1);"-")');
    aba.getRange(linha, 4).setValue('Aguardando...');
    aba.getRange(linha, 5).setFormula('=IFERROR(ÍNDICE(IMPORTXML(CONCATENAR(LINKD1;\'COD FUNDO\'!B' + linhaCodFundo + ';LINKD2);HTMLD2);1);"-")');
    aba.getRange(linha, 6).setValue('Aguardando...');
  }
  
  aba.setColumnWidth(1, 400);
  aba.setColumnWidth(2, 80);
  aba.setColumnWidth(3, 120);
  aba.setColumnWidth(4, 150);
  aba.setColumnWidth(5, 120);
  aba.setColumnWidth(6, 150);
  aba.setFrozenRows(3);
  
  Logger.log('  ✅ Diárias criada');
}

function criarFormulasLamina(ss) {
  var aba = ss.getSheetByName('Lâmina');
  aba.clear();
  
  aba.getRange('A1:D2').merge().setValue('Lâmina do Fundo')
    .setBackground('#667eea').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  
  aba.getRange('E1:E2').merge();
  aba.getRange('E1').setValue('AGUARDANDO DADOS')
    .setBackground('#fef3c7').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
  
  aba.getRange('A3:F3').setValues([['FUNDO', 'COD', 'COMPETÊNCIA 1', 'STATUS 1', 'COMPETÊNCIA 2', 'STATUS 2']])
    .setBackground('#667eea').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  
  var fundos = getFundos();
  for (var i = 0; i < fundos.length; i++) {
    var linha = i + 4;
    var linhaCodFundo = i + 2;
    
    aba.getRange(linha, 1).setFormula("='COD FUNDO'!A" + linhaCodFundo);
    aba.getRange(linha, 2).setFormula("='COD FUNDO'!C" + linhaCodFundo);
    aba.getRange(linha, 3).setFormula('=IFERROR(ÍNDICE(IMPORTXML(CONCATENAR(LINKL1;\'COD FUNDO\'!B' + linhaCodFundo + ';LINKL2);HTMLL);1);"-")');
    aba.getRange(linha, 4).setValue('Aguardando...');
    aba.getRange(linha, 5).setFormula('=IFERROR(ÍNDICE(IMPORTXML(CONCATENAR(LINKL1;\'COD FUNDO\'!B' + linhaCodFundo + ';LINKL2);HTMLL);2);"-")');
    aba.getRange(linha, 6).setValue('Aguardando...');
  }
  
  aba.setColumnWidth(1, 400);
  aba.setColumnWidth(2, 80);
  aba.setColumnWidth(3, 130);
  aba.setColumnWidth(4, 150);
  aba.setColumnWidth(5, 130);
  aba.setColumnWidth(6, 150);
  aba.setFrozenRows(3);
  
  Logger.log('  ✅ Lâmina criada (6 colunas com competências)');
}

function criarFormulasPerfilMensal(ss) {
  var aba = ss.getSheetByName('Perfil Mensal');
  aba.clear();
  
  aba.getRange('A1:D2').merge().setValue('Perfil Mensal')
    .setBackground('#667eea').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  
  aba.getRange('E1:E2').merge();
  aba.getRange('E1').setValue('AGUARDANDO DADOS')
    .setBackground('#fef3c7').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
  
  aba.getRange('A3:F3').setValues([['FUNDO', 'COD', 'COMPETÊNCIA 1', 'STATUS 1', 'COMPETÊNCIA 2', 'STATUS 2']])
    .setBackground('#667eea').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  
  var fundos = getFundos();
  for (var i = 0; i < fundos.length; i++) {
    var linha = i + 4;
    var linhaCodFundo = i + 2;
    
    aba.getRange(linha, 1).setFormula("='COD FUNDO'!A" + linhaCodFundo);
    aba.getRange(linha, 2).setFormula("='COD FUNDO'!C" + linhaCodFundo);
    aba.getRange(linha, 3).setFormula('=IFERROR(ÍNDICE(IMPORTXML(CONCATENAR(LINKP;\'COD FUNDO\'!B' + linhaCodFundo + ');HTMLP);1);"-")');
    aba.getRange(linha, 4).setValue('Aguardando...');
    aba.getRange(linha, 5).setFormula('=IFERROR(ÍNDICE(IMPORTXML(CONCATENAR(LINKP;\'COD FUNDO\'!B' + linhaCodFundo + ');HTMLP);2);"-")');
    aba.getRange(linha, 6).setValue('Aguardando...');
  }
  
  aba.setColumnWidth(1, 400);
  aba.setColumnWidth(2, 80);
  aba.setColumnWidth(3, 130);
  aba.setColumnWidth(4, 150);
  aba.setColumnWidth(5, 130);
  aba.setColumnWidth(6, 150);
  aba.setFrozenRows(3);
  
  Logger.log('  ✅ Perfil Mensal criado (6 colunas com competências)');
}

// ============================================
// CONFIGURAR ABA GERAL
// ============================================

function configurarAbaGeral(ss) {
  var aba = ss.getSheetByName('GERAL');
  aba.clear();
  
  aba.getRange('A1:F1').merge().setValue('📊 DASHBOARD GERAL')
    .setBackground('#667eea').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  aba.getRange('A3').setFormula('=Balancete!A1');
  aba.getRange('B3').setFormula('=Composição!A1');
  aba.getRange('C3:D3').merge().setFormula('=Diárias!A1');
  aba.getRange('E3').setFormula('=Lâmina!A1');
  aba.getRange('F3').setFormula('=\'Perfil Mensal\'!A1');
  
  aba.getRange('A4').setFormula('=Balancete!E1');
  aba.getRange('B4').setFormula('=Composição!E1');
  aba.getRange('C4').setFormula('=Diárias!E1');
  aba.getRange('D4').setFormula('=Diárias!F1');
  aba.getRange('E4').setFormula('=Lâmina!E1');
  aba.getRange('F4').setFormula('=\'Perfil Mensal\'!E1');
  
  Logger.log('  ✅ GERAL configurada');
}

// ============================================
// PREENCHER APOIO
// ============================================

function preencherAbaApoio(ss) {
  var aba = ss.getSheetByName('APOIO');
  aba.clear();
  
  aba.getRange('A1:G1').merge().setValue('⚙️ APOIO')
    .setBackground('#667eea').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  
  aba.getRange('D1').setFormula('=DATA(ANO(HOJE());MÊS(HOJE())-1;1)');
  aba.getRange('E1').setFormula('=DATA(ANO(HOJE());MÊS(HOJE());1)');
  aba.getRange('F1').setFormula('=E1+10');
  
  var urls = [
    ['BALANCETE', 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/Balancete/CPublicaBalancete.asp?PK_PARTIC=', '&SemFrame=', '/html/body/form/table/tbody/tr[1]/td/select'],
    ['COMPOSIÇÃO', 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/CDA/CPublicaCDA.aspx?PK_PARTIC=', '&SemFrame=', '/html/body/form/table/tbody/tr[1]/td/select'],
    ['DIÁRIAS', 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/InfDiario/CPublicaInfdiario.aspx?PK_PARTIC=', '&PK_SUBCLASSE=-1', ''],
    ['LÂMINA', 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/CPublicaLamina.aspx?PK_PARTIC=', '&PK_SUBCLASSE=-1', '/html/body/form/table[1]/tbody/tr[1]/td/select'],
    ['PERFIL MENSAL', 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/Regul/CPublicaRegulPerfilMensal.aspx?PK_PARTIC=', '', '/html/body/form/table[1]/tbody/tr[3]/td[2]/select']
  ];
  
  aba.getRange(2, 3, urls.length, 4).setValues(urls);
  
  aba.getRange('F4').setFormula('=CONCATENAR("/html/body/form/table[2]/tbody/tr[";2;"]";"/td[8]")');
  aba.getRange('G4').setFormula('=CONCATENAR("/html/body/form/table[2]/tbody/tr[";3;"]";"/td[8]")');
  
  aba.getRange('A11').setFormula('=HOJE()');
  
  Logger.log('  ✅ APOIO preenchida');
}

// ============================================
// DEFINIR NOMES
// ============================================

function definirNomesApoio(ss) {
  var aba = ss.getSheetByName('APOIO');
  
  var nomes = {
    'LINKB1': 'D2', 'LINKB2': 'E2', 'HTMLB': 'F2',
    'LINKC1': 'D3', 'LINKC2': 'E3', 'HTMLC': 'F3',
    'LINKD1': 'D4', 'LINKD2': 'E4', 'HTMLD1': 'F4', 'HTMLD2': 'G4',
    'LINKL1': 'D5', 'LINKL2': 'E5', 'HTMLL': 'F5',
    'LINKP': 'D6', 'HTMLP': 'F6'
  };
  
  for (var nome in nomes) {
    var range = aba.getRange(nomes[nome]);
    try {
      ss.setNamedRange(nome, range);
    } catch (e) {
      var existing = ss.getNamedRanges().filter(function(nr) { return nr.getName() === nome; })[0];
      if (existing) existing.remove();
      ss.setNamedRange(nome, range);
    }
  }
  
  Logger.log('  ✅ Nomes definidos');
}

// ============================================
// PREENCHER COD FUNDO
// ============================================

function preencherAbaCodFundo(ss) {
  var aba = ss.getSheetByName('COD FUNDO');
  aba.clear();
  
  // Cabeçalho COM 3 COLUNAS
  aba.getRange('A1:C1').setValues([['NOME FUNDO', 'Cod Fundo CVM', 'Cod Fundo']])
    .setBackground('#667eea').setFontColor('#FFFFFF').setFontWeight('bold');
  
  // Dados dos fundos COM 3 COLUNAS
  var fundos = getFundos();
  var dados = fundos.map(function(f) {
    return [f.nome, f.codigoCVM, f.codigoFundo]; // 3 COLUNAS: Nome | CVM | BANESTES
  });
  
  aba.getRange(2, 1, dados.length, 3).setValues(dados);
  
  // Formatação
  aba.setColumnWidth(1, 500); // Nome largo
  aba.setColumnWidth(2, 100); // CVM
  aba.setColumnWidth(3, 80);  // BANESTES
  aba.setFrozenRows(1);
  
  Logger.log('  ✅ COD FUNDO preenchida: ' + fundos.length + ' fundos com 3 colunas');
}

// ============================================
// PREENCHER FERIADOS
// ============================================

function preencherAbaFeriados(ss) {
  var aba = ss.getSheetByName('FERIADOS');
  aba.clear();
  
  aba.getRange('A1:C1').setValues([['Data', 'Dia da Semana', 'Feriado']])
    .setBackground('#667eea').setFontColor('#FFFFFF').setFontWeight('bold');
  
  var feriados = getFeriadosBrasileiros();
  
  if (feriados.length > 0) {
    aba.getRange(2, 1, feriados.length, 3).setValues(feriados);
    aba.getRange(2, 1, feriados.length, 1).setNumberFormat('dd/mm/yyyy');
  }
  
  aba.setColumnWidth(1, 120);
  aba.setColumnWidth(2, 150);
  aba.setColumnWidth(3, 300);
  aba.setFrozenRows(1);
  
  Logger.log('  ✅ FERIADOS preenchida');
}

function criarAbaDiarias(ss) {
  var aba = ss.getSheetByName('Diárias');
  if (!aba) {
    aba = ss.insertSheet('Diárias');
  }
  aba.clear();
  
  // Configurar cabeçalho linha 1
  aba.getRange('A1').setValue('Informações Diárias').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  aba.getRange('A1:B1').merge();
  
  // STATUS GERAL linha 1 (colunas E e F)
  aba.getRange('E1').setValue('DESCONFORMIDADE').setBackground('#f87171').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  aba.getRange('F1').setValue('A ATUALIZAR').setBackground('#fbbf24').setFontColor('#000000').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  
  // Cabeçalho linha 3 - Colunas
  aba.getRange('A3:F3').setValues([['FUNDO', 'COD', 'RETORNO 1', 'STATUS 1', 'RETORNO 2', 'STATUS 2']])
    .setBackground('#667eea').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  
  // Preencher lista de fundos a partir da aba COD FUNDO
  var abaCodFundo = ss.getSheetByName('COD FUNDO');
  if (!abaCodFundo) {
    Logger.log('  ⚠️ Aba COD FUNDO não encontrada. Criando aba vazia.');
    return;
  }
  
  var ultimaLinha = abaCodFundo.getLastRow();
  if (ultimaLinha < 2) {
    Logger.log('  ⚠️ Aba COD FUNDO está vazia.');
    return;
  }
  
  // Copiar dados: Coluna A (nome) e Coluna C (código BANESTES)
  var dadosFundos = abaCodFundo.getRange(2, 1, ultimaLinha - 1, 3).getValues().map(function(linha) {
    return [linha[0], linha[2]]; // A=Nome, C=Código BANESTES
  });
  
  aba.getRange(4, 1, dadosFundos.length, 2).setValues(dadosFundos);
  
  // Formatação de largura
  aba.setColumnWidth(1, 500); // Nome
  aba.setColumnWidth(2, 60);  // Código
  aba.setColumnWidth(3, 120); // RETORNO 1
  aba.setColumnWidth(4, 150); // STATUS 1
  aba.setColumnWidth(5, 120); // RETORNO 2
  aba.setColumnWidth(6, 150); // STATUS 2
  
  // Congelar linhas
  aba.setFrozenRows(3);
  
  // Centralizar colunas de dados
  aba.getRange('B4:F' + (3 + dadosFundos.length)).setHorizontalAlignment('center').setVerticalAlignment('middle');
  
  // Bordas
  var rangeComDados = aba.getRange(3, 1, dadosFundos.length + 1, 6);
  rangeComDados.setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  
  Logger.log('  ✅ Aba Diárias criada com ' + dadosFundos.length + ' fundos (6 colunas)');
}

function recriarAbaDiarias() {
  Logger.log('🔄 Recriando aba Diárias...');
  
  var ss = obterPlanilha(); // CORRIGIDO: obter planilha aqui
  
  // Deletar aba antiga
  var abaAntiga = ss.getSheetByName('Diárias');
  if (abaAntiga) {
    ss.deleteSheet(abaAntiga);
    Logger.log('  🗑️ Aba antiga deletada');
  }
  
  // Criar nova aba
  criarAbaDiarias(ss); // Passar ss como parâmetro
  
  Logger.log('✅ Aba Diárias recriada com sucesso!');
  Logger.log('📊 Layout: A=Nome, B=Código, C=Retorno1, D=Status1, E=Retorno2, F=Status2');
  
  return {
    success: true,
    message: 'Aba Diárias recriada!'
  };
}

/**
 * 🔧 Cria a célula G1 de controle em todas as abas mensais
 * Adicionar ao final de criarFormulasBalancete(), criarFormulasComposicao(), etc.
 */
function criarCelulaControleEmail(aba) {
  // Criar célula G1 com valor inicial "-"
  aba.getRange('G1').setValue('-');
  
  // Formatar
  aba.getRange('G1')
    .setBackground('#f3f4f6')
    .setFontColor('#6b7280')
    .setFontWeight('normal')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // Adicionar borda
  aba.getRange('G1').setBorder(
    true, true, true, true, 
    false, false, 
    '#9ca3af', 
    SpreadsheetApp.BorderStyle.SOLID
  );
}

function criarCelulasG1EmTodasAsAbas() {
  var ss = obterPlanilha();
  var abas = ['Balancete', 'Composição', 'Lâmina', 'Perfil Mensal'];
  
  abas.forEach(function(nomeAba) {
    var aba = ss.getSheetByName(nomeAba);
    if (aba) {
      criarCelulaControleEmail(aba);
      Logger.log('✅ ' + nomeAba + ': Célula G1 criada');
    }
  });
  
  Logger.log('✅ Todas as células G1 criadas!');
}