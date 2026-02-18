/**
 * Configurações e Dados Estáticos
 * Feriados e outras constantes
 */

// ============================================
// FERIADOS BRASILEIROS
// ============================================

function getFeriadosBrasileiros() {
  return [
    // 2025
    [new Date(2025, 0, 1), 'quarta-feira', 'Confraternização Universal'],
    [new Date(2025, 2, 3), 'segunda-feira', 'Carnaval'],
    [new Date(2025, 2, 4), 'terça-feira', 'Carnaval'],
    [new Date(2025, 3, 18), 'sexta-feira', 'Paixão de Cristo'],
    [new Date(2025, 3, 21), 'segunda-feira', 'Tiradentes'],
    [new Date(2025, 4, 1), 'quinta-feira', 'Dia do Trabalho'],
    [new Date(2025, 5, 19), 'quinta-feira', 'Corpus Christi'],
    [new Date(2025, 8, 7), 'domingo', 'Independência do Brasil'],
    [new Date(2025, 9, 12), 'domingo', 'Nossa Sr.a Aparecida - Padroeira do Brasil'],
    [new Date(2025, 10, 2), 'domingo', 'Finados'],
    [new Date(2025, 10, 15), 'sábado', 'Proclamação da República'],
    [new Date(2025, 10, 20), 'quinta-feira', 'Dia Nacional de Zumbi e da Consciência Negra'],
    [new Date(2025, 11, 25), 'quinta-feira', 'Natal'],
    
    // 2026
    [new Date(2026, 0, 1), 'quinta-feira', 'Confraternização Universal'],
    [new Date(2026, 1, 16), 'segunda-feira', 'Carnaval'],
    [new Date(2026, 1, 17), 'terça-feira', 'Carnaval'],
    [new Date(2026, 3, 3), 'sexta-feira', 'Paixão de Cristo'],
    [new Date(2026, 3, 21), 'terça-feira', 'Tiradentes'],
    [new Date(2026, 4, 1), 'sexta-feira', 'Dia do Trabalho'],
    [new Date(2026, 5, 4), 'quinta-feira', 'Corpus Christi'],
    [new Date(2026, 8, 7), 'segunda-feira', 'Independência do Brasil'],
    [new Date(2026, 9, 12), 'segunda-feira', 'Nossa Sr.a Aparecida - Padroeira do Brasil'],
    [new Date(2026, 10, 2), 'segunda-feira', 'Finados'],
    [new Date(2026, 10, 15), 'domingo', 'Proclamação da República'],
    [new Date(2026, 10, 20), 'sexta-feira', 'Dia Nacional de Zumbi e da Consciência Negra'],
    [new Date(2026, 11, 25), 'sexta-feira', 'Natal'],
    
    // 2027
    [new Date(2027, 0, 1), 'sexta-feira', 'Confraternização Universal'],
    [new Date(2027, 1, 8), 'segunda-feira', 'Carnaval'],
    [new Date(2027, 1, 9), 'terça-feira', 'Carnaval'],
    [new Date(2027, 2, 26), 'sexta-feira', 'Paixão de Cristo'],
    [new Date(2027, 3, 21), 'quarta-feira', 'Tiradentes'],
    [new Date(2027, 4, 1), 'sábado', 'Dia do Trabalho'],
    [new Date(2027, 4, 27), 'quinta-feira', 'Corpus Christi'],
    [new Date(2027, 8, 7), 'terça-feira', 'Independência do Brasil'],
    [new Date(2027, 9, 12), 'terça-feira', 'Nossa Sr.a Aparecida - Padroeira do Brasil'],
    [new Date(2027, 10, 2), 'terça-feira', 'Finados'],
    [new Date(2027, 10, 15), 'segunda-feira', 'Proclamação da República'],
    [new Date(2027, 10, 20), 'sábado', 'Dia Nacional de Zumbi e da Consciência Negra'],
    [new Date(2027, 11, 25), 'sábado', 'Natal'],
    
    // 2028
    [new Date(2028, 0, 1), 'sábado', 'Confraternização Universal'],
    [new Date(2028, 1, 28), 'segunda-feira', 'Carnaval'],
    [new Date(2028, 1, 29), 'terça-feira', 'Carnaval'],
    [new Date(2028, 3, 14), 'sexta-feira', 'Paixão de Cristo'],
    [new Date(2028, 3, 21), 'sexta-feira', 'Tiradentes'],
    [new Date(2028, 4, 1), 'segunda-feira', 'Dia do Trabalho'],
    [new Date(2028, 5, 15), 'quinta-feira', 'Corpus Christi'],
    [new Date(2028, 8, 7), 'quinta-feira', 'Independência do Brasil'],
    [new Date(2028, 9, 12), 'quinta-feira', 'Nossa Sr.a Aparecida - Padroeira do Brasil'],
    [new Date(2028, 10, 2), 'quinta-feira', 'Finados'],
    [new Date(2028, 10, 15), 'quarta-feira', 'Proclamação da República'],
    [new Date(2028, 10, 20), 'segunda-feira', 'Dia Nacional de Zumbi e da Consciência Negra'],
    [new Date(2028, 11, 25), 'segunda-feira', 'Natal'],
    
    // 2029
    [new Date(2029, 0, 1), 'segunda-feira', 'Confraternização Universal'],
    [new Date(2029, 1, 12), 'segunda-feira', 'Carnaval'],
    [new Date(2029, 1, 13), 'terça-feira', 'Carnaval'],
    [new Date(2029, 2, 30), 'sexta-feira', 'Paixão de Cristo'],
    [new Date(2029, 3, 21), 'sábado', 'Tiradentes'],
    [new Date(2029, 4, 1), 'terça-feira', 'Dia do Trabalho'],
    [new Date(2029, 4, 31), 'quinta-feira', 'Corpus Christi'],
    [new Date(2029, 8, 7), 'sexta-feira', 'Independência do Brasil'],
    [new Date(2029, 9, 12), 'sexta-feira', 'Nossa Sr.a Aparecida - Padroeira do Brasil'],
    [new Date(2029, 10, 2), 'sexta-feira', 'Finados'],
    [new Date(2029, 10, 15), 'quinta-feira', 'Proclamação da República'],
    [new Date(2029, 10, 20), 'terça-feira', 'Dia Nacional de Zumbi e da Consciência Negra'],
    [new Date(2029, 11, 25), 'terça-feira', 'Natal'],
    
    // 2030
    [new Date(2030, 0, 1), 'terça-feira', 'Confraternização Universal'],
    [new Date(2030, 2, 4), 'segunda-feira', 'Carnaval'],
    [new Date(2030, 2, 5), 'terça-feira', 'Carnaval'],
    [new Date(2030, 3, 19), 'sexta-feira', 'Paixão de Cristo'],
    [new Date(2030, 3, 21), 'domingo', 'Tiradentes'],
    [new Date(2030, 4, 1), 'quarta-feira', 'Dia do Trabalho'],
    [new Date(2030, 5, 20), 'quinta-feira', 'Corpus Christi'],
    [new Date(2030, 8, 7), 'sábado', 'Independência do Brasil'],
    [new Date(2030, 9, 12), 'sábado', 'Nossa Sr.a Aparecida - Padroeira do Brasil'],
    [new Date(2030, 10, 2), 'sábado', 'Finados'],
    [new Date(2030, 10, 15), 'sexta-feira', 'Proclamação da República'],
    [new Date(2030, 10, 20), 'quarta-feira', 'Dia Nacional de Zumbi e da Consciência Negra'],
    [new Date(2030, 11, 25), 'quarta-feira', 'Natal']
  ];
}

// Alias para compatibilidade
function getFeriados() {
  return getFeriadosBrasileiros();
}