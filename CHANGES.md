# Fix: Status Logic for Tabs (Balancete, Composição, Lâmina, Perfil Mensal)

## Problema

Os status das abas não estavam sendo atualizados corretamente para mostrar "EM CONFORMIDADE X DIAS RESTANTES".

## Causa Raiz

1. **`diasRestantes` sempre era 0**: As funções `getDatasReferencia()` e `calcularDatasManualmente()` retornavam sempre `diasRestantes: 0`, nunca calculando o valor real.

2. **`diaMesRef2` incorreto**: Era definido como o dia 15 do mês, mas deveria ser o 10º dia útil do mês.

## Solução Implementada

### 1. Adicionada função `calcularDiasUteisEntre()`

Nova função que calcula corretamente os dias úteis entre duas datas, considerando:
- Fins de semana (sábados e domingos)
- Feriados (da aba FERIADOS)
- Retorna valor negativo quando o prazo já expirou

```javascript
function calcularDiasUteisEntre(dataInicio, dataFim, ss) {
  // Calcula dias úteis entre as datas
  // Retorna positivo se prazo no futuro
  // Retorna 0 se hoje é o prazo
  // Retorna negativo se prazo já passou
}
```

### 2. Corrigida função `getDatasReferencia()`

Agora calcula corretamente:
- `diaMesRef2`: 10º dia útil do mês atual (não mais o dia 15)
- `diasRestantes`: dias úteis restantes até o prazo

### 3. Corrigida função `calcularDatasManualmente()`

Mesmas correções da função acima para o caso de fallback.

### 4. Atualizada lógica de status individual

```javascript
function calcularStatusIndividual(retorno, tipo) {
  if (tipo === 'mensal') {
    if (retorno === diaMesRef) return 'OK';
    if (diasRestantes >= 0) return 'EM CONFORMIDADE';
    return 'DESATUALIZADO';
  }
}
```

### 5. Atualizada lógica de status geral

```javascript
function calcularStatusGeralDaAba(dados, tipo) {
  if (todosOK) return 'OK';
  if (diasRestantes >= 0) {
    return 'EM CONFORMIDADE\n' + diasRestantes + ' DIAS RESTANTES';
  }
  return 'DESCONFORMIDADE';
}
```

## Fórmulas Implementadas

### Status Individual (Coluna D)
```
=SE(C4=DIAMESREF;"OK";SE(DIADDD<=DIAMESREF2;"EM CONFORMIDADE";"DESATUALIZADO"))
```

**Tradução:**
- Se data = 1º do mês anterior → "OK"
- Senão, se hoje <= 10º dia útil → "EM CONFORMIDADE"
- Senão → "DESATUALIZADO"

### Status Geral (Coluna E)
```
=SE(CONT.SE(D:D;"OK")=CONT.SE(A4:A;"<>"&VAZIO);"OK";SE(DIADDD<=DIAMESREF2;"EM CONFORMIDADE " & CARACT(10) & DATADIF(DIADDD;DIAMESREF2;"D") & " DIAS RESTANTES";"DESCONFORMIDADE"))
```

**Tradução:**
- Se todos na coluna D são "OK" → "OK"
- Senão, se hoje <= 10º dia útil → "EM CONFORMIDADE\n[X] DIAS RESTANTES"
- Senão → "DESCONFORMIDADE"

## Como Testar

### 1. No Editor do Apps Script

Execute as funções de teste no arquivo `TestStatusLogic.gs`:

```javascript
// Testar cálculo de dias
testarCalculoDiasRestantes();

// Testar status individual
testarStatusIndividual();

// Testar status geral
testarStatusGeral();

// Ou executar todos de uma vez
executarTodosTestes();
```

Veja os resultados em: **View > Logs** (Ctrl+Enter)

### 2. Na Planilha

1. Execute a função de atualização:
   ```javascript
   atualizarStatusNaPlanilha();
   ```

2. Verifique as abas:
   - **Balancete**: Coluna E deve mostrar "EM CONFORMIDADE X DIAS RESTANTES"
   - **Composição**: Coluna E deve mostrar "EM CONFORMIDADE X DIAS RESTANTES"
   - **Lâmina**: Coluna E deve mostrar "EM CONFORMIDADE X DIAS RESTANTES"
   - **Perfil Mensal**: Coluna E deve mostrar "EM CONFORMIDADE X DIAS RESTANTES"

3. A aba **GERAL** deve refletir estes status.

## Comportamento Esperado

### Cenário 1: Antes do 10º dia útil
- Status individual: "EM CONFORMIDADE"
- Status geral: "EM CONFORMIDADE\n7 DIAS RESTANTES" (exemplo)

### Cenário 2: No 10º dia útil
- Status individual: "EM CONFORMIDADE"
- Status geral: "EM CONFORMIDADE\n0 DIAS RESTANTES"

### Cenário 3: Após o 10º dia útil
- Status individual: "DESATUALIZADO"
- Status geral: "DESCONFORMIDADE"

### Cenário 4: Todos com data correta
- Status individual: "OK"
- Status geral: "OK"

## Arquivos Modificados

1. **DateUtils.gs**
   - Função `getDatasReferencia()` - corrigida
   - Função `calcularDatasManualmente()` - corrigida
   - Função `calcularDiasUteisEntre()` - NOVA

2. **Code.gs**
   - Função `calcularStatusIndividual()` - atualizada
   - Função `calcularStatusGeralDaAba()` - atualizada

3. **TestStatusLogic.gs** - NOVO
   - Testes manuais para validação

## Referências

- Issue: [Link para o issue]
- Pull Request: [Link para o PR]
- Fórmulas originais: Fornecidas na descrição do problema
