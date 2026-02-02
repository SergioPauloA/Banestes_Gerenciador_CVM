# Status Formula Fix - Implementation Summary

## Problem
The status formulas were not displaying "EM CONFORMIDADE X DIAS RESTANTES" in the general status column (E) and individual status column (D) of the sheets Balancete, Composição, Lâmina, and Perfil Mensal.

## Root Cause
1. The formulas in `onInstall.gs` were too simple and didn't include the "days remaining" logic
2. Named ranges (DIAMESREF, DIAMESREF2, DIADDD) were not being created in the APOIO sheet

## Solution Implemented

### 1. Created Named Ranges in DateUtils.gs
Added `criarNamedRangesDatas()` helper function that creates three named ranges in the APOIO sheet:
- **DIAMESREF** = APOIO!D1 (1º dia do mês anterior)
- **DIAMESREF2** = APOIO!F1 (10º dia útil do mês atual - prazo limite)
- **DIADDD** = APOIO!A17 (Data de hoje)

These named ranges are used by the formulas in all sheets and are created automatically when `criarAbaApoioComValores()` runs.

### 2. Updated Column E Formulas (General Status)
Created a constant `FORMULA_STATUS_GERAL` that is reused in all 4 sheets:
```javascript
var FORMULA_STATUS_GERAL = '=SE(CONT.SE(D:D;"OK")=CONT.SE(A4:A;"<>"&"");"OK";SE(DIADDD<=DIAMESREF2;"EM CONFORMIDADE " & CARACT(10) & DATADIF(DIADDD;DIAMESREF2;"D") & " DIAS RESTANTES";"DESCONFORMIDADE"))';
```

This formula now:
- Returns "OK" if all individual statuses are OK
- Returns "EM CONFORMIDADE\nX DIAS RESTANTES" if still within the deadline (uses DATEDIF to calculate days)
- Returns "DESCONFORMIDADE" if past the deadline

### 3. Updated Column D Formulas (Individual Status)
Created a helper function `criarFormulaStatusIndividual(linha)` that generates the formula for each row:
```javascript
function criarFormulaStatusIndividual(linha) {
  return '=SE(C' + linha + '=DIAMESREF;"OK";SE(DIADDD<=DIAMESREF2;"EM CONFORMIDADE";"DESATUALIZADO"))';
}
```

This formula:
- Returns "OK" if the return date equals the reference month (DIAMESREF)
- Returns "EM CONFORMIDADE" if still within the deadline
- Returns "DESATUALIZADO" if past the deadline

### 4. Code Quality Improvements
- Extracted duplicate formulas to constants and helper functions
- Moved `criarNamedRangesDatas()` to module level for better reusability
- Improved error handling with descriptive messages
- Preemptively removes existing named ranges to avoid conflicts

## Files Modified
1. **DateUtils.gs** - Added:
   - `criarNamedRangesDatas(ss, abaApoio)` helper function
   - Improved error handling in `criarAbaApoioComValores()`
   
2. **onInstall.gs** - Updated:
   - Added `FORMULA_STATUS_GERAL` constant
   - Added `criarFormulaStatusIndividual(linha)` helper function
   - Updated 4 functions:
     - `criarFormulasBalancete()`
     - `criarFormulasComposicao()`
     - `criarFormulasLamina()`
     - `criarFormulasPerfilMensal()`

## Expected Behavior
After running `setupCompletoAutomatico()` or `onInstall()`:
1. The APOIO sheet will have named ranges defined (DIAMESREF, DIAMESREF2, DIADDD)
2. Each sheet (Balancete, Composição, Lâmina, Perfil Mensal) will have:
   - Column D with individual status formulas that show "OK", "EM CONFORMIDADE", or "DESATUALIZADO"
   - Column E (row 1) with general status formula showing days remaining like "EM CONFORMIDADE\n11 DIAS RESTANTES"
3. Status will automatically update based on:
   - Current date (DIADDD in APOIO!A17)
   - Reference month (DIAMESREF in APOIO!D1)
   - Deadline (DIAMESREF2 in APOIO!F1)

## Testing
To test the implementation:
1. Run `criarAbaApoioComValores()` to create named ranges
2. Run `setupCompletoAutomatico()` to recreate all sheets with new formulas
3. Check that column E shows "EM CONFORMIDADE X DIAS RESTANTES" when appropriate
4. Verify column D shows correct individual statuses for each fund
5. Confirm the status updates automatically as dates change

## Notes
- The formula uses `CARACT(10)` which is the line break character in Google Sheets (CHAR(10))
- `DATADIF(DIADDD;DIAMESREF2;"D")` calculates the number of days between today and the deadline
- These formulas will automatically recalculate whenever the dates in APOIO sheet are updated
- The named ranges ensure consistency across all sheets
- All code improvements follow best practices with constants, helper functions, and proper error handling

## Example Output
When working correctly, the status cells will display:
- Column D (individual): "OK", "EM CONFORMIDADE", or "DESATUALIZADO"
- Column E (general): 
  - "OK" when all funds are updated
  - "EM CONFORMIDADE\n11 DIAS RESTANTES" when within deadline (11 is an example)
  - "DESCONFORMIDADE" when past deadline
