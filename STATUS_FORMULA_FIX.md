# Status Formula Fix - Implementation Summary

## Problem
The status formulas were not displaying "EM CONFORMIDADE X DIAS RESTANTES" in the general status column (E) and individual status column (D) of the sheets Balancete, Composição, Lâmina, and Perfil Mensal.

## Root Cause
1. The formulas in `onInstall.gs` were too simple and didn't include the "days remaining" logic
2. Named ranges (DIAMESREF, DIAMESREF2, DIADDD) were not being created in the APOIO sheet

## Solution Implemented

### 1. Created Named Ranges in DateUtils.gs (criarAbaApoioComValores function)
Added code to create three named ranges in the APOIO sheet:
- **DIAMESREF** = APOIO!D1 (1º dia do mês anterior)
- **DIAMESREF2** = APOIO!F1 (10º dia útil do mês atual - prazo limite)
- **DIADDD** = APOIO!A17 (Data de hoje)

These named ranges are used by the formulas in all sheets.

### 2. Updated Column E Formulas (General Status)
Changed the formula in all 4 sheets from:
```
=SE(CONT.SE(D:D;"OK")=CONT.SE(A4:A;"<>"&"");"OK";"DESCONFORMIDADE")
```

To:
```
=SE(CONT.SE(D:D;"OK")=CONT.SE(A4:A;"<>"&"");"OK";SE(DIADDD<=DIAMESREF2;"EM CONFORMIDADE " & CARACT(10) & DATADIF(DIADDD;DIAMESREF2;"D") & " DIAS RESTANTES";"DESCONFORMIDADE"))
```

This formula now:
- Returns "OK" if all individual statuses are OK
- Returns "EM CONFORMIDADE\nX DIAS RESTANTES" if still within the deadline (uses DATEDIF to calculate days)
- Returns "DESCONFORMIDADE" if past the deadline

### 3. Updated Column D Formulas (Individual Status)
Changed from a static value 'Aguardando...' to a formula:
```
=SE(C[linha]=DIAMESREF;"OK";SE(DIADDD<=DIAMESREF2;"EM CONFORMIDADE";"DESATUALIZADO"))
```

This formula:
- Returns "OK" if the return date equals the reference month (DIAMESREF)
- Returns "EM CONFORMIDADE" if still within the deadline
- Returns "DESATUALIZADO" if past the deadline

## Files Modified
1. **DateUtils.gs** - Added named range creation logic
2. **onInstall.gs** - Updated formulas in 4 functions:
   - `criarFormulasBalancete()`
   - `criarFormulasComposicao()`
   - `criarFormulasLamina()`
   - `criarFormulasPerfilMensal()`

## Expected Behavior
After running `setupCompletoAutomatico()` or `onInstall()`:
1. The APOIO sheet will have named ranges defined
2. Each sheet (Balancete, Composição, Lâmina, Perfil Mensal) will have:
   - Column D with individual status formulas
   - Column E (row 1) with general status formula showing days remaining
3. Status will automatically update based on:
   - Current date (DIADDD)
   - Reference month (DIAMESREF)
   - Deadline (DIAMESREF2)

## Testing
To test the implementation:
1. Run `criarAbaApoioComValores()` to create named ranges
2. Run `setupCompletoAutomatico()` to recreate all sheets with new formulas
3. Check that column E shows "EM CONFORMIDADE X DIAS RESTANTES" when appropriate
4. Verify column D shows correct individual statuses

## Notes
- The formula uses `CARACT(10)` which is the line break character in Google Sheets
- `DATADIF(DIADDD;DIAMESREF2;"D")` calculates the number of days between today and the deadline
- These formulas will automatically recalculate whenever the dates in APOIO sheet are updated
