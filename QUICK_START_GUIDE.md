# Quick Start Guide - Testing the Fix

## 🎯 What You'll Do
Test the status calculation fix in 4 simple steps.

## ⚡ Step-by-Step Instructions

### Step 1: Open Google Apps Script Editor
1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/1N6LP1ydsxnQO_Woatv9zWEccb0fOGaV_3EKK1GoSWZI
2. Click **Extensions** → **Apps Script**
3. Wait for the editor to load

### Step 2: Run Test Function ✅
```
1. In the function dropdown at the top, select: testarCalculoDeStatus
2. Click the "Run" button (▶️)
3. If prompted, authorize the script
4. Click View → Logs (or Ctrl+Enter)
```

**What to Look For**:
```
✅ See date references (diaMesRef: 01/12/2025)
✅ See "Comparando:" messages showing date matching
✅ See status calculations for 5 test rows
✅ See final "✅ Teste concluído!"
❌ No error messages
```

### Step 3: Update All Sheets ✅
```
1. In the function dropdown, select: atualizarStatusNaPlanilhaAutomatico
2. Click the "Run" button (▶️)
3. View Logs to see progress
```

**What to Look For**:
```
✅ See "Datas de referência:" section
✅ See "Processando Balancete..." with status counts
✅ See "Processando Composição..." 
✅ See "Processando Lâmina..."
✅ See "Processando Perfil Mensal..."
✅ See "✅ [TRIGGER] Atualização automática concluída!"
❌ No "❌ [TRIGGER] Erro" messages
```

### Step 4: Verify in Spreadsheet ✅
```
1. Go back to your Google Sheet
2. Check these tabs: Balancete, Composição, Lâmina, Perfil Mensal
```

**For Each Tab, Check**:

**Column D (rows 4-29)**:
- ✅ Should NOT be all "-" (dashes)
- ✅ Should show mix of:
  - "OK" (for rows with 01/12/2025)
  - "EM CONFORMIDADE" (for rows with other dates)
  - "DESATUALIZADO" (for empty rows)

**Cell E1**:
- ✅ Should show one of:
  - "OK" (if all funds are up to date)
  - "EM CONFORMIDADE" (on line 1)
    "11 DIAS RESTANTES" (on line 2)
  - "DESCONFORMIDADE" (if past deadline)

## 🎨 Visual Example

### BEFORE (Broken) ❌
```
┌─────────────────────────────────────────────┐
│ Balancete Tab                               │
├────┬─────────────┬────────────┬────────────┤
│Row │ Fund        │ Date (C)   │ Status (D) │
├────┼─────────────┼────────────┼────────────┤
│ 4  │ Fund 1      │ 01/12/2025 │ -          │
│ 5  │ Fund 2      │ 15/01/2026 │ -          │
│ 6  │ Fund 3      │ 20/01/2026 │ -          │
└────┴─────────────┴────────────┴────────────┘

Cell E1: [ OK ]
```

### AFTER (Fixed) ✅
```
┌──────────────────────────────────────────────────────┐
│ Balancete Tab                                        │
├────┬─────────────┬────────────┬─────────────────────┤
│Row │ Fund        │ Date (C)   │ Status (D)          │
├────┼─────────────┼────────────┼─────────────────────┤
│ 4  │ Fund 1      │ 01/12/2025 │ OK                  │
│ 5  │ Fund 2      │ 15/01/2026 │ EM CONFORMIDADE     │
│ 6  │ Fund 3      │ 20/01/2026 │ EM CONFORMIDADE     │
└────┴─────────────┴────────────┴─────────────────────┘

Cell E1: ┌──────────────────────┐
         │ EM CONFORMIDADE      │
         │ 11 DIAS RESTANTES    │
         └──────────────────────┘
```

## 🐛 Troubleshooting

### Problem: Authorization Required
**Solution**: Click "Review Permissions" → Select your account → Click "Advanced" → "Go to [Project Name]" → "Allow"

### Problem: "ReferenceError: getDatasReferencia is not defined"
**Solution**: 
1. Make sure all .gs files are saved
2. Check that DateUtils.gs is present in the project
3. Refresh the Apps Script Editor page

### Problem: Logs show errors
**Solution**:
1. Copy the error message
2. Check IMPLEMENTATION_SUMMARY.md for common issues
3. Verify SPREADSHEET_ID is correct in Code.gs line 6

### Problem: Column D still shows "-"
**Possible Causes**:
1. The update function hasn't been run yet → Run Step 3
2. The dates in Column C are in unexpected format → Check logs for "Comparando:" messages
3. DEBUG_MODE needs to be true for detailed logging → Check Code.gs line 9

## ✅ Success Checklist

After completing all steps, you should have:
- [ ] Test function ran without errors
- [ ] Full update function completed successfully
- [ ] Column D in Balancete shows status values (not "-")
- [ ] Column D in Composição shows status values
- [ ] Column D in Lâmina shows status values
- [ ] Column D in Perfil Mensal shows status values
- [ ] Cell E1 in each tab shows correct general status
- [ ] Dashboard (if web interface exists) displays correctly

## 🚀 Production Deployment

Once everything is verified:
```
1. Open Code.gs in Apps Script Editor
2. Find line 9: var DEBUG_MODE = true;
3. Change to: var DEBUG_MODE = false;
4. Click File → Save
5. Done! The automatic triggers will now run with minimal logging
```

## 📊 What Changed?

**3 New Functions**:
1. `normalizaDataParaComparacao()` - Better date handling
2. `testarCalculoDeStatus()` - Test function you just ran
3. Enhanced `calcularStatusIndividual()` - Fixed status logic

**1 Enhanced Function**:
1. `atualizarStatusNaPlanilhaAutomatico()` - Better logging

**Total Lines Changed**: +151 lines, -15 lines

## 🎉 Expected Timeline

- Step 1-2: 2 minutes (test function)
- Step 3: 2-3 minutes (full update)
- Step 4: 2 minutes (verification)
- **Total**: ~7-8 minutes

## 📞 Need Help?

If you encounter issues:
1. Check the execution logs for error messages
2. Review IMPLEMENTATION_SUMMARY.md for detailed documentation
3. Check BEFORE_AFTER_COMPARISON.md for technical details
4. Review FINAL_SUMMARY.md for deployment information

## 🎊 You're Done!

Once all checkboxes are marked, the fix is verified and ready for production use!
