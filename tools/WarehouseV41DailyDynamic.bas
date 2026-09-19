Attribute VB_Name = "WarehouseV41DailyDynamic"
Option Explicit

' ============================================================
' NEBRAS WAREHOUSE - DAILY PICKING DYNAMIC V41
' ASCII-SAFE VBA MODULE
'
' Fixes together:
'   1) No 5000-row Daily table. Table size is optimized to actual
'      used input rows + a small manual-entry buffer.
'   2) Bulk import grows the Daily table ONLY by the exact number
'      of additional rows required by the selected file.
'   3) Manual code entry shows stock/location without Occasion.
'   4) Location helper is generated only for the current table
'      plus a small helper reserve, not thousands of rows.
'   5) Green Excel error indicators are ignored ONLY inside the
'      Daily Picking table; global Excel error checking is not
'      disabled.
'
' Safety:
'   - Occasion column is hidden, not deleted.
'   - Existing finalization macro/button is not modified.
'   - Inventory bank data and quantities are never changed here.
'   - V11/mobile sync and production tracker are untouched.
'   - Existing input rows are preserved when table is optimized.
' ============================================================

Private Const DAILY_MIN_ROWS_V41 As Long = 20
Private Const DAILY_BLANK_BUFFER_V41 As Long = 20
Private Const HELPER_EXTRA_ROWS_V41 As Long = 50

Private Const DAILY_ROWNO_COL_V41 As Long = 1
Private Const DAILY_CUSTOMER_COL_V41 As Long = 2
Private Const DAILY_OCC_COL_V41 As Long = 3
Private Const DAILY_CODE_COL_V41 As Long = 4
Private Const DAILY_QTY_COL_V41 As Long = 5
Private Const DAILY_STOCK_COL_V41 As Long = 6
Private Const DAILY_SHELF_COL_V41 As Long = 7
Private Const DAILY_WHROW_COL_V41 As Long = 8
Private Const DAILY_BOX_COL_V41 As Long = 9
Private Const DAILY_LOCATION_COUNT_COL_V41 As Long = 10
Private Const DAILY_AUTO_WITHDRAW_COL_V41 As Long = 11
Private Const DAILY_RESULT_COL_V41 As Long = 12
Private Const DAILY_NOTE_COL_V41 As Long = 13

Private Const BANK_OCC_COL_V41 As Long = 1
Private Const BANK_CODE_COL_V41 As Long = 2
Private Const BANK_STOCK_COL_V41 As Long = 7
Private Const BANK_SHELF_COL_V41 As Long = 8
Private Const BANK_ROW_COL_V41 As Long = 9
Private Const BANK_BOX_COL_V41 As Long = 10
Private Const BANK_KEY_COL_V41 As Long = 13

Private Const HELPER_FIRST_ROW_V41 As Long = 2

Public Sub Nebras_Install_Daily_Dynamic_V41()
    Dim dailyTbl As ListObject
    Dim bankTbl As ListObject
    Dim rewired As Long
    Dim oldCalc As XlCalculation
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldAlerts As Boolean
    Dim oldAutomationSecurity As Long
    Dim failDescription As String

    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldAutomationSecurity = Application.AutomationSecurity

    On Error GoTo Failed

    Set dailyTbl = FindTableV41("WarehousePickingTable")
    Set bankTbl = FindTableV41("InventoryBankTable")

    If dailyTbl Is Nothing Then Err.Raise vbObjectError + 4101, , "WarehousePickingTable was not found."
    If bankTbl Is Nothing Then Err.Raise vbObjectError + 4102, , "InventoryBankTable was not found."

    ValidateSchemaV41 dailyTbl, bankTbl

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' If V40 had expanded the table, shrink it safely to actual
    ' used input rows + a small buffer.
    OptimizeDailyTableSizeV41 dailyTbl

    ' Install code-only formulas for all current rows and create
    ' only a small helper reserve.
    EnsureDailyArchitectureV41 dailyTbl, bankTbl

    ' Intentional product codes are text.
    dailyTbl.ListColumns(DAILY_CODE_COL_V41).DataBodyRange.NumberFormat = "@"

    ' Hide Occasion but keep the internal column for existing
    ' finalization/history compatibility.
    dailyTbl.ListColumns(DAILY_OCC_COL_V41).Range.EntireColumn.Hidden = True

    ' Remove green triangles only in this Daily table.
    IgnoreDailyErrorIndicatorsV41 dailyTbl.DataBodyRange

    ' Redirect the existing bulk-import button when it can be
    ' identified safely. No other buttons are touched.
    rewired = RewireBulkImportButtonV41(dailyTbl.Parent)

    Application.Calculation = oldCalc
    dailyTbl.DataBodyRange.Calculate

    ThisWorkbook.Save

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc

    MsgBox "Daily Picking V41 installed successfully." & vbCrLf & _
           "The table is now dynamic instead of being fixed at 5000 rows." & vbCrLf & _
           "Manual code entry now resolves stock and locations without Occasion." & vbCrLf & _
           "Green indicators were ignored only inside the Daily Picking table." & vbCrLf & _
           IIf(rewired > 0, "Bulk-import button was connected to V41.", _
               "Bulk-import button was not auto-detected; run Nebras_BulkDailyImport_V41 for group import."), _
           vbInformation, "Nebras Warehouse V41"
    Exit Sub

Failed:
    failDescription = Err.Description

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc

    MsgBox "V41 stopped safely." & vbCrLf & failDescription, _
           vbCritical, "Nebras Warehouse V41"
End Sub

Public Sub Nebras_BulkDailyImport_V41()
    Dim dailyTbl As ListObject
    Dim bankTbl As ListObject
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim pickedFile As String

    Dim headerRow As Long
    Dim codeCol As Long
    Dim qtyCol As Long
    Dim customerCol As Long
    Dim firstDataRow As Long
    Dim lastDataRow As Long

    Dim codes As Collection
    Dim qtys As Collection
    Dim customers As Collection

    Dim i As Long
    Dim rowIndex As Long
    Dim imported As Long

    Dim oldCalc As XlCalculation
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldAlerts As Boolean
    Dim failDescription As String

    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts

    On Error GoTo Failed

    Set dailyTbl = FindTableV41("WarehousePickingTable")
    Set bankTbl = FindTableV41("InventoryBankTable")

    If dailyTbl Is Nothing Then Err.Raise vbObjectError + 4110, , "WarehousePickingTable was not found."
    If bankTbl Is Nothing Then Err.Raise vbObjectError + 4111, , "InventoryBankTable was not found."

    ValidateSchemaV41 dailyTbl, bankTbl

    pickedFile = PickImportFileV41()
    If Len(pickedFile) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' Never run macros from the selected import workbook.
    Application.AutomationSecurity = 3

    Set srcWb = Workbooks.Open( _
        Filename:=pickedFile, _
        UpdateLinks:=False, _
        ReadOnly:=True, _
        AddToMru:=False)

    Application.AutomationSecurity = oldAutomationSecurity

    Set srcWs = FindImportSheetV41(srcWb, headerRow, codeCol, qtyCol, customerCol)

    If srcWs Is Nothing Or codeCol = 0 Then
        Err.Raise vbObjectError + 4112, , "A product-code column could not be detected in the selected file."
    End If

    If headerRow > 0 Then
        firstDataRow = headerRow + 1
    Else
        firstDataRow = FirstNonEmptyRowInColumnV41(srcWs, codeCol)
    End If

    lastDataRow = srcWs.Cells(srcWs.Rows.Count, codeCol).End(xlUp).Row

    If firstDataRow <= 0 Or lastDataRow < firstDataRow Then
        Err.Raise vbObjectError + 4113, , "No product codes were found in the selected file."
    End If

    Set codes = New Collection
    Set qtys = New Collection
    Set customers = New Collection

    ReadImportDataV41 srcWs, firstDataRow, lastDataRow, _
                      codeCol, qtyCol, customerCol, _
                      codes, qtys, customers

    If codes.Count = 0 Then
        Err.Raise vbObjectError + 4114, , "No usable product codes were found in the selected file."
    End If

    srcWb.Close SaveChanges:=False
    Set srcWb = Nothing

    ' Compact old unused rows first, then grow only by what this
    ' import actually needs.
    OptimizeDailyTableSizeV41 dailyTbl
    EnsureBlankRowsV41 dailyTbl, codes.Count
    EnsureDailyArchitectureV41 dailyTbl, bankTbl

    dailyTbl.ListColumns(DAILY_CODE_COL_V41).DataBodyRange.NumberFormat = "@"
    dailyTbl.ListColumns(DAILY_OCC_COL_V41).Range.EntireColumn.Hidden = True

    rowIndex = 1

    For i = 1 To codes.Count
        rowIndex = NextBlankCodeRowV41(dailyTbl, rowIndex)

        If rowIndex = 0 Then
            Err.Raise vbObjectError + 4115, , "No blank Daily Picking row remained after expansion."
        End If

        With dailyTbl
            .ListColumns(DAILY_OCC_COL_V41).DataBodyRange.Cells(rowIndex, 1).ClearContents

            .ListColumns(DAILY_CODE_COL_V41).DataBodyRange.Cells(rowIndex, 1).NumberFormat = "@"
            .ListColumns(DAILY_CODE_COL_V41).DataBodyRange.Cells(rowIndex, 1).Value = CStr(codes(i))

            If Len(Trim$(CStr(qtys(i)))) > 0 Then
                .ListColumns(DAILY_QTY_COL_V41).DataBodyRange.Cells(rowIndex, 1).Value = qtys(i)
            End If

            If Len(Trim$(CStr(customers(i)))) > 0 Then
                .ListColumns(DAILY_CUSTOMER_COL_V41).DataBodyRange.Cells(rowIndex, 1).Value = customers(i)
            End If
        End With

        imported = imported + 1
        rowIndex = rowIndex + 1
    Next i

    Application.Calculation = oldCalc

    EnsureDailyArchitectureV41 dailyTbl, bankTbl
    dailyTbl.DataBodyRange.Calculate

    IgnoreDailyErrorIndicatorsV41 dailyTbl.DataBodyRange

    ThisWorkbook.Save

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc
    Application.AutomationSecurity = oldAutomationSecurity

    MsgBox CStr(imported) & " product code(s) imported successfully." & vbCrLf & _
           "Only the required Daily Picking rows were added.", _
           vbInformation, "Nebras Warehouse V41"
    Exit Sub

Failed:
    failDescription = Err.Description

    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    On Error GoTo 0

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc
    Application.AutomationSecurity = oldAutomationSecurity

    MsgBox "Bulk import stopped safely." & vbCrLf & failDescription, _
           vbCritical, "Nebras Warehouse V41"
End Sub

Public Sub Nebras_Optimize_Daily_V41()
    Dim dailyTbl As ListObject
    Dim bankTbl As ListObject

    On Error GoTo Failed

    Set dailyTbl = FindTableV41("WarehousePickingTable")
    Set bankTbl = FindTableV41("InventoryBankTable")

    If dailyTbl Is Nothing Or bankTbl Is Nothing Then
        Err.Raise vbObjectError + 4120, , "Required warehouse tables were not found."
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    OptimizeDailyTableSizeV41 dailyTbl
    EnsureDailyArchitectureV41 dailyTbl, bankTbl
    IgnoreDailyErrorIndicatorsV41 dailyTbl.DataBodyRange

    ThisWorkbook.Save

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Daily Picking table optimized successfully.", _
           vbInformation, "Nebras Warehouse V41"
    Exit Sub

Failed:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Daily optimization stopped safely." & vbCrLf & Err.Description, _
           vbCritical, "Nebras Warehouse V41"
End Sub

Private Sub ValidateSchemaV41(ByVal dailyTbl As ListObject, ByVal bankTbl As ListObject)
    If dailyTbl.ListColumns.Count < DAILY_NOTE_COL_V41 Then
        Err.Raise vbObjectError + 4130, , "WarehousePickingTable schema is shorter than expected."
    End If

    If bankTbl.ListColumns.Count < BANK_KEY_COL_V41 Then
        Err.Raise vbObjectError + 4131, , "InventoryBankTable schema is shorter than expected."
    End If

    If dailyTbl.ShowTotals Then
        Err.Raise vbObjectError + 4132, , "Daily Picking table has a totals row. V41 stopped without changes."
    End If

    If dailyTbl.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 4133, , "WarehousePickingTable has no data rows."
    End If

    If bankTbl.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 4134, , "InventoryBankTable has no data rows."
    End If
End Sub

Private Sub OptimizeDailyTableSizeV41(ByVal dailyTbl As ListObject)
    Dim dailyWs As Worksheet
    Dim oldLastSheetRow As Long
    Dim headerRow As Long
    Dim usedRows As Long
    Dim desiredRows As Long
    Dim desiredLastSheetRow As Long
    Dim newRange As Range

    Set dailyWs = dailyTbl.Parent
    headerRow = dailyTbl.HeaderRowRange.Row
    oldLastSheetRow = dailyTbl.Range.Row + dailyTbl.Range.Rows.Count - 1

    usedRows = LastUsedInputRowIndexV41(dailyTbl)
    desiredRows = Application.Max(DAILY_MIN_ROWS_V41, usedRows + DAILY_BLANK_BUFFER_V41)

    If desiredRows = dailyTbl.ListRows.Count Then Exit Sub

    desiredLastSheetRow = headerRow + desiredRows

    Set newRange = dailyWs.Range( _
        dailyWs.Cells(headerRow, dailyTbl.Range.Column), _
        dailyWs.Cells(desiredLastSheetRow, dailyTbl.Range.Column + dailyTbl.Range.Columns.Count - 1))

    If desiredRows < dailyTbl.ListRows.Count Then
        ' Existing user inputs beyond desiredRows would have been
        ' detected by LastUsedInputRowIndexV41, so only calculated
        ' leftovers are below the new end.
        dailyTbl.Resize newRange

        ClearOldCalculatedTailV41 dailyWs, _
                                  desiredLastSheetRow + 1, _
                                  oldLastSheetRow, _
                                  dailyTbl.Range.Column
    Else
        dailyTbl.Resize newRange
    End If
End Sub

Private Function LastUsedInputRowIndexV41(ByVal dailyTbl As ListObject) As Long
    Dim r As Long

    For r = dailyTbl.ListRows.Count To 1 Step -1
        If Len(Trim$(CStr(dailyTbl.ListColumns(DAILY_CUSTOMER_COL_V41).DataBodyRange.Cells(r, 1).Value))) > 0 Or _
           Len(Trim$(CStr(dailyTbl.ListColumns(DAILY_CODE_COL_V41).DataBodyRange.Cells(r, 1).Value))) > 0 Or _
           Len(Trim$(CStr(dailyTbl.ListColumns(DAILY_QTY_COL_V41).DataBodyRange.Cells(r, 1).Value))) > 0 Or _
           Len(Trim$(CStr(dailyTbl.ListColumns(DAILY_NOTE_COL_V41).DataBodyRange.Cells(r, 1).Value))) > 0 Then

            LastUsedInputRowIndexV41 = r
            Exit Function
        End If
    Next r

    LastUsedInputRowIndexV41 = 0
End Function

Private Sub ClearOldCalculatedTailV41(ByVal ws As Worksheet, _
                                      ByVal firstRow As Long, _
                                      ByVal lastRow As Long, _
                                      ByVal firstTableCol As Long)
    If firstRow > lastRow Then Exit Sub

    ' Only calculated columns previously created by V40/V39 are
    ' cleared outside the resized table. Input columns are untouched.
    ws.Range( _
        ws.Cells(firstRow, firstTableCol + DAILY_ROWNO_COL_V41 - 1), _
        ws.Cells(lastRow, firstTableCol + DAILY_ROWNO_COL_V41 - 1)).ClearContents

    ws.Range( _
        ws.Cells(firstRow, firstTableCol + DAILY_STOCK_COL_V41 - 1), _
        ws.Cells(lastRow, firstTableCol + DAILY_RESULT_COL_V41 - 1)).ClearContents
End Sub

Private Sub EnsureBlankRowsV41(ByVal dailyTbl As ListObject, ByVal neededBlankRows As Long)
    Dim blankRows As Long
    Dim missingRows As Long
    Dim newRows As Long
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim newLastSheetRow As Long
    Dim newRange As Range

    blankRows = CountBlankCodeRowsV41(dailyTbl)

    If blankRows >= neededBlankRows Then Exit Sub

    missingRows = neededBlankRows - blankRows
    newRows = dailyTbl.ListRows.Count + missingRows

    Set ws = dailyTbl.Parent
    headerRow = dailyTbl.HeaderRowRange.Row
    newLastSheetRow = headerRow + newRows

    Set newRange = ws.Range( _
        ws.Cells(headerRow, dailyTbl.Range.Column), _
        ws.Cells(newLastSheetRow, dailyTbl.Range.Column + dailyTbl.Range.Columns.Count - 1))

    dailyTbl.Resize newRange
End Sub

Private Function CountBlankCodeRowsV41(ByVal dailyTbl As ListObject) As Long
    Dim r As Long

    For r = 1 To dailyTbl.ListRows.Count
        If Len(Trim$(CStr(dailyTbl.ListColumns(DAILY_CODE_COL_V41).DataBodyRange.Cells(r, 1).Value))) = 0 Then
            CountBlankCodeRowsV41 = CountBlankCodeRowsV41 + 1
        End If
    Next r
End Function

Private Function NextBlankCodeRowV41(ByVal dailyTbl As ListObject, ByVal startIndex As Long) As Long
    Dim r As Long

    If startIndex < 1 Then startIndex = 1

    For r = startIndex To dailyTbl.ListRows.Count
        If Len(Trim$(CStr(dailyTbl.ListColumns(DAILY_CODE_COL_V41).DataBodyRange.Cells(r, 1).Value))) = 0 Then
            NextBlankCodeRowV41 = r
            Exit Function
        End If
    Next r
End Function

Private Sub EnsureDailyArchitectureV41(ByVal dailyTbl As ListObject, ByVal bankTbl As ListObject)
    Dim helperWs As Worksheet
    Dim dailyWs As Worksheet

    Dim tableRows As Long
    Dim helperRowsNeeded As Long
    Dim helperLastRow As Long
    Dim oldHelperLastRow As Long

    Dim firstDailyRow As Long
    Dim codeSheetCol As Long

    Dim bankOccName As String
    Dim bankCodeName As String
    Dim bankStockName As String
    Dim bankShelfName As String
    Dim bankRowName As String
    Dim bankBoxName As String
    Dim bankKeyName As String

    Dim posCols As Variant
    Dim outCols As Variant
    Dim outputNames As Variant

    Dim rowFormula As String
    Dim col11Formula As String
    Dim col12Formula As String

    Dim f As String
    Dim k As Long

    Set helperWs = FindHelperSheetV41()
    If helperWs Is Nothing Then
        Err.Raise vbObjectError + 4140, , "Daily helper sheet was not found."
    End If

    Set dailyWs = dailyTbl.Parent

    tableRows = dailyTbl.ListRows.Count
    helperRowsNeeded = tableRows + HELPER_EXTRA_ROWS_V41
    helperLastRow = HELPER_FIRST_ROW_V41 + helperRowsNeeded - 1

    firstDailyRow = dailyTbl.DataBodyRange.Row
    codeSheetCol = dailyTbl.ListColumns(DAILY_CODE_COL_V41).Range.Column

    bankOccName = bankTbl.ListColumns(BANK_OCC_COL_V41).Name
    bankCodeName = bankTbl.ListColumns(BANK_CODE_COL_V41).Name
    bankStockName = bankTbl.ListColumns(BANK_STOCK_COL_V41).Name
    bankShelfName = bankTbl.ListColumns(BANK_SHELF_COL_V41).Name
    bankRowName = bankTbl.ListColumns(BANK_ROW_COL_V41).Name
    bankBoxName = bankTbl.ListColumns(BANK_BOX_COL_V41).Name
    bankKeyName = bankTbl.ListColumns(BANK_KEY_COL_V41).Name

    posCols = Array("L", "M", "N", "O", "P", "Q", "R", "S", "T", "U")
    outCols = Array("I", "J", "K")
    outputNames = Array(bankShelfName, bankRowName, bankBoxName)

    ' H = code-only lookup -> first matching Occasion|Code.
    f = HelperKeyFormulaV41( _
        dailyWs.Name, codeSheetCol, firstDailyRow, _
        bankTbl.Name, bankOccName, bankCodeName)

    helperWs.Range("H" & HELPER_FIRST_ROW_V41).Formula = f
    If helperLastRow > HELPER_FIRST_ROW_V41 Then
        helperWs.Range("H" & HELPER_FIRST_ROW_V41 & ":H" & helperLastRow).FillDown
    End If

    ' L:U = up to 10 matching rows for that selected Occasion|Code.
    For k = 0 To 9
        f = PositionFormulaV41(bankTbl.Name, bankKeyName, k + 1)

        helperWs.Range(posCols(k) & HELPER_FIRST_ROW_V41).Formula = f

        If helperLastRow > HELPER_FIRST_ROW_V41 Then
            helperWs.Range(posCols(k) & HELPER_FIRST_ROW_V41 & ":" & _
                           posCols(k) & helperLastRow).FillDown
        End If
    Next k

    ' I:K = shelf / warehouse row / box.
    For k = 0 To 2
        f = LocationJoinFormulaV41( _
            bankTbl.Name, CStr(outputNames(k)), HELPER_FIRST_ROW_V41, posCols)

        helperWs.Range(outCols(k) & HELPER_FIRST_ROW_V41).Formula = f

        If helperLastRow > HELPER_FIRST_ROW_V41 Then
            helperWs.Range(outCols(k) & HELPER_FIRST_ROW_V41 & ":" & _
                           outCols(k) & helperLastRow).FillDown
        End If
    Next k

    ' Remove any old V37/V39/V40 helper formulas far below the
    ' compact range. This is the key workbook-weight reduction.
    oldHelperLastRow = LastHelperFormulaRowV41(helperWs)

    If oldHelperLastRow > helperLastRow Then
        helperWs.Range("H" & (helperLastRow + 1) & ":U" & oldHelperLastRow).ClearContents
    End If

    ' Preserve the workbook's own row-number/result formulas.
    rowFormula = FindExistingFormulaR1C1V41(dailyTbl.ListColumns(DAILY_ROWNO_COL_V41).DataBodyRange)
    col11Formula = FindExistingFormulaR1C1V41(dailyTbl.ListColumns(DAILY_AUTO_WITHDRAW_COL_V41).DataBodyRange)
    col12Formula = FindExistingFormulaR1C1V41(dailyTbl.ListColumns(DAILY_RESULT_COL_V41).DataBodyRange)

    If Len(rowFormula) = 0 Then
        rowFormula = "=IF(RC[3]="""","""",ROW()-" & dailyTbl.HeaderRowRange.Row & ")"
    End If

    If Len(col11Formula) = 0 Then
        col11Formula = "=IF(RC[-7]="""","""",IF(RC[-1]=0,0,MIN(RC[-6],MAX(0,RC[-5]))))"
    End If

    If Len(col12Formula) = 0 Then
        Err.Raise vbObjectError + 4141, , "Existing Daily result formula could not be found."
    End If

    dailyTbl.ListColumns(DAILY_ROWNO_COL_V41).DataBodyRange.FormulaR1C1 = rowFormula

    ' Current stock.
    f = DailyStockFormulaV41( _
        dailyWs, firstDailyRow, codeSheetCol, helperWs.Name, _
        HELPER_FIRST_ROW_V41, bankTbl.Name, bankKeyName, bankStockName)

    dailyTbl.ListColumns(DAILY_STOCK_COL_V41).DataBodyRange.Cells(1, 1).Formula = f
    If tableRows > 1 Then dailyTbl.ListColumns(DAILY_STOCK_COL_V41).DataBodyRange.FillDown

    ' Shelf.
    f = DailyHelperValueFormulaV41( _
        firstDailyRow, codeSheetCol, helperWs.Name, "I", HELPER_FIRST_ROW_V41)

    dailyTbl.ListColumns(DAILY_SHELF_COL_V41).DataBodyRange.Cells(1, 1).Formula = f
    If tableRows > 1 Then dailyTbl.ListColumns(DAILY_SHELF_COL_V41).DataBodyRange.FillDown

    ' Warehouse row.
    f = DailyHelperValueFormulaV41( _
        firstDailyRow, codeSheetCol, helperWs.Name, "J", HELPER_FIRST_ROW_V41)

    dailyTbl.ListColumns(DAILY_WHROW_COL_V41).DataBodyRange.Cells(1, 1).Formula = f
    If tableRows > 1 Then dailyTbl.ListColumns(DAILY_WHROW_COL_V41).DataBodyRange.FillDown

    ' Box.
    f = DailyHelperValueFormulaV41( _
        firstDailyRow, codeSheetCol, helperWs.Name, "K", HELPER_FIRST_ROW_V41)

    dailyTbl.ListColumns(DAILY_BOX_COL_V41).DataBodyRange.Cells(1, 1).Formula = f
    If tableRows > 1 Then dailyTbl.ListColumns(DAILY_BOX_COL_V41).DataBodyRange.FillDown

    ' Location count.
    f = DailyLocationCountFormulaV41( _
        firstDailyRow, codeSheetCol, helperWs.Name, HELPER_FIRST_ROW_V41)

    dailyTbl.ListColumns(DAILY_LOCATION_COUNT_COL_V41).DataBodyRange.Cells(1, 1).Formula = f
    If tableRows > 1 Then dailyTbl.ListColumns(DAILY_LOCATION_COUNT_COL_V41).DataBodyRange.FillDown

    ' Existing auto-withdraw and result logic are preserved.
    dailyTbl.ListColumns(DAILY_AUTO_WITHDRAW_COL_V41).DataBodyRange.FormulaR1C1 = col11Formula
    dailyTbl.ListColumns(DAILY_RESULT_COL_V41).DataBodyRange.FormulaR1C1 = col12Formula

    dailyTbl.ListColumns(DAILY_CODE_COL_V41).DataBodyRange.NumberFormat = "@"
    dailyTbl.ListColumns(DAILY_OCC_COL_V41).Range.EntireColumn.Hidden = True

    ValidateDailyArchitectureV41 dailyTbl, helperWs, helperLastRow
End Sub

Private Function HelperKeyFormulaV41(ByVal dailySheetName As String, _
                                     ByVal codeSheetCol As Long, _
                                     ByVal firstDailyRow As Long, _
                                     ByVal tableName As String, _
                                     ByVal occName As String, _
                                     ByVal codeName As String) As String
    Dim codeAddress As String
    Dim t As String
    Dim o As String
    Dim c As String

    codeAddress = "'" & EscapeSheetV41(dailySheetName) & "'!" & _
                  ColLetterV41(codeSheetCol) & firstDailyRow

    t = EscStructV41(tableName)
    o = EscStructV41(occName)
    c = EscStructV41(codeName)

    HelperKeyFormulaV41 = _
        "=IF(" & codeAddress & "="""","""",IFERROR(" & _
        "INDEX(" & t & "[" & o & "],MATCH(" & codeAddress & "," & _
        t & "[" & c & "],0))&""|""&" & codeAddress & ",""""))"
End Function

Private Function PositionFormulaV41(ByVal tableName As String, _
                                    ByVal keyColumnName As String, _
                                    ByVal nthMatch As Long) As String
    Dim t As String
    Dim k As String

    t = EscStructV41(tableName)
    k = EscStructV41(keyColumnName)

    PositionFormulaV41 = _
        "=IF($H" & HELPER_FIRST_ROW_V41 & "="""","""",IFERROR(AGGREGATE(15,6," & _
        "(ROW(" & t & "[" & k & "])-ROW(INDEX(" & t & "[" & k & "],1,1))+1)/" & _
        "(" & t & "[" & k & "]=$H" & HELPER_FIRST_ROW_V41 & ")," & _
        nthMatch & "),""""))"
End Function

Private Function LocationJoinFormulaV41(ByVal tableName As String, _
                                        ByVal outputColumnName As String, _
                                        ByVal helperRow As Long, _
                                        ByVal posCols As Variant) As String
    Dim k As Long
    Dim s As String
    Dim t As String
    Dim outName As String

    t = EscStructV41(tableName)
    outName = EscStructV41(outputColumnName)

    s = "=IF($H" & helperRow & "="""","""",IFERROR(INDEX(" & _
        t & "[" & outName & "],$" & posCols(0) & helperRow & "),"""")"

    For k = 1 To 9
        s = s & "&IF($" & posCols(k) & helperRow & "="""","""","" | ""&INDEX(" & _
            t & "[" & outName & "],$" & posCols(k) & helperRow & "))"
    Next k

    LocationJoinFormulaV41 = s & ")"
End Function

Private Function DailyStockFormulaV41(ByVal dailyWs As Worksheet, _
                                      ByVal firstDailyRow As Long, _
                                      ByVal codeSheetCol As Long, _
                                      ByVal helperSheetName As String, _
                                      ByVal helperRow As Long, _
                                      ByVal tableName As String, _
                                      ByVal keyName As String, _
                                      ByVal stockName As String) As String
    Dim codeAddress As String
    Dim helperKeyAddress As String
    Dim t As String
    Dim k As String
    Dim s As String

    codeAddress = ColLetterV41(codeSheetCol) & firstDailyRow
    helperKeyAddress = "'" & EscapeSheetV41(helperSheetName) & "'!H" & helperRow

    t = EscStructV41(tableName)
    k = EscStructV41(keyName)
    s = EscStructV41(stockName)

    DailyStockFormulaV41 = _
        "=IF(" & codeAddress & "="""","""",IF(" & helperKeyAddress & "="""",0," & _
        "SUMIF(" & t & "[" & k & "]," & helperKeyAddress & "," & _
        t & "[" & s & "])))"
End Function

Private Function DailyHelperValueFormulaV41(ByVal firstDailyRow As Long, _
                                            ByVal codeSheetCol As Long, _
                                            ByVal helperSheetName As String, _
                                            ByVal helperCol As String, _
                                            ByVal helperRow As Long) As String
    Dim codeAddress As String

    codeAddress = ColLetterV41(codeSheetCol) & firstDailyRow

    DailyHelperValueFormulaV41 = _
        "=IF(" & codeAddress & "="""","""",IFERROR('" & _
        EscapeSheetV41(helperSheetName) & "'!" & helperCol & helperRow & ",""""))"
End Function

Private Function DailyLocationCountFormulaV41(ByVal firstDailyRow As Long, _
                                              ByVal codeSheetCol As Long, _
                                              ByVal helperSheetName As String, _
                                              ByVal helperRow As Long) As String
    Dim codeAddress As String

    codeAddress = ColLetterV41(codeSheetCol) & firstDailyRow

    DailyLocationCountFormulaV41 = _
        "=IF(" & codeAddress & "="""","""",COUNT('" & _
        EscapeSheetV41(helperSheetName) & "'!L" & helperRow & ":U" & helperRow & "))"
End Function

Private Function LastHelperFormulaRowV41(ByVal ws As Worksheet) As Long
    Dim col As Long
    Dim r As Long
    Dim lastFound As Long

    For col = 8 To 21
        r = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
        If r > lastFound Then lastFound = r
    Next col

    If lastFound < HELPER_FIRST_ROW_V41 Then lastFound = HELPER_FIRST_ROW_V41
    LastHelperFormulaRowV41 = lastFound
End Function

Private Function FindExistingFormulaR1C1V41(ByVal rng As Range) As String
    Dim c As Range

    If rng Is Nothing Then Exit Function

    For Each c In rng.Cells
        If c.HasFormula Then
            If InStr(1, CStr(c.Formula), "#REF!", vbTextCompare) = 0 Then
                FindExistingFormulaR1C1V41 = CStr(c.FormulaR1C1)
                Exit Function
            End If
        End If
    Next c
End Function

Private Sub ValidateDailyArchitectureV41(ByVal dailyTbl As ListObject, _
                                         ByVal helperWs As Worksheet, _
                                         ByVal helperLastRow As Long)
    Dim c As Range
    Dim colIndex As Variant

    For Each c In helperWs.Range("H2:U2").Cells
        If Not c.HasFormula Then
            Err.Raise vbObjectError + 4150, , _
                      "Missing helper formula at " & c.Address(False, False) & "."
        End If

        If InStr(1, CStr(c.Formula), "#REF!", vbTextCompare) > 0 Then
            Err.Raise vbObjectError + 4151, , _
                      "Broken helper reference at " & c.Address(False, False) & "."
        End If
    Next c

    If Not helperWs.Range("H" & helperLastRow).HasFormula Then
        Err.Raise vbObjectError + 4152, , "Helper reserve was not created correctly."
    End If

    For Each colIndex In Array( _
        DAILY_ROWNO_COL_V41, _
        DAILY_STOCK_COL_V41, _
        DAILY_SHELF_COL_V41, _
        DAILY_WHROW_COL_V41, _
        DAILY_BOX_COL_V41, _
        DAILY_LOCATION_COUNT_COL_V41, _
        DAILY_AUTO_WITHDRAW_COL_V41, _
        DAILY_RESULT_COL_V41)

        Set c = dailyTbl.ListColumns(CLng(colIndex)).DataBodyRange.Cells(1, 1)

        If Not c.HasFormula Then
            Err.Raise vbObjectError + 4153, , _
                      "Missing Daily formula in table column " & CStr(colIndex) & "."
        End If

        If InStr(1, CStr(c.Formula), "#REF!", vbTextCompare) > 0 Then
            Err.Raise vbObjectError + 4154, , _
                      "Broken Daily formula in table column " & CStr(colIndex) & "."
        End If
    Next colIndex
End Sub

Private Sub IgnoreDailyErrorIndicatorsV41(ByVal rng As Range)
    Dim c As Range
    Dim errorType As Long

    If rng Is Nothing Then Exit Sub

    For Each c In rng.Cells
        For errorType = 1 To 9
            On Error Resume Next

            If c.Errors(errorType).Value Then
                c.Errors(errorType).Ignore = True
            End If

            Err.Clear
            On Error GoTo 0
        Next errorType
    Next c
End Sub

Private Function PickImportFileV41() As String
    Dim fd As Object

    Set fd = Application.FileDialog(3)

    With fd
        .AllowMultiSelect = False
        .Title = "Select the Excel file containing product codes"
        .Filters.Clear
        .Filters.Add "Excel files", "*.xlsx;*.xlsm;*.xls;*.xlsb"

        If .Show <> -1 Then Exit Function

        PickImportFileV41 = CStr(.SelectedItems(1))
    End With
End Function

Private Function FindImportSheetV41(ByVal wb As Workbook, _
                                    ByRef headerRow As Long, _
                                    ByRef codeCol As Long, _
                                    ByRef qtyCol As Long, _
                                    ByRef customerCol As Long) As Worksheet
    Dim ws As Worksheet
    Dim r As Long
    Dim c As Long
    Dim maxRow As Long
    Dim maxCol As Long
    Dim txt As String

    Dim scanFirstRow As Long
    Dim scanLastRow As Long
    Dim scanFirstCol As Long
    Dim scanLastCol As Long

    For Each ws In wb.Worksheets
        scanFirstRow = Application.Max(1, ws.UsedRange.Row)
        scanLastRow = Application.Min(ws.Rows.Count, scanFirstRow + 19, _
                                      ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1)

        scanFirstCol = Application.Max(1, ws.UsedRange.Column)
        scanLastCol = Application.Min(ws.Columns.Count, scanFirstCol + 99, _
                                      ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1)

        maxRow = scanLastRow
        maxCol = scanLastCol

        For r = scanFirstRow To maxRow
            For c = scanFirstCol To maxCol
                txt = NormalizeHeaderV41(ws.Cells(r, c).Value)

                If codeCol = 0 And IsCodeHeaderV41(txt) Then
                    headerRow = r
                    codeCol = c
                End If

                If qtyCol = 0 And IsQtyHeaderV41(txt) Then
                    qtyCol = c
                End If

                If customerCol = 0 And IsCustomerHeaderV41(txt) Then
                    customerCol = c
                End If
            Next c

            If codeCol > 0 Then
                Set FindImportSheetV41 = ws
                Exit Function
            End If
        Next r

        headerRow = 0
        codeCol = 0
        qtyCol = 0
        customerCol = 0
    Next ws

    ' Headerless fallback: first non-empty column of the first
    ' non-empty worksheet. Quantity/customer remain optional.
    For Each ws In wb.Worksheets
        For c = 1 To Application.Min(100, ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1)
            If Application.WorksheetFunction.CountA(ws.Columns(c)) > 0 Then
                codeCol = c
                headerRow = 0
                Set FindImportSheetV41 = ws
                Exit Function
            End If
        Next c
    Next ws
End Function

Private Sub ReadImportDataV41(ByVal ws As Worksheet, _
                              ByVal firstDataRow As Long, _
                              ByVal lastDataRow As Long, _
                              ByVal codeCol As Long, _
                              ByVal qtyCol As Long, _
                              ByVal customerCol As Long, _
                              ByVal codes As Collection, _
                              ByVal qtys As Collection, _
                              ByVal customers As Collection)
    Dim r As Long
    Dim code As String
    Dim q As Variant
    Dim customer As Variant

    For r = firstDataRow To lastDataRow
        code = Trim$(CStr(ws.Cells(r, codeCol).Value))

        If Len(code) > 0 Then
            If Not IsLikelyHeaderValueV41(code) Then
                codes.Add code

                If qtyCol > 0 Then
                    q = ws.Cells(r, qtyCol).Value
                Else
                    q = vbNullString
                End If

                If customerCol > 0 Then
                    customer = ws.Cells(r, customerCol).Value
                Else
                    customer = vbNullString
                End If

                qtys.Add q
                customers.Add customer
            End If
        End If
    Next r
End Sub

Private Function FirstNonEmptyRowInColumnV41(ByVal ws As Worksheet, ByVal col As Long) As Long
    Dim r As Long
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    For r = 1 To lastRow
        If Len(Trim$(CStr(ws.Cells(r, col).Value))) > 0 Then
            FirstNonEmptyRowInColumnV41 = r
            Exit Function
        End If
    Next r
End Function

Private Function NormalizeHeaderV41(ByVal v As Variant) As String
    Dim s As String

    s = LCase$(Trim$(CStr(v)))
    s = Replace(s, ChrW(&H64A), ChrW(&H6CC))
    s = Replace(s, ChrW(&H643), ChrW(&H6A9))
    s = Replace(s, " ", "")
    s = Replace(s, "_", "")
    s = Replace(s, "-", "")

    NormalizeHeaderV41 = s
End Function

Private Function IsCodeHeaderV41(ByVal s As String) As Boolean
    If s = "code" Or s = "productcode" Or s = "sku" Then
        IsCodeHeaderV41 = True
        Exit Function
    End If

    If s = NormalizeHeaderV41(U("06A9 062F")) Or _
       s = NormalizeHeaderV41(U("06A9 062F 0020 0645 062D 0635 0648 0644")) Or _
       s = NormalizeHeaderV41(U("06A9 062F 0020 06A9 0627 0644 0627")) Then
        IsCodeHeaderV41 = True
    End If
End Function

Private Function IsQtyHeaderV41(ByVal s As String) As Boolean
    If s = "qty" Or s = "quantity" Or s = "requiredqty" Then
        IsQtyHeaderV41 = True
        Exit Function
    End If

    If s = NormalizeHeaderV41(U("062A 0639 062F 0627 062F")) Or _
       s = NormalizeHeaderV41(U("062A 0639 062F 0627 062F 0020 0645 0648 0631 062F 0646 06CC 0627 0632")) Then
        IsQtyHeaderV41 = True
    End If
End Function

Private Function IsCustomerHeaderV41(ByVal s As String) As Boolean
    If s = "customer" Or s = "customername" Then
        IsCustomerHeaderV41 = True
        Exit Function
    End If

    If s = NormalizeHeaderV41(U("0645 0634 062A 0631 06CC")) Then
        IsCustomerHeaderV41 = True
    End If
End Function

Private Function IsLikelyHeaderValueV41(ByVal s As String) As Boolean
    Dim n As String

    n = NormalizeHeaderV41(s)

    IsLikelyHeaderValueV41 = _
        IsCodeHeaderV41(n) Or _
        IsQtyHeaderV41(n) Or _
        IsCustomerHeaderV41(n)
End Function

Private Function RewireBulkImportButtonV41(ByVal ws As Worksheet) As Long
    Dim shp As Shape
    Dim actionName As String
    Dim caption As String
    Dim actionLower As String

    For Each shp In ws.Shapes
        actionName = vbNullString
        caption = vbNullString

        On Error Resume Next
        actionName = CStr(shp.OnAction)
        caption = ShapeCaptionV41(shp)
        On Error GoTo 0

        actionLower = LCase$(actionName)

        If InStr(actionLower, "final") = 0 And _
           InStr(actionLower, "sync") = 0 Then

            If InStr(actionLower, "bulk") > 0 Or _
               InStr(actionLower, "import") > 0 Or _
               InStr(actionLower, "v29") > 0 Or _
               InStr(1, caption, U("06AF 0631 0648 0647 06CC"), vbTextCompare) > 0 Then

                On Error Resume Next
                shp.OnAction = "Nebras_BulkDailyImport_V41"

                If Err.Number = 0 Then
                    RewireBulkImportButtonV41 = RewireBulkImportButtonV41 + 1
                End If

                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next shp
End Function

Private Function ShapeCaptionV41(ByVal shp As Shape) As String
    On Error Resume Next

    ShapeCaptionV41 = shp.TextFrame2.TextRange.Text

    If Len(ShapeCaptionV41) = 0 Then
        ShapeCaptionV41 = shp.TextFrame.Characters.Text
    End If

    On Error GoTo 0
End Function

Private Function FindTableV41(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing

        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0

        If Not lo Is Nothing Then
            Set FindTableV41 = lo
            Exit Function
        End If
    Next ws
End Function

Private Function FindHelperSheetV41() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(1)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    If Len(Trim$(CStr(ws.Range("H1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("I1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("J1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("K1").Value))) = 0 Then Exit Function

    Set FindHelperSheetV41 = ws
End Function

Private Function ColLetterV41(ByVal colNumber As Long) As String
    ColLetterV41 = Split(Cells(1, colNumber).Address(False, False), "1")(0)
End Function

Private Function EscapeSheetV41(ByVal s As String) As String
    EscapeSheetV41 = Replace(s, "'", "''")
End Function

Private Function EscStructV41(ByVal s As String) As String
    EscStructV41 = Replace(s, "]", "]]")
End Function

Private Function U(ByVal hexList As String) As String
    Dim parts() As String
    Dim i As Long
    Dim result As String

    If Len(Trim$(hexList)) = 0 Then Exit Function

    parts = Split(Trim$(hexList), " ")

    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            result = result & ChrW$(CLng("&H" & parts(i)))
        End If
    Next i

    U = result
End Function
