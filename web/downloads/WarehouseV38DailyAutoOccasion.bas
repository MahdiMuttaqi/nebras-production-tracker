Attribute VB_Name = "WarehouseV38DailyAutoOccasion"
Option Explicit

' ============================================================
' NEBRAS WAREHOUSE - DAILY PICKING AUTO OCCASION V38
' ASCII-SAFE VBA MODULE
'
' Purpose:
'   - Remove the need for the user to enter Occasion manually
'     in the Daily Picking sheet.
'   - Keep the Occasion table column internally for compatibility
'     with existing finalization/history/inventory logic.
'   - Hide that worksheet column from the user.
'   - Auto-fill Occasion from InventoryBankTable based on the
'     entered product code.
'
' Duplicate-code rule:
'   If one product code exists under more than one occasion,
'   V38 selects the bank row with the HIGHEST current stock and
'   uses that row's occasion. This keeps the existing
'   occasion+code downstream logic consistent.
'
' Scope:
'   - WarehousePickingTable column 3 only (internal Occasion)
'   - Hides that physical worksheet column
'   - Recalculates the already-installed V37 helper H:U
'
' It does NOT delete/reorder table columns.
' It does NOT change inventory quantities, product codes,
' requested quantities, finalization macros, history, V11,
' mobile sync, or production tracker.
' ============================================================

Private Const DAILY_OCC_COL_V38 As Long = 3
Private Const DAILY_CODE_COL_V38 As Long = 4

Private Const BANK_OCC_COL_V38 As Long = 1
Private Const BANK_CODE_COL_V38 As Long = 2
Private Const BANK_STOCK_COL_V38 As Long = 7

Private Const HELPER_FIRST_ROW_V38 As Long = 2
Private Const HELPER_TARGET_DAILY_ROW_V38 As Long = 5000

Public Sub Nebras_Install_Daily_AutoOccasion_V38()
    Dim dailyTbl As ListObject
    Dim bankTbl As ListObject
    Dim dailyWs As Worksheet
    Dim helperWs As Worksheet

    Dim occCol As ListColumn
    Dim codeCol As ListColumn
    Dim occRange As Range

    Dim bankOccName As String
    Dim bankCodeName As String
    Dim bankStockName As String

    Dim firstDailyRow As Long
    Dim helperLastRow As Long
    Dim dailyOffset As Long

    Dim formulaText As String
    Dim codeCellAddress As String

    Dim oldOccFormulas As Variant
    Dim oldHidden As Boolean
    Dim snapshotTaken As Boolean

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

    Set dailyTbl = FindTableV38("WarehousePickingTable")
    Set bankTbl = FindTableV38("InventoryBankTable")

    If dailyTbl Is Nothing Then Err.Raise vbObjectError + 3801, , "WarehousePickingTable was not found."
    If bankTbl Is Nothing Then Err.Raise vbObjectError + 3802, , "InventoryBankTable was not found."

    If dailyTbl.ListColumns.Count < DAILY_CODE_COL_V38 Then
        Err.Raise vbObjectError + 3803, , "WarehousePickingTable schema is shorter than expected."
    End If

    If bankTbl.ListColumns.Count < BANK_STOCK_COL_V38 Then
        Err.Raise vbObjectError + 3804, , "InventoryBankTable schema is shorter than expected."
    End If

    Set dailyWs = dailyTbl.Parent
    Set helperWs = FindHelperSheetV38()

    If helperWs Is Nothing Then
        Err.Raise vbObjectError + 3805, , "V37 Daily location helper sheet was not found."
    End If

    Set occCol = dailyTbl.ListColumns(DAILY_OCC_COL_V38)
    Set codeCol = dailyTbl.ListColumns(DAILY_CODE_COL_V38)

    bankOccName = bankTbl.ListColumns(BANK_OCC_COL_V38).Name
    bankCodeName = bankTbl.ListColumns(BANK_CODE_COL_V38).Name
    bankStockName = bankTbl.ListColumns(BANK_STOCK_COL_V38).Name

    If dailyTbl.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 3806, , "WarehousePickingTable has no data rows."
    End If

    Set occRange = occCol.DataBodyRange
    firstDailyRow = dailyTbl.DataBodyRange.Row

    dailyOffset = firstDailyRow - HELPER_FIRST_ROW_V38
    If dailyOffset < 1 Then Err.Raise vbObjectError + 3807, , "Unexpected Daily Picking table position."

    helperLastRow = HELPER_TARGET_DAILY_ROW_V38 - dailyOffset
    If helperLastRow < HELPER_FIRST_ROW_V38 Then
        Err.Raise vbObjectError + 3808, , "Invalid helper calculation range."
    End If

    codeCellAddress = dailyWs.Cells(firstDailyRow, codeCol.Range.Column).Address(False, False)

    formulaText = BuildAutoOccasionFormulaV38( _
        bankTbl.Name, bankOccName, bankCodeName, bankStockName, codeCellAddress)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' Rollback snapshot: only the hidden internal Occasion column.
    oldOccFormulas = occRange.Formula
    oldHidden = occCol.Range.EntireColumn.Hidden
    snapshotTaken = True

    ' Put the formula in the first table row, then FillDown.
    occRange.Cells(1, 1).Formula = formulaText
    If occRange.Rows.Count > 1 Then occRange.FillDown

    ' Hide, do not delete. Deleting the table column could break
    ' existing macros that depend on the current table schema.
    occCol.Range.EntireColumn.Hidden = True

    ' Validate that the formula is present and has no broken refs.
    ValidateOccasionFormulaV38 occRange

    Application.Calculation = oldCalc

    ' Recalculate only the relevant dependency chain.
    occRange.Calculate
    helperWs.Range("H" & HELPER_FIRST_ROW_V38 & ":U" & helperLastRow).Calculate
    dailyTbl.DataBodyRange.Calculate

    ThisWorkbook.Save

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc

    MsgBox "Daily Picking V38 installed successfully." & vbCrLf & _
           "Occasion is now automatic and hidden from the Daily Picking sheet." & vbCrLf & _
           "Enter only customer, product code and required quantity.", _
           vbInformation, "Nebras Warehouse V38"
    Exit Sub

Failed:
    failDescription = Err.Description

    On Error Resume Next
    If snapshotTaken Then
        occRange.Formula = oldOccFormulas
        occCol.Range.EntireColumn.Hidden = oldHidden
    End If

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc
    On Error GoTo 0

    MsgBox "V38 stopped safely and changes were rolled back." & vbCrLf & _
           failDescription, vbCritical, "Nebras Warehouse V38"
End Sub

Private Function BuildAutoOccasionFormulaV38(ByVal tableName As String, _
                                             ByVal occName As String, _
                                             ByVal codeName As String, _
                                             ByVal stockName As String, _
                                             ByVal codeCellAddress As String) As String
    Dim t As String
    Dim occN As String
    Dim codeN As String
    Dim stockN As String
    Dim maxStockExpr As String
    Dim positionExpr As String

    t = EscStructV38(tableName)
    occN = EscStructV38(occName)
    codeN = EscStructV38(codeName)
    stockN = EscStructV38(stockName)

    ' Highest current stock among matching product-code rows.
    maxStockExpr = "AGGREGATE(14,6," & _
                   t & "[" & stockN & "]/(" & _
                   t & "[" & codeN & "]=" & codeCellAddress & "),1)"

    ' First row-position whose code matches and whose stock equals
    ' the highest matching stock.
    positionExpr = "AGGREGATE(15,6," & _
                   "(ROW(" & t & "[" & codeN & "])-ROW(INDEX(" & _
                   t & "[" & codeN & "],1,1))+1)/" & _
                   "((" & t & "[" & codeN & "]=" & codeCellAddress & ")*(" & _
                   t & "[" & stockN & "]=" & maxStockExpr & ")),1)"

    BuildAutoOccasionFormulaV38 = _
        "=IF(" & codeCellAddress & "="""","""",IFERROR(INDEX(" & _
        t & "[" & occN & "]," & positionExpr & "),""""))"
End Function

Private Sub ValidateOccasionFormulaV38(ByVal rng As Range)
    Dim checkRows As Variant
    Dim i As Long
    Dim r As Long
    Dim c As Range
    Dim f As String

    If rng Is Nothing Then
        Err.Raise vbObjectError + 3820, , "Internal Occasion range is missing."
    End If

    checkRows = Array(1, _
                      Application.Min(10, rng.Rows.Count), _
                      Application.Min(50, rng.Rows.Count), _
                      rng.Rows.Count)

    For i = LBound(checkRows) To UBound(checkRows)
        r = CLng(checkRows(i))
        If r < 1 Then GoTo NextCheck

        Set c = rng.Cells(r, 1)

        If Not c.HasFormula Then
            Err.Raise vbObjectError + 3821, , _
                      "Missing auto-occasion formula at " & c.Address(False, False) & "."
        End If

        f = CStr(c.Formula)

        If InStr(1, f, "#REF!", vbTextCompare) > 0 Then
            Err.Raise vbObjectError + 3822, , _
                      "Broken reference at " & c.Address(False, False) & "."
        End If

NextCheck:
    Next i
End Sub

Private Function FindTableV38(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0

        If Not lo Is Nothing Then
            Set FindTableV38 = lo
            Exit Function
        End If
    Next ws
End Function

Private Function FindHelperSheetV38() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(1)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    ' Known V37 helper block safety gate.
    If Len(Trim$(CStr(ws.Range("H1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("I1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("J1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("K1").Value))) = 0 Then Exit Function

    Set FindHelperSheetV38 = ws
End Function

Private Function EscStructV38(ByVal s As String) As String
    EscStructV38 = Replace(s, "]", "]]")
End Function
