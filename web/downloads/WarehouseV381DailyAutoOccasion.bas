Attribute VB_Name = "WarehouseV381DailyAutoOccasion"
Option Explicit

' ============================================================
' NEBRAS WAREHOUSE - DAILY PICKING AUTO OCCASION V38.1
' ASCII-SAFE VBA MODULE
'
' Purpose:
'   - User no longer needs to enter Occasion in Daily Picking.
'   - The existing Occasion table column is kept internally and
'     hidden, so older finalization/history logic is not broken.
'   - Occasion is calculated from the entered product code by a
'     small VBA worksheet function.
'
' Duplicate-code rule:
'   If the same product code exists under multiple occasions,
'   the bank row with the HIGHEST current stock is selected.
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

Private Const DAILY_OCC_COL_V381 As Long = 3
Private Const DAILY_CODE_COL_V381 As Long = 4

Private Const BANK_OCC_COL_V381 As Long = 1
Private Const BANK_CODE_COL_V381 As Long = 2
Private Const BANK_STOCK_COL_V381 As Long = 7

Private Const HELPER_FIRST_ROW_V381 As Long = 2
Private Const HELPER_TARGET_DAILY_ROW_V381 As Long = 5000

Public Sub Nebras_Install_Daily_AutoOccasion_V381()
    Dim dailyTbl As ListObject
    Dim bankTbl As ListObject
    Dim helperWs As Worksheet

    Dim occCol As ListColumn
    Dim occRange As Range

    Dim firstDailyRow As Long
    Dim helperLastRow As Long
    Dim dailyOffset As Long

    Dim oldOccFormulas As Variant
    Dim oldHidden As Boolean
    Dim snapshotTaken As Boolean

    Dim oldCalc As XlCalculation
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldAlerts As Boolean

    Dim failDescription As String
    Dim formulaText As String

    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts

    On Error GoTo Failed

    Set dailyTbl = FindTableV381("WarehousePickingTable")
    Set bankTbl = FindTableV381("InventoryBankTable")

    If dailyTbl Is Nothing Then Err.Raise vbObjectError + 3811, , "WarehousePickingTable was not found."
    If bankTbl Is Nothing Then Err.Raise vbObjectError + 3812, , "InventoryBankTable was not found."

    If dailyTbl.ListColumns.Count < DAILY_CODE_COL_V381 Then
        Err.Raise vbObjectError + 3813, , "WarehousePickingTable schema is shorter than expected."
    End If

    If bankTbl.ListColumns.Count < BANK_STOCK_COL_V381 Then
        Err.Raise vbObjectError + 3814, , "InventoryBankTable schema is shorter than expected."
    End If

    If dailyTbl.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 3815, , "WarehousePickingTable has no data rows."
    End If

    If bankTbl.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 3816, , "InventoryBankTable has no data rows."
    End If

    Set helperWs = FindHelperSheetV381()
    If helperWs Is Nothing Then
        Err.Raise vbObjectError + 3817, , "V37 Daily location helper sheet was not found."
    End If

    Set occCol = dailyTbl.ListColumns(DAILY_OCC_COL_V381)
    Set occRange = occCol.DataBodyRange

    firstDailyRow = dailyTbl.DataBodyRange.Row
    dailyOffset = firstDailyRow - HELPER_FIRST_ROW_V381

    If dailyOffset < 1 Then
        Err.Raise vbObjectError + 3818, , "Unexpected Daily Picking table position."
    End If

    helperLastRow = HELPER_TARGET_DAILY_ROW_V381 - dailyOffset
    If helperLastRow < HELPER_FIRST_ROW_V381 Then
        Err.Raise vbObjectError + 3819, , "Invalid helper calculation range."
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' Rollback snapshot: only the internal Occasion column.
    oldOccFormulas = occRange.Formula
    oldHidden = occCol.Range.EntireColumn.Hidden
    snapshotTaken = True

    ' Simple and robust R1C1 formula:
    ' column 3 looks one cell to the right for product code.
    ' The UDF below resolves the occasion from InventoryBankTable.
    formulaText = "=IF(RC[1]="""","""",NebrasDailyOccasionV381(RC[1]))"
    occRange.FormulaR1C1 = formulaText

    ' Hide, never delete, to preserve the existing table schema.
    occCol.Range.EntireColumn.Hidden = True

    ValidateOccasionFormulaV381 occRange

    Application.Calculation = oldCalc

    ' Calculate only the relevant chain.
    occRange.Calculate
    helperWs.Range("H" & HELPER_FIRST_ROW_V381 & ":U" & helperLastRow).Calculate
    dailyTbl.DataBodyRange.Calculate

    ThisWorkbook.Save

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc

    MsgBox "Daily Picking V38.1 installed successfully." & vbCrLf & _
           "Occasion is automatic and hidden." & vbCrLf & _
           "Enter customer, product code and required quantity only.", _
           vbInformation, "Nebras Warehouse V38.1"
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

    MsgBox "V38.1 stopped safely and changes were rolled back." & vbCrLf & _
           failDescription, vbCritical, "Nebras Warehouse V38.1"
End Sub

Public Function NebrasDailyOccasionV381(ByVal productCode As Variant) As Variant
    Dim bankTbl As ListObject
    Dim data As Variant
    Dim i As Long

    Dim wanted As String
    Dim rowCode As String
    Dim rowOccasion As Variant
    Dim rowStock As Double

    Dim bestOccasion As Variant
    Dim bestStock As Double
    Dim found As Boolean

    On Error GoTo SafeBlank

    Application.Volatile True

    wanted = Trim$(CStr(productCode))
    If Len(wanted) = 0 Then
        NebrasDailyOccasionV381 = vbNullString
        Exit Function
    End If

    Set bankTbl = FindTableV381("InventoryBankTable")
    If bankTbl Is Nothing Then GoTo SafeBlank
    If bankTbl.DataBodyRange Is Nothing Then GoTo SafeBlank
    If bankTbl.ListColumns.Count < BANK_STOCK_COL_V381 Then GoTo SafeBlank

    data = bankTbl.DataBodyRange.Value2

    bestStock = -1E+308
    found = False

    For i = 1 To UBound(data, 1)
        rowCode = Trim$(CStr(data(i, BANK_CODE_COL_V381)))

        If StrComp(rowCode, wanted, vbTextCompare) = 0 Then
            rowOccasion = data(i, BANK_OCC_COL_V381)

            If IsNumeric(data(i, BANK_STOCK_COL_V381)) Then
                rowStock = CDbl(data(i, BANK_STOCK_COL_V381))
            Else
                rowStock = 0
            End If

            If (Not found) Or rowStock > bestStock Then
                bestStock = rowStock
                bestOccasion = rowOccasion
                found = True
            End If
        End If
    Next i

    If found Then
        NebrasDailyOccasionV381 = bestOccasion
    Else
        NebrasDailyOccasionV381 = vbNullString
    End If

    Exit Function

SafeBlank:
    NebrasDailyOccasionV381 = vbNullString
End Function

Private Sub ValidateOccasionFormulaV381(ByVal rng As Range)
    Dim checkRows As Variant
    Dim i As Long
    Dim r As Long
    Dim c As Range
    Dim f As String

    If rng Is Nothing Then
        Err.Raise vbObjectError + 3830, , "Internal Occasion range is missing."
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
            Err.Raise vbObjectError + 3831, , _
                      "Missing auto-occasion formula at " & c.Address(False, False) & "."
        End If

        f = CStr(c.FormulaR1C1)

        If InStr(1, f, "#REF!", vbTextCompare) > 0 Then
            Err.Raise vbObjectError + 3832, , _
                      "Broken reference at " & c.Address(False, False) & "."
        End If

        If InStr(1, f, "NebrasDailyOccasionV381", vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 3833, , _
                      "Unexpected Occasion formula at " & c.Address(False, False) & "."
        End If

NextCheck:
    Next i
End Sub

Private Function FindTableV381(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0

        If Not lo Is Nothing Then
            Set FindTableV381 = lo
            Exit Function
        End If
    Next ws
End Function

Private Function FindHelperSheetV381() As Worksheet
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

    Set FindHelperSheetV381 = ws
End Function
