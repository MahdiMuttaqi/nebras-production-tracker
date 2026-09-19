Attribute VB_Name = "WarehouseV40DailyCapacity5000"
Option Explicit

' ============================================================
' NEBRAS WAREHOUSE - DAILY PICKING CAPACITY V40
' ASCII-SAFE VBA MODULE
'
' Purpose:
'   Expand WarehousePickingTable safely through worksheet row 5000.
'   This fixes bulk-import failures caused by the table itself
'   having too few empty rows.
'
' Important:
'   V37/V39 already prepared the location helper through row 5000.
'   V40 expands the ACTUAL Daily Picking table to the same limit.
'
' Safety:
'   - No existing data rows are deleted or moved.
'   - Input columns are not overwritten.
'   - Only calculated columns are extended into NEW rows.
'   - Existing Occasion column remains hidden if it was hidden.
'   - If validation fails, the table is resized back.
'
' Not changed:
'   inventory bank data, quantities, finalization logic, history,
'   V11/mobile sync, production tracker, buttons, or customer data.
' ============================================================

Private Const TARGET_DAILY_SHEET_ROW_V40 As Long = 5000
Private Const DAILY_ROWNO_COL_V40 As Long = 1
Private Const DAILY_FIRST_CALC_COL_V40 As Long = 6
Private Const DAILY_LAST_CALC_COL_V40 As Long = 12
Private Const DAILY_AUTO_WITHDRAW_COL_V40 As Long = 11

Public Sub Nebras_Install_Daily_Capacity_V40()
    Dim dailyTbl As ListObject
    Dim dailyWs As Worksheet
    Dim helperWs As Worksheet

    Dim oldTableRange As Range
    Dim newTableRange As Range
    Dim oldLastRow As Long
    Dim targetLastRow As Long
    Dim helperRowForTarget As Long
    Dim dailyOffset As Long

    Dim formulaR1C1(1 To 13) As String
    Dim hasFormula(1 To 13) As Boolean

    Dim oldCalc As XlCalculation
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldAlerts As Boolean

    Dim resized As Boolean
    Dim c As Long
    Dim failDescription As String

    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts

    On Error GoTo Failed

    Set dailyTbl = FindTableV40("WarehousePickingTable")
    If dailyTbl Is Nothing Then
        Err.Raise vbObjectError + 4001, , "WarehousePickingTable was not found."
    End If

    If dailyTbl.ShowTotals Then
        Err.Raise vbObjectError + 4002, , "Daily Picking table has a totals row. V40 stopped without changes."
    End If

    If dailyTbl.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 4003, , "WarehousePickingTable has no data rows."
    End If

    If dailyTbl.ListColumns.Count < DAILY_LAST_CALC_COL_V40 Then
        Err.Raise vbObjectError + 4004, , "WarehousePickingTable schema is shorter than expected."
    End If

    Set dailyWs = dailyTbl.Parent
    Set helperWs = FindHelperSheetV40()

    If helperWs Is Nothing Then
        Err.Raise vbObjectError + 4005, , "V37/V39 helper sheet was not found."
    End If

    oldLastRow = dailyTbl.DataBodyRange.Row + dailyTbl.DataBodyRange.Rows.Count - 1
    targetLastRow = TARGET_DAILY_SHEET_ROW_V40

    If oldLastRow >= targetLastRow Then
        IgnoreAutoWithdrawIndicatorsV40 dailyTbl.ListColumns(DAILY_AUTO_WITHDRAW_COL_V40).DataBodyRange
        MsgBox "Daily Picking table already reaches row " & oldLastRow & "." & vbCrLf & _
               "No capacity change was required.", _
               vbInformation, "Nebras Warehouse V40"
        Exit Sub
    End If

    dailyOffset = dailyTbl.DataBodyRange.Row - 2
    helperRowForTarget = targetLastRow - dailyOffset

    ' Safety gate: V37/V39 helper must already cover the target row.
    If helperRowForTarget < 2 Then
        Err.Raise vbObjectError + 4006, , "Invalid Daily/helper row mapping."
    End If

    If Not helperWs.Range("H" & helperRowForTarget).HasFormula Then
        Err.Raise vbObjectError + 4007, , _
                  "Location helper does not yet cover Daily row 5000. V37/V39 must remain installed."
    End If

    If InStr(1, CStr(helperWs.Range("H" & helperRowForTarget).Formula), "#REF!", vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 4008, , "Location helper has a broken reference at the target row."
    End If

    ' Snapshot the formulas used by calculated columns before resize.
    For c = 1 To dailyTbl.ListColumns.Count
        If c = DAILY_ROWNO_COL_V40 Or _
           (c >= DAILY_FIRST_CALC_COL_V40 And c <= DAILY_LAST_CALC_COL_V40) Then

            If dailyTbl.ListColumns(c).DataBodyRange.Cells(1, 1).HasFormula Then
                formulaR1C1(c) = CStr(dailyTbl.ListColumns(c).DataBodyRange.Cells(1, 1).FormulaR1C1)
                hasFormula(c) = True
            Else
                Err.Raise vbObjectError + 4009, , _
                          "Expected calculated formula is missing in Daily column " & c & "."
            End If
        End If
    Next c

    Set oldTableRange = dailyTbl.Range

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set newTableRange = dailyWs.Range( _
        dailyWs.Cells(dailyTbl.HeaderRowRange.Row, dailyTbl.Range.Column), _
        dailyWs.Cells(targetLastRow, dailyTbl.Range.Column + dailyTbl.Range.Columns.Count - 1))

    dailyTbl.Resize newTableRange
    resized = True

    ' Extend ONLY calculated columns. Input columns remain untouched.
    For c = 1 To dailyTbl.ListColumns.Count
        If hasFormula(c) Then
            dailyTbl.ListColumns(c).DataBodyRange.FormulaR1C1 = formulaR1C1(c)
        End If
    Next c

    ' Recalculate only the Daily Picking table.
    Application.Calculation = oldCalc
    dailyTbl.DataBodyRange.Calculate

    ' Hide green indicators only in auto-withdraw calculated column.
    IgnoreAutoWithdrawIndicatorsV40 dailyTbl.ListColumns(DAILY_AUTO_WITHDRAW_COL_V40).DataBodyRange

    ValidateCapacityV40 dailyTbl, targetLastRow

    ThisWorkbook.Save

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc

    MsgBox "Daily Picking capacity V40 installed successfully." & vbCrLf & _
           "The actual table now reaches worksheet row 5000." & vbCrLf & _
           "Bulk import is no longer limited to the old small set of empty rows." & vbCrLf & _
           "Existing data and finalization logic were not changed.", _
           vbInformation, "Nebras Warehouse V40"
    Exit Sub

Failed:
    failDescription = Err.Description

    On Error Resume Next

    If resized Then
        dailyTbl.Resize oldTableRange
    End If

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc

    On Error GoTo 0

    MsgBox "V40 stopped safely and the table was returned to its previous size." & vbCrLf & _
           failDescription, vbCritical, "Nebras Warehouse V40"
End Sub

Private Sub ValidateCapacityV40(ByVal dailyTbl As ListObject, ByVal targetLastRow As Long)
    Dim actualLastRow As Long
    Dim c As Long

    If dailyTbl Is Nothing Then
        Err.Raise vbObjectError + 4020, , "Daily table disappeared during validation."
    End If

    If dailyTbl.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 4021, , "Daily table data range disappeared during validation."
    End If

    actualLastRow = dailyTbl.DataBodyRange.Row + dailyTbl.DataBodyRange.Rows.Count - 1

    If actualLastRow <> targetLastRow Then
        Err.Raise vbObjectError + 4022, , _
                  "Daily table did not reach the requested target row."
    End If

    For c = 1 To dailyTbl.ListColumns.Count
        If c = DAILY_ROWNO_COL_V40 Or _
           (c >= DAILY_FIRST_CALC_COL_V40 And c <= DAILY_LAST_CALC_COL_V40) Then

            If Not dailyTbl.ListColumns(c).DataBodyRange.Cells( _
                dailyTbl.ListColumns(c).DataBodyRange.Rows.Count, 1).HasFormula Then

                Err.Raise vbObjectError + 4023, , _
                          "Calculated formula did not extend through Daily column " & c & "."
            End If

            If InStr(1, CStr(dailyTbl.ListColumns(c).DataBodyRange.Cells( _
                dailyTbl.ListColumns(c).DataBodyRange.Rows.Count, 1).Formula), _
                "#REF!", vbTextCompare) > 0 Then

                Err.Raise vbObjectError + 4024, , _
                          "Broken reference found in Daily column " & c & "."
            End If
        End If
    Next c
End Sub

Private Sub IgnoreAutoWithdrawIndicatorsV40(ByVal rng As Range)
    Dim cell As Range
    Dim errorType As Long

    If rng Is Nothing Then Exit Sub

    For Each cell In rng.Cells
        For errorType = 1 To 9
            On Error Resume Next
            If cell.Errors(errorType).Value Then
                cell.Errors(errorType).Ignore = True
            End If
            Err.Clear
            On Error GoTo 0
        Next errorType
    Next cell
End Sub

Private Function FindTableV40(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing

        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0

        If Not lo Is Nothing Then
            Set FindTableV40 = lo
            Exit Function
        End If
    Next ws
End Function

Private Function FindHelperSheetV40() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(1)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    If Len(Trim$(CStr(ws.Range("H1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("I1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("J1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("K1").Value))) = 0 Then Exit Function

    Set FindHelperSheetV40 = ws
End Function
