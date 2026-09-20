Attribute VB_Name = "WarehouseV42ManualDailyLive"
Option Explicit

' ============================================================
' NEBRAS WAREHOUSE - MANUAL DAILY LIVE FIX V42
' ASCII-SAFE VBA MODULE
'
' Surgical purpose:
'   Fix the case where group import shows stock/location but
'   manual code entry does not update the Daily Picking row.
'
' V42 does three narrow things:
'   1) Ensures every CURRENT Daily Picking calculated row has the
'      existing calculated-column formulas.
'   2) Enables normal live table behavior:
'        - Automatic calculation
'        - Events enabled
'        - Excel table formula auto-fill enabled
'      so typing a product code manually recalculates immediately
'      and newly appended table rows inherit formulas.
'   3) Ignores green Excel error indicators ONLY inside the
'      WarehousePickingTable.
'
' V42 does NOT change:
'   - Inventory bank data or quantities
'   - Product/customer/required-quantity inputs
'   - Hidden Occasion values
'   - Finalization logic/button
'   - Bulk-import logic
'   - History
'   - V11/mobile sync
'   - Production tracker
' ============================================================

Private Const DAILY_ROWNO_COL_V42 As Long = 1
Private Const DAILY_CODE_COL_V42 As Long = 4
Private Const DAILY_STOCK_COL_V42 As Long = 6
Private Const DAILY_RESULT_COL_V42 As Long = 12

Public Sub Nebras_Install_Manual_Daily_Live_V42()
    Dim dailyTbl As ListObject
    Dim helperWs As Worksheet

    Dim calcCols As Variant
    Dim colIndex As Variant
    Dim lc As ListColumn
    Dim firstFormulaR1C1 As String

    Dim formulaSnapshots As Object
    Dim oldCalc As XlCalculation
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldAlerts As Boolean
    Dim oldAutoFillLists As Boolean

    Dim helperLastRow As Long
    Dim failDescription As String
    Dim snapshotTaken As Boolean

    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts

    On Error Resume Next
    oldAutoFillLists = Application.AutoCorrect.AutoFillFormulasInLists
    On Error GoTo 0

    On Error GoTo Failed

    Set dailyTbl = FindTableV42("WarehousePickingTable")

    If dailyTbl Is Nothing Then
        Err.Raise vbObjectError + 4201, , "WarehousePickingTable was not found."
    End If

    If dailyTbl.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 4202, , "WarehousePickingTable has no data rows."
    End If

    If dailyTbl.ListColumns.Count < DAILY_RESULT_COL_V42 Then
        Err.Raise vbObjectError + 4203, , "WarehousePickingTable schema is shorter than expected."
    End If

    Set helperWs = FindHelperSheetV42()

    If helperWs Is Nothing Then
        Err.Raise vbObjectError + 4204, , "Daily location helper sheet was not found."
    End If

    calcCols = Array(1, 6, 7, 8, 9, 10, 11, 12)
    Set formulaSnapshots = CreateObject("Scripting.Dictionary")

    ' Snapshot only the calculated columns V42 may repair.
    For Each colIndex In calcCols
        Set lc = dailyTbl.ListColumns(CLng(colIndex))
        formulaSnapshots(CStr(colIndex)) = lc.DataBodyRange.FormulaR1C1
    Next colIndex

    snapshotTaken = True

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Temporarily stop events only while repairing formulas.
    Application.EnableEvents = False

    ' --------------------------------------------------------
    ' Repair calculated columns from their own existing formula.
    ' No new business logic is introduced here.
    ' --------------------------------------------------------
    For Each colIndex In calcCols
        Set lc = dailyTbl.ListColumns(CLng(colIndex))

        firstFormulaR1C1 = FindFormulaR1C1V42(lc.DataBodyRange)

        If Len(firstFormulaR1C1) = 0 Then
            Err.Raise vbObjectError + 4210 + CLng(colIndex), , _
                      "Existing calculated formula was not found in Daily column " & CStr(colIndex) & "."
        End If

        lc.DataBodyRange.FormulaR1C1 = firstFormulaR1C1
    Next colIndex

    ' Product codes must remain text when entered manually.
    dailyTbl.ListColumns(DAILY_CODE_COL_V42).DataBodyRange.NumberFormat = "@"

    ' --------------------------------------------------------
    ' These are the normal Excel settings required for live
    ' manual entry. Group import works without them because it
    ' explicitly recalculates after import.
    ' --------------------------------------------------------
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    On Error Resume Next
    Application.AutoCorrect.AutoFillFormulasInLists = True
    On Error GoTo Failed

    ' Calculate only the relevant dependency chain first.
    helperLastRow = LastHelperFormulaRowV42(helperWs)

    If helperLastRow >= 2 Then
        helperWs.Range("H2:U" & helperLastRow).Calculate
    End If

    dailyTbl.DataBodyRange.Calculate

    ' One full dependency refresh after the repair. This is
    ' one-time during installation, not a permanent heavy mode.
    Application.CalculateFull

    ' Remove green indicators only inside Daily Picking table.
    IgnoreDailyErrorIndicatorsV42 dailyTbl.DataBodyRange

    ValidateLiveDailyV42 dailyTbl

    ThisWorkbook.Save

    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts

    ' Intentionally keep these NORMAL live-entry settings:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    On Error Resume Next
    Application.AutoCorrect.AutoFillFormulasInLists = True
    On Error GoTo 0

    MsgBox "Daily Picking V42 installed successfully." & vbCrLf & _
           "Manual product-code entry will now recalculate stock and locations live." & vbCrLf & _
           "New table rows will inherit calculated formulas automatically." & vbCrLf & _
           "Green indicators were ignored only inside the Daily Picking table.", _
           vbInformation, "Nebras Warehouse V42"
    Exit Sub

Failed:
    failDescription = Err.Description

    On Error Resume Next

    If snapshotTaken Then
        For Each colIndex In calcCols
            dailyTbl.ListColumns(CLng(colIndex)).DataBodyRange.FormulaR1C1 = _
                formulaSnapshots(CStr(colIndex))
        Next colIndex
    End If

    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.AutoCorrect.AutoFillFormulasInLists = oldAutoFillLists

    On Error GoTo 0

    MsgBox "V42 stopped safely and Daily formulas were returned to their previous state." & vbCrLf & _
           failDescription, vbCritical, "Nebras Warehouse V42"
End Sub

Private Function FindFormulaR1C1V42(ByVal rng As Range) As String
    Dim c As Range

    If rng Is Nothing Then Exit Function

    For Each c In rng.Cells
        If c.HasFormula Then
            If InStr(1, CStr(c.Formula), "#REF!", vbTextCompare) = 0 Then
                FindFormulaR1C1V42 = CStr(c.FormulaR1C1)
                Exit Function
            End If
        End If
    Next c
End Function

Private Sub ValidateLiveDailyV42(ByVal dailyTbl As ListObject)
    Dim calcCols As Variant
    Dim colIndex As Variant
    Dim lc As ListColumn
    Dim firstCell As Range
    Dim lastCell As Range

    calcCols = Array(1, 6, 7, 8, 9, 10, 11, 12)

    For Each colIndex In calcCols
        Set lc = dailyTbl.ListColumns(CLng(colIndex))
        Set firstCell = lc.DataBodyRange.Cells(1, 1)
        Set lastCell = lc.DataBodyRange.Cells(lc.DataBodyRange.Rows.Count, 1)

        If Not firstCell.HasFormula Then
            Err.Raise vbObjectError + 4230, , _
                      "Calculated formula is missing at " & firstCell.Address(False, False) & "."
        End If

        If Not lastCell.HasFormula Then
            Err.Raise vbObjectError + 4231, , _
                      "Calculated formula is missing at " & lastCell.Address(False, False) & "."
        End If

        If InStr(1, CStr(firstCell.Formula), "#REF!", vbTextCompare) > 0 Or _
           InStr(1, CStr(lastCell.Formula), "#REF!", vbTextCompare) > 0 Then

            Err.Raise vbObjectError + 4232, , _
                      "Broken calculated reference was found in Daily column " & CStr(colIndex) & "."
        End If
    Next colIndex

    If Application.Calculation <> xlCalculationAutomatic Then
        Err.Raise vbObjectError + 4233, , "Excel calculation did not switch to Automatic."
    End If

    If Application.EnableEvents = False Then
        Err.Raise vbObjectError + 4234, , "Excel events are still disabled."
    End If
End Sub

Private Sub IgnoreDailyErrorIndicatorsV42(ByVal rng As Range)
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

Private Function LastHelperFormulaRowV42(ByVal ws As Worksheet) As Long
    Dim col As Long
    Dim r As Long
    Dim lastFound As Long

    For col = 8 To 21
        r = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

        If r > lastFound Then lastFound = r
    Next col

    LastHelperFormulaRowV42 = lastFound
End Function

Private Function FindTableV42(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing

        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0

        If Not lo Is Nothing Then
            Set FindTableV42 = lo
            Exit Function
        End If
    Next ws
End Function

Private Function FindHelperSheetV42() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(1)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    If Len(Trim$(CStr(ws.Range("H1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("I1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("J1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("K1").Value))) = 0 Then Exit Function

    Set FindHelperSheetV42 = ws
End Function
