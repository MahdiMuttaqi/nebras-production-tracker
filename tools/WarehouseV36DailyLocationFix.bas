Attribute VB_Name = "WarehouseV36DailyLocationFix"
Option Explicit

' Nebras warehouse - surgical repair for Daily Picking location helper only.
' Scope: first worksheet helper range H:U only. No other sheet data or VBA module is changed.

Public Sub Nebras_Repair_Daily_Locations_V36()
    Dim bankTbl As ListObject
    Dim dailyTbl As ListObject
    Dim bankWs As Worksheet
    Dim dailyWs As Worksheet
    Dim helperWs As Worksheet
    Dim firstDailyRow As Long
    Dim lastDailyRow As Long
    Dim helperFirstRow As Long
    Dim helperLastRow As Long
    Dim lastBankRow As Long
    Dim r As Long, k As Long
    Dim posCols As Variant
    Dim outCols As Variant
    Dim bankCols As Variant
    Dim oldFormulas As Variant
    Dim formulaText As String
    Dim c As Range
    Dim oldCalc As XlCalculation
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldAlerts As Boolean

    On Error GoTo Failed

    Set bankTbl = FindTableV36("InventoryBankTable")
    Set dailyTbl = FindTableV36("WarehousePickingTable")

    If bankTbl Is Nothing Then Err.Raise vbObjectError + 3601, , "InventoryBankTable was not found."
    If dailyTbl Is Nothing Then Err.Raise vbObjectError + 3602, , "WarehousePickingTable was not found."
    If dailyTbl.DataBodyRange Is Nothing Then Err.Raise vbObjectError + 3603, , "WarehousePickingTable has no data rows."

    Set bankWs = bankTbl.Parent
    Set dailyWs = dailyTbl.Parent
    Set helperWs = FindHelperSheetV36()

    If helperWs Is Nothing Then Err.Raise vbObjectError + 3604, , "Daily location helper sheet was not found."

    firstDailyRow = dailyTbl.DataBodyRange.Row
    lastDailyRow = firstDailyRow + dailyTbl.DataBodyRange.Rows.Count - 1

    ' Existing workbook design maps Daily row 4 -> helper row 2.
    helperFirstRow = firstDailyRow - 2
    helperLastRow = lastDailyRow - 2

    If helperFirstRow < 2 Then Err.Raise vbObjectError + 3605, , "Unexpected Daily Picking table position."
    If helperLastRow < helperFirstRow Then Err.Raise vbObjectError + 3606, , "Invalid helper row range."

    ' Match the current Daily Picking formulas' forward-capacity range.
    lastBankRow = Application.Max(2957, bankTbl.Range.Row + bankTbl.ListRows.Count + 500)
    If lastBankRow > bankWs.Rows.Count Then lastBankRow = bankWs.Rows.Count

    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' Keep a rollback snapshot. Only this range is ever changed.
    oldFormulas = helperWs.Range("H" & helperFirstRow & ":U" & helperLastRow).Formula

    posCols = Array("L", "M", "N", "O", "P", "Q", "R", "S", "T", "U")
    outCols = Array("I", "J", "K")
    bankCols = Array("H", "I", "J")

    For r = helperFirstRow To helperLastRow
        helperWs.Range("H" & r).Formula = _
            "=IF('" & EscapeSheetV36(dailyWs.Name) & "'!D" & (r + 2) & "="""","""",'" & _
            EscapeSheetV36(dailyWs.Name) & "'!C" & (r + 2) & "&""|""&'" & _
            EscapeSheetV36(dailyWs.Name) & "'!D" & (r + 2) & ")"

        For k = 0 To 9
            helperWs.Range(posCols(k) & r).Formula = _
                "=IF($H" & r & "="""","""",IFERROR(AGGREGATE(15,6,(ROW('" & _
                EscapeSheetV36(bankWs.Name) & "'!$M$4:$M$" & lastBankRow & ")-3)/('" & _
                EscapeSheetV36(bankWs.Name) & "'!$M$4:$M$" & lastBankRow & "=$H" & r & ")," & _
                (k + 1) & "),""""))"
        Next k

        For k = 0 To 2
            formulaText = LocationJoinFormulaV36(bankWs.Name, bankCols(k), r, posCols, lastBankRow)
            helperWs.Range(outCols(k) & r).Formula = formulaText
        Next k
    Next r

    ' Formula-structure validation before saving.
    For Each c In helperWs.Range("H" & helperFirstRow & ":U" & helperLastRow).Cells
        If Not c.HasFormula Then
            Err.Raise vbObjectError + 3607, , "A helper formula is missing at " & c.Address(False, False) & "."
        End If
        If InStr(1, c.Formula, "#REF!", vbTextCompare) > 0 Then
            Err.Raise vbObjectError + 3608, , "A broken reference remains at " & c.Address(False, False) & "."
        End If
    Next c

    Application.Calculation = oldCalc
    helperWs.Range("H" & helperFirstRow & ":U" & helperLastRow).Calculate
    dailyTbl.Range.Calculate

    ThisWorkbook.Save

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc

    MsgBox "Daily location helper repaired safely." & vbCrLf & _
           "Helper rows repaired: " & helperFirstRow & " to " & helperLastRow & vbCrLf & _
           "No other worksheet data or VBA module was changed.", vbInformation, "Nebras Warehouse"
    Exit Sub

Failed:
    On Error Resume Next
    If Not helperWs Is Nothing Then
        If helperLastRow >= helperFirstRow And helperFirstRow >= 2 Then
            helperWs.Range("H" & helperFirstRow & ":U" & helperLastRow).Formula = oldFormulas
        End If
    End If
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc
    On Error GoTo 0

    MsgBox "Daily location repair stopped and rolled back." & vbCrLf & Err.Description, _
           vbCritical, "Nebras Warehouse"
End Sub

Private Function FindTableV36(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0

        If Not lo Is Nothing Then
            Set FindTableV36 = lo
            Exit Function
        End If
    Next ws
End Function

Private Function FindHelperSheetV36() As Worksheet
    Dim ws As Worksheet

    ' The original Nebras warehouse design stores the helper block on worksheet 1.
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(1)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    ' Safety check: H1:K1 must be the existing helper header block (four non-empty cells).
    If Len(Trim$(CStr(ws.Range("H1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("I1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("J1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("K1").Value))) = 0 Then Exit Function

    Set FindHelperSheetV36 = ws
End Function

Private Function LocationJoinFormulaV36(ByVal bankName As String, _
                                        ByVal bankCol As String, _
                                        ByVal helperRow As Long, _
                                        ByVal posCols As Variant, _
                                        ByVal lastBankRow As Long) As String
    Dim k As Long
    Dim s As String

    s = "=IF($H" & helperRow & "="""","""",IFERROR(INDEX('" & _
        EscapeSheetV36(bankName) & "'!$" & bankCol & "$4:$" & bankCol & "$" & lastBankRow & _
        ",$" & posCols(0) & helperRow & "),"""")"

    For k = 1 To 9
        s = s & "&IF($" & posCols(k) & helperRow & "="""","""","" | ""&INDEX('" & _
            EscapeSheetV36(bankName) & "'!$" & bankCol & "$4:$" & bankCol & "$" & lastBankRow & _
            ",$" & posCols(k) & helperRow & "))"
    Next k

    LocationJoinFormulaV36 = s & ")"
End Function

Private Function EscapeSheetV36(ByVal s As String) As String
    EscapeSheetV36 = Replace(s, "'", "''")
End Function
