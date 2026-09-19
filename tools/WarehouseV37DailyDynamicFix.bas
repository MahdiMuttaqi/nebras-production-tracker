Attribute VB_Name = "WarehouseV37DailyDynamicFix"
Option Explicit

' ============================================================
' NEBRAS WAREHOUSE - DAILY LOCATION DYNAMIC FIX V37.1
' ASCII-SAFE VBA MODULE
'
' This module intentionally contains NO Persian text literals.
' Reason: VBA .bas imports can corrupt UTF-8 Persian strings on
' some Windows/Office installations.
'
' Surgical scope only:
'   1) Rebuild the Daily Picking location helper H:U.
'   2) Prepare helper rows through Daily sheet row 5000, or more
'      if the current Daily table is already larger.
'   3) Use dynamic InventoryBankTable structured references.
'   4) Hide Excel green error indicators ONLY in Daily table
'      column 11 (auto withdraw quantity).
'
' It does NOT change inventory quantities, bank data, customer
' inputs, finalization macros, mobile sync, V11, production
' tracker, history sheets, or any other VBA module.
' ============================================================

Private Const MIN_DAILY_CAPACITY_ROW_V37 As Long = 5000
Private Const HELPER_FIRST_ROW_V37 As Long = 2

' WarehousePickingTable fixed schema positions.
Private Const DAILY_OCC_COL_V37 As Long = 3
Private Const DAILY_CODE_COL_V37 As Long = 4
Private Const DAILY_AUTO_WITHDRAW_COL_V37 As Long = 11

' InventoryBankTable fixed schema positions.
Private Const BANK_SHELF_COL_V37 As Long = 8
Private Const BANK_ROW_COL_V37 As Long = 9
Private Const BANK_BOX_COL_V37 As Long = 10
Private Const BANK_KEY_COL_V37 As Long = 13

Public Sub Nebras_Install_Daily_Dynamic_Fix_V37()
    Dim bankTbl As ListObject
    Dim dailyTbl As ListObject
    Dim bankWs As Worksheet
    Dim dailyWs As Worksheet
    Dim helperWs As Worksheet

    Dim firstDailyRow As Long
    Dim currentDailyLastRow As Long
    Dim targetDailyLastRow As Long
    Dim helperLastRow As Long
    Dim dailyOffset As Long

    Dim dailyOccCol As Long
    Dim dailyCodeCol As Long

    Dim bankKeyName As String
    Dim bankShelfName As String
    Dim bankRowName As String
    Dim bankBoxName As String

    Dim posCols As Variant
    Dim outCols As Variant
    Dim bankOutputNames As Variant

    Dim oldHelper As Variant
    Dim snapshotTaken As Boolean

    Dim oldCalc As XlCalculation
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldAlerts As Boolean

    Dim k As Long
    Dim f As String
    Dim failDescription As String
    Dim autoCol As ListColumn

    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts

    On Error GoTo Failed

    Set bankTbl = FindTableV37("InventoryBankTable")
    Set dailyTbl = FindTableV37("WarehousePickingTable")

    If bankTbl Is Nothing Then Err.Raise vbObjectError + 3701, , "InventoryBankTable was not found."
    If dailyTbl Is Nothing Then Err.Raise vbObjectError + 3702, , "WarehousePickingTable was not found."

    ' Schema safety gates. Stop before changing anything if a table
    ' structure is not the expected Nebras structure.
    If dailyTbl.ListColumns.Count < DAILY_AUTO_WITHDRAW_COL_V37 Then
        Err.Raise vbObjectError + 3703, , "WarehousePickingTable schema is shorter than expected."
    End If

    If bankTbl.ListColumns.Count < BANK_KEY_COL_V37 Then
        Err.Raise vbObjectError + 3704, , "InventoryBankTable schema is shorter than expected."
    End If

    Set bankWs = bankTbl.Parent
    Set dailyWs = dailyTbl.Parent
    Set helperWs = FindHelperSheetV37()

    If helperWs Is Nothing Then Err.Raise vbObjectError + 3705, , "Daily location helper sheet was not found."

    If dailyTbl.DataBodyRange Is Nothing Then
        firstDailyRow = dailyTbl.HeaderRowRange.Row + 1
        currentDailyLastRow = firstDailyRow
    Else
        firstDailyRow = dailyTbl.DataBodyRange.Row
        currentDailyLastRow = dailyTbl.DataBodyRange.Row + dailyTbl.DataBodyRange.Rows.Count - 1
    End If

    ' Existing Nebras mapping: Daily row 4 -> helper row 2.
    dailyOffset = firstDailyRow - HELPER_FIRST_ROW_V37
    If dailyOffset < 1 Then Err.Raise vbObjectError + 3706, , "Unexpected Daily Picking table position."

    targetDailyLastRow = Application.Max(MIN_DAILY_CAPACITY_ROW_V37, currentDailyLastRow)
    If targetDailyLastRow > dailyWs.Rows.Count Then targetDailyLastRow = dailyWs.Rows.Count

    helperLastRow = targetDailyLastRow - dailyOffset
    If helperLastRow < HELPER_FIRST_ROW_V37 Then
        Err.Raise vbObjectError + 3707, , "Invalid helper range."
    End If

    ' Use schema positions, not Persian header literals.
    dailyOccCol = dailyTbl.ListColumns(DAILY_OCC_COL_V37).Range.Column
    dailyCodeCol = dailyTbl.ListColumns(DAILY_CODE_COL_V37).Range.Column

    ' Read the real workbook header names at runtime. These may be
    ' Persian internally, but no Persian literal is stored in this BAS.
    bankShelfName = bankTbl.ListColumns(BANK_SHELF_COL_V37).Name
    bankRowName = bankTbl.ListColumns(BANK_ROW_COL_V37).Name
    bankBoxName = bankTbl.ListColumns(BANK_BOX_COL_V37).Name
    bankKeyName = bankTbl.ListColumns(BANK_KEY_COL_V37).Name

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' Rollback snapshot: only the helper area touched by this repair.
    oldHelper = helperWs.Range("H" & HELPER_FIRST_ROW_V37 & ":U" & helperLastRow).Formula
    snapshotTaken = True

    posCols = Array("L", "M", "N", "O", "P", "Q", "R", "S", "T", "U")
    outCols = Array("I", "J", "K")
    bankOutputNames = Array(bankShelfName, bankRowName, bankBoxName)

    ' H = occasion|product code
    f = "=IF('" & EscapeSheetV37(dailyWs.Name) & "'!" & _
        ColLetterV37(dailyCodeCol) & (HELPER_FIRST_ROW_V37 + dailyOffset) & "="""","""",'" & _
        EscapeSheetV37(dailyWs.Name) & "'!" & _
        ColLetterV37(dailyOccCol) & (HELPER_FIRST_ROW_V37 + dailyOffset) & _
        "&""|""&'" & EscapeSheetV37(dailyWs.Name) & "'!" & _
        ColLetterV37(dailyCodeCol) & (HELPER_FIRST_ROW_V37 + dailyOffset) & ")"

    helperWs.Range("H" & HELPER_FIRST_ROW_V37).Formula = f
    If helperLastRow > HELPER_FIRST_ROW_V37 Then
        helperWs.Range("H" & HELPER_FIRST_ROW_V37 & ":H" & helperLastRow).FillDown
    End If

    ' L:U = up to 10 matching InventoryBankTable row positions.
    ' Structured references make this dynamic when the bank grows.
    For k = 0 To 9
        f = PositionFormulaV37(bankTbl.Name, bankKeyName, k + 1)
        helperWs.Range(posCols(k) & HELPER_FIRST_ROW_V37).Formula = f
        If helperLastRow > HELPER_FIRST_ROW_V37 Then
            helperWs.Range(posCols(k) & HELPER_FIRST_ROW_V37 & ":" & _
                           posCols(k) & helperLastRow).FillDown
        End If
    Next k

    ' I:K = shelf / warehouse row / box.
    For k = 0 To 2
        f = LocationJoinFormulaV37(bankTbl.Name, CStr(bankOutputNames(k)), _
                                   HELPER_FIRST_ROW_V37, posCols)
        helperWs.Range(outCols(k) & HELPER_FIRST_ROW_V37).Formula = f
        If helperLastRow > HELPER_FIRST_ROW_V37 Then
            helperWs.Range(outCols(k) & HELPER_FIRST_ROW_V37 & ":" & _
                           outCols(k) & helperLastRow).FillDown
        End If
    Next k

    ValidateHelperV37 helperWs, helperLastRow, dailyOffset

    Application.Calculation = oldCalc
    helperWs.Range("H" & HELPER_FIRST_ROW_V37 & ":U" & helperLastRow).Calculate
    If Not dailyTbl.DataBodyRange Is Nothing Then dailyTbl.DataBodyRange.Calculate

    ' Hide green Excel error indicators only in column 11.
    ' This does not change values, formulas, formats, or global settings.
    Set autoCol = dailyTbl.ListColumns(DAILY_AUTO_WITHDRAW_COL_V37)
    If Not autoCol.DataBodyRange Is Nothing Then
        IgnoreErrorIndicatorsV37 autoCol.DataBodyRange
    End If

    ThisWorkbook.Save

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc

    MsgBox "Daily Picking V37.1 installed successfully." & vbCrLf & _
           "Location helper is dynamic and prepared through Daily row " & _
           targetDailyLastRow & "." & vbCrLf & _
           "Green indicators were hidden only in auto-withdraw column.", _
           vbInformation, "Nebras Warehouse V37.1"
    Exit Sub

Failed:
    failDescription = Err.Description

    On Error Resume Next
    If snapshotTaken Then
        helperWs.Range("H" & HELPER_FIRST_ROW_V37 & ":U" & helperLastRow).Formula = oldHelper
    End If

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc
    On Error GoTo 0

    MsgBox "V37.1 stopped safely and helper changes were rolled back." & vbCrLf & _
           failDescription, vbCritical, "Nebras Warehouse V37.1"
End Sub

Private Function PositionFormulaV37(ByVal tableName As String, _
                                    ByVal keyColumnName As String, _
                                    ByVal nthMatch As Long) As String
    Dim t As String
    Dim keyName As String

    t = EscStructV37(tableName)
    keyName = EscStructV37(keyColumnName)

    PositionFormulaV37 = _
        "=IF($H" & HELPER_FIRST_ROW_V37 & "="""","""",IFERROR(AGGREGATE(15,6," & _
        "(ROW(" & t & "[" & keyName & "])-ROW(INDEX(" & t & "[" & keyName & "],1,1))+1)/" & _
        "(" & t & "[" & keyName & "]=$H" & HELPER_FIRST_ROW_V37 & ")," & nthMatch & "),""""))"
End Function

Private Function LocationJoinFormulaV37(ByVal tableName As String, _
                                        ByVal outputColumnName As String, _
                                        ByVal helperRow As Long, _
                                        ByVal posCols As Variant) As String
    Dim k As Long
    Dim s As String
    Dim t As String
    Dim outName As String

    t = EscStructV37(tableName)
    outName = EscStructV37(outputColumnName)

    s = "=IF($H" & helperRow & "="""","""",IFERROR(INDEX(" & _
        t & "[" & outName & "],$" & posCols(0) & helperRow & "),"""")"

    For k = 1 To 9
        s = s & "&IF($" & posCols(k) & helperRow & "="""","""","" | ""&INDEX(" & _
            t & "[" & outName & "],$" & posCols(k) & helperRow & "))"
    Next k

    LocationJoinFormulaV37 = s & ")"
End Function

Private Sub ValidateHelperV37(ByVal ws As Worksheet, _
                              ByVal helperLastRow As Long, _
                              ByVal dailyOffset As Long)
    Dim checkRows As Variant
    Dim i As Long
    Dim r As Long
    Dim c As Range
    Dim expectedDailyRow As Long
    Dim f As String

    checkRows = Array(HELPER_FIRST_ROW_V37, _
                      Application.Min(100, helperLastRow), _
                      Application.Min(500, helperLastRow), _
                      Application.Min(1000, helperLastRow), _
                      helperLastRow)

    For i = LBound(checkRows) To UBound(checkRows)
        r = CLng(checkRows(i))
        If r < HELPER_FIRST_ROW_V37 Then GoTo NextCheck

        expectedDailyRow = r + dailyOffset

        If Not ws.Range("H" & r).HasFormula Then
            Err.Raise vbObjectError + 3720, , "Missing helper formula at H" & r & "."
        End If

        f = CStr(ws.Range("H" & r).Formula)

        If InStr(1, f, "#REF!", vbTextCompare) > 0 Then
            Err.Raise vbObjectError + 3721, , "Broken reference at H" & r & "."
        End If

        If InStr(1, f, CStr(expectedDailyRow), vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 3722, , "Helper/Daily row mapping is invalid at row " & r & "."
        End If

        For Each c In ws.Range("I" & r & ":U" & r).Cells
            If Not c.HasFormula Then
                Err.Raise vbObjectError + 3723, , _
                          "Missing helper formula at " & c.Address(False, False) & "."
            End If

            If InStr(1, CStr(c.Formula), "#REF!", vbTextCompare) > 0 Then
                Err.Raise vbObjectError + 3724, , _
                          "Broken reference at " & c.Address(False, False) & "."
            End If
        Next c

NextCheck:
    Next i
End Sub

Private Sub IgnoreErrorIndicatorsV37(ByVal rng As Range)
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

Private Function FindTableV37(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0

        If Not lo Is Nothing Then
            Set FindTableV37 = lo
            Exit Function
        End If
    Next ws
End Function

Private Function FindHelperSheetV37() As Worksheet
    Dim ws As Worksheet

    ' Known Nebras workbook design: helper block is on worksheet 1.
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(1)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    ' Safety gate: known helper headers H1:K1 must exist.
    If Len(Trim$(CStr(ws.Range("H1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("I1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("J1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("K1").Value))) = 0 Then Exit Function

    Set FindHelperSheetV37 = ws
End Function

Private Function ColLetterV37(ByVal colNumber As Long) As String
    ColLetterV37 = Split(Cells(1, colNumber).Address(False, False), "1")(0)
End Function

Private Function EscapeSheetV37(ByVal s As String) As String
    EscapeSheetV37 = Replace(s, "'", "''")
End Function

Private Function EscStructV37(ByVal s As String) As String
    EscStructV37 = Replace(s, "]", "]]")
End Function
