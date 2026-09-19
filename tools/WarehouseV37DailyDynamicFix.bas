Attribute VB_Name = "WarehouseV37DailyDynamicFix"
Option Explicit

' ============================================================
' NEBRAS WAREHOUSE - DAILY LOCATION DYNAMIC FIX V37
'
' Surgical scope only:
'   1) Rebuilds the hidden Daily Picking location helper H:U
'      on "راهنما و تنظیمات" with dynamic references to InventoryBankTable.
'   2) Prepares helper rows through Daily sheet row 5000 (or farther if the
'      current WarehousePickingTable is already larger), so future Daily rows
'      do not lose shelf/row/box merely because the old helper ended.
'   3) Removes Excel green error indicators ONLY from the calculated
'      "تعداد خودکار قابل برداشت" column.
'
' It does NOT change:
'   - inventory quantities
'   - bank data
'   - customer/order input columns
'   - finalization macros/buttons
'   - mobile sync / V11
'   - production tracker
'   - output/history sheets
'
' Safe behavior:
'   - snapshots only H:U helper range before writing
'   - rolls H:U back if validation fails
'   - preserves Excel calculation/events/screen state
' ============================================================

Private Const MIN_DAILY_CAPACITY_ROW_V37 As Long = 5000
Private Const HELPER_FIRST_ROW_V37 As Long = 2
Private Const HELPER_FIRST_COL_V37 As String = "H"
Private Const HELPER_LAST_COL_V37 As String = "U"

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
    Dim changedIndicatorCount As Long

    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts

    On Error GoTo Failed

    Set bankTbl = FindTableV37("InventoryBankTable")
    Set dailyTbl = FindTableV37("WarehousePickingTable")

    If bankTbl Is Nothing Then Err.Raise vbObjectError + 3701, , "InventoryBankTable was not found."
    If dailyTbl Is Nothing Then Err.Raise vbObjectError + 3702, , "WarehousePickingTable was not found."

    Set bankWs = bankTbl.Parent
    Set dailyWs = dailyTbl.Parent
    Set helperWs = FindHelperSheetV37()

    If helperWs Is Nothing Then Err.Raise vbObjectError + 3703, , "Daily location helper sheet was not found."

    If dailyTbl.DataBodyRange Is Nothing Then
        firstDailyRow = dailyTbl.HeaderRowRange.Row + 1
        currentDailyLastRow = firstDailyRow
    Else
        firstDailyRow = dailyTbl.DataBodyRange.Row
        currentDailyLastRow = dailyTbl.DataBodyRange.Row + dailyTbl.DataBodyRange.Rows.Count - 1
    End If

    ' Existing Nebras layout maps Daily row 4 to helper row 2.
    dailyOffset = firstDailyRow - HELPER_FIRST_ROW_V37
    If dailyOffset < 1 Then Err.Raise vbObjectError + 3704, , "Unexpected Daily Picking table position."

    ' Permanent practical capacity: never tied to the current number of filled rows.
    ' If the Daily table is already larger than row 5000, honor the larger table.
    targetDailyLastRow = Application.Max(MIN_DAILY_CAPACITY_ROW_V37, currentDailyLastRow)
    If targetDailyLastRow > dailyWs.Rows.Count Then targetDailyLastRow = dailyWs.Rows.Count

    helperLastRow = targetDailyLastRow - dailyOffset
    If helperLastRow < HELPER_FIRST_ROW_V37 Then Err.Raise vbObjectError + 3705, , "Invalid helper range."

    dailyOccCol = FindListColumnV37(dailyTbl, "مناسبت").Range.Column
    dailyCodeCol = FindListColumnV37(dailyTbl, "کد محصول").Range.Column

    If dailyOccCol = 0 Then Err.Raise vbObjectError + 3706, , "Daily column 'مناسبت' was not found."
    If dailyCodeCol = 0 Then Err.Raise vbObjectError + 3707, , "Daily column 'کد محصول' was not found."

    bankKeyName = FindListColumnV37(bankTbl, "کلید لوکیشن").Name
    bankShelfName = FindListColumnV37(bankTbl, "قفسه").Name
    bankRowName = FindListColumnV37(bankTbl, "ردیف").Name
    bankBoxName = FindListColumnV37(bankTbl, "جعبه").Name

    If Len(bankKeyName) = 0 Then Err.Raise vbObjectError + 3708, , "Bank column 'کلید لوکیشن' was not found."
    If Len(bankShelfName) = 0 Then Err.Raise vbObjectError + 3709, , "Bank column 'قفسه' was not found."
    If Len(bankRowName) = 0 Then Err.Raise vbObjectError + 3710, , "Bank column 'ردیف' was not found."
    If Len(bankBoxName) = 0 Then Err.Raise vbObjectError + 3711, , "Bank column 'جعبه' was not found."

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' Rollback snapshot: ONLY the helper area touched by this repair.
    oldHelper = helperWs.Range(HELPER_FIRST_COL_V37 & HELPER_FIRST_ROW_V37 & ":" & _
                               HELPER_LAST_COL_V37 & helperLastRow).Formula
    snapshotTaken = True

    posCols = Array("L", "M", "N", "O", "P", "Q", "R", "S", "T", "U")
    outCols = Array("I", "J", "K")
    bankOutputNames = Array(bankShelfName, bankRowName, bankBoxName)

    ' --------------------------------------------------------
    ' H: key = occasion|product code
    ' Write first helper row once, then FillDown so row mapping
    ' remains consistent and fast.
    ' --------------------------------------------------------
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

    ' --------------------------------------------------------
    ' L:U = matching row positions inside InventoryBankTable.
    ' Uses STRUCTURED REFERENCES, so adding rows to the bank
    ' does not require changing a fixed row number ever again.
    ' --------------------------------------------------------
    For k = 0 To 9
        f = PositionFormulaV37(bankTbl.Name, bankKeyName, posCols(k), k + 1)
        helperWs.Range(posCols(k) & HELPER_FIRST_ROW_V37).Formula = f
        If helperLastRow > HELPER_FIRST_ROW_V37 Then
            helperWs.Range(posCols(k) & HELPER_FIRST_ROW_V37 & ":" & posCols(k) & helperLastRow).FillDown
        End If
    Next k

    ' --------------------------------------------------------
    ' I:K = shelf / warehouse row / box list.
    ' Also uses structured references to InventoryBankTable.
    ' --------------------------------------------------------
    For k = 0 To 2
        f = LocationJoinFormulaV37(bankTbl.Name, CStr(bankOutputNames(k)), HELPER_FIRST_ROW_V37, posCols)
        helperWs.Range(outCols(k) & HELPER_FIRST_ROW_V37).Formula = f
        If helperLastRow > HELPER_FIRST_ROW_V37 Then
            helperWs.Range(outCols(k) & HELPER_FIRST_ROW_V37 & ":" & outCols(k) & helperLastRow).FillDown
        End If
    Next k

    ' --------------------------------------------------------
    ' Validate only the helper architecture we changed.
    ' --------------------------------------------------------
    ValidateHelperV37 helperWs, helperLastRow, dailyOffset

    ' Restore calculation before final recalc.
    Application.Calculation = oldCalc

    helperWs.Range("H" & HELPER_FIRST_ROW_V37 & ":U" & helperLastRow).Calculate
    If Not dailyTbl.DataBodyRange Is Nothing Then dailyTbl.DataBodyRange.Calculate

    ' --------------------------------------------------------
    ' Remove the green Excel error triangles ONLY from
    ' "تعداد خودکار قابل برداشت".
    ' No value/formula/format is changed.
    ' --------------------------------------------------------
    Set autoCol = FindListColumnV37(dailyTbl, "تعداد خودکار قابل برداشت")
    If Not autoCol Is Nothing Then
        changedIndicatorCount = IgnoreErrorIndicatorsV37(autoCol.DataBodyRange)
    End If

    ThisWorkbook.Save

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc

    MsgBox "اصلاح داینامیک لیست برداشت با موفقیت انجام شد." & vbCrLf & vbCrLf & _
           "• لوکیشن‌ها تا ردیف " & targetDailyLastRow & " آماده هستند." & vbCrLf & _
           "• اضافه‌شدن ردیف‌های جدید به بانک به‌صورت داینامیک در فرمول‌ها دیده می‌شود." & vbCrLf & _
           "• علامت‌های سبز فقط از ستون «تعداد خودکار قابل برداشت» مخفی شدند." & vbCrLf & _
           "• موجودی، سوابق، همگام‌سازی و سایر ماژول‌ها تغییر نکردند.", _
           vbInformation, "Nebras Warehouse V37"
    Exit Sub

Failed:
    failDescription = Err.Description

    On Error Resume Next
    If snapshotTaken Then
        helperWs.Range(HELPER_FIRST_COL_V37 & HELPER_FIRST_ROW_V37 & ":" & _
                       HELPER_LAST_COL_V37 & helperLastRow).Formula = oldHelper
    End If

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc
    On Error GoTo 0

    MsgBox "اصلاح متوقف شد و بخش کمکی به حالت قبل برگشت." & vbCrLf & _
           failDescription, vbCritical, "Nebras Warehouse V37"
End Sub

Private Function PositionFormulaV37(ByVal tableName As String, _
                                    ByVal keyColumnName As String, _
                                    ByVal posCol As String, _
                                    ByVal nthMatch As Long) As String
    Dim t As String, k As String

    t = EscStructV37(tableName)
    k = EscStructV37(keyColumnName)

    PositionFormulaV37 = _
        "=IF($H" & HELPER_FIRST_ROW_V37 & "="""","""",IFERROR(AGGREGATE(15,6," & _
        "(ROW(" & t & "[" & k & "])-ROW(INDEX(" & t & "[" & k & "],1,1))+1)/" & _
        "(" & t & "[" & k & "]=$H" & HELPER_FIRST_ROW_V37 & ")," & nthMatch & "),""""))"
End Function

Private Function LocationJoinFormulaV37(ByVal tableName As String, _
                                        ByVal outputColumnName As String, _
                                        ByVal helperRow As Long, _
                                        ByVal posCols As Variant) As String
    Dim k As Long
    Dim s As String
    Dim t As String, outName As String

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
    Dim i As Long, r As Long
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
                Err.Raise vbObjectError + 3723, , "Missing helper formula at " & c.Address(False, False) & "."
            End If

            If InStr(1, CStr(c.Formula), "#REF!", vbTextCompare) > 0 Then
                Err.Raise vbObjectError + 3724, , "Broken reference at " & c.Address(False, False) & "."
            End If
        Next c
NextCheck:
    Next i
End Sub

Private Function IgnoreErrorIndicatorsV37(ByVal rng As Range) As Long
    Dim c As Range
    Dim errorType As Long
    Dim changed As Long

    If rng Is Nothing Then Exit Function

    For Each c In rng.Cells
        For errorType = 1 To 9
            On Error Resume Next
            If c.Errors(errorType).Value Then
                c.Errors(errorType).Ignore = True
                If Err.Number = 0 Then changed = changed + 1
            End If
            Err.Clear
            On Error GoTo 0
        Next errorType
    Next c

    IgnoreErrorIndicatorsV37 = changed
End Function

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

Private Function FindListColumnV37(ByVal lo As ListObject, _
                                   ByVal headerName As String) As ListColumn
    Dim lc As ListColumn

    If lo Is Nothing Then Exit Function

    For Each lc In lo.ListColumns
        If Trim$(CStr(lc.Name)) = Trim$(headerName) Then
            Set FindListColumnV37 = lc
            Exit Function
        End If
    Next lc
End Function

Private Function FindHelperSheetV37() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("راهنما و تنظیمات")
    On Error GoTo 0

    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(1)
        On Error GoTo 0
    End If

    If ws Is Nothing Then Exit Function

    ' Safety gate: this must be the known Nebras helper block.
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
    ' Excel structured-reference escaping for a closing bracket.
    EscStructV37 = Replace(s, "]", "]]")
End Function
