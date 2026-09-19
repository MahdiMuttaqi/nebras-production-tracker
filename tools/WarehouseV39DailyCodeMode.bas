Attribute VB_Name = "WarehouseV39DailyCodeMode"
Option Explicit

' ============================================================
' NEBRAS WAREHOUSE - DAILY PICKING CODE-ONLY MODE V39
' ASCII-SAFE VBA MODULE
'
' Goal:
'   The user does NOT need to know or enter Occasion in the
'   Daily Picking sheet.
'
' Safety design:
'   - Occasion column is HIDDEN, not deleted.
'   - Existing finalization logic is preserved.
'   - Existing finalizer already fills a blank Occasion from the
'     first matching product-code row in InventoryBankTable.
'   - V39 uses that SAME "first matching product code" rule for
'     displayed stock and locations, so preview and finalization
'     stay consistent.
'
' Scope changed by this installer:
'   1) Hide Daily Picking internal Occasion column (column 3).
'   2) Daily current-stock formula (column 6): use the selected
'      internal key derived from product code only.
'   3) Daily location display (columns 7:9): read helper outputs.
'   4) Daily location-count formula (column 10): count helper hits.
'   5) Helper H:U: derive Occasion|Code automatically from the
'      first matching product-code row in InventoryBankTable,
'      then locate rows dynamically.
'
' NOT changed:
'   - Occasion cell values themselves
'   - product codes
'   - requested quantities
'   - auto-withdraw formula (column 11)
'   - result formula (column 12)
'   - notes
'   - inventory quantities
'   - finalization button/macro
'   - history
'   - V11/mobile sync
'   - production tracker
'
' Capacity:
'   Helper is prepared through Daily row 5000.
' ============================================================

Private Const DAILY_OCC_COL_V39 As Long = 3
Private Const DAILY_CODE_COL_V39 As Long = 4
Private Const DAILY_STOCK_COL_V39 As Long = 6
Private Const DAILY_SHELF_COL_V39 As Long = 7
Private Const DAILY_ROW_COL_V39 As Long = 8
Private Const DAILY_BOX_COL_V39 As Long = 9
Private Const DAILY_LOCATION_COUNT_COL_V39 As Long = 10

Private Const BANK_OCC_COL_V39 As Long = 1
Private Const BANK_CODE_COL_V39 As Long = 2
Private Const BANK_STOCK_COL_V39 As Long = 7
Private Const BANK_SHELF_COL_V39 As Long = 8
Private Const BANK_ROW_COL_V39 As Long = 9
Private Const BANK_BOX_COL_V39 As Long = 10
Private Const BANK_KEY_COL_V39 As Long = 13

Private Const HELPER_FIRST_ROW_V39 As Long = 2
Private Const TARGET_DAILY_ROW_V39 As Long = 5000

Public Sub Nebras_Install_Daily_CodeOnly_V39()
    Dim dailyTbl As ListObject
    Dim bankTbl As ListObject
    Dim dailyWs As Worksheet
    Dim helperWs As Worksheet

    Dim firstDailyRow As Long
    Dim dailyOffset As Long
    Dim helperLastRow As Long

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

    Dim oldHelper As Variant
    Dim oldDaily6To10 As Variant
    Dim oldOccHidden As Boolean
    Dim snapshotTaken As Boolean

    Dim oldCalc As XlCalculation
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldAlerts As Boolean

    Dim k As Long
    Dim f As String
    Dim failDescription As String

    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts

    On Error GoTo Failed

    Set dailyTbl = FindTableV39("WarehousePickingTable")
    Set bankTbl = FindTableV39("InventoryBankTable")

    If dailyTbl Is Nothing Then Err.Raise vbObjectError + 3901, , "WarehousePickingTable was not found."
    If bankTbl Is Nothing Then Err.Raise vbObjectError + 3902, , "InventoryBankTable was not found."

    If dailyTbl.ListColumns.Count < DAILY_LOCATION_COUNT_COL_V39 Then
        Err.Raise vbObjectError + 3903, , "WarehousePickingTable schema is shorter than expected."
    End If

    If bankTbl.ListColumns.Count < BANK_KEY_COL_V39 Then
        Err.Raise vbObjectError + 3904, , "InventoryBankTable schema is shorter than expected."
    End If

    If dailyTbl.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 3905, , "WarehousePickingTable has no data rows."
    End If

    If bankTbl.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 3906, , "InventoryBankTable has no data rows."
    End If

    Set dailyWs = dailyTbl.Parent
    Set helperWs = FindHelperSheetV39()

    If helperWs Is Nothing Then
        Err.Raise vbObjectError + 3907, , "V37 helper sheet was not found."
    End If

    firstDailyRow = dailyTbl.DataBodyRange.Row
    dailyOffset = firstDailyRow - HELPER_FIRST_ROW_V39

    If dailyOffset < 1 Then
        Err.Raise vbObjectError + 3908, , "Unexpected Daily Picking table position."
    End If

    helperLastRow = TARGET_DAILY_ROW_V39 - dailyOffset

    If helperLastRow < HELPER_FIRST_ROW_V39 Then
        Err.Raise vbObjectError + 3909, , "Invalid helper range."
    End If

    codeSheetCol = dailyTbl.ListColumns(DAILY_CODE_COL_V39).Range.Column

    bankOccName = bankTbl.ListColumns(BANK_OCC_COL_V39).Name
    bankCodeName = bankTbl.ListColumns(BANK_CODE_COL_V39).Name
    bankStockName = bankTbl.ListColumns(BANK_STOCK_COL_V39).Name
    bankShelfName = bankTbl.ListColumns(BANK_SHELF_COL_V39).Name
    bankRowName = bankTbl.ListColumns(BANK_ROW_COL_V39).Name
    bankBoxName = bankTbl.ListColumns(BANK_BOX_COL_V39).Name
    bankKeyName = bankTbl.ListColumns(BANK_KEY_COL_V39).Name

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' Rollback snapshots: only ranges V39 changes.
    oldHelper = helperWs.Range("H" & HELPER_FIRST_ROW_V39 & ":U" & helperLastRow).Formula
    oldDaily6To10 = dailyTbl.DataBodyRange.Columns(DAILY_STOCK_COL_V39).Resize(, 5).Formula
    oldOccHidden = dailyTbl.ListColumns(DAILY_OCC_COL_V39).Range.EntireColumn.Hidden
    snapshotTaken = True

    posCols = Array("L", "M", "N", "O", "P", "Q", "R", "S", "T", "U")
    outCols = Array("I", "J", "K")
    outputNames = Array(bankShelfName, bankRowName, bankBoxName)

    ' --------------------------------------------------------
    ' Helper H:
    '   product code -> first matching bank Occasion -> Occasion|Code
    ' This matches the workbook's existing finalizer behavior.
    ' --------------------------------------------------------
    f = HelperKeyFormulaV39( _
        dailyWs.Name, _
        codeSheetCol, _
        firstDailyRow, _
        bankTbl.Name, _
        bankOccName, _
        bankCodeName)

    helperWs.Range("H" & HELPER_FIRST_ROW_V39).Formula = f
    If helperLastRow > HELPER_FIRST_ROW_V39 Then
        helperWs.Range("H" & HELPER_FIRST_ROW_V39 & ":H" & helperLastRow).FillDown
    End If

    ' L:U = up to 10 matching bank positions using dynamic table refs.
    For k = 0 To 9
        f = PositionFormulaV39(bankTbl.Name, bankKeyName, k + 1)
        helperWs.Range(posCols(k) & HELPER_FIRST_ROW_V39).Formula = f

        If helperLastRow > HELPER_FIRST_ROW_V39 Then
            helperWs.Range(posCols(k) & HELPER_FIRST_ROW_V39 & ":" & _
                           posCols(k) & helperLastRow).FillDown
        End If
    Next k

    ' I:K = shelf / warehouse row / box.
    For k = 0 To 2
        f = LocationJoinFormulaV39( _
            bankTbl.Name, CStr(outputNames(k)), HELPER_FIRST_ROW_V39, posCols)

        helperWs.Range(outCols(k) & HELPER_FIRST_ROW_V39).Formula = f

        If helperLastRow > HELPER_FIRST_ROW_V39 Then
            helperWs.Range(outCols(k) & HELPER_FIRST_ROW_V39 & ":" & _
                           outCols(k) & helperLastRow).FillDown
        End If
    Next k

    ' --------------------------------------------------------
    ' Daily column 6 = stock of the SAME internal Occasion|Code
    ' selected in helper H.
    ' --------------------------------------------------------
    f = DailyStockFormulaV39( _
        dailyWs, firstDailyRow, codeSheetCol, helperWs.Name, _
        HELPER_FIRST_ROW_V39, bankTbl.Name, bankKeyName, bankStockName)

    dailyTbl.ListColumns(DAILY_STOCK_COL_V39).DataBodyRange.Cells(1, 1).Formula = f
    If dailyTbl.ListRows.Count > 1 Then
        dailyTbl.ListColumns(DAILY_STOCK_COL_V39).DataBodyRange.FillDown
    End If

    ' Daily columns 7:9 directly read the matching helper row.
    f = DailyHelperValueFormulaV39(dailyWs, firstDailyRow, codeSheetCol, helperWs.Name, "I", HELPER_FIRST_ROW_V39)
    dailyTbl.ListColumns(DAILY_SHELF_COL_V39).DataBodyRange.Cells(1, 1).Formula = f
    If dailyTbl.ListRows.Count > 1 Then dailyTbl.ListColumns(DAILY_SHELF_COL_V39).DataBodyRange.FillDown

    f = DailyHelperValueFormulaV39(dailyWs, firstDailyRow, codeSheetCol, helperWs.Name, "J", HELPER_FIRST_ROW_V39)
    dailyTbl.ListColumns(DAILY_ROW_COL_V39).DataBodyRange.Cells(1, 1).Formula = f
    If dailyTbl.ListRows.Count > 1 Then dailyTbl.ListColumns(DAILY_ROW_COL_V39).DataBodyRange.FillDown

    f = DailyHelperValueFormulaV39(dailyWs, firstDailyRow, codeSheetCol, helperWs.Name, "K", HELPER_FIRST_ROW_V39)
    dailyTbl.ListColumns(DAILY_BOX_COL_V39).DataBodyRange.Cells(1, 1).Formula = f
    If dailyTbl.ListRows.Count > 1 Then dailyTbl.ListColumns(DAILY_BOX_COL_V39).DataBodyRange.FillDown

    ' Daily column 10 = number of matching bank locations.
    f = DailyLocationCountFormulaV39( _
        dailyWs, firstDailyRow, codeSheetCol, helperWs.Name, _
        HELPER_FIRST_ROW_V39)

    dailyTbl.ListColumns(DAILY_LOCATION_COUNT_COL_V39).DataBodyRange.Cells(1, 1).Formula = f
    If dailyTbl.ListRows.Count > 1 Then
        dailyTbl.ListColumns(DAILY_LOCATION_COUNT_COL_V39).DataBodyRange.FillDown
    End If

    ' Hide Occasion. Do not delete it, because existing finalization
    ' still uses column 3 internally.
    dailyTbl.ListColumns(DAILY_OCC_COL_V39).Range.EntireColumn.Hidden = True

    ValidateV39 helperWs, helperLastRow, dailyTbl

    Application.Calculation = oldCalc

    helperWs.Range("H" & HELPER_FIRST_ROW_V39 & ":U" & helperLastRow).Calculate
    dailyTbl.DataBodyRange.Calculate

    ThisWorkbook.Save

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc

    MsgBox "Daily Picking V39 installed successfully." & vbCrLf & _
           "Occasion is hidden and no manual Occasion entry is required." & vbCrLf & _
           "Stock and locations now resolve from product code only." & vbCrLf & _
           "Existing finalization logic was not changed.", _
           vbInformation, "Nebras Warehouse V39"
    Exit Sub

Failed:
    failDescription = Err.Description

    On Error Resume Next
    If snapshotTaken Then
        helperWs.Range("H" & HELPER_FIRST_ROW_V39 & ":U" & helperLastRow).Formula = oldHelper
        dailyTbl.DataBodyRange.Columns(DAILY_STOCK_COL_V39).Resize(, 5).Formula = oldDaily6To10
        dailyTbl.ListColumns(DAILY_OCC_COL_V39).Range.EntireColumn.Hidden = oldOccHidden
    End If

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc
    On Error GoTo 0

    MsgBox "V39 stopped safely and all V39 changes were rolled back." & vbCrLf & _
           failDescription, vbCritical, "Nebras Warehouse V39"
End Sub

Private Function HelperKeyFormulaV39(ByVal dailySheetName As String, _
                                     ByVal codeSheetCol As Long, _
                                     ByVal firstDailyRow As Long, _
                                     ByVal tableName As String, _
                                     ByVal occName As String, _
                                     ByVal codeName As String) As String
    Dim codeAddress As String
    Dim t As String
    Dim o As String
    Dim c As String

    codeAddress = "'" & EscapeSheetV39(dailySheetName) & "'!" & _
                  ColLetterV39(codeSheetCol) & firstDailyRow

    t = EscStructV39(tableName)
    o = EscStructV39(occName)
    c = EscStructV39(codeName)

    HelperKeyFormulaV39 = _
        "=IF(" & codeAddress & "="""","""",IFERROR(" & _
        "INDEX(" & t & "[" & o & "],MATCH(" & codeAddress & "," & _
        t & "[" & c & "],0))&""|""&" & codeAddress & ",""""))"
End Function

Private Function PositionFormulaV39(ByVal tableName As String, _
                                    ByVal keyColumnName As String, _
                                    ByVal nthMatch As Long) As String
    Dim t As String
    Dim k As String

    t = EscStructV39(tableName)
    k = EscStructV39(keyColumnName)

    PositionFormulaV39 = _
        "=IF($H" & HELPER_FIRST_ROW_V39 & "="""","""",IFERROR(AGGREGATE(15,6," & _
        "(ROW(" & t & "[" & k & "])-ROW(INDEX(" & t & "[" & k & "],1,1))+1)/" & _
        "(" & t & "[" & k & "]=$H" & HELPER_FIRST_ROW_V39 & ")," & _
        nthMatch & "),""""))"
End Function

Private Function LocationJoinFormulaV39(ByVal tableName As String, _
                                        ByVal outputColumnName As String, _
                                        ByVal helperRow As Long, _
                                        ByVal posCols As Variant) As String
    Dim k As Long
    Dim s As String
    Dim t As String
    Dim outName As String

    t = EscStructV39(tableName)
    outName = EscStructV39(outputColumnName)

    s = "=IF($H" & helperRow & "="""","""",IFERROR(INDEX(" & _
        t & "[" & outName & "],$" & posCols(0) & helperRow & "),"""")"

    For k = 1 To 9
        s = s & "&IF($" & posCols(k) & helperRow & "="""","""","" | ""&INDEX(" & _
            t & "[" & outName & "],$" & posCols(k) & helperRow & "))"
    Next k

    LocationJoinFormulaV39 = s & ")"
End Function

Private Function DailyStockFormulaV39(ByVal dailyWs As Worksheet, _
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

    codeAddress = ColLetterV39(codeSheetCol) & firstDailyRow
    helperKeyAddress = "'" & EscapeSheetV39(helperSheetName) & "'!H" & helperRow

    t = EscStructV39(tableName)
    k = EscStructV39(keyName)
    s = EscStructV39(stockName)

    DailyStockFormulaV39 = _
        "=IF(" & codeAddress & "="""","""",IF(" & helperKeyAddress & "="""",0," & _
        "SUMIF(" & t & "[" & k & "]," & helperKeyAddress & "," & _
        t & "[" & s & "])))"
End Function

Private Function DailyHelperValueFormulaV39(ByVal dailyWs As Worksheet, _
                                            ByVal firstDailyRow As Long, _
                                            ByVal codeSheetCol As Long, _
                                            ByVal helperSheetName As String, _
                                            ByVal helperCol As String, _
                                            ByVal helperRow As Long) As String
    Dim codeAddress As String

    codeAddress = ColLetterV39(codeSheetCol) & firstDailyRow

    DailyHelperValueFormulaV39 = _
        "=IF(" & codeAddress & "="""","""",IFERROR('" & _
        EscapeSheetV39(helperSheetName) & "'!" & helperCol & helperRow & ",""""))"
End Function

Private Function DailyLocationCountFormulaV39(ByVal dailyWs As Worksheet, _
                                              ByVal firstDailyRow As Long, _
                                              ByVal codeSheetCol As Long, _
                                              ByVal helperSheetName As String, _
                                              ByVal helperRow As Long) As String
    Dim codeAddress As String

    codeAddress = ColLetterV39(codeSheetCol) & firstDailyRow

    DailyLocationCountFormulaV39 = _
        "=IF(" & codeAddress & "="""","""",COUNT('" & _
        EscapeSheetV39(helperSheetName) & "'!L" & helperRow & ":U" & helperRow & "))"
End Function

Private Sub ValidateV39(ByVal helperWs As Worksheet, _
                        ByVal helperLastRow As Long, _
                        ByVal dailyTbl As ListObject)
    Dim c As Range
    Dim checkRows As Variant
    Dim i As Long
    Dim r As Long

    checkRows = Array(HELPER_FIRST_ROW_V39, _
                      Application.Min(100, helperLastRow), _
                      Application.Min(1000, helperLastRow), _
                      helperLastRow)

    For i = LBound(checkRows) To UBound(checkRows)
        r = CLng(checkRows(i))

        For Each c In helperWs.Range("H" & r & ":U" & r).Cells
            If Not c.HasFormula Then
                Err.Raise vbObjectError + 3920, , _
                          "Missing helper formula at " & c.Address(False, False) & "."
            End If

            If InStr(1, CStr(c.Formula), "#REF!", vbTextCompare) > 0 Then
                Err.Raise vbObjectError + 3921, , _
                          "Broken reference at " & c.Address(False, False) & "."
            End If
        Next c
    Next i

    For i = DAILY_STOCK_COL_V39 To DAILY_LOCATION_COUNT_COL_V39
        Set c = dailyTbl.ListColumns(i).DataBodyRange.Cells(1, 1)

        If Not c.HasFormula Then
            Err.Raise vbObjectError + 3922, , _
                      "Missing Daily formula at " & c.Address(False, False) & "."
        End If

        If InStr(1, CStr(c.Formula), "#REF!", vbTextCompare) > 0 Then
            Err.Raise vbObjectError + 3923, , _
                      "Broken Daily formula at " & c.Address(False, False) & "."
        End If
    Next i
End Sub

Private Function FindTableV39(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0

        If Not lo Is Nothing Then
            Set FindTableV39 = lo
            Exit Function
        End If
    Next ws
End Function

Private Function FindHelperSheetV39() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(1)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    If Len(Trim$(CStr(ws.Range("H1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("I1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("J1").Value))) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("K1").Value))) = 0 Then Exit Function

    Set FindHelperSheetV39 = ws
End Function

Private Function ColLetterV39(ByVal colNumber As Long) As String
    ColLetterV39 = Split(Cells(1, colNumber).Address(False, False), "1")(0)
End Function

Private Function EscapeSheetV39(ByVal s As String) As String
    EscapeSheetV39 = Replace(s, "'", "''")
End Function

Private Function EscStructV39(ByVal s As String) As String
    EscStructV39 = Replace(s, "]", "]]")
End Function
