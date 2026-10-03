Attribute VB_Name = "NebrasPrintReconcileV1"
Option Explicit

' ============================================================
' Nebras unified shortage / reconciliation print module V1.
'
' This is an add-on only.  It never reads or writes stock,
' InventoryBankTable, final registration, output history, or the
' mobile-picking queue.  It only reads the existing shortage print
' form and projects the already-recorded mobile exceptions into that
' same form after a successful server response.
' ============================================================

Private Const NEBRAS_PRINT_URL_V1 As String = "https://script.google.com/macros/s/AKfycbwaz6vrezdBIX0klg4th7_G4dYZIm05AvLHGMmzJ35Pxcoxp_ceshrPlw42cnk0Z1l3/exec"
Private Const NEBRAS_PRINT_KEY_V1 As String = "565bcdf6f75dbfa57dc467be8fd02910ea1c2cdc90592903"
Private Const NEBRAS_PRINT_QUEUE_SHEET_V1 As String = "NebrasPickQueue"
Private Const NEBRAS_PRINT_STATE_SHEET_V1 As String = "NebrasPrintSyncState"
Private Const NEBRAS_PRINT_BUTTON_V1 As String = "btnNebrasUnifiedPrintV1"
Private Const NEBRAS_PRINT_FIRST_ROW_V1 As Long = 6
Private Const NEBRAS_PRINT_MIN_LAST_ROW_V1 As Long = 102
Private Const NEBRAS_PRINT_BUFFER_ROWS_V1 As Long = 20

Private mNebrasPrintRunningV1 As Boolean
Private mNebrasPrintShowMessageV1 As Boolean
Private mNebrasPrintModeV1 As String
Private mNebrasPrintBatchV1 As String
Private mNebrasPrintFingerprintV1 As String
Private mNebrasPrintRenderedFingerprintV1 As String
Private mNebrasPrintSourceRowsV1 As Long
Private mNebrasPrintStartedAtV1 As Date
Private mNebrasPrintPollAtV1 As Date
Private mNebrasPrintHttpV1 As Object

Public Sub Nebras_Install_PrintReconcile_V1()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim anchor As Range

    Set ws = NebrasPrintFindShortageSheetV1()
    If ws Is Nothing Then
        MsgBox "The shortage print sheet was not found. No changes were made.", vbCritical, "Nebras"
        Exit Sub
    End If

    On Error Resume Next
    Set shp = ws.Shapes(NEBRAS_PRINT_BUTTON_V1)
    On Error GoTo 0

    Set anchor = ws.Range("N3")
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchor.Left, anchor.Top, 180, 36)
        shp.Name = NEBRAS_PRINT_BUTTON_V1
    Else
        shp.Left = anchor.Left
        shp.Top = anchor.Top
        shp.Width = 180
        shp.Height = 36
    End If

    With shp
        .OnAction = "Nebras_Sync_UnifiedPrint_V1"
        .TextFrame2.TextRange.Text = "Sync unified print list"
        .TextFrame2.TextRange.Font.Name = "Tahoma"
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Fill.ForeColor.RGB = RGB(92, 72, 168)
        .Line.Visible = msoFalse
        .Visible = msoTrue
    End With

    MsgBox "The unified print sync button is ready on the shortage sheet.", vbInformation, "Nebras"
End Sub

Public Sub Nebras_Sync_UnifiedPrint_V1()
    Dim records As Collection
    Dim fingerprint As String
    Dim previousBatch As String
    Dim previousFingerprint As String
    Dim previousRenderedFingerprint As String
    Dim latestBatch As String
    Dim payload As String
    Dim body As String
    Dim useCurrentShortageRows As Boolean

    If mNebrasPrintRunningV1 Then
        MsgBox "The unified print sync is already running.", vbInformation, "Nebras"
        Exit Sub
    End If

    mNebrasPrintShowMessageV1 = True
    On Error GoTo FailedToStart

    Set records = New Collection
    NebrasPrintReadShortageRowsV1 records, fingerprint
    NebrasPrintReadStateV1 previousBatch, previousFingerprint, previousRenderedFingerprint
    latestBatch = NebrasPrintLatestQueueBatchV1()

    ' After a successful render the form itself contains the unified rows.
    ' Recognize that exact rendered fingerprint on later clicks so it is not
    ' re-imported as a new Excel shortage source.  A newer queue batch or a
    ' changed form is a genuine new source and is imported once.
    useCurrentShortageRows = (records.Count > 0 And _
        (fingerprint <> previousRenderedFingerprint Or _
         (Len(latestBatch) > 0 And StrComp(latestBatch, previousBatch, vbTextCompare) <> 0) Or _
         Len(previousBatch) = 0))

    mNebrasPrintBatchV1 = previousBatch
    If Len(latestBatch) > 0 Then
        If records.Count = 0 Or useCurrentShortageRows Or Len(mNebrasPrintBatchV1) = 0 Then
            mNebrasPrintBatchV1 = latestBatch
        End If
    End If
    If Len(mNebrasPrintBatchV1) = 0 Then mNebrasPrintBatchV1 = "XLS-SHORTAGE-" & Format$(Now, "yyyymmdd-hhnnss")

    If useCurrentShortageRows Then
        mNebrasPrintFingerprintV1 = fingerprint
        mNebrasPrintSourceRowsV1 = records.Count
    Else
        mNebrasPrintFingerprintV1 = previousFingerprint
        mNebrasPrintSourceRowsV1 = 0
    End If
    mNebrasPrintRunningV1 = True
    NebrasPrintSetButtonBusyV1 True

    ' Import is idempotent.  If the same request is retried, the server keeps
    ' one source row per batch/customer/code rather than duplicating print rows.
    If mNebrasPrintSourceRowsV1 > 0 Then
        payload = NebrasPrintBuildImportPayloadV1(mNebrasPrintBatchV1, records)
        body = "action=pickQueuePrintImport"
        body = body & "&key=" & NebrasPrintUrlEncodeV1(NEBRAS_PRINT_KEY_V1)
        body = body & "&payload=" & NebrasPrintUrlEncodeV1(payload)
        NebrasPrintStartHttpV1 "POST", NEBRAS_PRINT_URL_V1, body, "IMPORT"
    Else
        NebrasPrintStartExportV1
    End If
    Exit Sub

FailedToStart:
    NebrasPrintFinishV1 False, "Unified print sync could not start: " & Err.Description
End Sub

Public Sub Nebras_PrintReconcile_Poll_V1()
    Dim finished As Boolean
    Dim response As String
    Dim statusCode As Long
    Dim renderedRows As Long

    If Not mNebrasPrintRunningV1 Then Exit Sub

    If DateDiff("s", mNebrasPrintStartedAtV1, Now) > 35 Then
        On Error Resume Next
        mNebrasPrintHttpV1.Abort
        On Error GoTo 0
        NebrasPrintFinishV1 False, "The server did not answer in time. The shortage sheet was left unchanged."
        Exit Sub
    End If

    On Error Resume Next
    finished = mNebrasPrintHttpV1.WaitForResponse(0)
    If Err.Number <> 0 Then
        Err.Clear
        finished = False
    End If
    On Error GoTo 0

    If Not finished Then
        NebrasPrintSchedulePollV1
        Exit Sub
    End If

    On Error Resume Next
    statusCode = CLng(mNebrasPrintHttpV1.Status)
    response = CStr(mNebrasPrintHttpV1.responseText)
    If Err.Number <> 0 Then
        response = Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    If mNebrasPrintModeV1 = "IMPORT" Then
        If statusCode >= 200 And statusCode < 300 And InStr(1, response, """ok"":true", vbTextCompare) > 0 Then
            On Error GoTo StartExportFailed
            NebrasPrintStartExportV1
        Else
            NebrasPrintFinishV1 False, "The server did not confirm the import. The shortage sheet was left unchanged."
        End If
        Exit Sub
    End If

    If statusCode < 200 Or statusCode >= 300 Then
        NebrasPrintFinishV1 False, "The unified print list could not be received. The shortage sheet was left unchanged."
        Exit Sub
    End If

    On Error GoTo RenderFailed
    renderedRows = NebrasPrintRenderExportV1(response, mNebrasPrintRenderedFingerprintV1)
    If renderedRows = 0 Then
        If mNebrasPrintSourceRowsV1 > 0 Then
            Err.Raise vbObjectError + 7141, , "No print rows were returned for the shortage rows."
        End If
        NebrasPrintFinishV1 True, "There are no open print rows for this batch. The shortage sheet was not cleared."
    Else
        NebrasPrintWriteStateV1 mNebrasPrintBatchV1, mNebrasPrintFingerprintV1, mNebrasPrintRenderedFingerprintV1, renderedRows
        NebrasPrintFinishV1 True, CStr(renderedRows) & " unified print row(s) are ready on the shortage sheet."
    End If
    Exit Sub

RenderFailed:
    NebrasPrintFinishV1 False, "The server answered, but the shortage sheet was left unchanged: " & Err.Description
    Exit Sub

StartExportFailed:
    NebrasPrintFinishV1 False, "The server import was accepted, but the print list could not be requested: " & Err.Description
End Sub

Private Sub NebrasPrintStartExportV1()
    Dim url As String
    url = NEBRAS_PRINT_URL_V1 & "?action=pickQueuePrintExport&key=" & NebrasPrintUrlEncodeV1(NEBRAS_PRINT_KEY_V1) & _
          "&batch=" & NebrasPrintUrlEncodeV1(mNebrasPrintBatchV1) & "&_=" & CStr(CLng(Timer * 100))
    NebrasPrintStartHttpV1 "GET", url, "", "EXPORT"
End Sub

Private Sub NebrasPrintStartHttpV1(ByVal methodName As String, ByVal url As String, _
                                   ByVal body As String, ByVal modeName As String)
    Set mNebrasPrintHttpV1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    mNebrasPrintModeV1 = modeName
    mNebrasPrintStartedAtV1 = Now
    mNebrasPrintHttpV1.SetTimeouts 8000, 8000, 15000, 30000
    mNebrasPrintHttpV1.Open methodName, url, True
    mNebrasPrintHttpV1.SetRequestHeader "Cache-Control", "no-cache"
    If methodName = "POST" Then
        mNebrasPrintHttpV1.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=UTF-8"
        mNebrasPrintHttpV1.Send body
    Else
        mNebrasPrintHttpV1.Send
    End If
    NebrasPrintSchedulePollV1
End Sub

Private Sub NebrasPrintSchedulePollV1()
    mNebrasPrintPollAtV1 = Now + TimeSerial(0, 0, 1)
    Application.OnTime EarliestTime:=mNebrasPrintPollAtV1, Procedure:="Nebras_PrintReconcile_Poll_V1", Schedule:=True
End Sub

Private Sub NebrasPrintFinishV1(ByVal success As Boolean, ByVal messageText As String)
    mNebrasPrintRunningV1 = False
    Set mNebrasPrintHttpV1 = Nothing
    mNebrasPrintModeV1 = ""
    NebrasPrintSetButtonBusyV1 False
    If mNebrasPrintShowMessageV1 Then
        MsgBox messageText, IIf(success, vbInformation, vbExclamation), "Nebras"
    End If
    mNebrasPrintShowMessageV1 = False
End Sub

Private Sub NebrasPrintReadShortageRowsV1(ByRef records As Collection, ByRef fingerprint As String)
    Dim ws As Worksheet
    Dim grouped As Object
    Dim lastRow As Long
    Dim r As Long
    Dim customer As String
    Dim currentCustomer As String
    Dim codeText As String
    Dim qty As Double
    Dim key As String
    Dim rowData As Variant
    Dim item As Variant
    Dim itemKey As Variant

    Set ws = NebrasPrintFindShortageSheetV1()
    If ws Is Nothing Then Err.Raise vbObjectError + 7142, , "The shortage print sheet was not found."

    Set grouped = CreateObject("Scripting.Dictionary")
    grouped.CompareMode = vbTextCompare
    lastRow = NebrasPrintLastShortageRowV1(ws)

    For r = NEBRAS_PRINT_FIRST_ROW_V1 To lastRow
        customer = Trim$(CStr(ws.Cells(r, "A").Value2))
        codeText = Trim$(CStr(ws.Cells(r, "B").Text))
        qty = NebrasPrintNumberV1(ws.Cells(r, "F").Value2)

        If Len(customer) > 0 And Len(codeText) = 0 Then currentCustomer = customer
        If Len(codeText) > 0 And qty > 0 And Len(currentCustomer) > 0 Then
            key = currentCustomer & ChrW$(30) & codeText
            If grouped.Exists(key) Then
                rowData = grouped(key)
                rowData(2) = CDbl(rowData(2)) + qty
                grouped(key) = rowData
            Else
                grouped.Add key, Array(currentCustomer, codeText, qty)
            End If
        End If
    Next r

    For Each itemKey In grouped.Keys
        item = grouped(itemKey)
        records.Add item
    Next itemKey
    fingerprint = NebrasPrintFingerprintV1(records)
End Sub

Private Function NebrasPrintLastShortageRowV1(ByVal ws As Worksheet) As Long
    Dim lastA As Long
    Dim lastB As Long
    Dim lastF As Long
    lastA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    lastF = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    NebrasPrintLastShortageRowV1 = Application.Max(NEBRAS_PRINT_MIN_LAST_ROW_V1, lastA, lastB, lastF)
End Function

Private Function NebrasPrintLatestQueueBatchV1() As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim batchNo As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NEBRAS_PRINT_QUEUE_SHEET_V1)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = 2 To lastRow
        batchNo = Trim$(CStr(ws.Cells(r, 2).Value2))
        If StrComp(batchNo, NebrasPrintLatestQueueBatchV1, vbTextCompare) > 0 Then
            NebrasPrintLatestQueueBatchV1 = batchNo
        End If
    Next r
End Function

Private Function NebrasPrintBuildImportPayloadV1(ByVal batchNo As String, ByVal records As Collection) As String
    Dim i As Long
    Dim item As Variant
    Dim itemsText As String
    Dim qtyText As String

    For i = 1 To records.Count
        item = records(i)
        If Len(itemsText) > 0 Then itemsText = itemsText & ","
        qtyText = Replace$(CStr(CDbl(item(2))), Application.DecimalSeparator, ".")
        itemsText = itemsText & "{""customer"":" & NebrasPrintJsonV1(CStr(item(0))) & _
                    ",""code"":" & NebrasPrintJsonV1(CStr(item(1))) & _
                    ",""qty"":" & qtyText & "}"
    Next i

    NebrasPrintBuildImportPayloadV1 = "{""batch"":" & NebrasPrintJsonV1(batchNo) & ",""items"":[" & itemsText & "]}"
End Function

Private Function NebrasPrintRenderExportV1(ByVal tsv As String, ByRef renderedFingerprint As String) As Long
    Dim ws As Worksheet
    Dim lines As Variant
    Dim fields As Variant
    Dim i As Long
    Dim nextRow As Long
    Dim customer As String
    Dim previousCustomer As String
    Dim codeText As String
    Dim qty As Double
    Dim headerRange As Range
    Dim validRows As Collection
    Dim rowText As String
    Dim requiredLastRow As Long
    Dim clearLastRow As Long
    Dim record As Variant
    Dim oldScreenUpdating As Boolean
    Dim errorText As String

    Set ws = NebrasPrintFindShortageSheetV1()
    If ws Is Nothing Then Err.Raise vbObjectError + 7143, , "The shortage print sheet was not found."

    Set validRows = New Collection
    lines = Split(Replace$(tsv, vbCr, ""), vbLf)
    For i = LBound(lines) + 1 To UBound(lines)
        rowText = CStr(lines(i))
        If Len(rowText) > 0 Then
            fields = Split(rowText, vbTab)
            If UBound(fields) >= 3 Then
                customer = Trim$(CStr(fields(1)))
                codeText = Trim$(CStr(fields(2)))
                qty = NebrasPrintNumberV1(fields(3))
                If Len(customer) > 0 And Len(codeText) > 0 And qty > 0 Then
                    validRows.Add Array(customer, codeText, qty)
                End If
            End If
        End If
    Next i

    If validRows.Count = 0 Then Exit Function
    renderedFingerprint = NebrasPrintFingerprintV1(validRows)

    ' In the worst case every print row starts a new customer section, so
    ' reserve one customer heading plus one code row per item.
    requiredLastRow = NEBRAS_PRINT_FIRST_ROW_V1 + validRows.Count * 2 + NEBRAS_PRINT_BUFFER_ROWS_V1
    clearLastRow = Application.Max(NebrasPrintLastShortageRowV1(ws), requiredLastRow)

    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo RestoreScreen
    Application.DisplayAlerts = False
    ws.Range("A" & NEBRAS_PRINT_FIRST_ROW_V1 & ":N" & clearLastRow).UnMerge
    Application.DisplayAlerts = True
    ws.Range("A7:N7").Copy
    ws.Range("A" & NEBRAS_PRINT_FIRST_ROW_V1 & ":N" & clearLastRow).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    ws.Range("A" & NEBRAS_PRINT_FIRST_ROW_V1 & ":N" & clearLastRow).ClearContents
    ws.Rows(NEBRAS_PRINT_FIRST_ROW_V1 & ":" & clearLastRow).RowHeight = ws.Rows(7).RowHeight

    nextRow = NEBRAS_PRINT_FIRST_ROW_V1
    previousCustomer = ""
    For Each record In validRows
        customer = CStr(record(0))
        If StrComp(customer, previousCustomer, vbTextCompare) <> 0 Then
            Set headerRange = ws.Range("A" & nextRow & ":N" & nextRow)
            headerRange.ClearContents
            headerRange.Merge
            headerRange.Value2 = customer
            headerRange.Interior.Color = RGB(217, 217, 217)
            headerRange.Font.Bold = True
            headerRange.HorizontalAlignment = xlRight
            previousCustomer = customer
            nextRow = nextRow + 1
        End If
        ws.Cells(nextRow, "B").NumberFormat = "@"
        ws.Cells(nextRow, "B").Value2 = CStr(record(1))
        ws.Cells(nextRow, "F").Value2 = CDbl(record(2))
        nextRow = nextRow + 1
        NebrasPrintRenderExportV1 = NebrasPrintRenderExportV1 + 1
    Next record

RestoreScreen:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = oldScreenUpdating
    If Err.Number <> 0 Then
        errorText = Err.Description
        Err.Clear
        Err.Raise vbObjectError + 7144, , errorText
    End If
End Function

Private Function NebrasPrintFindShortageSheetV1() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.CodeName, "Sheet6", vbTextCompare) = 0 Then
            Set NebrasPrintFindShortageSheetV1 = ws
            Exit Function
        End If
    Next ws
End Function

Private Function NebrasPrintStateSheetV1() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NEBRAS_PRINT_STATE_SHEET_V1)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = NEBRAS_PRINT_STATE_SHEET_V1
        ws.Cells(1, 1).Value2 = "Batch"
        ws.Cells(1, 2).Value2 = "SourceFingerprint"
        ws.Cells(1, 3).Value2 = "RenderedFingerprint"
        ws.Cells(1, 4).Value2 = "UpdatedAt"
        ws.Cells(1, 5).Value2 = "RenderedRows"
        ws.Visible = xlSheetVeryHidden
    End If
    Set NebrasPrintStateSheetV1 = ws
End Function

Private Sub NebrasPrintReadStateV1(ByRef batchNo As String, ByRef fingerprint As String, _
                                   ByRef renderedFingerprint As String)
    Dim ws As Worksheet
    Set ws = NebrasPrintStateSheetV1()
    batchNo = Trim$(CStr(ws.Cells(2, 1).Value2))
    fingerprint = Trim$(CStr(ws.Cells(2, 2).Value2))
    renderedFingerprint = Trim$(CStr(ws.Cells(2, 3).Value2))
End Sub

Private Sub NebrasPrintWriteStateV1(ByVal batchNo As String, ByVal fingerprint As String, _
                                    ByVal renderedFingerprint As String, ByVal renderedRows As Long)
    Dim ws As Worksheet
    Set ws = NebrasPrintStateSheetV1()
    ws.Cells(2, 1).Value2 = batchNo
    ws.Cells(2, 2).Value2 = fingerprint
    ws.Cells(2, 3).Value2 = renderedFingerprint
    ws.Cells(2, 4).Value2 = Now
    ws.Cells(2, 5).Value2 = renderedRows
End Sub

Private Function NebrasPrintFingerprintV1(ByVal records As Collection) As String
    Dim i As Long
    Dim item As Variant
    Dim rowHash As Double
    Dim totalHash As Double
    totalHash = 17#
    For i = 1 To records.Count
        item = records(i)
        rowHash = NebrasPrintTextHashV1(CStr(item(0)) & ChrW$(31) & CStr(item(1)) & ChrW$(31) & _
                                        Replace$(CStr(CDbl(item(2))), Application.DecimalSeparator, "."))
        totalHash = totalHash + rowHash
        totalHash = totalHash - Int(totalHash / 2147483629#) * 2147483629#
    Next i
    NebrasPrintFingerprintV1 = CStr(records.Count) & "-" & Format$(totalHash, "0")
End Function

Private Function NebrasPrintTextHashV1(ByVal value As String) As Double
    Dim i As Long
    Dim charCode As Long
    Dim result As Double
    result = 5381#
    For i = 1 To Len(value)
        charCode = AscW(Mid$(value, i, 1))
        If charCode < 0 Then charCode = charCode + 65536
        result = result * 33# + charCode
        result = result - Int(result / 2147483629#) * 2147483629#
    Next i
    NebrasPrintTextHashV1 = result
End Function

Private Function NebrasPrintNumberV1(ByVal value As Variant) As Double
    If IsNumeric(value) Then NebrasPrintNumberV1 = CDbl(value)
End Function

Private Function NebrasPrintJsonV1(ByVal value As String) As String
    value = Replace$(value, "\", "\\")
    value = Replace$(value, Chr$(34), "\" & Chr$(34))
    value = Replace$(value, vbCr, "\r")
    value = Replace$(value, vbLf, "\n")
    NebrasPrintJsonV1 = Chr$(34) & value & Chr$(34)
End Function

Private Function NebrasPrintUrlEncodeV1(ByVal value As String) As String
    Dim stream As Object
    Dim bytes As Variant
    Dim i As Long
    Dim byteValue As Long

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText value
    stream.Position = 0
    stream.Type = 1
    stream.Position = 3
    bytes = stream.Read
    stream.Close

    For i = LBound(bytes) To UBound(bytes)
        byteValue = bytes(i)
        If (byteValue >= 48 And byteValue <= 57) Or _
           (byteValue >= 65 And byteValue <= 90) Or _
           (byteValue >= 97 And byteValue <= 122) Or _
           byteValue = 45 Or byteValue = 46 Or byteValue = 95 Or byteValue = 126 Then
            NebrasPrintUrlEncodeV1 = NebrasPrintUrlEncodeV1 & Chr$(byteValue)
        Else
            NebrasPrintUrlEncodeV1 = NebrasPrintUrlEncodeV1 & "%" & Right$("0" & Hex$(byteValue), 2)
        End If
    Next i
End Function

Private Sub NebrasPrintSetButtonBusyV1(ByVal busy As Boolean)
    Dim ws As Worksheet
    Dim shp As Shape
    Set ws = NebrasPrintFindShortageSheetV1()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Set shp = ws.Shapes(NEBRAS_PRINT_BUTTON_V1)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub
    If busy Then
        shp.TextFrame2.TextRange.Text = "Syncing print list..."
        shp.Fill.ForeColor.RGB = RGB(120, 120, 120)
    Else
        shp.TextFrame2.TextRange.Text = "Sync unified print list"
        shp.Fill.ForeColor.RGB = RGB(92, 72, 168)
    End If
End Sub
