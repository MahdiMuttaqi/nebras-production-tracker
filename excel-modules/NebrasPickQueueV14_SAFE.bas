Attribute VB_Name = "NebrasPickQueueV13"
Option Explicit

' ============================================================
' Nebras Pick Queue V14 - resilient, idempotent mobile dispatch
' Independent mobile picking queue only.
' It NEVER changes InventoryBankTable quantities.
' Final stock deduction remains inside the existing finalizer.
' ============================================================

Private Const PICK_URL As String = "https://script.google.com/macros/s/AKfycbwaz6vrezdBIX0klg4th7_G4dYZIm05AvLHGMmzJ35Pxcoxp_ceshrPlw42cnk0Z1l3/exec"
Private Const PICK_KEY As String = "565bcdf6f75dbfa57dc467be8fd02910ea1c2cdc90592903"
Private Const QUEUE_SHEET As String = "NebrasPickQueue"
Private Const STATUS_STAGED As String = "STAGED"
Private Const STATUS_PENDING As String = "PENDING"
Private Const STATUS_SENDING As String = "SENDING"
Private Const STATUS_SENT As String = "SENT"
Private Const RESPONSE_LIMIT As Long = 500
Private Const MAX_AUTO_ATTEMPTS As Long = 3

Private mBatch As String
Private mCustomers As Object
Private mLineNo As Long

Private mSendRunning As Boolean
Private mSendShowMessage As Boolean
Private mSendBatch As String
Private mSendOrderIds As Object
Private mSendHttp As Object
Private mSendStartedAt As Date
Private mSendPollAt As Date
Private mSendMode As String
Private mSendAttemptId As String
Private mSendAttemptNo As Long
Private mRetryBatch As String
Private mRetryAt As Date
Private mDeferredBatch As String

Public Sub Nebras_Install_PickQueue_V13()
    Dim ws As Worksheet
    Dim historyTbl As ListObject
    Dim pickingTbl As ListObject
    Dim shp As Shape
    Dim anchor As Range

    Set ws = NebrasPickEnsureSheet()
    ws.Visible = xlSheetVeryHidden

    Set historyTbl = NebrasPickFindTable("WarehouseHistoryTable")
    If historyTbl Is Nothing Then
        MsgBox "ÃœÊ· ”Ê«»ﬁ «‰»«— ÅÌœ« ‰‘œ. ÂÌç  €ÌÌ—Ì «‰Ã«„ ‰‘œ.", vbCritical, "’› ”›«—‘ ‰»—«”"
        Exit Sub
    End If

    Set pickingTbl = NebrasPickFindTable("WarehousePickingTable")
    If Not pickingTbl Is Nothing Then
        On Error Resume Next
        pickingTbl.Parent.Shapes("btnNebrasSendPickQueueV13").Delete
        On Error GoTo 0
    End If

    On Error Resume Next
    Set shp = historyTbl.Parent.Shapes("btnNebrasSendPickQueueV13")
    On Error GoTo 0

    Set anchor = historyTbl.Range.Cells(1, historyTbl.Range.Columns.Count).Offset(0, 2)

    If shp Is Nothing Then
        Set shp = historyTbl.Parent.Shapes.AddShape(msoShapeRoundedRectangle, _
                  anchor.Left, anchor.Top, 205, 42)
        shp.Name = "btnNebrasSendPickQueueV13"
    Else
        shp.Left = anchor.Left
        shp.Top = anchor.Top
        shp.Width = 205
        shp.Height = 42
    End If

    With shp
        .OnAction = "Nebras_Send_PickQueue_V13"
        .TextFrame2.TextRange.Text = "«—”«· ¬Œ—Ì‰ ”›«—‘ ‰«„Ê›ﬁ »Â „Ê»«Ì·"
        .TextFrame2.TextRange.Font.Name = "Tahoma"
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Fill.ForeColor.RGB = RGB(92, 72, 168)
        .Line.Visible = msoFalse
        .Visible = msoTrue
    End With

    MsgBox "’› ”›«—‘ „Ê»«Ì· »« «—”«· ŒÊœò«— Ê  ò—«— «Ì„‰ ›⁄«· ‘œ.", vbInformation, "’› ”›«—‘ ‰»—«”"
End Sub

Public Sub Nebras_Install_All_Pick_V13()
    Nebras_Install_PickQueue_V13
    Application.Run "Nebras_Install_FinalizeStable_V13"
End Sub

Public Sub NebrasPick_BeginBatch_V13(ByVal batchNo As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = NebrasPickEnsureSheet()

    mBatch = batchNo
    Set mCustomers = CreateObject("Scripting.Dictionary")
    mCustomers.CompareMode = vbTextCompare
    mLineNo = 0

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For r = lastRow To 2 Step -1
        If CStr(ws.Cells(r, 2).Value2) = batchNo Then
            ws.rows(r).Delete
        End If
    Next r
End Sub

Public Sub NebrasPick_RecordAllocation_V13(ByVal batchNo As String, ByVal customer As String, _
                                           ByVal codeText As String, ByVal qty As Double, _
                                           ByVal shelf As String, ByVal rowNo As String, _
                                           ByVal boxNo As String)
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim orderId As String
    Dim customerNo As Long

    If qty <= 0 Then Exit Sub
    If Len(Trim$(customer)) = 0 Then Exit Sub

    If mCustomers Is Nothing Or mBatch <> batchNo Then
        NebrasPick_BeginBatch_V13 batchNo
    End If

    If Not mCustomers.Exists(Trim$(customer)) Then
        customerNo = mCustomers.Count + 1
        mCustomers.Add Trim$(customer), customerNo
    Else
        customerNo = CLng(mCustomers(Trim$(customer)))
    End If

    mLineNo = mLineNo + 1
    orderId = "XLS-PICK-" & batchNo & "-" & Format$(customerNo, "000")

    Set ws = NebrasPickEnsureSheet()
    nextRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value2 = orderId
    ws.Cells(nextRow, 2).Value2 = batchNo
    ws.Cells(nextRow, 3).Value2 = customer
    ws.Cells(nextRow, 4).Value2 = Format$(mLineNo, "00000")
    ws.Cells(nextRow, 5).NumberFormat = "@"
    ws.Cells(nextRow, 5).Value2 = codeText
    ws.Cells(nextRow, 6).Value2 = qty
    ws.Cells(nextRow, 7).NumberFormat = "@"
    ws.Cells(nextRow, 7).Value2 = NebrasPickLocation(shelf, rowNo, boxNo)
    ws.Cells(nextRow, 8).Value2 = STATUS_STAGED
    ws.Cells(nextRow, 9).Value2 = ""
    ws.Cells(nextRow, 10).Value2 = Now
    ws.Cells(nextRow, 11).Value2 = 0
    ws.Cells(nextRow, 12).Value2 = ""
    ws.Cells(nextRow, 13).Value2 = ""
    ws.Cells(nextRow, 14).Value2 = ""
End Sub

Public Sub NebrasPick_CommitBatch_V13(ByVal batchNo As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = NebrasPickEnsureSheet()
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If CStr(ws.Cells(r, 2).Value2) = batchNo Then
            If CStr(ws.Cells(r, 8).Value2) = STATUS_STAGED Then
                ws.Cells(r, 8).Value2 = STATUS_PENDING
                ws.Cells(r, 13).Value2 = Now
            End If
        End If
    Next r
End Sub

Public Sub NebrasPick_AbortBatch_V13(ByVal batchNo As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = NebrasPickEnsureSheet()
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = lastRow To 2 Step -1
        If CStr(ws.Cells(r, 2).Value2) = batchNo Then
            If CStr(ws.Cells(r, 8).Value2) = STATUS_STAGED Then
                ws.rows(r).Delete
            End If
        End If
    Next r
End Sub

Public Sub Nebras_Send_PickQueue_V13()
    Dim batchNo As String
    NebrasPickRecoverStale_V13
    batchNo = NebrasPickLatestPendingBatch_V13()
    If Len(batchNo) = 0 Then
        MsgBox "ÂÌç œ” Â ‰«„Ê›ﬁÌ »—«Ì «—”«· „Ãœœ ÊÃÊœ ‰œ«—œ.", vbInformation, "’› ”›«—‘ ‰»—«”"
        Exit Sub
    End If
    NebrasPickStartSend_V13 batchNo, True
End Sub

Public Sub Nebras_AutoSend_PickBatch_V13(ByVal batchNo As String)
    NebrasPickRecoverStale_V13
    NebrasPickStartSend_V13 batchNo, False
End Sub

Public Sub Nebras_Check_PickQueue_V13()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim staged As Long
    Dim pending As Long
    Dim sending As Long
    Dim sent As Long
    Dim lastInfo As String

    Set ws = NebrasPickEnsureSheet()
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        Select Case UCase$(Trim$(CStr(ws.Cells(r, 8).Value2)))
            Case STATUS_STAGED
                staged = staged + 1
            Case STATUS_PENDING
                pending = pending + 1
            Case STATUS_SENDING
                sending = sending + 1
            Case STATUS_SENT
                sent = sent + 1
        End Select

        If Len(Trim$(CStr(ws.Cells(r, 9).Value2))) > 0 Then
            lastInfo = CStr(ws.Cells(r, 9).Value2)
        End If
    Next r

    MsgBox "Ê÷⁄Ì  ’› »—œ«‘ :" & vbCrLf & _
           "œ— Õ«· ¬„«œÂ ”«“Ì: " & CStr(staged) & vbCrLf & _
           "œ— «‰ Ÿ«— «—”«·: " & CStr(pending) & vbCrLf & _
           "œ— Õ«· «—”«·: " & CStr(sending) & vbCrLf & _
           "«—”«· ‘œÂ: " & CStr(sent) & _
           IIf(Len(lastInfo) > 0, vbCrLf & vbCrLf & "¬Œ—Ì‰ Å«”Œ:" & vbCrLf & lastInfo, ""), _
           vbInformation, "»——”Ì ’› ”›«—‘ ‰»—«”"
End Sub

Private Sub NebrasPickStartSend_V13(ByVal onlyBatch As String, ByVal showMessage As Boolean)
    Dim payload As String
    Dim orderCount As Long
    Dim body As String

    If mSendRunning Then
        If Len(onlyBatch) > 0 And StrComp(onlyBatch, mSendBatch, vbTextCompare) <> 0 Then mDeferredBatch = onlyBatch
        If showMessage Then MsgBox "«—”«· ﬁ»·Ì Â‰Ê“ œ— Õ«· »——”Ì «” ∫  ò—«—Ì «—”«· ‰‘œ.", vbInformation, "’› ”›«—‘ ‰»—«”"
        Exit Sub
    End If

    If Len(onlyBatch) = 0 Then onlyBatch = NebrasPickLatestPendingBatch_V13()
    If Len(onlyBatch) = 0 Then Exit Sub

    NebrasPickCancelScheduledRetry_V13 onlyBatch

    Set mSendOrderIds = CreateObject("Scripting.Dictionary")
    payload = NebrasPickBuildBatchPayload(onlyBatch, orderCount)

    If orderCount = 0 Then
        If showMessage Then
            MsgBox "«Ì‰ œ” Â ”›«—‘Ì œ— «‰ Ÿ«— «—”«· ‰œ«—œ.", vbInformation, "’› ”›«—‘ ‰»—«”"
        End If
        Exit Sub
    End If

    mSendRunning = True
    mSendShowMessage = showMessage
    mSendBatch = onlyBatch
    mSendStartedAt = Now
    mSendMode = "POST"
    mSendAttemptNo = NebrasPickNextAttemptNo_V13(onlyBatch)
    mSendAttemptId = onlyBatch & "-" & Format$(Now, "hhmmss")
    NebrasPickMarkSending_V13
    NebrasPickSetButtonBusy True

    body = "action=pickQueuePutMany"
    body = body & "&key=" & NebrasPickUrl(PICK_KEY)
    body = body & "&payload=" & NebrasPickUrl(payload)

    Set mSendHttp = CreateObject("WinHttp.WinHttpRequest.5.1")

    On Error GoTo StartFailed

    mSendHttp.SetTimeouts 8000, 8000, 15000, 30000
    mSendHttp.Open "POST", PICK_URL, True
    mSendHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=UTF-8"
    mSendHttp.SetRequestHeader "Cache-Control", "no-cache"
    mSendHttp.Send body

    NebrasPickSchedulePoll_V13

    If showMessage Then
        MsgBox CStr(orderCount) & " ”›«—‘ »—«Ì «—”«· œ— Å” “„Ì‰Â ﬁ—«— ê—› ." & vbCrLf & _
               "«ﬂ”· ¬“«œ «”  Ê „Ì  Ê«‰Ìœ »Â ﬂ«— «œ«„Â œÂÌœ.", _
               vbInformation, "’› ”›«—‘ ‰»—«”"
    End If

    Exit Sub

StartFailed:
    NebrasPickStartConfirm_V13 "‘—Ê⁄ «—”«· ‰«„Ê›ﬁ »Êœ: " & Err.description
End Sub

Public Sub NebrasPick_SendPoll_V13()
    Dim finished As Boolean
    Dim response As String
    Dim statusCode As Long
    Dim ok As Boolean

    If Not mSendRunning Then Exit Sub

    If DateDiff("s", mSendStartedAt, Now) > 35 Then
        On Error Resume Next
        mSendHttp.Abort
        On Error GoTo 0
        If mSendMode = "POST" Then
            NebrasPickStartConfirm_V13 "Å«”Œ «—”«· »«  «ŒÌ— „Ê«ÃÂ ‘œ."
        Else
            NebrasPickStoreResponse False, " «ÌÌœ ”—Ê— œ— „Â·  „ﬁ—— Å«”Œ ‰œ«œ."
            NebrasPickRetryOrStop_V13 " «ÌÌœ «—”«· «‰Ã«„ ‰‘œ."
        End If
        Exit Sub
    End If

    On Error Resume Next
    finished = mSendHttp.WaitForResponse(0)
    If Err.number <> 0 Then
        Err.Clear
        finished = False
    End If
    On Error GoTo 0

    If Not finished Then
        NebrasPickSchedulePoll_V13
        Exit Sub
    End If

    On Error Resume Next
    statusCode = CLng(mSendHttp.Status)
    response = CStr(mSendHttp.responseText)
    If Err.number <> 0 Then
        response = "Œÿ« œ— ŒÊ«‰œ‰ Å«”Œ: " & Err.description
        Err.Clear
    End If
    On Error GoTo 0

    ok = (statusCode >= 200 And statusCode < 300 And InStr(1, response, """ok"":true", vbTextCompare) > 0)
    If mSendMode = "POST" Then
        If ok Then
            NebrasPickStoreResponse True, response
            NebrasPickStopSend_V13 True, "«—”«· ”›«—‘ Â« »« „Ê›ﬁÌ  «‰Ã«„ ‘œ."
        Else
            NebrasPickStartConfirm_V13 "Å«”Œ «—”«· ﬁÿ⁄Ì ‰»Êœ: " & Left$(response, RESPONSE_LIMIT)
        End If
    Else
        If ok Then
            If NebrasPickStoreConfirmation_V13(response) = 0 Then
                NebrasPickRetryOrStop_V13 "”—Ê— Â‰Ê“ ”›«—‘ —«  «ÌÌœ ‰ò—œ."
            ElseIf NebrasPickPendingCountForBatch_V13(mSendBatch) > 0 Then
                NebrasPickRetryOrStop_V13 "»Œ‘Ì «“ ”›«—‘ Â« ‰Ì«“ »Â «—”«· „Ãœœ œ«—‰œ."
            Else
                NebrasPickStopSend_V13 True, "«—”«· »«  «ÌÌœ ”—Ê— «‰Ã«„ ‘œ."
            End If
        Else
            NebrasPickStoreResponse False, " «ÌÌœ ”—Ê— ‰«„Ê›ﬁ »Êœ: " & Left$(response, RESPONSE_LIMIT)
            NebrasPickRetryOrStop_V13 " «ÌÌœ ”—Ê— ‰«„Ê›ﬁ »Êœ."
        End If
    End If
End Sub

Private Sub NebrasPickStopSend_V13(ByVal success As Boolean, ByVal messageText As String)
    Dim deferredBatch As String
    deferredBatch = mDeferredBatch
    mDeferredBatch = ""
    mSendRunning = False
    NebrasPickSetButtonBusy False

    If mSendShowMessage Then
        If success Then
            MsgBox messageText, vbInformation, "’› ”›«—‘ ‰»—«”"
        Else
            MsgBox messageText, vbExclamation, "’› ”›«—‘ ‰»—«”"
        End If
    End If

    Set mSendHttp = Nothing
    Set mSendOrderIds = Nothing
    mSendBatch = ""
    mSendMode = ""
    mSendShowMessage = False
    If Len(deferredBatch) > 0 And mRetryAt = 0 Then
        NebrasPickScheduleRetry_V13 deferredBatch, 2
    ElseIf Len(deferredBatch) > 0 Then
        mDeferredBatch = deferredBatch
    End If
End Sub

Private Sub NebrasPickStoreResponse(ByVal ok As Boolean, ByVal response As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim orderId As String

    Set ws = NebrasPickEnsureSheet()
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        orderId = CStr(ws.Cells(r, 1).Value2)

        If Not mSendOrderIds Is Nothing Then
            If mSendOrderIds.Exists(orderId) And CStr(ws.Cells(r, 2).Value2) = mSendBatch Then
                If CStr(ws.Cells(r, 8).Value2) = STATUS_SENDING Then
                    If ok Then
                        ws.Cells(r, 8).Value2 = STATUS_SENT Else ws.Cells(r, 8).Value2 = STATUS_PENDING
                    ws.Cells(r, 9).Value2 = Left$(response, RESPONSE_LIMIT)
                    If Not ok Then ws.Cells(r, 13).Value2 = Now + TimeSerial(0, 0, 5)
                End If
            End If
        End If
    Next r
End Sub

Private Sub NebrasPickRecoverStale_V13()
    Dim ws As Worksheet, lastRow As Long, r As Long
    Set ws = NebrasPickEnsureSheet()
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If CStr(ws.Cells(r, 8).Value2) = STATUS_SENDING Then
            If IsDate(ws.Cells(r, 12).Value) Then
                If DateDiff("s", CDate(ws.Cells(r, 12).Value), Now) > 120 Then
                    ws.Cells(r, 8).Value2 = STATUS_PENDING
                    ws.Cells(r, 9).Value2 = "«—”«· ﬁ»·Ì »œÊ‰ Å«”Œ „«‰œ∫ »—«Ì »——”Ì «Ì„‰ ¬„«œÂ ‘œ."
                    ws.Cells(r, 13).Value2 = Now
                End If
            Else
                ws.Cells(r, 8).Value2 = STATUS_PENDING
            End If
        End If
    Next r
End Sub

Private Function NebrasPickLatestPendingBatch_V13() As String
    Dim ws As Worksheet, lastRow As Long, r As Long, batchNo As String
    Set ws = NebrasPickEnsureSheet()
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If CStr(ws.Cells(r, 8).Value2) = STATUS_PENDING Then
            batchNo = CStr(ws.Cells(r, 2).Value2)
            If StrComp(batchNo, NebrasPickLatestPendingBatch_V13, vbTextCompare) > 0 Then NebrasPickLatestPendingBatch_V13 = batchNo
        End If
    Next r
End Function

Private Function NebrasPickNextAttemptNo_V13(ByVal batchNo As String) As Long
    Dim ws As Worksheet, lastRow As Long, r As Long, largest As Long
    Set ws = NebrasPickEnsureSheet()
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If CStr(ws.Cells(r, 2).Value2) = batchNo Then largest = Application.Max(largest, Val(ws.Cells(r, 11).Value2))
    Next r
    NebrasPickNextAttemptNo_V13 = largest + 1
End Function

Private Sub NebrasPickMarkSending_V13()
    Dim ws As Worksheet, lastRow As Long, r As Long, orderId As String
    Set ws = NebrasPickEnsureSheet()
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        orderId = CStr(ws.Cells(r, 1).Value2)
        If mSendOrderIds.Exists(orderId) And CStr(ws.Cells(r, 2).Value2) = mSendBatch Then
            If CStr(ws.Cells(r, 8).Value2) = STATUS_PENDING Then
                ws.Cells(r, 8).Value2 = STATUS_SENDING
                ws.Cells(r, 9).Value2 = "œ— Õ«· «—”«· Ê  «ÌÌœ ”—Ê—..."
                ws.Cells(r, 11).Value2 = mSendAttemptNo
                ws.Cells(r, 12).Value2 = Now
                ws.Cells(r, 13).Value2 = ""
                ws.Cells(r, 14).Value2 = mSendAttemptId
            End If
        End If
    Next r
End Sub

Private Sub NebrasPickStartConfirm_V13(ByVal reasonText As String)
    Dim url As String
    On Error GoTo ConfirmFailed
    Set mSendHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    mSendMode = "CONFIRM"
    mSendStartedAt = Now
    url = PICK_URL & "?action=pickQueueConfirm&key=" & NebrasPickUrl(PICK_KEY) & _
          "&orderIds=" & NebrasPickUrl(NebrasPickOrderIdList_V13())
    mSendHttp.SetTimeouts 8000, 8000, 15000, 30000
    mSendHttp.Open "GET", url, True
    mSendHttp.SetRequestHeader "Cache-Control", "no-cache"
    mSendHttp.Send
    NebrasPickSchedulePoll_V13
    Exit Sub
ConfirmFailed:
    NebrasPickStoreResponse False, reasonText & " /  «ÌÌœ: " & Err.description
    NebrasPickRetryOrStop_V13 " «ÌÌœ «—”«· «‰Ã«„ ‰‘œ."
End Sub

Private Function NebrasPickOrderIdList_V13() As String
    Dim key As Variant
    If mSendOrderIds Is Nothing Then Exit Function
    For Each key In mSendOrderIds.Keys
        If Len(NebrasPickOrderIdList_V13) > 0 Then NebrasPickOrderIdList_V13 = NebrasPickOrderIdList_V13 & "|"
        NebrasPickOrderIdList_V13 = NebrasPickOrderIdList_V13 & CStr(key)
    Next key
End Function

Private Function NebrasPickStoreConfirmation_V13(ByVal response As String) As Long
    Dim ws As Worksheet, lastRow As Long, r As Long, orderId As String
    Set ws = NebrasPickEnsureSheet()
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        orderId = CStr(ws.Cells(r, 1).Value2)
        If mSendOrderIds.Exists(orderId) And CStr(ws.Cells(r, 2).Value2) = mSendBatch Then
            If CStr(ws.Cells(r, 8).Value2) = STATUS_SENDING Then
                If NebrasPickOrderIsConfirmed_V13(response, orderId) Then
                    ws.Cells(r, 8).Value2 = STATUS_SENT
                    ws.Cells(r, 9).Value2 = "«—”«· »«  «ÌÌœ ”—Ê— «‰Ã«„ ‘œ."
                    NebrasPickStoreConfirmation_V13 = NebrasPickStoreConfirmation_V13 + 1
                Else
                    ws.Cells(r, 8).Value2 = STATUS_PENDING
                    ws.Cells(r, 9).Value2 = "”—Ê— Â‰Ê“ «Ì‰ ”›«—‘ —«  «ÌÌœ ‰ò—œ."
                    ws.Cells(r, 13).Value2 = Now + TimeSerial(0, 0, 5)
                End If
            End If
        End If
    Next r
End Function

Private Function NebrasPickOrderIsConfirmed_V13(ByVal response As String, ByVal orderId As String) As Boolean
    Dim startAt As Long, endAt As Long, presentText As String
    startAt = InStr(1, response, """present""", vbTextCompare)
    If startAt = 0 Then Exit Function
    startAt = InStr(startAt, response, "[")
    endAt = InStr(startAt, response, "]")
    If startAt = 0 Or endAt = 0 Then Exit Function
    presentText = Mid$(response, startAt, endAt - startAt + 1)
    NebrasPickOrderIsConfirmed_V13 = (InStr(1, presentText, """" & orderId & """", vbTextCompare) > 0)
End Function

Private Function NebrasPickPendingCountForBatch_V13(ByVal batchNo As String) As Long
    Dim ws As Worksheet, lastRow As Long, r As Long
    Set ws = NebrasPickEnsureSheet()
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If CStr(ws.Cells(r, 2).Value2) = batchNo And CStr(ws.Cells(r, 8).Value2) = STATUS_PENDING Then
            NebrasPickPendingCountForBatch_V13 = NebrasPickPendingCountForBatch_V13 + 1
        End If
    Next r
End Function

Private Sub NebrasPickRetryOrStop_V13(ByVal messageText As String)
    If mSendAttemptNo < MAX_AUTO_ATTEMPTS And NebrasPickPendingCountForBatch_V13(mSendBatch) > 0 Then
        NebrasPickScheduleRetry_V13 mSendBatch, 5 * mSendAttemptNo
        NebrasPickStopSend_V13 False, messageText
    Else
        NebrasPickStopSend_V13 False, messageText & " «—”«· ŒÊœò«— „ Êﬁ› ‘œ∫ ›ﬁÿ ¬Œ—Ì‰ œ” Â ‰«„Ê›ﬁ »« œò„Â ﬁ«»· «—”«· «” ."
    End If
End Sub

Public Sub NebrasPick_RetryBatch_V13()
    Dim batchNo As String
    batchNo = mRetryBatch
    mRetryBatch = ""
    mRetryAt = 0
    If Len(batchNo) > 0 Then NebrasPickStartSend_V13 batchNo, False
End Sub

Private Sub NebrasPickScheduleRetry_V13(ByVal batchNo As String, ByVal delaySeconds As Long)
    If Len(Trim$(batchNo)) = 0 Then Exit Sub
    If delaySeconds < 1 Then delaySeconds = 1
    mRetryBatch = batchNo
    mRetryAt = Now + TimeSerial(0, 0, delaySeconds)
    Application.OnTime EarliestTime:=mRetryAt, Procedure:="NebrasPick_RetryBatch_V13", Schedule:=True
End Sub

Private Sub NebrasPickCancelScheduledRetry_V13(ByVal batchNo As String)
    If mRetryAt = 0 Then Exit Sub
    If StrComp(batchNo, mRetryBatch, vbTextCompare) <> 0 Then Exit Sub
    On Error Resume Next
    Application.OnTime EarliestTime:=mRetryAt, Procedure:="NebrasPick_RetryBatch_V13", Schedule:=False
    On Error GoTo 0
    mRetryBatch = ""
    mRetryAt = 0
End Sub

Private Sub NebrasPickSchedulePoll_V13()
    mSendPollAt = Now + TimeSerial(0, 0, 1)
    Application.OnTime EarliestTime:=mSendPollAt, Procedure:="NebrasPick_SendPoll_V13", Schedule:=True
End Sub

Private Function NebrasPickBuildBatchPayload(ByVal onlyBatch As String, ByRef orderCount As Long) As String
    Dim ws As Worksheet
    Dim orders As Object
    Dim rows As Collection
    Dim c As Collection
    Dim lastRow As Long
    Dim r As Long
    Dim orderId As String
    Dim key As Variant
    Dim jsonOrders As String

    Set ws = NebrasPickEnsureSheet()
    Set orders = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If CStr(ws.Cells(r, 8).Value2) = STATUS_PENDING Then
            If CStr(ws.Cells(r, 2).Value2) = onlyBatch Then
                orderId = CStr(ws.Cells(r, 1).Value2)

                If Not orders.Exists(orderId) Then
                    Set c = New Collection
                    orders.Add orderId, c
                    mSendOrderIds.Add orderId, True
                End If

                Set c = orders(orderId)
                c.Add r
            End If
        End If
    Next r

    orderCount = orders.Count

    For Each key In orders.Keys
        Set rows = orders(key)

        If Len(jsonOrders) > 0 Then
            jsonOrders = jsonOrders & ","
        End If

        jsonOrders = jsonOrders & NebrasPickOrderPayload(ws, rows)
    Next key

    NebrasPickBuildBatchPayload = "{""orders"":[" & jsonOrders & "]}"
End Function

Private Function NebrasPickOrderPayload(ByVal ws As Worksheet, ByVal rows As Collection) As String
    Dim i As Long
    Dim r As Long
    Dim firstRow As Long
    Dim items As String
    Dim qtyText As String

    firstRow = CLng(rows(1))

    For i = 1 To rows.Count
        r = CLng(rows(i))

        If Len(items) > 0 Then
            items = items & ","
        End If

        qtyText = Replace(CStr(CDbl(ws.Cells(r, 6).Value2)), Application.DecimalSeparator, ".")

        items = items & "{"
        items = items & """lineId"":" & NebrasPickJson(CStr(ws.Cells(r, 4).Value2))
        items = items & ",""code"":" & NebrasPickJson(CStr(ws.Cells(r, 5).Value2))
        items = items & ",""qty"":" & qtyText
        items = items & ",""location"":" & NebrasPickJson(CStr(ws.Cells(r, 7).Value2))
        items = items & "}"
    Next i

    NebrasPickOrderPayload = "{"
    NebrasPickOrderPayload = NebrasPickOrderPayload & """orderId"":" & NebrasPickJson(CStr(ws.Cells(firstRow, 1).Value2))
    NebrasPickOrderPayload = NebrasPickOrderPayload & ",""batch"":" & NebrasPickJson(CStr(ws.Cells(firstRow, 2).Value2))
    NebrasPickOrderPayload = NebrasPickOrderPayload & ",""customer"":" & NebrasPickJson(CStr(ws.Cells(firstRow, 3).Value2))
    NebrasPickOrderPayload = NebrasPickOrderPayload & ",""items"":[" & items & "]"
    NebrasPickOrderPayload = NebrasPickOrderPayload & "}"
End Function

Private Function NebrasPickEnsureSheet() As Worksheet
    Dim ws As Worksheet
    Dim headers As Variant
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(QUEUE_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = QUEUE_SHEET

        ws.rows(1).Font.Bold = True
    End If

    headers = Array("OrderId", "Batch", "Customer", "LineId", "Code", "Qty", "Location", "Status", "Server", "CreatedAt", _
                    "AttemptCount", "LastAttemptAt", "NextRetryAt", "AttemptId")
    For i = 0 To UBound(headers)
        If Len(Trim$(CStr(ws.Cells(1, i + 1).Value2))) = 0 Then
            ws.Cells(1, i + 1).Value2 = headers(i)
        End If
    Next i

    Set NebrasPickEnsureSheet = ws
End Function

Private Function NebrasPickFindTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing

        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0

        If Not lo Is Nothing Then
            Set NebrasPickFindTable = lo
            Exit Function
        End If
    Next ws
End Function

Private Function NebrasPickLocation(ByVal shelf As String, ByVal rowNo As String, ByVal boxNo As String) As String
    Dim s As String
    Dim p As Variant
    Dim result As String

    s = Trim$(boxNo)
    s = Replace(s, ChrW$(&H60C), ",")
    s = Replace(s, ".", ",")
    s = Replace(s, " ", "")

    If Val(shelf) <= 0 Or Val(rowNo) <= 0 Or Len(s) = 0 Then
        NebrasPickLocation = "Shelf " & shelf & " / Row " & rowNo & " / Box " & boxNo
        Exit Function
    End If

    p = Split(s, ",")

    result = "NBR-S" & Format$(Val(shelf), "00")
    result = result & "-R" & Format$(Val(rowNo), "00")
    result = result & "-B" & Format$(Val(p(0)), "00")

    If UBound(p) >= 1 Then
        result = result & "-" & Format$(Val(p(1)), "00")
    End If

    NebrasPickLocation = result
End Function

Private Function NebrasPickJson(ByVal value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, Chr$(34), "\" & Chr$(34))
    value = Replace(value, vbCr, "\r")
    value = Replace(value, vbLf, "\n")

    NebrasPickJson = Chr$(34) & value & Chr$(34)
End Function

Private Function NebrasPickUrl(ByVal value As String) As String
    Dim stream As Object
    Dim bytes As Variant
    Dim i As Long
    Dim b As Long

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
        b = bytes(i)

        If (b >= 48 And b <= 57) Or _
           (b >= 65 And b <= 90) Or _
           (b >= 97 And b <= 122) Or _
           b = 45 Or b = 46 Or b = 95 Or b = 126 Then

            NebrasPickUrl = NebrasPickUrl & Chr$(b)
        Else
            NebrasPickUrl = NebrasPickUrl & "%" & Right$("0" & Hex$(b), 2)
        End If
    Next i
End Function

Private Sub NebrasPickSetButtonBusy(ByVal busy As Boolean)
    Dim tbl As ListObject
    Dim shp As Shape

    Set tbl = NebrasPickFindTable("WarehouseHistoryTable")
    If tbl Is Nothing Then Exit Sub

    On Error Resume Next
    Set shp = tbl.Parent.Shapes("btnNebrasSendPickQueueV13")
    On Error GoTo 0

    If shp Is Nothing Then Exit Sub

    If busy Then
        shp.TextFrame2.TextRange.Text = "œ— Õ«· «—”«·..."
        shp.Fill.ForeColor.RGB = RGB(120, 120, 120)
    Else
        shp.TextFrame2.TextRange.Text = "«—”«· ¬Œ—Ì‰ ”›«—‘ ‰«„Ê›ﬁ »Â „Ê»«Ì·"
        shp.Fill.ForeColor.RGB = RGB(92, 72, 168)
    End If
End Sub
