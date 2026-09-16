Attribute VB_Name = "NebrasWarehouseBackgroundV10"
Option Explicit

' Nebras V10 - responsive two-way warehouse sync.
' Network is asynchronous. Excel writes are split into very short callbacks.
' Excel is authoritative after each completed sync snapshot.

Private Const NB10_URL As String = "https://script.google.com/macros/s/AKfycbx3Kl16l2U21knbjy-7BVWvB0u0H2gNMvpp8nIr-SrnFLe9WplVYjdNxFDYmqymZPsrHA/exec"
Private Const NB10_KEY As String = "565bcdf6f75dbfa57dc467be8fd02910ea1c2cdc90592903"
Private Const NB10_EPS As Double = 0.0001

Private nb10Req As Object
Private nb10Busy As Boolean
Private nb10Stage As String
Private nb10Due As Date
Private nb10Proc As String
Private nb10Deadline As Date
Private nb10Lines() As String
Private nb10Index As Long
Private nb10Acks As Collection
Private nb10AckIndex As Long
Private nb10Applied As Long
Private nb10Errors As Long
Private nb10BlockedCodes As Object
Private nb10Dirty As Boolean
Private nb10Saved As Boolean
Private nb10SavedStatus As Variant
Private nb10SaveAttempts As Long

Public Sub InstallNebrasBackgroundV10()
    Dim ws As Worksheet, shp As Shape
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Set shp = Nothing
        Set shp = ws.Shapes("btnNebrasWarehouseSync")
        If Not shp Is Nothing Then Exit For
    Next ws
    On Error GoTo Failed
    If shp Is Nothing Then Err.Raise vbObjectError + 900, , "Warehouse sync button not found"
    shp.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!Nebras_Background_Sync_V10"
    Application.StatusBar = "Nebras: background sync V10 installed"
    Exit Sub
Failed:
    MsgBox "Nebras V10 install:" & vbCrLf & Err.Description, vbExclamation, "Nebras"
End Sub

Public Sub Nebras_Background_Sync_V10()
    NB10Start
End Sub

Public Sub Nebras_Stop_Background_Sync_V10()
    If nb10Busy Then NB10Cleanup
End Sub

Private Sub NB10Start()
    On Error GoTo Failed
    If nb10Busy Then Exit Sub
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 901, , "Workbook is read-only"
    If Len(ThisWorkbook.Path) = 0 Then Err.Raise vbObjectError + 902, , "Save the workbook once before synchronization"
    NB10EnsureSyncSheet
    nb10SavedStatus = Application.StatusBar
    nb10Busy = True
    nb10Dirty = False
    nb10Saved = False
    nb10SaveAttempts = 0
    nb10Applied = 0
    nb10Errors = 0
    Set nb10Acks = New Collection
    Set nb10BlockedCodes = CreateObject("Scripting.Dictionary")
    nb10BlockedCodes.CompareMode = vbTextCompare
    nb10Stage = "START"
    Application.StatusBar = "Nebras: synchronization started in background..."
    NB10Schedule
    Exit Sub
Failed:
    NB10Fail Err.Description
End Sub

Private Sub NB10Schedule()
    nb10Due = Now + TimeSerial(0, 0, 1)
    nb10Proc = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!Nebras_Background_Sync_V10_Tick"
    Application.OnTime EarliestTime:=nb10Due, Procedure:=nb10Proc
End Sub

Private Sub NB10Request(ByVal stage As String, ByVal method As String, ByVal url As String, Optional ByVal body As String = "")
    nb10Stage = stage
    Set nb10Req = CreateObject("WinHttp.WinHttpRequest.5.1")
    nb10Req.Option(6) = True
    nb10Req.SetTimeouts 7000, 10000, 20000, 45000
    nb10Req.Open method, url, True
    nb10Req.SetRequestHeader "Cache-Control", "no-cache"
    If method = "POST" Then nb10Req.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=UTF-8"
    nb10Deadline = DateAdd("s", 70, Now)
    nb10Req.Send body
    NB10Schedule
End Sub

Public Sub Nebras_Background_Sync_V10_Tick()
    Dim response As String, f() As String, ack As Variant, countOps As Long
    Dim saveErr As String
    On Error GoTo Failed
    nb10Due = 0
    If Not nb10Busy Then Exit Sub

    Select Case nb10Stage
        Case "START"
            NB10Request "PULL", "GET", NB10_URL & "?action=warehousePull&limit=200&key=" & NB10UrlEncodeUtf8(NB10_KEY)
            Exit Sub

        Case "APPLY"
            NB10ApplyChunk
            Exit Sub

        Case "SAVE"
            If Not nb10Dirty Then
                nb10AckIndex = 1
                NB10NextAck
                Exit Sub
            End If
            On Error Resume Next
            Err.Clear
            ThisWorkbook.Save
            saveErr = Err.Description
            On Error GoTo Failed
            If Len(saveErr) > 0 Or Not ThisWorkbook.Saved Then
                nb10SaveAttempts = nb10SaveAttempts + 1
                If nb10SaveAttempts < 3 Then
                    Application.StatusBar = "Nebras: save retry " & CStr(nb10SaveAttempts) & "..."
                    NB10Schedule
                    Exit Sub
                End If
                Err.Raise vbObjectError + 903, , "Document not saved after 3 attempts. " & saveErr
            End If
            nb10Saved = True
            nb10AckIndex = 1
            NB10NextAck
            Exit Sub

        Case "BUILD"
            Application.StatusBar = "Nebras: publishing current Excel inventory..."
            response = NB10SnapshotBody()
            NB10Request "PUBLISH", "POST", NB10_URL, response
            Exit Sub
    End Select

    If Now > nb10Deadline Then Err.Raise vbObjectError + 904, , "Network deadline exceeded"
    If Not nb10Req.WaitForResponse(0) Then
        NB10Schedule
        Exit Sub
    End If
    If nb10Req.Status < 200 Or nb10Req.Status >= 300 Then Err.Raise vbObjectError + 905, , "HTTP " & CStr(nb10Req.Status)
    response = nb10Req.responseText
    Set nb10Req = Nothing

    Select Case nb10Stage
        Case "PULL"
            response = NB10NormalizeResponse(response)
            nb10Lines = Split(Replace(response, vbCr, ""), vbLf)
            f = Split(nb10Lines(0), vbTab)
            If UBound(f) < 1 Or UCase$(f(0)) <> "OK" Then Err.Raise vbObjectError + 906, , "Invalid pull response"
            countOps = CLng(Val(f(1)))
            nb10Index = 1
            If countOps <= 0 Then
                nb10Stage = "BUILD"
            Else
                nb10Stage = "APPLY"
            End If
            NB10Schedule

        Case "ACK"
            If InStr(1, Replace(Replace(response, " ", ""), vbCr, ""), """ok"":true", vbTextCompare) = 0 Then _
                Err.Raise vbObjectError + 907, , "Server did not accept acknowledgement"
            nb10AckIndex = nb10AckIndex + 1
            NB10NextAck

        Case "PUBLISH"
            If InStr(1, Replace(Replace(response, " ", ""), vbCr, ""), """ok"":true", vbTextCompare) = 0 Then _
                Err.Raise vbObjectError + 908, , "Snapshot not accepted"
            NB10Cleanup
            Application.StatusBar = "Nebras: sync complete | applied " & CStr(nb10Applied) & " | review " & CStr(nb10Errors)
    End Select
    Exit Sub
Failed:
    NB10Fail Err.Description
End Sub

Private Sub NB10ApplyChunk()
    Dim f() As String, issue As String, codeKey As String, done As Long
    Dim oldEvents As Boolean
    On Error GoTo Failed

    Do While nb10Index <= UBound(nb10Lines) And done < 2
        If Len(Trim$(nb10Lines(nb10Index))) > 0 Then
            f = Split(nb10Lines(nb10Index), vbTab)
            If UBound(f) < 9 Then Err.Raise vbObjectError + 909, , "Incomplete operation"
            issue = ""
            codeKey = NB10NormalizeKey(f(2))

            If NB10OperationAlreadyApplied(f(0)) Then
                nb10Acks.Add Array(f(0), True, "")
            ElseIf nb10BlockedCodes.Exists(codeKey) Then
                issue = "Previous operation for this product needs review; dependent operation was not applied."
                NB10Review f, issue
                nb10Dirty = True
                nb10Errors = nb10Errors + 1
                nb10Acks.Add Array(f(0), False, issue)
            Else
                oldEvents = Application.EnableEvents
                Application.EnableEvents = False
                issue = NB10TryApply(f)
                Application.EnableEvents = oldEvents

                If Len(issue) = 0 Then
                    NB10LogApplied f
                    nb10Dirty = True
                    nb10Applied = nb10Applied + 1
                    nb10Acks.Add Array(f(0), True, "")
                Else
                    nb10BlockedCodes(codeKey) = True
                    NB10Review f, issue
                    nb10Dirty = True
                    nb10Errors = nb10Errors + 1
                    nb10Acks.Add Array(f(0), False, issue)
                End If
            End If
        End If
        nb10Index = nb10Index + 1
        done = done + 1
    Loop

    If nb10Index > UBound(nb10Lines) Then nb10Stage = "SAVE"
    Application.StatusBar = "Nebras: background apply " & CStr(nb10Applied) & " | review " & CStr(nb10Errors)
    NB10Schedule
    Exit Sub
Failed:
    Application.EnableEvents = True
    NB10Fail Err.Description
End Sub

Private Function NB10TryApply(ByRef f() As String) As String
    Dim n As Long, d As String
    On Error GoTo Rejected
    NB10ApplyOperation f
    Exit Function
Rejected:
    n = Err.Number: d = Err.Description
    Select Case n
        Case vbObjectError + 710, vbObjectError + 711, vbObjectError + 712, vbObjectError + 713, _
             vbObjectError + 714, vbObjectError + 716, vbObjectError + 718, vbObjectError + 719, _
             vbObjectError + 830
            NB10TryApply = d
        Case Else
            Err.Raise n, "NB10TryApply", d
    End Select
End Function

Private Sub NB10ApplyOperation(ByRef f() As String)
    Dim tbl As ListObject, rowIndex As Long, targetRow As Long
    Dim qty As Double, beforeQty As Double, afterQty As Double, actualQty As Double
    Dim opType As String, code As String, fromLoc As String, toLoc As String
    Dim occasionText As String, customerText As String, descText As String, auditQty As Double

    Set tbl = NB10SheetByTable("InventoryBankTable").ListObjects("InventoryBankTable")
    opType = f(1): code = f(2): qty = Val(f(3)): fromLoc = f(4): toLoc = f(5)
    beforeQty = Val(f(8)): afterQty = Val(f(9)): auditQty = qty

    Select Case opType
        Case NB10U("062E 0631 0648 062C") ' output
            rowIndex = NB10FindLocationRow(tbl, code, fromLoc)
            If rowIndex = 0 Then Err.Raise vbObjectError + 710, , "Source location no longer exists in Excel."
            NB10RecalcRow tbl, rowIndex
            actualQty = Val(tbl.ListColumns(7).DataBodyRange.Cells(rowIndex, 1).Value2)
            NB10RequireBefore actualQty, beforeQty
            If actualQty < qty - NB10_EPS Then Err.Raise vbObjectError + 711, , "Insufficient Excel stock for output."
            tbl.ListColumns(5).DataBodyRange.Cells(rowIndex, 1).Value2 = Val(tbl.ListColumns(5).DataBodyRange.Cells(rowIndex, 1).Value2) + qty
            NB10RecalcRow tbl, rowIndex
            occasionText = CStr(tbl.ListColumns(1).DataBodyRange.Cells(rowIndex, 1).Value2)
            customerText = NB10MobileCustomer(opType, f(6))
            descText = NB10Description(opType, fromLoc, toLoc, f(6), beforeQty, afterQty)
            NB10AppendOutputHistory code, qty, customerText, descText, f(7), rowIndex, tbl

        Case NB10U("0648 0631 0648 062F") ' input
            rowIndex = NB10FindLocationRow(tbl, code, toLoc)
            If rowIndex = 0 Then
                If Abs(beforeQty) > NB10_EPS Then Err.Raise vbObjectError + 830, , "Excel changed after this mobile operation; refresh mobile before retrying."
                rowIndex = NB10CreateLocationRow(tbl, code, toLoc, NB10OperationOccasion(f(6)))
                NB10RecalcRow tbl, rowIndex
            Else
                NB10RecalcRow tbl, rowIndex
                actualQty = Val(tbl.ListColumns(7).DataBodyRange.Cells(rowIndex, 1).Value2)
                NB10RequireBefore actualQty, beforeQty
            End If
            tbl.ListColumns(3).DataBodyRange.Cells(rowIndex, 1).Value2 = Val(tbl.ListColumns(3).DataBodyRange.Cells(rowIndex, 1).Value2) + qty
            NB10RecalcRow tbl, rowIndex
            occasionText = CStr(tbl.ListColumns(1).DataBodyRange.Cells(rowIndex, 1).Value2)
            customerText = ""
            descText = NB10Description(opType, fromLoc, toLoc, f(6), beforeQty, afterQty)

        Case NB10U("0627 0646 062A 0642 0627 0644") ' transfer
            rowIndex = NB10FindLocationRow(tbl, code, fromLoc)
            If rowIndex = 0 Then Err.Raise vbObjectError + 712, , "Source location no longer exists in Excel."
            NB10RecalcRow tbl, rowIndex
            actualQty = Val(tbl.ListColumns(7).DataBodyRange.Cells(rowIndex, 1).Value2)
            NB10RequireBefore actualQty, beforeQty
            If actualQty < qty - NB10_EPS Then Err.Raise vbObjectError + 713, , "Insufficient Excel stock for transfer."
            tbl.ListColumns(5).DataBodyRange.Cells(rowIndex, 1).Value2 = Val(tbl.ListColumns(5).DataBodyRange.Cells(rowIndex, 1).Value2) + qty
            NB10RecalcRow tbl, rowIndex
            targetRow = NB10GetLocationRow(tbl, code, toLoc, "")
            tbl.ListColumns(3).DataBodyRange.Cells(targetRow, 1).Value2 = Val(tbl.ListColumns(3).DataBodyRange.Cells(targetRow, 1).Value2) + qty
            NB10RecalcRow tbl, targetRow
            occasionText = CStr(tbl.ListColumns(1).DataBodyRange.Cells(targetRow, 1).Value2)
            customerText = ""
            descText = NB10Description(opType, fromLoc, toLoc, f(6), beforeQty, afterQty)

        Case NB10U("0627 0635 0644 0627 062D") ' adjust
            rowIndex = NB10FindLocationRow(tbl, code, fromLoc)
            If rowIndex = 0 Then Err.Raise vbObjectError + 714, , "Adjusted location no longer exists in Excel."
            NB10RecalcRow tbl, rowIndex
            actualQty = Val(tbl.ListColumns(7).DataBodyRange.Cells(rowIndex, 1).Value2)
            NB10RequireBefore actualQty, beforeQty
            If afterQty > beforeQty Then
                tbl.ListColumns(3).DataBodyRange.Cells(rowIndex, 1).Value2 = Val(tbl.ListColumns(3).DataBodyRange.Cells(rowIndex, 1).Value2) + (afterQty - beforeQty)
            ElseIf beforeQty > afterQty Then
                tbl.ListColumns(5).DataBodyRange.Cells(rowIndex, 1).Value2 = Val(tbl.ListColumns(5).DataBodyRange.Cells(rowIndex, 1).Value2) + (beforeQty - afterQty)
            End If
            NB10RecalcRow tbl, rowIndex
            occasionText = CStr(tbl.ListColumns(1).DataBodyRange.Cells(rowIndex, 1).Value2)
            customerText = "": auditQty = Abs(afterQty - beforeQty)
            descText = NB10Description(opType, fromLoc, toLoc, f(6), beforeQty, afterQty)

        Case NB10U("062A 0641 06A9 06CC 06A9") ' split
            rowIndex = NB10FindSharedRow(tbl, code, fromLoc)
            If rowIndex = 0 Then Err.Raise vbObjectError + 718, , "Shared Excel location no longer exists. Refresh mobile."
            NB10RecalcRow tbl, rowIndex
            actualQty = Val(tbl.ListColumns(7).DataBodyRange.Cells(rowIndex, 1).Value2)
            NB10RequireBefore actualQty, beforeQty
            If actualQty < qty - NB10_EPS Then Err.Raise vbObjectError + 719, , "Insufficient shared Excel stock."
            tbl.ListColumns(5).DataBodyRange.Cells(rowIndex, 1).Value2 = Val(tbl.ListColumns(5).DataBodyRange.Cells(rowIndex, 1).Value2) + qty
            NB10RecalcRow tbl, rowIndex
            targetRow = NB10GetLocationRow(tbl, code, toLoc, "")
            tbl.ListColumns(3).DataBodyRange.Cells(targetRow, 1).Value2 = Val(tbl.ListColumns(3).DataBodyRange.Cells(targetRow, 1).Value2) + qty
            NB10RecalcRow tbl, targetRow
            occasionText = CStr(tbl.ListColumns(1).DataBodyRange.Cells(targetRow, 1).Value2)
            customerText = ""
            descText = NB10Description(opType, fromLoc, toLoc, f(6), beforeQty, afterQty)

        Case Else
            Err.Raise vbObjectError + 715, , "Unknown mobile operation type."
    End Select

    NB10AppendMobileAudit f(0), opType, code, occasionText, auditQty, customerText, fromLoc, toLoc, beforeQty, afterQty, descText, f(7)
End Sub

Private Sub NB10RequireBefore(ByVal actualQty As Double, ByVal beforeQty As Double)
    If Abs(actualQty - beforeQty) > NB10_EPS Then _
        Err.Raise vbObjectError + 830, , "Excel changed after this mobile operation; Excel was kept as the authoritative state. Refresh mobile before retrying."
End Sub

Private Sub NB10RecalcRow(ByVal tbl As ListObject, ByVal rowIndex As Long)
    If rowIndex < 1 Or rowIndex > tbl.ListRows.Count Then Exit Sub
    tbl.DataBodyRange.Cells(rowIndex, 6).Resize(1, 2).Calculate
End Sub

Private Sub NB10NextAck()
    Dim a As Variant
    If nb10AckIndex < 1 Then nb10AckIndex = 1
    If nb10AckIndex > nb10Acks.Count Then
        nb10Stage = "BUILD"
        NB10Schedule
        Exit Sub
    End If
    a = nb10Acks(nb10AckIndex)
    NB10Request "ACK", "GET", NB10_URL & "?action=warehouseAck&key=" & NB10UrlEncodeUtf8(NB10_KEY) & _
        "&id=" & NB10UrlEncodeUtf8(CStr(a(0))) & "&ok=" & IIf(CBool(a(1)), "1", "0") & _
        "&message=" & NB10UrlEncodeUtf8(CStr(a(2)))
End Sub

Private Function NB10SnapshotBody() As String
    Dim tbl As ListObject, bank As Variant, r As Long
    Dim code As String, occasion As String, shelf As String, rowNo As String, boxNo As String, loc As String
    Dim qty As Double, products As String, stock As String, locations As String
    Dim productKeys As Object, locationKeys As Object, payload As String

    Set tbl = NB10SheetByTable("InventoryBankTable").ListObjects("InventoryBankTable")
    If tbl.DataBodyRange Is Nothing Then Err.Raise vbObjectError + 924, , "Inventory table is empty"
    bank = tbl.DataBodyRange.Value2
    Set productKeys = CreateObject("Scripting.Dictionary")
    Set locationKeys = CreateObject("Scripting.Dictionary")
    productKeys.CompareMode = vbTextCompare: locationKeys.CompareMode = vbTextCompare

    For r = 1 To UBound(bank, 1)
        code = NB10NormalizeKey(bank(r, 2))
        If Len(code) > 0 Then
            occasion = Trim$(CStr(bank(r, 1)))
            If Not productKeys.Exists(code) Then
                productKeys.Add code, True
                If Len(products) > 0 Then products = products & ","
                products = products & "{""code"":" & NB10JsonString(code) & ",""name"":" & NB10JsonString(occasion) & ",""occasion"":" & NB10JsonString(occasion) & "}"
            End If
            qty = Val(bank(r, 7))
            If qty > 0 Then
                shelf = Trim$(CStr(bank(r, 8))): rowNo = Trim$(CStr(bank(r, 9))): boxNo = Trim$(CStr(bank(r, 10)))
                loc = NB10PhysicalLocationCode(shelf, rowNo, boxNo)
                If Len(loc) > 0 Then
                    If Not locationKeys.Exists(loc) Then
                        locationKeys.Add loc, True
                        If Len(locations) > 0 Then locations = locations & ","
                        locations = locations & NB10JsonString(loc)
                    End If
                    If Len(stock) > 0 Then stock = stock & ","
                    stock = stock & "{""p"":" & NB10JsonString(code) & ",""l"":" & NB10JsonString(loc) & ",""q"":" & NB10Number(qty) & ",""display"":" & NB10JsonString(loc) & "}"
                ElseIf Val(shelf) > 0 And Val(rowNo) > 0 And Len(NB10NormalizeBox(boxNo)) > 0 Then
                    loc = NB10SharedLocationCode(shelf, rowNo, boxNo)
                    If Len(stock) > 0 Then stock = stock & ","
                    stock = stock & "{""p"":" & NB10JsonString(code) & ",""l"":" & NB10JsonString(loc) & ",""q"":" & NB10Number(qty) & ",""display"":" & NB10JsonString(shelf & " / " & rowNo & " / " & boxNo) & ",""shared"":true,""unassigned"":true}"
                End If
            End If
        End If
    Next r

    payload = "{""revision"":" & NB10JsonString(Format$(Now, "yyyy-mm-dd\Thh:nn:ss")) & ",""products"":[" & products & "],""locations"":[" & locations & "],""stock"":[" & stock & "]}"
    NB10SnapshotBody = "action=warehouseSnapshotPut&key=" & NB10UrlEncodeUtf8(NB10_KEY) & "&payload=" & NB10UrlEncodeUtf8(payload)
End Function

Private Function NB10PhysicalLocationCode(ByVal shelf As String, ByVal rowNo As String, ByVal boxNo As String) As String
    Dim p() As String, s As String, result As String
    If Val(shelf) <= 0 Or Val(rowNo) <= 0 Then Exit Function
    s = NB10NormalizeBox(boxNo)
    If InStr(1, s, "-", vbBinaryCompare) > 0 Then Exit Function
    p = Split(s, ",")
    If UBound(p) > 1 Or Val(p(0)) <= 0 Then Exit Function
    result = "NBR-S" & Format$(Val(shelf), "00") & "-R" & Format$(Val(rowNo), "00") & "-B" & Format$(Val(p(0)), "00")
    If UBound(p) = 1 Then
        If Val(p(1)) <= 0 Then Exit Function
        result = result & "-" & Format$(Val(p(1)), "00")
    End If
    NB10PhysicalLocationCode = result
End Function

Private Function NB10SharedLocationCode(ByVal shelf As String, ByVal rowNo As String, ByVal boxNo As String) As String
    Dim s As String
    s = NB10NormalizeBox(boxNo)
    If Val(shelf) <= 0 Or Val(rowNo) <= 0 Or Len(s) = 0 Then Exit Function
    s = Replace(s, ",", "-")
    NB10SharedLocationCode = "NBR-S" & Format$(Val(shelf), "00") & "-R" & Format$(Val(rowNo), "00") & "-MULTI-" & s
End Function

Private Function NB10NormalizeBox(ByVal v As String) As String
    Dim s As String
    s = Trim$(v)
    s = Replace(s, ChrW$(&H60C), ",")
    s = Replace(s, ".", ",")
    s = Replace(s, " ", "")
    s = Replace(s, ChrW$(8204), "")
    NB10NormalizeBox = s
End Function

Private Function NB10FindSharedRow(ByVal tbl As ListObject, ByVal code As String, ByVal sharedLoc As String) As Long
    Dim p() As String, rowIndex As Long, r As Long, candidate As String
    If UCase$(Left$(Trim$(sharedLoc), 7)) = "SHARED-" Then
        p = Split(sharedLoc, "-")
        rowIndex = Val(p(UBound(p)))
        If rowIndex >= 1 And rowIndex <= tbl.ListRows.Count Then
            If NB10NormalizeKey(tbl.ListColumns(2).DataBodyRange.Cells(rowIndex, 1).Value2) = NB10NormalizeKey(code) Then
                NB10FindSharedRow = rowIndex: Exit Function
            End If
        End If
    End If
    If InStr(1, UCase$(sharedLoc), "-MULTI-", vbTextCompare) > 0 Then
        For r = 1 To tbl.ListRows.Count
            If NB10NormalizeKey(tbl.ListColumns(2).DataBodyRange.Cells(r, 1).Value2) = NB10NormalizeKey(code) Then
                candidate = NB10SharedLocationCode(CStr(tbl.ListColumns(8).DataBodyRange.Cells(r, 1).Value2), _
                    CStr(tbl.ListColumns(9).DataBodyRange.Cells(r, 1).Value2), CStr(tbl.ListColumns(10).DataBodyRange.Cells(r, 1).Value2))
                If NB10NormalizeKey(candidate) = NB10NormalizeKey(sharedLoc) Then NB10FindSharedRow = r: Exit Function
            End If
        Next r
    End If
End Function

Private Function NB10FindLocationRow(ByVal tbl As ListObject, ByVal code As String, ByVal loc As String) As Long
    Dim shelf As String, rowNo As String, boxNo As String, r As Long
    NB10ParseLocation loc, shelf, rowNo, boxNo
    For r = 1 To tbl.ListRows.Count
        If NB10NormalizeKey(tbl.ListColumns(2).DataBodyRange.Cells(r, 1).Value2) = NB10NormalizeKey(code) _
           And NB10NormalizeKey(tbl.ListColumns(8).DataBodyRange.Cells(r, 1).Value2) = NB10NormalizeKey(shelf) _
           And NB10NormalizeKey(tbl.ListColumns(9).DataBodyRange.Cells(r, 1).Value2) = NB10NormalizeKey(rowNo) _
           And NB10NormalizeBox(CStr(tbl.ListColumns(10).DataBodyRange.Cells(r, 1).Value2)) = NB10NormalizeBox(boxNo) Then
            NB10FindLocationRow = r: Exit Function
        End If
    Next r
End Function

Private Function NB10GetLocationRow(ByVal tbl As ListObject, ByVal code As String, ByVal loc As String, Optional ByVal occasion As String = "") As Long
    Dim r As Long
    r = NB10FindLocationRow(tbl, code, loc)
    If r > 0 Then NB10GetLocationRow = r Else NB10GetLocationRow = NB10CreateLocationRow(tbl, code, loc, occasion)
End Function

Private Function NB10CreateLocationRow(ByVal tbl As ListObject, ByVal code As String, ByVal loc As String, Optional ByVal occasion As String = "") As Long
    Dim sourceRow As Long, shelf As String, rowNo As String, boxNo As String, lr As ListRow, r As Long
    For sourceRow = 1 To tbl.ListRows.Count
        If NB10NormalizeKey(tbl.ListColumns(2).DataBodyRange.Cells(sourceRow, 1).Value2) = NB10NormalizeKey(code) Then Exit For
    Next sourceRow
    NB10ParseLocation loc, shelf, rowNo, boxNo
    Set lr = tbl.ListRows.Add: r = lr.Index
    If sourceRow <= tbl.ListRows.Count - 1 Then
        tbl.ListColumns(1).DataBodyRange.Cells(r, 1).Value2 = tbl.ListColumns(1).DataBodyRange.Cells(sourceRow, 1).Value2
    ElseIf Len(Trim$(occasion)) > 0 Then
        tbl.ListColumns(1).DataBodyRange.Cells(r, 1).Value2 = occasion
    Else
        Err.Raise vbObjectError + 716, , "Product code does not exist in Excel and occasion was not supplied."
    End If
    tbl.ListColumns(2).DataBodyRange.Cells(r, 1).NumberFormat = "@"
    tbl.ListColumns(2).DataBodyRange.Cells(r, 1).Value2 = code
    tbl.ListColumns(8).DataBodyRange.Cells(r, 1).Value2 = shelf
    tbl.ListColumns(9).DataBodyRange.Cells(r, 1).Value2 = rowNo
    tbl.ListColumns(10).DataBodyRange.Cells(r, 1).Value2 = boxNo
    NB10CreateLocationRow = r
End Function

Private Sub NB10ParseLocation(ByVal v As String, ByRef shelf As String, ByRef rowNo As String, ByRef boxNo As String)
    Dim p() As String
    p = Split(UCase$(Trim$(v)), "-")
    If UBound(p) < 3 Or p(0) <> "NBR" Or Left$(p(1), 1) <> "S" Or Left$(p(2), 1) <> "R" Or Left$(p(3), 1) <> "B" Then _
        Err.Raise vbObjectError + 717, , "Invalid physical location: " & v
    shelf = CStr(Val(Mid$(p(1), 2)))
    rowNo = CStr(Val(Mid$(p(2), 2)))
    boxNo = CStr(Val(Mid$(p(3), 2)))
    If UBound(p) >= 4 Then boxNo = boxNo & "," & CStr(Val(p(4)))
End Sub

Private Function NB10OperationOccasion(ByVal note As String) As String
    Dim e As Long
    If UCase$(Left$(note, 4)) <> "OCC=" Then Exit Function
    e = InStr(5, note, ";"): If e = 0 Then e = Len(note) + 1
    NB10OperationOccasion = Trim$(Mid$(note, 5, e - 5))
End Function

Private Function NB10MobileCustomer(ByVal opType As String, ByVal note As String) As String
    If opType = NB10U("062E 0631 0648 062C") Then NB10MobileCustomer = Trim$(note)
End Function

Private Function NB10Description(ByVal opType As String, ByVal fromLoc As String, ByVal toLoc As String, ByVal note As String, ByVal beforeQty As Double, ByVal afterQty As Double) As String
    Dim s As String
    s = opType
    If Len(fromLoc) > 0 Then s = s & " | FROM=" & fromLoc
    If Len(toLoc) > 0 Then s = s & " | TO=" & toLoc
    If Len(Trim$(note)) > 0 Then s = s & " | " & note
    NB10Description = s
End Function

Private Sub NB10AppendOutputHistory(ByVal code As String, ByVal qty As Double, ByVal customer As String, ByVal description As String, ByVal operationTime As String, ByVal bankRow As Long, ByVal bankTbl As ListObject)
    Dim h As ListObject, lr As ListRow, d As Date, r As Long, foundBlank As Boolean
    Set h = NB10SheetByTable("WarehouseHistoryTable").ListObjects("WarehouseHistoryTable")
    If Not h.DataBodyRange Is Nothing Then
        For r = 1 To h.ListRows.Count
            If Application.WorksheetFunction.CountA(h.ListRows(r).Range) = 0 Then Set lr = h.ListRows(r): foundBlank = True: Exit For
        Next r
    End If
    If Not foundBlank Then Set lr = h.ListRows.Add
    d = NB10OperationDate(operationTime)
    lr.Range.Cells(1, 1).NumberFormat = "@": lr.Range.Cells(1, 1).Value2 = NB10ShamsiDateTime(d)
    lr.Range.Cells(1, 2).Value2 = customer
    lr.Range.Cells(1, 3).Value2 = bankTbl.ListColumns(1).DataBodyRange.Cells(bankRow, 1).Value2
    lr.Range.Cells(1, 4).NumberFormat = "@": lr.Range.Cells(1, 4).Value2 = code
    lr.Range.Cells(1, 5).Value2 = qty
    lr.Range.Cells(1, 6).Value2 = description
    lr.Range.Cells(1, 7).NumberFormat = "@": lr.Range.Cells(1, 7).Value2 = Format$(d, "yyyymmdd-hhnnss")
    lr.Range.Cells(1, 8).Value2 = Environ$("Username")
End Sub

Private Sub NB10AppendMobileAudit(ByVal opId As String, ByVal opType As String, ByVal code As String, ByVal occasion As String, ByVal qty As Double, ByVal customer As String, ByVal fromLoc As String, ByVal toLoc As String, ByVal beforeQty As Double, ByVal afterQty As Double, ByVal description As String, ByVal operationTime As String)
    Dim ws As Worksheet, r As Long, d As Date
    Set ws = NB10AuditSheet(False)
    If ws Is Nothing Then Exit Sub
    r = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1: If r < 4 Then r = 4
    d = NB10OperationDate(operationTime)
    ws.Cells(r, 1).NumberFormat = "@": ws.Cells(r, 1).Value2 = NB10ShamsiDateTime(d)
    ws.Cells(r, 2).NumberFormat = "@": ws.Cells(r, 2).Value2 = Format$(d, "yyyymmdd-hhnnss")
    ws.Cells(r, 3).Value2 = opType: ws.Cells(r, 4).NumberFormat = "@": ws.Cells(r, 4).Value2 = code
    ws.Cells(r, 5).Value2 = occasion: ws.Cells(r, 6).Value2 = qty: ws.Cells(r, 7).Value2 = customer
    ws.Cells(r, 8).Value2 = fromLoc: ws.Cells(r, 9).Value2 = toLoc: ws.Cells(r, 10).Value2 = beforeQty: ws.Cells(r, 11).Value2 = afterQty
    ws.Cells(r, 12).Value2 = description: ws.Cells(r, 13).Value2 = Environ$("Username")
    ws.Cells(r, 14).NumberFormat = "@": ws.Cells(r, 14).Value2 = opId
End Sub

Private Function NB10AuditSheet(ByVal createIfMissing As Boolean) As Worksheet
    Dim ws As Worksheet, nm As String
    nm = NB10U("0633 0648 0627 0628 0642 0020 0639 0645 0644 06CC 0627 062A 0020 0645 0648 0628 0627 06CC 0644")
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(nm): On Error GoTo 0
    If ws Is Nothing And createIfMissing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If
    Set NB10AuditSheet = ws
End Function

Private Sub NB10Review(ByRef f() As String, ByVal issue As String)
    Dim ws As Worksheet, hit As Range, r As Long, nm As String
    nm = NB10U("0628 0631 0631 0633 06CC 0020 0647 0645 06AF 0627 0645 0020 0633 0627 0632 06CC")
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(nm): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)): ws.Name = nm
        ws.Range("A1:H1").Value = Array("Operation ID", "Review time", "Code", "Type", "Qty", "From", "Status", "Reason")
    End If
    Set hit = ws.Columns(1).Find(What:=f(0), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If hit Is Nothing Then r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1 Else r = hit.Row
    ws.Cells(r, 1).Resize(1, 8).Value = Array(f(0), NB10ShamsiDateTime(Now), f(2), f(1), f(3), f(4), "REVIEW", issue)
End Sub

Private Function NB10SyncSheetName() As String
    NB10SyncSheetName = NB10U("067E 0644 0020 0627 0646 0628 0627 0631")
End Function

Private Sub NB10EnsureSyncSheet()
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(NB10SyncSheetName()): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = NB10SyncSheetName()
        ws.Range("A1:J1").Value = Array("Operation ID", "Applied", "Type", "Code", "Qty", "From", "To", "Note", "Remote time", "Status")
        ws.Visible = xlSheetVeryHidden
    End If
End Sub

Private Function NB10OperationAlreadyApplied(ByVal opId As String) As Boolean
    Dim ws As Worksheet, hit As Range
    Set ws = ThisWorkbook.Worksheets(NB10SyncSheetName())
    Set hit = ws.Columns(1).Find(What:=opId, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
    NB10OperationAlreadyApplied = Not hit Is Nothing
End Function

Private Sub NB10LogApplied(ByRef f() As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(NB10SyncSheetName())
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).NumberFormat = "@": ws.Cells(r, 1).Value2 = f(0)
    ws.Cells(r, 2).Value2 = Now
    ws.Cells(r, 3).Resize(1, 7).Value = Array(f(1), f(2), Val(f(3)), f(4), f(5), f(6), f(7))
    ws.Cells(r, 10).Value2 = "APPLIED"
End Sub

Private Function NB10SheetByTable(ByVal tableName As String) As Worksheet
    Dim ws As Worksheet, tbl As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If StrComp(tbl.Name, tableName, vbTextCompare) = 0 Then Set NB10SheetByTable = ws: Exit Function
        Next tbl
    Next ws
    Err.Raise vbObjectError + 925, , "Table not found: " & tableName
End Function

Private Function NB10NormalizeKey(ByVal v As Variant) As String
    Dim s As String
    s = UCase$(Trim$(CStr(v)))
    s = Replace(s, " ", ""): s = Replace(s, ChrW$(8204), "")
    s = Replace(s, ChrW$(8211), "-"): s = Replace(s, ChrW$(8212), "-")
    NB10NormalizeKey = s
End Function

Private Function NB10NormalizeResponse(ByVal v As String) As String
    NB10NormalizeResponse = Trim$(Replace(v, ChrW$(65279), ""))
End Function

Private Function NB10JsonString(ByVal v As String) As String
    v = Replace(v, "\", "\\"): v = Replace(v, Chr$(34), "\" & Chr$(34))
    v = Replace(v, vbCr, "\r"): v = Replace(v, vbLf, "\n"): v = Replace(v, vbTab, "\t")
    NB10JsonString = Chr$(34) & v & Chr$(34)
End Function

Private Function NB10Number(ByVal v As Double) As String
    NB10Number = Replace(CStr(v), Application.DecimalSeparator, ".")
End Function

Private Function NB10UrlEncodeUtf8(ByVal v As String) As String
    Dim st As Object, bytes As Variant, i As Long, b As Long
    Set st = CreateObject("ADODB.Stream"): st.Type = 2: st.Charset = "utf-8": st.Open: st.WriteText v
    st.Position = 0: st.Type = 1: st.Position = 3: bytes = st.Read: st.Close
    For i = LBound(bytes) To UBound(bytes)
        b = bytes(i)
        If (b >= 48 And b <= 57) Or (b >= 65 And b <= 90) Or (b >= 97 And b <= 122) Or b = 45 Or b = 46 Or b = 95 Or b = 126 Then
            NB10UrlEncodeUtf8 = NB10UrlEncodeUtf8 & Chr$(b)
        Else
            NB10UrlEncodeUtf8 = NB10UrlEncodeUtf8 & "%" & Right$("0" & Hex$(b), 2)
        End If
    Next i
End Function

Private Function NB10OperationDate(ByVal isoText As String) As Date
    Dim s As String, y As Long, m As Long, d As Long, hh As Long, nn As Long, ss As Long
    On Error GoTo Fallback
    s = Trim$(isoText): If Len(s) < 19 Then GoTo Fallback
    y = CLng(Mid$(s, 1, 4)): m = CLng(Mid$(s, 6, 2)): d = CLng(Mid$(s, 9, 2))
    hh = CLng(Mid$(s, 12, 2)): nn = CLng(Mid$(s, 15, 2)): ss = CLng(Mid$(s, 18, 2))
    NB10OperationDate = DateSerial(y, m, d) + TimeSerial(hh, nn, ss) + TimeSerial(3, 30, 0)
    Exit Function
Fallback:
    NB10OperationDate = Now
End Function

Private Function NB10ShamsiDateTime(ByVal dt As Date) As String
    Dim gy As Long, gm As Long, gd As Long, gy2 As Long, jy As Long, jm As Long, jd As Long, days As Long, i As Long
    Dim md As Variant
    gy = Year(dt): gm = Month(dt): gd = Day(dt): md = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    If gy > 1600 Then jy = 979: gy2 = gy - 1600 Else jy = 0: gy2 = gy - 621
    days = 365 * gy2 + Int((gy2 + 3) / 4) - Int((gy2 + 99) / 100) + Int((gy2 + 399) / 400) - 80 + gd
    If gm > 2 Then days = days + 1
    For i = 0 To gm - 2: days = days + CLng(md(i)): Next i
    jy = jy + 33 * Int(days / 12053): days = days Mod 12053
    jy = jy + 4 * Int(days / 1461): days = days Mod 1461
    If days > 365 Then jy = jy + Int((days - 1) / 365): days = (days - 1) Mod 365
    If days < 186 Then jm = 1 + Int(days / 31): jd = 1 + (days Mod 31) Else jm = 7 + Int((days - 186) / 30): jd = 1 + ((days - 186) Mod 30)
    NB10ShamsiDateTime = CStr(jy) & "/" & Format$(jm, "00") & "/" & Format$(jd, "00") & " " & Format$(dt, "hh:nn")
End Function

Private Function NB10U(ByVal hexValues As String) As String
    Dim p() As String, i As Long
    p = Split(hexValues, " ")
    For i = LBound(p) To UBound(p): NB10U = NB10U & ChrW$(CLng("&H" & p(i))): Next i
End Function

Private Sub NB10Cleanup()
    On Error Resume Next
    If nb10Due <> 0 Then Application.OnTime EarliestTime:=nb10Due, Procedure:=nb10Proc, Schedule:=False
    If Not nb10Req Is Nothing Then nb10Req.Abort
    Set nb10Req = Nothing
    nb10Due = 0: nb10Busy = False
End Sub

Private Sub NB10Fail(ByVal description As String)
    Dim msg As String
    msg = "Nebras V10 stage: " & nb10Stage & vbCrLf & description
    If nb10Saved Then msg = msg & vbCrLf & "Excel changes were saved; remote synchronization is incomplete. Do not repeat mobile operations."
    NB10Cleanup
    Application.StatusBar = nb10SavedStatus
    MsgBox msg, vbExclamation, "Nebras"
End Sub
