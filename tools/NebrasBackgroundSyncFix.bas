Attribute VB_Name = "NebrasBackgroundSyncFix"
Option Explicit

' Surgical fix: the warehouse sync button starts only the async bridge.
' It does NOT repair formulas, calculate the whole workbook, or save immediately.

Public Sub InstallNebrasBackgroundSyncFix()
    Dim ws As Worksheet
    Dim shp As Shape

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Set shp = Nothing
        Set shp = ws.Shapes("btnNebrasWarehouseSync")
        If Not shp Is Nothing Then Exit For
    Next ws
    On Error GoTo Failed

    If shp Is Nothing Then
        MsgBox "دکمه همگام‌سازی موبایل پیدا نشد.", vbExclamation, "نبراس"
        Exit Sub
    End If

    shp.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!NebrasBackgroundSync"
    MsgBox "دکمه همگام‌سازی به حالت پس‌زمینه متصل شد.", vbInformation, "نبراس"
    Exit Sub

Failed:
    MsgBox "اتصال دکمه انجام نشد:" & vbCrLf & Err.Description, vbCritical, "نبراس"
End Sub

Public Sub NebrasBackgroundSync()
    On Error GoTo Failed
    Application.Run "SyncBridgeFA_V7"
    Exit Sub
Failed:
    MsgBox "شروع همگام‌سازی انجام نشد:" & vbCrLf & Err.Description, vbCritical, "نبراس"
End Sub
