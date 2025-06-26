Option Explicit

Public Sub UpdateOrderEntry()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Order Entry")
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(1)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If lo.ListRows.Count = 0 Then Exit Sub

    SyncMasterFromRoleSheet ws, lo.DataBodyRange
    ClearOrderEntryRows ws, lo.DataBodyRange
    AppendChangeLog "", "Order Entry synced"

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim mLo As ListObject
    Set mLo = ThisWorkbook.Worksheets("Master").ListObjects(1)
    Dim c As Range
    For Each c In mLo.ListColumns("Stage").DataBodyRange
        If Not dict.Exists(c.Value) Then dict.Add c.Value, 1
    Next c
    Dim stage
    For Each stage In dict.Keys
        RefreshStageSheet CStr(stage)
    Next stage
End Sub
