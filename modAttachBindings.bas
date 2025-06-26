Option Explicit

Public Sub Attach_Design_Click()
    Dim rowIndex As Long
    If Not ActiveCell Is Nothing Then
        rowIndex = ActiveCell.Row
    Else
        MsgBox "Select a row first", vbExclamation
        Exit Sub
    End If
    ShowAttachForm rowIndex, "Design"
End Sub

Public Sub Attach_OrderEntry_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Order Entry")

    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(1)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    Dim rowIndex As Long
    If Not Intersect(ActiveCell, lo.DataBodyRange) Is Nothing Then
        rowIndex = ActiveCell.Row
    Else
        MsgBox "Select a row in the Order Entry table first", vbExclamation
        Exit Sub
    End If
    ShowAttachForm rowIndex, "Order Entry"
End Sub
