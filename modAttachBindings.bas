Option Explicit

Public Sub Attach_Design_Click()
    ShowAttachForm 0, "Design"
End Sub

Public Sub Attach_OrderEntry_Click()
    Dim r As Long
    r = ActiveCell.Row
    ShowAttachForm r, "Order Entry"
End Sub
