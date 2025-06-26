Option Explicit

Public Sub UpdateCellHyperlink(ByVal c As Range, ByVal filePath As Variant)
    If IsMissing(filePath) Then Exit Sub
    c.Hyperlinks.Delete
    If Len(CStr(filePath)) = 0 Then
        c.Value = ""
    Else
        c.Value = filePath
        c.Hyperlinks.Add Anchor:=c, Address:=CStr(filePath), TextToDisplay:=CStr(filePath)
    End If
End Sub

Public Function GetHyperlinkAddress(ByVal c As Range) As String
    On Error Resume Next
    If c.Hyperlinks.Count > 0 Then
        GetHyperlinkAddress = c.Hyperlinks(1).Address
    Else
        GetHyperlinkAddress = c.Value
    End If
    On Error GoTo 0
End Function
