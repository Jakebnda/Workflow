Option Explicit

Public Sub ShowAttachForm(Optional ByVal rowIndex As Variant, Optional ByVal sheetName As Variant)
    Dim ws As Worksheet
    If IsMissing(sheetName) Then
        Set ws = ActiveSheet
    Else
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        On Error GoTo 0
        If ws Is Nothing Then Exit Sub
    End If

    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(1)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If lo.ListRows.Count = 0 Then Exit Sub

    Dim r As Long
    If IsMissing(rowIndex) Or rowIndex <= 0 Then Exit Sub
    r = CLng(rowIndex)

    Dim rowPos As Long
    rowPos = r - lo.DataBodyRange.Row + 1
    If rowPos < 1 Or rowPos > lo.ListRows.Count Then Exit Sub

    Dim wo As Variant
    wo = lo.DataBodyRange.Cells(rowPos, lo.ListColumns(COL_WO).Index).Value

    Dim f As Object
    Set f = VBA.UserForms.Add("frmAttach")
    f.Tag = ws.Name & "|" & CStr(wo)
    f.Show
    Unload f
End Sub
