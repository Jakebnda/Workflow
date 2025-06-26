Option Explicit

Public Sub AddDesignAttachLinks()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Design")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(1)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If lo.ListRows.Count = 0 Then Exit Sub

    Dim mWs As Worksheet
    On Error Resume Next
    Set mWs = ThisWorkbook.Worksheets("Master")
    On Error GoTo 0
    If mWs Is Nothing Then Exit Sub
    Dim mLo As ListObject
    On Error Resume Next
    Set mLo = mWs.ListObjects(1)
    On Error GoTo 0
    If mLo Is Nothing Then Exit Sub

    Dim cols As Variant
    cols = Array("ProofPath", "EmailPath", "PrintPath")

    Dim r As Range, colName As Variant
    For Each r In lo.DataBodyRange.Rows
        Dim wo As Variant
        wo = r.Cells(1, lo.ListColumns(COL_WO).Index).Value
        If Len(Trim(CStr(wo))) = 0 Then GoTo NextRow
        Dim mRow As Range
        If mLo.ListRows.Count > 0 Then
            Set mRow = mLo.ListColumns(COL_WO).DataBodyRange.Find(wo, LookIn:=xlValues, LookAt:=xlWhole)
            If Not mRow Is Nothing Then
                Set mRow = mLo.DataBodyRange.Rows(mRow.Row - mLo.DataBodyRange.Row + 1)
            End If
        End If
        For Each colName In cols
            Dim c As Range
            Set c = r.Cells(lo.ListColumns(colName).Index)
            Dim masterMissing As Boolean
            masterMissing = True
            If Not mRow Is Nothing Then
                Dim mVal As String
                mVal = GetHyperlinkAddress(mRow.Cells(1, mLo.ListColumns(colName).Index))
                If Len(mVal) > 0 Then masterMissing = False
            End If
            If masterMissing Then
                If Len(c.Value) = 0 And c.Hyperlinks.Count = 0 Then
                    c.Hyperlinks.Add Anchor:=c, Address:="", TextToDisplay:="Attach"
                End If
            End If
        Next colName
NextRow:
    Next r
End Sub
