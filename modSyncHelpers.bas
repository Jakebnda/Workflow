Option Explicit

Public Sub ClearOrderEntryRows(ByVal sh As Worksheet, ByVal rng As Range)
    rng.Rows.ClearContents
End Sub

Public Sub AppendChangeLog(ByVal wo As Variant, ByVal msg As String)
    Dim logWs As Worksheet
    On Error Resume Next
    Set logWs = ThisWorkbook.Worksheets("ChangeLog")
    On Error GoTo 0
    If logWs Is Nothing Then Exit Sub
    Dim lr As Long
    lr = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row + 1
    logWs.Cells(lr, 1).Value = Now
    logWs.Cells(lr, 2).Value = wo
    logWs.Cells(lr, 3).Value = msg
End Sub

Public Sub RefreshStageSheet(ByVal sheetName As String)
    Dim master As Worksheet
    Set master = ThisWorkbook.Worksheets("Master")
    Dim mLo As ListObject
    On Error Resume Next
    Set mLo = master.ListObjects(1)
    On Error GoTo 0
    If mLo Is Nothing Then Exit Sub

    Dim target As Worksheet
    On Error Resume Next
    Set target = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If target Is Nothing Then Exit Sub
    Dim tLo As ListObject
    On Error Resume Next
    Set tLo = target.ListObjects(1)
    On Error GoTo 0
    If tLo Is Nothing Then Exit Sub

    If Not tLo.DataBodyRange Is Nothing Then tLo.DataBodyRange.ClearContents
    If mLo.ListRows.Count = 0 Then Exit Sub
    mLo.Range.AutoFilter Field:=mLo.ListColumns("Stage").Index, Criteria1:=sheetName
    Dim vis As Range
    On Error Resume Next
    Set vis = mLo.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If Not vis Is Nothing Then
        vis.Columns(1).Resize(, 9).Copy tLo.DataBodyRange.Cells(1, 1)
    End If
    If master.AutoFilterMode Then mLo.AutoFilter.ShowAllData
    If sheetName = "Design" Then AddDesignAttachLinks
End Sub
