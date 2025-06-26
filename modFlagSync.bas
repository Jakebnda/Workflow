Option Explicit

Public Sub SyncFlagToMaster(ByVal sh As Worksheet, ByVal Target As Range)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = sh.ListObjects(1)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    Dim mLo As ListObject
    On Error Resume Next
    Set mLo = ThisWorkbook.Worksheets("Master").ListObjects(1)
    On Error GoTo 0
    If mLo Is Nothing Then Exit Sub
    If lo.ListRows.Count = 0 Then Exit Sub

    Dim colName As String
    colName = lo.HeaderRowRange.Cells(1, Target.Column - lo.Range.Column + 1).Value
    Dim prefix As String
    prefix = sh.Name & "_"
    If Left(colName, Len(prefix)) <> prefix Then Exit Sub
    Dim flagName As String
    flagName = Mid(colName, Len(prefix) + 1)

    Dim destIndex As Long
    On Error Resume Next
    destIndex = mLo.ListColumns(flagName).Index
    On Error GoTo 0
    If destIndex = 0 Then Exit Sub

    Dim rowIdx As Long
    rowIdx = Target.Row - lo.DataBodyRange.Row + 1
    Dim wo As Variant
    wo = lo.DataBodyRange.Cells(rowIdx, lo.ListColumns(COL_WO).Index).Value
    If Len(Trim(CStr(wo))) = 0 Then Exit Sub

    Dim f As Range
    Set f = mLo.ListColumns(COL_WO).DataBodyRange.Find(wo, LookIn:=xlValues, LookAt:=xlWhole)
    If Not f Is Nothing Then
        mLo.DataBodyRange.Rows(f.Row - mLo.DataBodyRange.Row + 1).Cells(1, destIndex).Value = Target.Value
    End If
End Sub
