Option Explicit

Public Sub SyncMasterFromRoleSheet(ByVal ws As Worksheet, ByVal rng As Range)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(1)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    Dim mWs As Worksheet
    Set mWs = ThisWorkbook.Worksheets("Master")
    Dim mLo As ListObject
    On Error Resume Next
    Set mLo = mWs.ListObjects(1)
    On Error GoTo 0
    If mLo Is Nothing Then Exit Sub

    Dim srcMap As Object, destMap As Object
    Set srcMap = CreateObject("Scripting.Dictionary")
    Set destMap = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        srcMap(lo.ListColumns(i).Name) = i
    Next i
    For i = 1 To mLo.ListColumns.Count
        destMap(mLo.ListColumns(i).Name) = i
    Next i

    Dim rw As Range
    If rng Is Nothing Then Exit Sub
    For Each rw In rng.Rows
        Dim wo As Variant
        wo = rw.Cells(1, srcMap(COL_WO)).Value
        If Len(Trim(CStr(wo))) = 0 Then GoTo NextRw
        Dim f As Range
        If mLo.ListRows.Count > 0 Then
            Set f = mLo.ListColumns(COL_WO).DataBodyRange.Find(wo, LookIn:=xlValues, LookAt:=xlWhole)
        Else
            Set f = Nothing
        End If
        Dim destRow As Range
        If f Is Nothing Then
            Set destRow = mLo.ListRows.Add.Range
        Else
            Set destRow = mLo.DataBodyRange.Rows(f.Row - mLo.DataBodyRange.Row + 1)
        End If
        Dim key
        For Each key In srcMap.Keys
            If destMap.Exists(key) Then
                If key = "ProofPath" Or key = "EmailPath" Or key = "PrintPath" Then
                    UpdateCellHyperlink destRow.Cells(destMap(key)), GetHyperlinkAddress(rw.Cells(srcMap(key)))
                Else
                    destRow.Cells(destMap(key)).Value = rw.Cells(srcMap(key)).Value
                End If
            End If
        Next key
        AppendChangeLog wo, "Sync from " & ws.Name
NextRw:
    Next rw
End Sub
