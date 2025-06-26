VERSION 1.0 CLASS
Begin VB.UserForm frmAttach
   Caption         =   "Attach Files"
   ClientHeight    =   3192
   ClientWidth     =   4680
   Begin VB.TextBox txtProof
      Height          =   288
      Left            =   120
      Top             =   240
      Width           =   3000
   End
   Begin VB.CommandButton cmdBrowseProof
      Caption         =   "Browse"
      Height          =   288
      Left            =   3240
      Top             =   240
      Width           =   1200
   End
   Begin VB.TextBox txtEmail
      Height          =   288
      Left            =   120
      Top             =   720
      Width           =   3000
   End
   Begin VB.CommandButton cmdBrowseEmail
      Caption         =   "Browse"
      Height          =   288
      Left            =   3240
      Top             =   720
      Width           =   1200
   End
   Begin VB.TextBox txtPrint
      Height          =   288
      Left            =   120
      Top             =   1200
      Width           =   3000
   End
   Begin VB.CommandButton cmdBrowsePrint
      Caption         =   "Browse"
      Height          =   288
      Left            =   3240
      Top             =   1200
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK
      Caption         =   "OK"
      Height          =   288
      Left            =   1800
      Top             =   1800
      Width           =   1200
   End
End
Attribute VB_Name = "frmAttach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sheetName As String
Private workOrder As String

Private Sub UserForm_Initialize()
    Dim parts
    parts = Split(Me.Tag, "|")
    If UBound(parts) >= 1 Then
        sheetName = parts(0)
        workOrder = parts(1)
        LoadExistingPaths
    End If
End Sub

Private Sub LoadExistingPaths()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(1)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If lo.ListRows.Count = 0 Then Exit Sub
    Dim f As Range
    Set f = lo.ListColumns(COL_WO).DataBodyRange.Find(workOrder, LookIn:=xlValues, LookAt:=xlWhole)
    If Not f Is Nothing Then
        Dim row As Range
        Set row = f.EntireRow
        txtProof.Text = GetHyperlinkAddress(row.Cells(lo.ListColumns("ProofPath").Index))
        txtEmail.Text = GetHyperlinkAddress(row.Cells(lo.ListColumns("EmailPath").Index))
        txtPrint.Text = GetHyperlinkAddress(row.Cells(lo.ListColumns("PrintPath").Index))
    End If
    On Error GoTo 0
End Sub

Private Sub cmdBrowseProof_Click()
    Dim f As Variant
    f = Application.GetOpenFilename("PDF Files (*.pdf), *.pdf")
    If f <> False Then txtProof.Text = f
End Sub

Private Sub cmdBrowseEmail_Click()
    Dim f As Variant
    f = Application.GetOpenFilename("PDF Files (*.pdf), *.pdf")
    If f <> False Then txtEmail.Text = f
End Sub

Private Sub cmdBrowsePrint_Click()
    Dim f As Variant
    f = Application.GetOpenFilename("PDF Files (*.pdf), *.pdf")
    If f <> False Then txtPrint.Text = f
End Sub

Private Sub cmdOK_Click()
    If sheetName = "Design" Then
        UpdateRoleSheet sheetName
    End If
    UpdateRoleSheet "Master"
    Unload Me
End Sub

Private Sub UpdateRoleSheet(ByVal targetSheet As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(targetSheet)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(1)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If lo.ListRows.Count = 0 Then Exit Sub
    Dim f As Range
    Set f = lo.ListColumns(COL_WO).DataBodyRange.Find(workOrder, LookIn:=xlValues, LookAt:=xlWhole)
    If Not f Is Nothing Then
        Dim row As Range
        Set row = lo.DataBodyRange.Rows(f.Row - lo.DataBodyRange.Row + 1)
        UpdateCellHyperlink row.Cells(lo.ListColumns("ProofPath").Index), txtProof.Text
        UpdateCellHyperlink row.Cells(lo.ListColumns("EmailPath").Index), txtEmail.Text
        UpdateCellHyperlink row.Cells(lo.ListColumns("PrintPath").Index), txtPrint.Text
        AppendChangeLog workOrder, "Attachments updated on " & targetSheet
    End If
End Sub
