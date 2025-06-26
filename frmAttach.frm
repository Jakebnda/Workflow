VERSION 1.0 CLASS
Begin VB.UserForm frmAttach
   Caption         =   "Attach Files"
   ClientHeight    =   3192
   ClientWidth     =   4680
   Begin VB.Label lblProof
      Caption         =   "Proof File:"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   900
   End
   Begin VB.TextBox txtProof
      Height          =   288
      Left            =   120
      Top             =   240
      Width           =   3000
      TabIndex        =   1
   End
   Begin VB.CommandButton cmdBrowseProof
      Caption         =   "Browse"
      Height          =   288
      Left            =   3240
      Top             =   240
      Width           =   1200
      TabIndex        =   2
   End
   Begin VB.Label lblEmail
      Caption         =   "Email File:"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Width           =   900
   End
   Begin VB.TextBox txtEmail
      Height          =   288
      Left            =   120
      Top             =   720
      Width           =   3000
      TabIndex        =   4
   End
   Begin VB.CommandButton cmdBrowseEmail
      Caption         =   "Browse"
      Height          =   288
      Left            =   3240
      Top             =   720
      Width           =   1200
      TabIndex        =   5
   End
   Begin VB.Label lblPrint
      Caption         =   "Print File:"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   900
   End
   Begin VB.TextBox txtPrint
      Height          =   288
      Left            =   120
      Top             =   1200
      Width           =   3000
      TabIndex        =   7
   End
   Begin VB.CommandButton cmdBrowsePrint
      Caption         =   "Browse"
      Height          =   288
      Left            =   3240
      Top             =   1200
      Width           =   1200
      TabIndex        =   8
   End
   Begin VB.CommandButton cmdOK
      Caption         =   "OK"
      Height          =   288
      Left            =   1800
      Top             =   1800
      Width           =   1200
      TabIndex        =   9
   End
   Begin VB.CommandButton cmdCancel
      Caption         =   "Cancel"
      Height          =   288
      Left            =   1800
      Top             =   2280
      Width           =   1200
      TabIndex        =   10
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
    Dim parts As Variant
    parts = Split(Me.Tag, "|")
    If UBound(parts) >= 1 Then
        sheetName = parts(0)
        workOrder = parts(1)
        Me.Caption = "Attach Files - " & sheetName & " " & workOrder
        LoadExistingPaths
    End If
    cmdOK.Default = True
    cmdCancel.Cancel = True
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
    Dim f As String
    f = BrowseForPDF()
    If Len(f) > 0 Then txtProof.Text = f
End Sub

Private Sub cmdBrowseEmail_Click()
    Dim f As String
    f = BrowseForPDF()
    If Len(f) > 0 Then txtEmail.Text = f
End Sub

Private Sub cmdBrowsePrint_Click()
    Dim f As String
    f = BrowseForPDF()
    If Len(f) > 0 Then txtPrint.Text = f
End Sub

Private Function BrowseForPDF() As String
    Dim f As Variant
    f = Application.GetOpenFilename("PDF Files (*.pdf), *.pdf")
    If f = False Then
        BrowseForPDF = ""
    Else
        BrowseForPDF = CStr(f)
    End If
End Function

Private Sub cmdOK_Click()
    If sheetName = "Design" Then
        UpdateRoleSheet sheetName
    End If
    UpdateRoleSheet "Master"
    Unload Me
End Sub

Private Sub cmdCancel_Click()
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
