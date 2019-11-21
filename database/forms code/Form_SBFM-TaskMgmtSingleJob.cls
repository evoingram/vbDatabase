VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SBFM-TaskMgmtSingleJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command15_Click()
Dim intWindowHeight As Integer
Dim intWindowWidth As Integer
Dim intTotalFormHeight As Integer
Dim intTotalFormWidth As Integer
Dim intHeightHeader As Integer
Dim intHeightDetail As Integer
Dim intHeightFooter As Integer
If Me.Text44.Height = 2200 Then
    Me.Text44.Height = 1
    Me.Section(acDetail).Height = 1600
    Me.Text44.Visible = False
ElseIf Me.Text44.Height = 1 Then
    Me.Text44.Height = 2200
    Me.Section(acDetail).Height = 4000
    Me.Text44.Visible = True
Else
    Me.Text44.Height = 2200
    Me.Section(acDetail).Height = 4000
    Me.Text44.Visible = True
End If

Me.Requery

End Sub
