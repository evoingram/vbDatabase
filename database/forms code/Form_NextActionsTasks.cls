VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_NextActionsTasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command15_Click()

If Me.Description.Height = 2200 Then

    Me.Description.Height = 1
    Me.Section(acDetail).Height = 1600
    Me.Description.Visible = False
    
ElseIf Me.Description.Height = 1 Then

    Me.Description.Height = 2200
    Me.Section(acDetail).Height = 4000
    Me.Description.Visible = True
    
Else

    Me.Description.Height = 2200
    Me.Section(acDetail).Height = 4000
    Me.Description.Visible = True
    
End If

Me.Requery

End Sub
