VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SBFMInvShipDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Sub Form_Current()

Dim sParentName As String

On Error Resume Next
sParentName = Me.Parent.Name

If Err <> 0 Then
    GoTo eHandlerX
Else
    On Error GoTo eHandler
    Me.Parent![INVInvTemplate].Requery
End If

eHandlerX:
Exit Sub

eHandler:
Resume eHandlerX

End Sub


