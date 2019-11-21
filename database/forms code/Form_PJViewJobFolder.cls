VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PJViewJobFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub WBCViewJobFolder_OnUpdated()
Dim sTranscriptDir As String, sJobNumberDirectory As String, sAppName As String

sJobNumberDirectory = "\\hubcloud\evoingram\production\2inprogress\" & Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField] & "\"
sTranscriptDir = "file:///" & sJobNumberDirectory
sAppName = "C:\Windows\explorer.exe " & sTranscriptDir 'used to be sJobNumberDirectory but probably wrong

Me.WBCViewJobFolder.Navigate2 sTranscriptDir

End Sub
Private Sub Command5_Click()

Dim sTranscriptDir As String, sJobNumberDirectory As String, sAppName As String

sJobNumberDirectory = "\\hubcloud\evoingram\production\2inprogress\" & Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField] & "\"
sTranscriptDir = "file:///" & sJobNumberDirectory
sAppName = "C:\Windows\explorer.exe " & sJobNumberDirectory 'used to be sJobNumberDirectory but probably wrong

Me.WBCViewJobFolder.Navigate2 sTranscriptDir

End Sub
