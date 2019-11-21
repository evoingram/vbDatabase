VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PJType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub btnStartWork_Click()
Call pfOpenWork
End Sub

Private Sub Form_Load()

Dim sNotesPath As String, wbViewNotes As String
Dim sRoughDraftPath As String
sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
sNotesPath = "I:\" & sCourtDatesID & "\Notes\" & sCourtDatesID & "-Notes.pdf"
wbViewNotes = "Me.Parent.wbNotes"
Debug.Print Len(Dir(sNotesPath))
If Len(Dir(sNotesPath)) = 0 Then

    MsgBox "Job number " & sCourtDatesID & " doesn't have notes to open." & Chr(13) & "Notes were supposed to be found at " & sNotesPath
    Me.wbNotes.Visible = True
    
Else

    Me.wbNotes.Visible = True
    Me.wbNotes.Navigate2 sNotesPath
    Me.wbNotes.Requery
    
End If

'sRoughDraftPath = "I:\" & sCourtDatesID & "\" & sCourtDatesID & "RoughDraft.docx"
'  Me.oleShowDoc.SourceDoc = sRoughDraftPath
'Call DisplayDoc(Me.oleShowDoc, sRoughDraftPath)


End Sub


  Private Sub Form_Current()
Dim sRoughDraftPath As String

sRoughDraftPath = "I:\" & sCourtDatesID & "\" & sCourtDatesID & "RoughDraft.docx"
 ' Me.oleShowDoc.SourceDoc = sRoughDraftPath
  'Me.oleShowDoc.Requery
  
  End Sub
Public Function DisplayDoc(ctlDocControl As Control, sRoughDraftPath As Variant) As String
On Error GoTo Err_DisplayDoc

Dim strResult As String
Dim strDatabasePath As String
Dim intSlashLocation As Integer
  
  With ctlDocControl
      If Trim(sRoughDraftPath & " ") = "" Then
          .Visible = False
          strResult = "No document name specified."
      Else
          'If InStr(1, sRoughDraftPath, "\") = 0 Then
              ' Path is relative
              'strDatabasePath = CurrentProject.FullName
              'intSlashLocation = InStrRev(strDatabasePath, "\", Len(strDatabasePath))
              'strDatabasePath = Left(strDatabasePath, intSlashLocation)
              'sRoughDraftPath = strDatabasePath & strImagePath
          'End If
          '.Visible = True
          '.Enabled = True
          '.Locked = False
          ' Specify what kind of object can appear in the field.
          .SourceDoc = sRoughDraftPath
          '.Action = acOLECreateLink
          ' Optional size adjustment.
          '.SizeMode = acOLESizeZoom
          strResult = "Document found and displayed."
      End If
  End With
      
Exit_DisplayDoc:
      DisplayDoc = strResult
      Debug.Print strResult
      Exit Function
  
Err_DisplayDoc:
      Select Case Err.Number
          Case 2101       ' Can't find the picture.
              ctlDocControl.Visible = False
              strResult = "Can't find document."
              Resume Exit_DisplayDoc:
          Case Else       ' Some other error.
              MsgBox Err.Number & " " & Err.Description
              Debug.Print Err.Number & " " & Err.Description
              strResult = "An error occurred displaying document."
              Resume Exit_DisplayDoc:
      End Select
  End Function
  
  

Public Function pfOpenWork()


Dim sRoughDraftPath As String, sFTRFolder As String, sCourtSmartFolder As String
Dim sCaseViewerFolder As String, sExpressScribeFolder As String
Dim oFSO As Object, oFolder As Object, oFiles As Object
Dim oSingleFile
'get courtdatesid
sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

sFTRFolder = "C:\Program Files (x86)\FTR\ForTheRecord\TheRecordPlayer.exe"
sAudioFolder = "T:\Production\2InProgress\" & sCourtDatesID & "\Audio\"
sCourtSmartFolder = "T:\Production\2InProgress\" & sCourtDatesID & "\Audio\bin\cplayer.exe"
sCaseViewerFolder = "C:\Program Files (x86)\JAVS\CaseViewer\CaseViewer3.exe"
sExpressScribeFolder = "C:\Program Files (x86)\NCH Software\Scribe\scribe.exe"
sRoughDraftPath = "I:\" & sCourtDatesID & "\" & sCourtDatesID & "RoughDraft.docx"


Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFolder = oFSO.GetFolder(sAudioFolder)
Set oFiles = oFolder.Files


'open correct audio exe
For Each oSingleFile In oFiles

    If oSingleFile.Name Like "*.trm" Then
        Debug.Print "FTR"
        Shell sFTRFolder, vbNormalFocus
        GoTo NextStep
    ElseIf oSingleFile.Name Like "*.csx" Then
        Debug.Print "CourtSmart"
        Shell sCourtSmartFolder, vbNormalFocus
        GoTo NextStep
    ElseIf oSingleFile.Name Like "*.trs" Then
        Debug.Print "FTR"
        Shell sFTRFolder, vbNormalFocus
        GoTo NextStep
    Else
        Debug.Print "CaseViewer/ExpressScribe"
        Shell sCaseViewerFolder, vbNormalFocus
        Shell sExpressScribeFolder, vbNormalFocus
        GoTo NextStep
    End If
        
Next

NextStep:
'open roughdraft doc
Shell "C:\WINDOWS\explorer.exe """ & sRoughDraftPath & "", vbNormalFocus

'open audio folder  'T:\Production\2InProgress\1915\Audio
Shell "C:\WINDOWS\explorer.exe """ & sAudioFolder & "", vbNormalFocus
    

End Function
