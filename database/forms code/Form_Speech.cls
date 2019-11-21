VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Speech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnUSPSTracking_Click()


Call pfPrepareAudio
End Sub

Private Sub Command174_Click()


Call pfCopyTranscriptFromCompletedToPrepared
'Call RenameBaseFiles
End Sub

Private Sub Command177_Click()


Call pfSplitAudio
End Sub

Private Sub Command183_Click()


Call pfCorpusUpload
End Sub

Private Sub Command7_Click()


Call pfPrepareTranscript
End Sub

Private Sub Command8_Click()


Call pfTrainAudio
End Sub

Private Sub Command9_Click()


Call pfSRTranscribe
End Sub
