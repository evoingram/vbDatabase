VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_TaskMgmtForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnViewReport_Click()
'============================================================================
' Name        : btnViewReport_Click
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call btnViewReport_Click
' Description : generates task list from combo box selection
'============================================================================


If Me.cmbListAllDocuments = "Stage 1" Then
    Me.YearlySalesReport.SourceObject = "SBFM-TaskMgmtQS1"
    
ElseIf Me.cmbListAllDocuments = "Stage 2" Then
    Me.YearlySalesReport.SourceObject = "SBFM-TaskMgmtQS2"
    
ElseIf Me.cmbListAllDocuments = "Stage 3" Then
    Me.YearlySalesReport.SourceObject = "SBFM-TaskMgmtQS3"
        
ElseIf Me.cmbListAllDocuments = "Stage 4" Then
    Me.YearlySalesReport.SourceObject = "SBFM-TaskMgmtQS4"
Else
End If

End Sub

Private Sub Command44_Click()
Call fCompleteTimeMgmtTasks
End Sub

Private Sub Command45_Click()
Call fCompleteStatusBoxes
End Sub

Private Sub Command46_Click()
Call fCompleteStage1Tasks

End Sub


Private Sub Command47_Click()
Call fCompleteStage2Tasks

End Sub

Private Sub Command48_Click()
Call fCompleteStage3Tasks

End Sub
Private Sub Command49_Click()
Call fCompleteStage4Tasks

End Sub

Private Sub Command50_Click()
Call fUnCompleteTimeMgmtTasks
End Sub

Private Sub Command73_Click()

Forms![NewMainMenu]![ProcessJobSubformNMM].SourceObject = "fPayments"
Forms![NewMainMenu]![ProcessJobSubformNMM].Requery
End Sub


