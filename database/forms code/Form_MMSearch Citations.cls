VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MMSearch Citations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command40_Click()
Me.SearchJobsSubform.SourceObject = "Query." & "SCB-CitationHyperlinks"
'Me.SearchJobsSubform.Requery
End Sub

Private Sub SearchQueryBtnC_Click()
Me.SearchJobsSubform.SourceObject = "Query." & "SCB-Company"
Me.SearchJobsSubform.Requery
End Sub
Private Sub SearchQueryBtnAFN_Click()
Me.SearchJobsSubform.SourceObject = "Query." & "SCB-AttorneyFirstName"
Me.SearchJobsSubform.Requery
End Sub
Private Sub SearchQueryBtnOAFN_Click()
Me.SearchJobsSubform.SourceObject = "Query." & "SCB-OrderingAttorneysFirstName"
Me.SearchJobsSubform.Requery
End Sub
Private Sub SearchQueryBtnOALN_Click()
Me.SearchJobsSubform.SourceObject = "Query." & "SCB-OrderingAttorneysLastName"
Me.SearchJobsSubform.Requery
End Sub

Private Sub SearchQueryBtnOAC_Click()
Me.SearchJobsSubform.SourceObject = "Query." & "SCB-OrderingAttorneysCompany"
Me.SearchJobsSubform.Requery
End Sub

Private Sub SearchQueryBtnE_Click()
Me.SearchJobsSubform.SourceObject = "Query." & "SCB-Email"
Me.SearchJobsSubform.Requery
End Sub
Private Sub SearchQueryBtnHT_Click()
Me.SearchJobsSubform.SourceObject = "Query." & "SCB-HearingTitle"
Me.SearchJobsSubform.Requery
End Sub
Private Sub SearchQueryBtnCN_Click()
Me.SearchJobsSubform.SourceObject = "Query." & "SCB-CaseNumber"
Me.SearchJobsSubform.Requery
End Sub
Private Sub SearchQueryBtnJN_Click()
Me.SearchJobsSubform.SourceObject = "Query." & "SCB-JudgeName"
Me.SearchJobsSubform.Requery
End Sub
Private Sub SearchQueryBtnPoD_Click()
Me.SearchJobsSubform.SourceObject = "Query." & "SCB-PlaintiffDefendants"
Me.SearchJobsSubform.Requery
End Sub
Private Sub SearchQueryBtnD_Click()
Me.SearchJobsSubform.SourceObject = "Query." & "SCB-Defendants"
Me.SearchJobsSubform.Requery
End Sub
Private Sub SearchQueryBtnJ_Click()
Me.SearchJobsSubform.SourceObject = "Query." & "SCB-Jurisdiction"
Me.SearchJobsSubform.Requery
End Sub
Private Sub SearchQueryBtnP_Click()
Me.SearchJobsSubform.SourceObject = "Query." & "SCB-Plaintiffs"
Me.SearchJobsSubform.Requery
End Sub

