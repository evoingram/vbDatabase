VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MMScheduling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command3_Click()


Call pfPriorityPointsAlgorithm
End Sub

Private Sub Command40_Click()
Me.SearchJobsSubform.Visible = True
Me.TaskSchedule.Visible = False
Me.SearchJobsSubform.SourceObject = "SBFM-Availability"
Me.TaskSchedule.Requery
Me.SearchJobsSubform.Requery
End Sub

Private Sub SearchQueryBtnC_Click()
Me.SearchJobsSubform.Visible = False
Me.SearchJobsSubform.Requery
Me.TaskSchedule.Visible = True
Me.TaskSchedule.SourceObject = "Report.Tasks Scheduled"
Me.TaskSchedule.Requery
End Sub
Private Sub SearchQueryBtnAFN_Click()
Me.SearchJobsSubform.Visible = True
Me.TaskSchedule.Visible = False
Me.SearchJobsSubform.SourceObject = "NextActionsTasks"
Me.TaskSchedule.Requery
Me.SearchJobsSubform.Requery
End Sub
Private Sub SearchQueryBtnOAFN_Click()
Me.SearchJobsSubform.Visible = True
Me.TaskSchedule.Visible = False
Me.SearchJobsSubform.SourceObject = "TaskMgmtForm"
Me.TaskSchedule.Requery
Me.SearchJobsSubform.Requery
End Sub

