VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SBFMUncompletedStatuses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Combo57_AfterUpdate()
Me.Requery
End Sub

Private Sub ContactsEntered_AfterUpdate()
Me.Requery
End Sub

Private Sub CourtDatesID_AfterUpdate()


End Sub

Private Sub CoverPage_AfterUpdate()
Me.Requery
End Sub

Private Sub Form_AfterUpdate()
Me.Requery
End Sub
