VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PJStatuses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CourtDatesID_AfterUpdate()
Me.Filter = "CourtDatesID= " & Me.CourtDatesID
Me.FilterOn = True
End Sub
