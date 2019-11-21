VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SBFM-Availability"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCheckAvailability1_Click()
'============================================================================
' Name        : btnCheckAvailability_Click
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call btnCheckAvailability_Click
' Description : compares calendar against current tasks length sum to see if i'm available to complete work within a certain deadline
'============================================================================

Dim iTimeLength As Integer, iAudioLength As Integer, iCalcJobLength As Integer
Dim iWorkTimeAvail As Integer, iAvailTimePreDeadline As Integer
Dim dDeadline As Date, dDateDifference As Date
Dim qdf As QueryDef
Dim rstTasksB4DdlnQ As DAO.Recordset

iTimeLength = 0
iAudioLength = Forms![NewMainMenu]![ProcessJobSubformNMM]![SearchJobsSubform].Form![txtAudioLength].Value
dDeadline = Forms![NewMainMenu]![ProcessJobSubformNMM]![SearchJobsSubform].Form![txtDeadline].Value
iCalcJobLength = (55 + (iAudioLength * 3) + (iAudioLength * 1.5))
dDateDifference = DateDiff("d", Now, dDeadline)
iWorkTimeAvail = 600 * dDateDifference

Set qdf = CurrentDb.QueryDefs("TasksBeforeDeadlineQuery")
qdf.Parameters(0) = dDeadline
Set rstTasksB4DdlnQ = qdf.OpenRecordset()

If Not (rstTasksB4DdlnQ.EOF And rstTasksB4DdlnQ.BOF) Then  'For each CourtDates.ID

    rstTasksB4DdlnQ.MoveFirst
    
    Do Until rstTasksB4DdlnQ.EOF = True
        iTimeLength = iTimeLength + rstTasksB4DdlnQ.Fields("TimeLength").Value
        rstTasksB4DdlnQ.MoveNext
    Loop
    
Else
End If

'when complete
iAvailTimePreDeadline = iWorkTimeAvail - iCalcJobLength - iTimeLength
If iAvailTimePreDeadline >= -30 Then

    Me.txtAvailabilityYN.ControlSource = "=""Yes"""
    
ElseIf iAvailTimePreDeadline < -30 Then

    Me.txtAvailabilityYN.ControlSource = "=""No"""
    
Else
End If
End Sub
