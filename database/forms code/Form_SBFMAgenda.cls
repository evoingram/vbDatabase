VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SBFMAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub List66_AfterUpdate()

'    CurrentDb.Execute "DELETE FROM TempTasksDay", dbFailOnError
'OrderGroupIncompleteTasks
'Do while rSumTimeLength<=600
    'For each row in '=Int([TimeLength])
    '"GroupTasksIncompleteProduction Query"

Dim rstGroupTasksICPQ As DAO.Recordset

Set rstGroupTasksICPQ = CurrentDb.OpenRecordset("SELECT * FROM GroupTasksIncompleteProduction Query")

If Me.List66.Value = "Today" Then

    Me.Text68.ControlSource = "=Now"
    Me.RecordSource = "TempTasksDay1"
    Me.Requery
    
ElseIf Me.List66.Value = "Tomorrow" Then

    Me.Text68.ControlSource = "=Now+1"
    Me.RecordSource = "TempTasksDay2"
    Me.Requery
    
ElseIf Me.List66.Value = "Day 3" Then

    Me.Text68.ControlSource = "=Now+2"
    Me.RecordSource = "TempTasksDay3"
    Me.Requery
    
ElseIf Me.List66.Value = "Day 4" Then

    Me.Text68.ControlSource = "=Now+3"
    Me.RecordSource = "TempTasksDay4"
    Me.Requery
    
ElseIf Me.List66.Value = "Day 5" Then

    Me.Text68.ControlSource = "=Now+4"
    Me.RecordSource = "TempTasksDay5"
    Me.Requery
    
ElseIf Me.List66.Value = "Day 6" Then

   Me.Text68.ControlSource = "=Now+5"
    Me.RecordSource = "TempTasksDay6"
    Me.Requery
    
ElseIf Me.List66.Value = "Day 7" Then

    Me.Text68.ControlSource = "=Now+6"
    Me.RecordSource = "TempTasksDay7"
    Me.Requery
    
ElseIf Me.List66.Value = "Day 8" Then

    Me.Text68.ControlSource = "=Now+7"
    Me.RecordSource = "TempTasksDay8"
    Me.Requery

Else

End If

If Not (rstGroupTasksICPQ.EOF And rstGroupTasksICPQ.BOF) Then

    rstGroupTasksICPQ.MoveFirst
    
    Do Until rstGroupTasksICPQ.EOF = True
    
        rstGroupTasksICPQ.Edit
        rstGroupTasksICPQ!VendorYN = True
        rstGroupTasksICPQ("VendorYN") = True
        rstGroupTasksICPQ.Update

        sContactName = rstGroupTasksICPQ!FirstName & " " & rstGroupTasksICPQ!LastName

        rstGroupTasksICPQ.MoveNext
        
    Loop
    
Else

    MsgBox "There are no records in the recordset."
    
End If

MsgBox "Finished looping through records."

rstGroupTasksICPQ.Close
Set rstGroupTasksICPQ = Nothing

End Sub

