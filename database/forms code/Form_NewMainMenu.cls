VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_NewMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub bCheckFTPFiles_Click()

Call pfDownloadfromFTP
End Sub

Private Sub bEnterNewJob_Click()

Call pfEnterNewJob

End Sub

Private Sub bPriceQuote_Click()
Forms![NewMainMenu]![ProcessJobSubformNMM].SourceObject = "MMPriceQuote"
Forms![NewMainMenu]![ProcessJobSubformNMM].Requery

End Sub


Private Sub bScheduling1_Click()
Forms![NewMainMenu]![ProcessJobSubformNMM].SourceObject = "TaskMgmtForm"
Forms![NewMainMenu]![ProcessJobSubformNMM].Requery
End Sub

Private Sub btnSearchCitations_Click()
'search jobs button MMSearch Jobs
'prcess a job button
Forms![NewMainMenu]![ProcessJobSubformNMM].SourceObject = "MMSearch Citations"
Forms![NewMainMenu]![ProcessJobSubformNMM].Requery
End Sub

Private Sub bTranscriptPrdn_Click()
 'process a job button
Forms![NewMainMenu]![ProcessJobSubformNMM].SourceObject = "MMProcess Jobs"
Forms![NewMainMenu]![ProcessJobSubformNMM].Requery
End Sub
Private Sub bUntaggedCmcn_Click()
'communication bucket button CommunicationBucket
'prcess a job button
Forms![NewMainMenu]![ProcessJobSubformNMM].SourceObject = "CommunicationBucket"
Forms![NewMainMenu]![ProcessJobSubformNMM].Requery
End Sub
Private Sub bFinances_Click()
'prcess a job button
svURL = "https://go.xero.com/AccountsReceivable/Search.aspx?invoiceStatus=INVOICESTATUS%2fDRAFT&graphSearch=False"
Forms![NewMainMenu]![ProcessJobSubformNMM].SourceObject = "MMFinances"
Forms![NewMainMenu]![ProcessJobSubformNMM].Requery
'Application.FollowHyperlink (svURL)
End Sub
Private Sub bSearchJobs_Click()
'search jobs button MMSearch Jobs
'prcess a job button
Forms![NewMainMenu]![ProcessJobSubformNMM].SourceObject = "MMSearch Jobs"
Forms![NewMainMenu]![ProcessJobSubformNMM].Requery
End Sub

Private Sub bBarAddresses_Click()

Call pfScrapingBALoop(InputBox("Enter begining bar number"), InputBox("Enter ending bar number"))

End Sub

Private Sub bPaypalPayment_Click()


Call fApplyPayPalPayment

End Sub

Private Sub bPriorityPoints_Click()

Call pfPriorityPointsAlgorithm
End Sub

Private Sub bSpeechRecognition_Click()
Forms![NewMainMenu]![ProcessJobSubformNMM].SourceObject = "Speech"
Forms![NewMainMenu]![ProcessJobSubformNMM].Requery
End Sub

Private Sub bRoman_Click()
Forms![NewMainMenu]![ProcessJobSubformNMM].SourceObject = "RomanNConv"
Forms![NewMainMenu]![ProcessJobSubformNMM].Requery
End Sub
Private Sub bScheduling_Click()
Forms![NewMainMenu]![ProcessJobSubformNMM].SourceObject = "MMScheduling"
Forms![NewMainMenu]![ProcessJobSubformNMM].Requery
End Sub

Private Sub bEnterPayment_Click()
Forms![NewMainMenu]![ProcessJobSubformNMM].SourceObject = "fPayments"
Forms![NewMainMenu]![ProcessJobSubformNMM].Requery
End Sub

Private Sub bReset_Click()
'unused timer
Me.TimerInterval = 0
Me![lblTime].Value = "00:00:00"
lngNumOfSecs = 0
End Sub

Private Sub bStop_Click()
'unused timer
Me.TimerInterval = 0
End Sub

Private Sub bStart_Click()
'unused timer
Me.TimerInterval = 1000
End Sub
Private Sub Form_AfterUpdate()
'unused timer
Forms![NewMainMenu]![lblTime].Text = sTime
Forms![NewMainMenu]![lblTime].Caption = sTime
Forms![NewMainMenu]![lblTime].Value = sTime
End Sub
Public Sub Form_Timer()
'unused timer
lngNumOfSecs = lngNumOfSecs + 1


Select Case lngNumOfSecs
  Case Is > 86400        '>1 day - not equipped for that
  Case Is >= 3600        '>1 hour
    lngNumOfHrs = lngNumOfSecs \ 3600
    lngNumOfMins = ((lngNumOfSecs - (lngNumOfHrs * 3600)) \ 60)
    lngNumOfSecsRem = lngNumOfSecs - ((lngNumOfHrs * 3600) + (lngNumOfMins * 60))
  Case Is >= 60          '>1 minute
    lngNumOfMins = ((lngNumOfSecs - (lngNumOfHrs * 3600)) \ 60)
    lngNumOfSecsRem = lngNumOfSecs - ((lngNumOfHrs * 3600) + (lngNumOfMins * 60))
  Case Is > 0            '< 1 minute
    lngNumOfSecsRem = lngNumOfSecs - ((lngNumOfHrs * 3600) + (lngNumOfMins * 60))
  Case Else              'shouldn't happen, but who knows?
End Select

sTime = Format$(lngNumOfHrs, "00") & ":" & Format$(lngNumOfMins, "00") & _
                             ":" & Format$(lngNumOfSecsRem, "00")
                             
 Forms![NewMainMenu]![lblTime].Value = Format$(lngNumOfHrs, "00") & ":" & Format$(lngNumOfMins, "00") & _
                             ":" & Format$(lngNumOfSecsRem, "00")
  
                         
                             
                             
End Sub
Private Sub NewMainMenu_Timer()
'unused timer
lngNumOfSecs = lngNumOfSecs + 1

Select Case lngNumOfSecs
  Case Is > 86400        '>1 day - not equipped for that
  Case Is >= 3600        '>1 hour
    lngNumOfHrs = lngNumOfSecs \ 3600
    lngNumOfMins = ((lngNumOfSecs - (lngNumOfHrs * 3600)) \ 60)
    lngNumOfSecsRem = lngNumOfSecs - ((lngNumOfHrs * 3600) + (lngNumOfMins * 60))
  Case Is >= 60          '>1 minute
    lngNumOfMins = ((lngNumOfSecs - (lngNumOfHrs * 3600)) \ 60)
    lngNumOfSecsRem = lngNumOfSecs - ((lngNumOfHrs * 3600) + (lngNumOfMins * 60))
  Case Is > 0            '< 1 minute
    lngNumOfSecsRem = lngNumOfSecs - ((lngNumOfHrs * 3600) + (lngNumOfMins * 60))
  Case Else              'shouldn't happen, but who knows?
End Select
sTime = Format$(lngNumOfHrs, "00") & ":" & Format$(lngNumOfMins, "00") & _
                             ":" & Format$(lngNumOfSecsRem, "00")
Forms![NewMainMenu]![lblTime].Text = Format$(lngNumOfHrs, "00") & ":" & Format$(lngNumOfMins, "00") & _
                             ":" & Format$(lngNumOfSecsRem, "00")
                             
      

Forms![NewMainMenu]![lblTime].Value = Format$(lngNumOfHrs, "00") & ":" & Format$(lngNumOfMins, "00") & _
                             ":" & Format$(lngNumOfSecsRem, "00")
                             
End Sub


