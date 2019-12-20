Attribute VB_Name = "Stage2"
'@Folder("Database.Production.Modules")
Option Compare Database
Option Explicit

'============================================================================
'module Stage2

'variables:
'   NONE

'functions:

'pfStage2Ppwk:                               Description:  completes all stage 2 tasks
'                                        Arguments:    NONE
'pfAutoCorrect:                              Description:  adds entries as listed on form to rough draft autocorrect in Word
'                                        Arguments:    NONE
'pfRoughDraftToCoverF:                       Description:  Adds rough draft to courtcover
'                                                      does find/replacements of static speakers 1-17, all dynamic speakers, Q&A, : a-z, various AQC & AVT headings
'                                        Arguments:    NONE
'pfStaticSpeakersFindReplace:                Description:  finds and replaces static speakers in CourtCover after rough draft is inserted
'                                        Arguments:    NONE
'pfReplaceColonUndercasewithColonUppercase:  Description:  replaces : a-z with : A-Z, applies styles to fixed phrases in transcript
'                                        Arguments:    NONE
'pfTypeRoughDraftF:                          Description:  copies correct roughdraft template to job folder
'                                        Arguments:    NONE
'pfReplaceWeberOR:                           Description:  Adds rough draft to courtcover
'                                                      does find/replacements of static speakers 1-17, all dynamic speakers, Q&A, : a-z, various AQC & Weber headings
'                                        Arguments:    NONE
'pfReplaceWeberNV:                           Description:  Adds rough draft to courtcover
'                                                      does find/replacements of static speakers 1-17, all dynamic speakers, Q&A, : a-z, various AQC & Weber headings
'                                        Arguments:    NONE
'pfReplaceWeberBR:                           Description:  Adds rough draft to courtcover
'                                                      does find/replacements of static speakers 1-17, all dynamic speakers, Q&A, : a-z, various AQC & Weber headings
'                                        Arguments:    NONE
'pfReplaceAVT:                               Description:  Adds rough draft to courtcover
'                                                      does find/replacements of static speakers 1-17, all dynamic speakers, Q&A, : a-z, various AQC & AVT headings
'                                        Arguments:    NONE
'pfReplaceAQC:                               Description:  Adds rough draft to courtcover
'                                                      does find/replacements of static speakers 1-17, all dynamic speakers, Q&A, : a-z, various AQC & AVT headings
'                                        Arguments:    NONE
        
'============================================================================

Public Sub pfStage2Ppwk()
    '============================================================================
    ' Name        : pfStage2Ppwk
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfStage2Ppwk
    ' Description : completes all stage 2 tasks
    '============================================================================

    Dim sAnswer As String
    Dim sQuestion As String
    Dim cJob As New Job

    'refresh transcript info
    Call pfCheckFolderExistence                  'checks for job folder and creates it if not exists
    Call pfTypeRoughDraftF                       'Add RD template to job folder
    Call pfUpdateCheckboxStatus("AddRDtoCover")
    Call pfUpdateCheckboxStatus("FindReplaceRD")
    Call pfUpdateCheckboxStatus("Transcribe")

    Call pfCurrentCaseInfo

    If sJurisdiction = "*AVT*" Then

        Call pfReplaceAVT
        MsgBox "Stage 2 complete."
        Application.FollowHyperlink cJob.DocPath.CourtCover
    
    ElseIf sJurisdiction Like "FDA" Then

        Call pfReplaceFDA
        Application.FollowHyperlink cJob.DocPath.CourtCover
        
    ElseIf sJurisdiction Like "Food and Drug Administration" Then

        Call pfReplaceFDA
        Application.FollowHyperlink cJob.DocPath.CourtCover

    ElseIf sJurisdiction Like "Weber Oregon" Then

        Call wwReplaceWeberOR
        Call FPJurors
        MsgBox "Stage 2 complete."
        Application.FollowHyperlink cJob.DocPath.RoughDraft

    ElseIf sJurisdiction Like "Weber Bankruptcy" Then

        Call wwReplaceWeberBR
        Application.FollowHyperlink cJob.DocPath.RoughDraft

    ElseIf sJurisdiction Like "Weber Nevada" Then

        Call wwReplaceWeberNV
        Application.FollowHyperlink cJob.DocPath.RoughDraft
 
    ElseIf sJurisdiction Like "Massachusetts" Then

        Call pfReplaceMass
        Application.FollowHyperlink cJob.DocPath.RoughDraft
       
    Else

        Call pfReplaceAQC
    
    End If

    sQuestion = "Need to send an information-needed e-mail?" 'information needed email prompt
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    If sAnswer = vbNo Then                       'Code for No

        GoTo EndIf1
    
    Else                                         'Code for yes
    
        Call pfCheckFolderExistence              'checks for job folder and creates it if not exists
        Call pfSendWordDocAsEmail("InfoNeeded", "Spellings/Information Needed")
        Call pfCommunicationHistoryAdd("InfoNeeded") 'save in commhistory
        'Call fInfoNeededEmailF
        Call pfUpdateCheckboxStatus("SpellingsEmail")
    
EndIf1:
    End If

    MsgBox "Stage 2 complete."

    Call pfCurrentCaseInfo                       'refresh transcript info
    Application.FollowHyperlink cJob.DocPath.CourtCover
    Call pfClearGlobals
End Sub

Public Sub pfAutoCorrect()
    '============================================================================
    ' Name        : pfAutoCorrect
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfAutoCorrect
    ' Description : adds entries as listed on form to rough draft autocorrect in Word
    '============================================================================

    Dim db As Database
    Dim flCurrentField As DAO.Field
    Dim sFieldName As String
    Dim sACShortcutsSQL As String
    Dim sFieldValue As String
    Dim rstAGShortcuts As DAO.Recordset
    Dim oWordDoc As Word.Document
    Dim oWordApp As Word.Application
    Dim cJob As New Job

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

    sACShortcutsSQL = "SELECT * FROM AGShortcuts WHERE [CourtDatesID] = " & sCourtDatesID & ";"

    Set db = CurrentDb
    Set rstAGShortcuts = db.OpenRecordset(sACShortcutsSQL)



    On Error Resume Next

    Set oWordApp = GetObject(, "Word.Application")
    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    'oWordApp.Visible = True
    Set oWordDoc = GetObject(cJob.DocPath.RoughDraft, "Word.Document")

    With oWordDoc                                'insert rough draft at RoughBKMK bookmark


        For Each flCurrentField In rstAGShortcuts.Fields
    
            sFieldName = LCase(flCurrentField.Name)
        
            If sFieldName = "CourtDatesID" Then
                GoTo NextAGShortcut
            
            ElseIf sFieldName = "CasesID" Then
                GoTo NextAGShortcut
            
            ElseIf sFieldName = "ID" Then
                GoTo NextAGShortcut
            Else
            
            End If
        
            If IsNull(rstAGShortcuts.Fields(sFieldName).Value) Or rstAGShortcuts.Fields(sFieldName).Value = "" Or rstAGShortcuts.Fields(sFieldName).Value = " " Then
                GoTo NextAGShortcut
            
            Else
                sFieldValue = rstAGShortcuts.Fields(sFieldName).Value
                .Application.AutoCorrect.Entries.Add sFieldName, sFieldValue
        
            End If
    
NextAGShortcut:
        Next

    End With
    
    rstAGShortcuts.Close
    Set flCurrentField = Nothing
    Set rstAGShortcuts = Nothing

End Sub

Public Sub pfRoughDraftToCoverF()
    '============================================================================
    ' Name        : pfRoughDraftToCoverF
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfRoughDraftToCoverF
    ' Description : Adds rough draft to courtcover,
    '               does find/replacements of static speakers 1-17
    '                   all dynamic speakers, Q&A, : a-z, various AQC & AVT headings
    '============================================================================

    Dim sSpeakerName As String
    Dim oWordDoc As New Word.Document
    Dim oWordApp As New Word.Application
    Dim sTextToFind As String
    Dim sReplacementText As String
    Dim x As Long
    Dim drSpeakerName As DAO.Recordset
    Dim qdf As QueryDef
    Dim wsyWordStyle As String
    Dim bMatchCase As Boolean
    Dim cJob As New Job


    Call pfCurrentCaseInfo                       'refresh transcript info
    

    On Error Resume Next

    Set oWordApp = GetObject(, "Word.Application")
    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If
    Set oWordDoc = GetObject(cJob.DocPath.CourtCover, "Word.Document")

    oWordApp.Visible = True
    On Error GoTo 0



    With oWordDoc                                'insert rough draft at RoughBKMK bookmark

        If .bookmarks.Exists("RoughBKMK") = True Then
    
            .bookmarks("RoughBKMK").Select
            .Application.Selection.InsertFile FileName:=cJob.DocPath.RoughDraft
        
        Else
            MsgBox "Bookmark ""RoughBKMK"" does not exist!"
        End If
        .MailMerge.MainDocumentType = wdNotAMergeDocument
        .SaveAs2 FileName:=cJob.DocPath.CourtCover
        .Close
    End With

    'Documents("RoughDraft.docx").Close wdDoNotSaveChanges
    
    'Set oWordDoc = Documents.Open(cJob.DocPath.CourtCover)
    
    On Error Resume Next
    Set oWordApp = GetObject(, "Word.Application")
    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0

    Set oWordApp = CreateObject("Word.Application")

    Set oWordDoc = GetObject(cJob.DocPath.CourtCover, "Word.Document")
    oWordApp.Visible = True

    x = 18  '18 is number of first dynamic speaker
    
    '@Ignore UnassignedVariableUsage
    DoCmd.OpenQuery qnViewJobFormAppearancesQ, acViewNormal, acReadOnly
    
    
    'file name to do find replaces in
    
    '@Ignore UnassignedVariableUsage
    Set qdf = CurrentDb.QueryDefs(qnViewJobFormAppearancesQ)
    qdf.Parameters(0) = sCourtDatesID
    Set drSpeakerName = qdf.OpenRecordset
    
    If Not (drSpeakerName.EOF And drSpeakerName.BOF) Then
        drSpeakerName.MoveFirst
        Do Until drSpeakerName.EOF = True
           
            sMrMs = drSpeakerName!MrMs           'get MrMs & LastName variables
            sLastName = drSpeakerName!LastName
            sSpeakerName = UCase(sMrMs & ". " & sLastName & ":  ") 'store together in variable as a string
            
            
       
            'Do find/replaces
            sTextToFind = " snl" & x & Chr(32)
            sReplacementText = ".^p" & sSpeakerName
            Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
                
            pfDelay 3
                    
                
                
            sTextToFind = " dnl" & x & Chr(32)
            sReplacementText = " --^p" & sSpeakerName
            Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
                
            pfDelay 3
                    
                
                
            sTextToFind = " qnl" & x & Chr(32)
            sReplacementText = "?^p" & sSpeakerName
            Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
                
            pfDelay 3
                    
                
                
            sTextToFind = " sbl" & x & Chr(32)
            sReplacementText = ".^pBY " & sSpeakerName & "^pQ.  "
            Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
                
            pfDelay 3
                    
                
                
            sTextToFind = " dbl" & x & Chr(32)
            sReplacementText = " --^pBY " & sSpeakerName & "^pQ.  "
            Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
                
            pfDelay 3
                    
                
                
            sTextToFind = " qbl" & x & Chr(32)
            sReplacementText = "?^pBY " & sSpeakerName & "^pQ.  "
            Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
                
            pfDelay 3
                    
                
                
            sTextToFind = " sqnl" & x & Chr(32)
            sReplacementText = "." & Chr(34) & "^p" & sSpeakerName
            Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
                
            pfDelay 3
                    
                
                
            sTextToFind = " dqnl" & x & Chr(32)
            sReplacementText = " --" & Chr(34) & "^p" & sSpeakerName
            Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
                
            pfDelay 3
                    
                
                
                
            sTextToFind = " qqnl" & x & Chr(32)
            sReplacementText = "?" & Chr(34) & "^p" & sSpeakerName
            Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
                
            pfDelay 3
                    
                
        
            'clear variables before
            sMrMs = ""
            sLastName = ""
            sSpeakerName = ""
            sTextToFind = ""
            sReplacementText = ""
            
            x = x + 1                            'add 1 to x for next speaker name
            drSpeakerName.MoveNext               'go to next speaker name
            
            'back up to the top
            DoEvents
        Loop
    
    
    
    
        'MsgBox "Finished ing through dynamic speakers."
        
        drSpeakerName.Close                      'Close the recordset
        Set drSpeakerName = Nothing              'Clean up
        
        
        '@Ignore UnassignedVariableUsage
        DoCmd.OpenQuery qnViewJobFormAppearancesQ, acViewNormal, acReadOnly
        sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
        
        
        sTextToFind = " --"
        sReplacementText = " --"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 3
            
        
    
        sTextToFind = "  --"
        sReplacementText = " --"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 3
            
        
    
        sTextToFind = " --"
        sReplacementText = "^s--"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 3
            
        
    
        sTextToFind = "i'"
        sReplacementText = "I'"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        '**********************************Question and Answer / Q&A
        
        sTextToFind = " snlq "
        sReplacementText = ".^pQ.  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        sTextToFind = " dnlq "
        sReplacementText = "^s--^pQ.  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        sTextToFind = " qnlq "
        sReplacementText = "?^pQ.  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " snla "
        sReplacementText = ".^pA.  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
                
        sTextToFind = " dnla "
        sReplacementText = "^s--^pA.  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnla "
        sReplacementText = "?^pA.  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE COURT 1
        sTextToFind = " snl1 "
        sReplacementText = ".^pTHE COURT:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl1 "
        sReplacementText = "^s--^pTHE COURT:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl1 "
        sReplacementText = "?^pTHE COURT:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE CLERK 2
        
        sTextToFind = " dnl2 "
        sReplacementText = "^s--^pTHE CLERK:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl2 "
        sReplacementText = "?^pTHE CLERK:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " snl2 "
        sReplacementText = ".^pTHE CLERK:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE WITNESS 3
        
        sTextToFind = " snl3 "
        sReplacementText = ".^pTHE WITNESS:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl3 "
        sReplacementText = "^s--^pTHE WITNESS:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl3 "
        sReplacementText = "?^pTHE WITNESS:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE BAILIFF 4
        
        sTextToFind = " snl4 "
        sReplacementText = ".^pTHE BAILIFF:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl4 "
        sReplacementText = "^s--^pTHE BAILIFF:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
                
        sTextToFind = " qnl4 "
        sReplacementText = "?^pTHE BAILIFF:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
                
        '**********************************THE COURT REPORTER 5
        
        sTextToFind = " snl5 "
        sReplacementText = ".^pTHE COURT REPORTER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl5 "
        sReplacementText = "^s--^pTHE COURT REPORTER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl5 "
        sReplacementText = "?^pTHE COURT REPORTER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE REPORTER 6
        
        sTextToFind = " snl6 "
        sReplacementText = ".^pTHE REPORTER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl6 "
        sReplacementText = "^s--^pTHE REPORTER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl6 "
        sReplacementText = "?^pTHE REPORTER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE MONITOR 7
        
        sTextToFind = " snl7 "
        sReplacementText = ".^pTHE MONITOR:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl7 "
        sReplacementText = "^s--^pTHE MONITOR:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl7 "
        sReplacementText = "?^pTHE MONITOR:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE INTERPRETER 8
        
        sTextToFind = " snl8 "
        sReplacementText = ".^pTHE INTERPRETER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
                
        sTextToFind = " dnl8 "
        sReplacementText = "^s--^pTHE INTERPRETER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl8 "
        sReplacementText = "?^pTHE INTERPRETER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE PLAINTIFF 9
        
        sTextToFind = " snl9 "
        sReplacementText = ".^pTHE PLAINTIFF:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl9 "
        sReplacementText = "^s--^pTHE PLAINTIFF:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl9 "
        sReplacementText = "?^pTHE PLAINTIFF:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
    
        '**********************************THE DEFENDANT 10
    
        sTextToFind = " snl10 "
        sReplacementText = ".^pTHE DEFENDANT:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl10 "
        sReplacementText = "^s--^pTHE DEFENDANT:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl10 "
        sReplacementText = "?^pTHE DEFENDANT:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE PETITIONER 11
        
        sTextToFind = " snl11 "
        sReplacementText = ".^pTHE PETITIONER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl11 "
        sReplacementText = "^s--^pTHE PETITIONER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl11 "
        sReplacementText = "?^pTHE PETITIONER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
    
        '**********************************THE RESPONDENT 12
    
        sTextToFind = " snl12 "
        sReplacementText = ".^pTHE RESPONDENT:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl12 "
        sReplacementText = "^s--^pTHE RESPONDENT:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl12 "
        sReplacementText = "?^pTHE RESPONDENT:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
    
        '**********************************THE DEBTOR 13
        
        sTextToFind = " snl13 "
        sReplacementText = ".^pTHE DEBTOR:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl13 "
        sReplacementText = "^s--^pTHE DEBTOR:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl13 "
        sReplacementText = "?^pTHE DEBTOR:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
    
        '**********************************THE MOTHER 14
        
        sTextToFind = " snl14 "
        sReplacementText = ".^pTHE MOTHER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl14 "
        sReplacementText = "^s--^pTHE MOTHER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl14 "
        sReplacementText = "?^pTHE MOTHER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE JURY 15
        
        sTextToFind = " snl15 "
        sReplacementText = ".^pTHE JURY:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl15 "
        sReplacementText = "^s--^pTHE JURY:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl15 "
        sReplacementText = "?^pTHE JURY:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE UNIDENTIFIED SPEAKER 16
        
        sTextToFind = " snl16 "
        sReplacementText = ".^pTHE UNIDENTIFIED SPEAKER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl16 "
        sReplacementText = "^s--^pTHE UNIDENTIFIED SPEAKER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl16 "
        sReplacementText = "?^pTHE UNIDENTIFIED SPEAKER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
    
        '**********************************THE FATHER 17
        
        sTextToFind = " snl17 "
        sReplacementText = ".^pTHE FATHER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl17 "
        sReplacementText = "^s--^pTHE FATHER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl17 "
        sReplacementText = "?^pTHE FATHER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************
        'MsgBox "Finished ing through static speakers!"
    
        sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
        
    
        '********************************** :  A through Z
        sTextToFind = ":  a"
        sReplacementText = " :  A"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  b"
        sReplacementText = " :  B"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  c"
        sReplacementText = " :  C"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  d"
        sReplacementText = " :  D"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  e"
        sReplacementText = " :  E"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  f"
        sReplacementText = " :  F"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  g"
        sReplacementText = " :  G"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
                
        sTextToFind = ":  h"
        sReplacementText = " :  H"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  i"
        sReplacementText = " :  I"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
                
        sTextToFind = ":  j"
        sReplacementText = " :  J"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  k"
        sReplacementText = " :  K"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  l"
        sReplacementText = " :  L"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  m"
        sReplacementText = " :  M"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  n"
        sReplacementText = " :  N"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  o"
        sReplacementText = " :  O"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  p"
        sReplacementText = " :  P"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  q"
        sReplacementText = " :  Q"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  r"
        sReplacementText = " :  R"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  s"
        sReplacementText = " :  S"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  t"
        sReplacementText = " :  T"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  u"
        sReplacementText = " :  U"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  v"
        sReplacementText = " :  V"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  w"
        sReplacementText = " :  W"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  x"
        sReplacementText = " :  X"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  y"
        sReplacementText = " :  Y"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  z"
        sReplacementText = " :  Z"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        x = 18                                   '18 is number of first dynamic speaker
        
        '@Ignore UnassignedVariableUsage
        DoCmd.OpenQuery qnViewJobFormAppearancesQ, acViewNormal, acReadOnly
        
        sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
        
        '@Ignore UnassignedVariableUsage
        Set qdf = CurrentDb.QueryDefs(qnViewJobFormAppearancesQ)
        qdf.Parameters(0) = sCourtDatesID
        Set drSpeakerName = qdf.OpenRecordset
        
        If Not (drSpeakerName.EOF And drSpeakerName.BOF) Then
        
            drSpeakerName.MoveFirst
            Do Until drSpeakerName.EOF = True
                
                sMrMs = drSpeakerName!MrMs
                sLastName = drSpeakerName!LastName
                sSpeakerName = UCase(sMrMs & " " & sLastName & ":  ")
            
                sTextToFind = " snl" & x & Chr(32)
                sReplacementText = ".^p" & sSpeakerName
                Call pfSingleFindReplace(sTextToFind, sReplacementText)
            
                sTextToFind = " dnl" & x & Chr(32)
                sReplacementText = "^s--^p" & sSpeakerName
                Call pfSingleFindReplace(sTextToFind, sReplacementText)
            
                sTextToFind = " qnl" & x & Chr(32)
                sReplacementText = "?^p" & sSpeakerName
                Call pfSingleFindReplace(sTextToFind, sReplacementText)
            
                sTextToFind = " sbl" & x & Chr(32)
                sReplacementText = ".^pBY " & sSpeakerName
                Call pfSingleFindReplace(sTextToFind, sReplacementText)
            
                sTextToFind = " dbl" & x & Chr(32)
                sReplacementText = "^s--^pBY " & sSpeakerName
                Call pfSingleFindReplace(sTextToFind, sReplacementText)
            
                sTextToFind = " qbl" & x & Chr(32)
                sReplacementText = "?^pBY " & sSpeakerName
                Call pfSingleFindReplace(sTextToFind, sReplacementText)
            
                sTextToFind = " sqnl" & x & Chr(32)
                sReplacementText = "." & Chr(34) & "^p" & sSpeakerName
                Call pfSingleFindReplace(sTextToFind, sReplacementText)
            
                sTextToFind = " dqnl" & x & Chr(32)
                sReplacementText = "^s--" & Chr(34) & "^p" & sSpeakerName
                Call pfSingleFindReplace(sTextToFind, sReplacementText)
            
                sTextToFind = " qqnl" & x & Chr(32)
                sReplacementText = "?" & Chr(34) & "^p" & sSpeakerName
                Call pfSingleFindReplace(sTextToFind, sReplacementText)
            
                sMrMs = ""                       'clear variables before loop
                sLastName = ""
                sSpeakerName = ""
            
                x = x + 1                        'add 1 to x for next speaker name
                drSpeakerName.MoveNext           'go to next speaker name
                
            Loop                                 'back up to the top
            
        Else                                     'upon completion
        
            MsgBox "There are no records in the recordset."
        End If
        
        
        
        'MsgBox "Finished looping through A: to Z:."
        
        '********************************** various style-related F/Rs
        
        If InStr(sJurisdiction, "AVT") > 0 Or InStr(sJurisdiction, "eScribers") > 0 Then
        
            sTextToFind = "Q.  "
            sReplacementText = "Q.  "
            wsyWordStyle = "ESQandA"
            Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
            
            sTextToFind = "A.  "
            sReplacementText = "A.  "
            wsyWordStyle = "ESQandA"
            Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
            
            sTextToFind = "BY M"
            sReplacementText = "BY M"
            wsyWordStyle = "ESBYLawyer"
            Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
            
            sTextToFind = ":  "
            sReplacementText = ":  "
            wsyWordStyle = "ESColloquy"
            Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
            
            sTextToFind = "^p("
            sReplacementText = "^p("
            wsyWordStyle = "ESParen"
            Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
            
            sTextToFind = "SWORN"
            sReplacementText = "SWORN"
            wsyWordStyle = "ESHeading"
            Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
            
            sTextToFind = "DIRECT EXAMINATION"
            sReplacementText = "DIRECT EXAMINATION"
            wsyWordStyle = "ESHeading"
            Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
            
            sTextToFind = "CROSS-EXAMINATION"
            sReplacementText = "CROSS-EXAMINATION"
            wsyWordStyle = "ESHeading"
            Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
            
            sTextToFind = "REDIRECT EXAMINATION"
            sReplacementText = "REDIRECT EXAMINATION"
            wsyWordStyle = "ESHeading"
            Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
            
            sTextToFind = "RECROSS-EXAMINATION"
            sReplacementText = "RECROSS-EXAMINATION"
            wsyWordStyle = "ESHeading"
            Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
            
            sTextToFind = "FURTHER REDIRECT EXAMINATION"
            sReplacementText = "FURTHER REDIRECT EXAMINATION"
            wsyWordStyle = "ESHeading"
            Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
            
            sTextToFind = "FURTHER RECROSS-EXAMINATION"
            sReplacementText = "FURTHER RECROSS-EXAMINATION"
            wsyWordStyle = "ESHeading"
            Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
            
            sTextToFind = "OPENING STATEMENT"
            sReplacementText = "OPENING STATEMENT"
            wsyWordStyle = "ESHeading"
            Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
            
            sTextToFind = "CLOSING ARGUMENT"
            sReplacementText = "CLOSING ARGUMENT"
            wsyWordStyle = "ESHeading"
            Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
            
            sTextToFind = "VERDICT"
            sReplacementText = "VERDICT"
            wsyWordStyle = "ESHeading"
            Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
            
            sTextToFind = "SENTENCING"
            sReplacementText = "SENTENCING"
            wsyWordStyle = "ESHeading"
            Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
            
            sTextToFind = "COURT'S RULING"
            sReplacementText = "COURT'S RULING"
            wsyWordStyle = "ESHeading"
            Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
                
        Else
        End If
        
        Call pfTCEntryReplacement
        
        sTextToFind = "^^p"
        sReplacementText = "^p"
        wsyWordStyle = ""
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, bFormat:=True, bMatchCase:=True)
        bMatchCase = ""
        
        sTextToFind = "Q.  "
        sReplacementText = "Q.  "
        wsyWordStyle = "AQC-QA"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True, bMatchCase:=True)
        bMatchCase = ""
        sTextToFind = "A.  "
        sReplacementText = "A.  "
        wsyWordStyle = "AQC-QA"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True, bMatchCase:=True)
        bMatchCase = ""
        
        sTextToFind = ":  "
        sReplacementText = ":  "
        wsyWordStyle = "AQC-Colloquy"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "^p("
        sReplacementText = "^p("
        wsyWordStyle = "AQC-Parenthesis"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "admitted.)"
        sReplacementText = "admitted.)"
        wsyWordStyle = "Heading 3"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "received.)"
        sReplacementText = "received.)"
        wsyWordStyle = "Heading 3"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "marked.)"
        sReplacementText = "marked.)"
        wsyWordStyle = "Heading 3"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "SWORN"
        sReplacementText = "SWORN"
        wsyWordStyle = "Heading 1"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True, bMatchCase:=True)
        bMatchCase = ""
        
        sTextToFind = "DIRECT EXAMINATION"
        sReplacementText = "DIRECT EXAMINATION"
        wsyWordStyle = "Heading 2"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True, bMatchCase:=True)
        bMatchCase = ""
        
        sTextToFind = "CROSS-EXAMINATION"
        sReplacementText = "CROSS-EXAMINATION"
        wsyWordStyle = "Heading 2"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True, bMatchCase:=True)
        bMatchCase = ""
        
        sTextToFind = "REDIRECT EXAMINATION"
        sReplacementText = "REDIRECT EXAMINATION"
        wsyWordStyle = "Heading 2"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True, bMatchCase:=True)
        bMatchCase = ""
        
        sTextToFind = "RECROSS-EXAMINATION"
        sReplacementText = "RECROSS-EXAMINATION"
        wsyWordStyle = "Heading 2"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True, bMatchCase:=True)
        bMatchCase = ""
        
        sTextToFind = "FURTHER REDIRECT EXAMINATION"
        sReplacementText = "FURTHER REDIRECT EXAMINATION"
        wsyWordStyle = "Heading 2"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True, bMatchCase:=True)
        bMatchCase = ""
        
        sTextToFind = "FURTHER RECROSS-EXAMINATION"
        sReplacementText = "FURTHER RECROSS-EXAMINATION"
        wsyWordStyle = "Heading 2"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True, bMatchCase:=True)
        bMatchCase = ""
        
        sTextToFind = "OPENING STATEMENT"
        sReplacementText = "OPENING STATEMENT"
        wsyWordStyle = "Heading 1"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True, bMatchCase:=True)
        bMatchCase = ""
        
        sTextToFind = "CLOSING ARGUMENT"
        sReplacementText = "CLOSING ARGUMENT"
        wsyWordStyle = "Heading 1"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True, bMatchCase:=True)
        bMatchCase = ""
        
        sTextToFind = "VERDICT"
        sReplacementText = "VERDICT"
        wsyWordStyle = "Heading 1"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True, bMatchCase:=True)
        bMatchCase = ""
        
        sTextToFind = "SENTENCING"
        sReplacementText = "SENTENCING"
        wsyWordStyle = "Heading 1"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True, bMatchCase:=True)
        bMatchCase = ""
        
        sTextToFind = "COURT'S RULING"
        sReplacementText = "COURT'S RULING"
        wsyWordStyle = "Heading 1"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True, bMatchCase:=True)
        bMatchCase = ""
        
        sTextToFind = "ARGUMENT"
        sReplacementText = "ARGUMENT"
        wsyWordStyle = "Heading 1"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True, bMatchCase:=True)
        bMatchCase = ""
        
        sTextToFind = "BY M"
        sReplacementText = "BY M"
        wsyWordStyle = "Normal"
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True, bMatchCase:=True)
        bMatchCase = ""
        
        sTextToFind = "TC " & Chr(34) & "TC" & Chr(34) & " "
        sReplacementText = "TC "
        wsyWordStyle = ""
        Call pfSingleTCReplaceAll(sTextToFind, sReplacementText, bFormat:=True, bMatchCase:=True)
        bMatchCase = ""
        
        
    End If
    
    drSpeakerName.Close                          'close the recordset
    Set drSpeakerName = Nothing                  'clean up



    On Error Resume Next
    oWordDoc.Close (wdSaveChanges)
    oWordApp.Quit

    qdf.Close

    If sJurisdiction <> "FDA" And sJurisdiction <> "Food and Drug Administration" And sJurisdiction <> "eScribers NH" And sJurisdiction <> "eScribers Bankruptcy" And sJurisdiction <> "eScribers New Jersey" Then
        Call pfHeaders
        Call fDynamicHeaders
    End If

    On Error GoTo 0

    Call pfClearGlobals
End Sub

Public Sub pfStaticSpeakersFindReplace()
    '============================================================================
    ' Name        : pfStaticSpeakersFindReplace
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfStaticSpeakersFindReplace
    ' Description : finds and replaces static speakers in CourtCover after rough draft is inserted
    '============================================================================

    Dim sTextToFind As String
    Dim sReplacementText As String
    Dim oWordApp As New Word.Application
    Dim oWordDoc As New Word.Document
    Dim cJob As New Job

    DoCmd.OpenQuery qnViewJobFormAppearancesQ, acViewNormal, acReadOnly

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    
    Set oWordApp = GetObject(, "Word.Application")
    If oWordApp Is Nothing Then
        Set oWordApp = CreateObject("Word.Application")
    End If
    oWordApp.Visible = False
    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.CourtCover)
    oWordDoc.Activate

    sTextToFind = " --"
    sReplacementText = " --"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = "  --"
    sReplacementText = " --"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " --"
    sReplacementText = "^s--"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)
    
    sTextToFind = "i'"
    sReplacementText = "I'"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    '**********************************Question and Answer / Q&A

    sTextToFind = " snlq "
    sReplacementText = ".^pQ.  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " dnlq "
    sReplacementText = "^s--^pQ.  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " qnlq "
    sReplacementText = "?^pQ.  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " snla "
    sReplacementText = ".^pA.  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)
        
    sTextToFind = " dnla "
    sReplacementText = "^s--^pA.  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " qnla "
    sReplacementText = "?^pA.  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    '**********************************THE COURT 1
    sTextToFind = " snl1 "
    sReplacementText = ".^pTHE COURT:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " dnl1 "
    sReplacementText = "^s--^pTHE COURT:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " qnl1 "
    sReplacementText = "?^pTHE COURT:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    '**********************************THE CLERK 2

    sTextToFind = " dnl2 "
    sReplacementText = "^s--^pTHE CLERK:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " qnl2 "
    sReplacementText = "?^pTHE CLERK:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " snl2 "
    sReplacementText = ".^pTHE CLERK:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    '**********************************THE WITNESS 3

    sTextToFind = " snl3 "
    sReplacementText = ".^pTHE WITNESS:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " dnl3 "
    sReplacementText = "^s--^pTHE WITNESS:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " qnl3 "
    sReplacementText = "?^pTHE WITNESS:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    '**********************************THE BAILIFF 4

    sTextToFind = " snl4 "
    sReplacementText = ".^pTHE BAILIFF:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " dnl4 "
    sReplacementText = "^s--^pTHE BAILIFF:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)
        
    sTextToFind = " qnl4 "
    sReplacementText = "?^pTHE BAILIFF:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)
        
    '**********************************THE COURT REPORTER 5

    sTextToFind = " snl5 "
    sReplacementText = ".^pTHE COURT REPORTER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " dnl5 "
    sReplacementText = "^s--^pTHE COURT REPORTER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " qnl5 "
    sReplacementText = "?^pTHE COURT REPORTER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    '**********************************THE REPORTER 6

    sTextToFind = " snl6 "
    sReplacementText = ".^pTHE REPORTER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " dnl6 "
    sReplacementText = "^s--^pTHE REPORTER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " qnl6 "
    sReplacementText = "?^pTHE REPORTER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    '**********************************THE MONITOR 7

    sTextToFind = " snl7 "
    sReplacementText = ".^pTHE MONITOR:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " dnl7 "
    sReplacementText = "^s--^pTHE MONITOR:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " qnl7 "
    sReplacementText = "?^pTHE MONITOR:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    '**********************************THE INTERPRETER 8

    sTextToFind = " snl8 "
    sReplacementText = ".^pTHE INTERPRETER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)
        
    sTextToFind = " dnl8 "
    sReplacementText = "^s--^pTHE INTERPRETER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " qnl8 "
    sReplacementText = "?^pTHE INTERPRETER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    '**********************************THE PLAINTIFF 9

    sTextToFind = " snl9 "
    sReplacementText = ".^pTHE PLAINTIFF:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " dnl9 "
    sReplacementText = "^s--^pTHE PLAINTIFF:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " qnl9 "
    sReplacementText = "?^pTHE PLAINTIFF:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    '**********************************THE DEFENDANT 10

    sTextToFind = " snl10 "
    sReplacementText = ".^pTHE DEFENDANT:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " dnl10 "
    sReplacementText = "^s--^pTHE DEFENDANT:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " qnl10 "
    sReplacementText = "?^pTHE DEFENDANT:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    '**********************************THE PETITIONER 11

    sTextToFind = " snl11 "
    sReplacementText = ".^pTHE PETITIONER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " dnl11 "
    sReplacementText = "^s--^pTHE PETITIONER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " qnl11 "
    sReplacementText = "?^pTHE PETITIONER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    '**********************************THE RESPONDENT 12

    sTextToFind = " snl12 "
    sReplacementText = ".^pTHE RESPONDENT:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " dnl12 "
    sReplacementText = "^s--^pTHE RESPONDENT:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " qnl12 "
    sReplacementText = "?^pTHE RESPONDENT:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    '**********************************THE DEBTOR 13

    sTextToFind = " snl13 "
    sReplacementText = ".^pTHE DEBTOR:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " dnl13 "
    sReplacementText = "^s--^pTHE DEBTOR:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " qnl13 "
    sReplacementText = "?^pTHE DEBTOR:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    '**********************************THE MOTHER 14

    sTextToFind = " snl14 "
    sReplacementText = ".^pTHE MOTHER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " dnl14 "
    sReplacementText = "^s--^pTHE MOTHER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " qnl14 "
    sReplacementText = "?^pTHE MOTHER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    '**********************************THE JURY 15

    sTextToFind = " snl15 "
    sReplacementText = ".^pTHE JURY:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " dnl15 "
    sReplacementText = "^s--^pTHE JURY:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " qnl15 "
    sReplacementText = "?^pTHE JURY:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    '**********************************THE UNIDENTIFIED SPEAKER 16

    sTextToFind = " snl16 "
    sReplacementText = ".^pTHE UNIDENTIFIED SPEAKER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " dnl16 "
    sReplacementText = "^s--^pTHE UNIDENTIFIED SPEAKER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " qnl16 "
    sReplacementText = "?^pTHE UNIDENTIFIED SPEAKER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    '**********************************THE FATHER 17

    sTextToFind = " snl17 "
    sReplacementText = ".^pTHE FATHER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " dnl17 "
    sReplacementText = "^s--^pTHE FATHER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = " qnl17 "
    sReplacementText = "?^pTHE FATHER:  "
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    'MsgBox "Finished looping through static speakers!"

    oWordDoc.Save

    oWordDoc.Close
    oWordApp.Quit
    Set oWordDoc = Nothing
    Set oWordApp = Nothing

    DoCmd.Close (qnViewJobFormAppearancesQ)
End Sub

Public Sub pfReplaceColonUndercasewithColonUppercase()
    '============================================================================
    ' Name        : pfReplaceColonUndercasewithColonUppercase
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfReplaceColonUndercasewithColonUppercase
    ' Description:  replaces : a-z with : A-Z, applies styles to fixed phrases in transcript
    '============================================================================
    Dim oWordApp As New Word.Application
    Dim oWordDoc As New Word.Document
    Dim sTextToFind As String
    Dim sReplacementText As String
    Dim wsyWordStyle As Word.Style
    Dim cJob As New Job

    DoCmd.OpenQuery qnViewJobFormAppearancesQ, acViewNormal, acReadOnly
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

    Set oWordApp = GetObject(, "Word.Application")

    If oWordApp Is Nothing Then
        Set oWordApp = CreateObject("Word.Application")
    End If
    oWordApp.Visible = False
    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.CourtCover)
    oWordDoc.Activate

    '********************************** :  A through Z
    sTextToFind = ":  a"
    sReplacementText = " :  A"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  b"
    sReplacementText = " :  B"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  c"
    sReplacementText = " :  C"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  d"
    sReplacementText = " :  D"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  e"
    sReplacementText = " :  E"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  f"
    sReplacementText = " :  F"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  g"
    sReplacementText = " :  G"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)
        
    sTextToFind = ":  h"
    sReplacementText = " :  H"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  i"
    sReplacementText = " :  I"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)
        
    sTextToFind = ":  j"
    sReplacementText = " :  J"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  k"
    sReplacementText = " :  K"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  l"
    sReplacementText = " :  L"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  m"
    sReplacementText = " :  M"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  n"
    sReplacementText = " :  N"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  o"
    sReplacementText = " :  O"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  p"
    sReplacementText = " :  P"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  q"
    sReplacementText = " :  Q"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  r"
    sReplacementText = " :  R"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  s"
    sReplacementText = " :  S"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  t"
    sReplacementText = " :  T"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  u"
    sReplacementText = " :  U"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  v"
    sReplacementText = " :  V"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  w"
    sReplacementText = " :  W"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  x"
    sReplacementText = " :  X"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  y"
    sReplacementText = " :  Y"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    sTextToFind = ":  z"
    sReplacementText = " :  Z"
    Call pfSingleFindReplace(sTextToFind, sReplacementText)

    'MsgBox "Finished looping through A: to Z:."

    '********************************** Q/A Question and Answer Q&A

    sTextToFind = "Q.  "
    sReplacementText = "Q.  "
    wsyWordStyle = "AQC-QA"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    sTextToFind = "A.  "
    sReplacementText = "A.  "
    wsyWordStyle = "AQC-QA"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    '********************************** Colloquy

    sTextToFind = ":  "
    sReplacementText = ":  "
    wsyWordStyle = "AQC-Colloquy"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    '********************************** Exhibits and Parens

    sTextToFind = "^p("
    sReplacementText = "^p("
    wsyWordStyle = "AQC-Parenthesis"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    sTextToFind = "admitted.)"
    sReplacementText = "admitted.)"
    wsyWordStyle = "Heading 3"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    sTextToFind = "received.)"
    sReplacementText = "received.)"
    wsyWordStyle = "Heading 3"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    sTextToFind = "marked.)"
    sReplacementText = "marked.)"
    wsyWordStyle = "Heading 3"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    '********************************** Various Main Headings

    sTextToFind = "DIRECT EXAMINATION"
    sReplacementText = "DIRECT EXAMINATION"
    wsyWordStyle = "Heading 2"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    sTextToFind = "CROSS-EXAMINATION"
    sReplacementText = "CROSS-EXAMINATION"
    wsyWordStyle = "Heading 2"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    sTextToFind = "REDIRECT EXAMINATION"
    sReplacementText = "REDIRECT EXAMINATION"
    wsyWordStyle = "Heading 2"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    sTextToFind = "RECROSS-EXAMINATION"
    sReplacementText = "RECROSS-EXAMINATION"
    wsyWordStyle = "Heading 2"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    sTextToFind = "FURTHER REDIRECT EXAMINATION"
    sReplacementText = "FURTHER REDIRECT EXAMINATION"
    wsyWordStyle = "Heading 2"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    sTextToFind = "FURTHER RECROSS-EXAMINATION"
    sReplacementText = "FURTHER RECROSS-EXAMINATION"
    wsyWordStyle = "Heading 2"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    sTextToFind = "SWORN"
    sReplacementText = "SWORN"
    wsyWordStyle = "Heading 1"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    sTextToFind = "OPENING STATEMENT"
    sReplacementText = "OPENING STATEMENT"
    wsyWordStyle = "Heading 1"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    sTextToFind = "CLOSING ARGUMENT"
    sReplacementText = "CLOSING ARGUMENT"
    wsyWordStyle = "Heading 1"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    sTextToFind = "VERDICT"
    sReplacementText = "VERDICT"
    wsyWordStyle = "Heading 1"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    sTextToFind = "SENTENCING"
    sReplacementText = "SENTENCING"
    wsyWordStyle = "Heading 1"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    sTextToFind = "COURT'S RULING"
    sReplacementText = "COURT'S RULING"
    wsyWordStyle = "Heading 1"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    sTextToFind = "ARGUMENT"
    sReplacementText = "ARGUMENT"
    wsyWordStyle = "Heading 1"
    Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)

    oWordDoc.Save
    oWordDoc.Close

    oWordApp.Quit
    Set oWordDoc = Nothing
    Set oWordApp = Nothing

    DoCmd.Close (qnViewJobFormAppearancesQ)
End Sub

Public Sub pfTypeRoughDraftF()
    '============================================================================
    ' Name        : pfTypeRoughDraftF
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfTypeRoughDraftF
    ' Description : copies correct roughdraft template to job folder
    '============================================================================

    Dim oRoughDraft As Object
    Dim cJob As New Job

    Call pfCurrentCaseInfo                       'refresh transcript info
    Call pfCheckFolderExistence

    Set oRoughDraft = CreateObject("Scripting.FileSystemObject")

    If sJurisdiction = "Weber Nevada" Then
    
        If Not oRoughDraft.FileExists(cJob.DocPath.RoughDraft) Then
            FileCopy cJob.DocPath.TemplateFolder2 & "RoughDraft-WeberNV.docx", cJob.DocPath.RoughDraft
        End If
        If Not oRoughDraft.FileExists(cJob.DocPath.JobDirectoryN & "Transcribing Manual.PDF") Then
            FileCopy cJob.DocPath.TemplateFolder1 & "Transcribing Manual.PDF", cJob.DocPath.JobDirectoryN & "Transcribing Manual.PDF"
        End If
        If Not oRoughDraft.FileExists(cJob.DocPath.JobDirectoryN & "Proofreading Manual - nevada.PDF") Then
            FileCopy cJob.DocPath.TemplateFolder3 & "Proofreading Manual - nevada.PDF", cJob.DocPath.JobDirectoryN & "Proofreading Manual - nevada.PDF"
        End If
        If Not oRoughDraft.FileExists(cJob.DocPath.JobDirectoryN & "WeberNVSample.docx") Then
            FileCopy cJob.DocPath.TemplateFolder2 & "WeberNVSample.docx", cJob.DocPath.JobDirectoryN & "WeberNVSample.docx"
        End If
    Else
    End If

    If sJurisdiction = "Weber Bankruptcy" Then
        If Not oRoughDraft.FileExists(cJob.DocPath.JobDirectoryN & "WeberBKSample.docx") Then
            FileCopy cJob.DocPath.TemplateFolder2 & "WeberNVSample.docx", cJob.DocPath.JobDirectoryN & "WeberNVSample.docx"
        End If
    Else
    End If

    If sJurisdiction = "Weber Oregon" Then
        If Not oRoughDraft.FileExists(cJob.DocPath.RoughDraft) Then
            FileCopy cJob.DocPath.TemplateFolder2 & "RoughDraft-WeberOR.docx", cJob.DocPath.RoughDraft
        End If
        If Not oRoughDraft.FileExists(cJob.DocPath.JobDirectoryN & "WeberORSample.docx") Then
            FileCopy cJob.DocPath.TemplateFolder2 & "WeberORSample.docx", cJob.DocPath.JobDirectoryN & "WeberORSample.docx"
        End If
        If Not oRoughDraft.FileExists(cJob.DocPath.JobDirectoryN & "WeberORSample1.docx") Then
            FileCopy cJob.DocPath.TemplateFolder2 & "WeberORSample1.docx", cJob.DocPath.JobDirectoryN & "WeberORSample1.docx"
        End If
        If Not oRoughDraft.FileExists(cJob.DocPath.JobDirectoryN & "WeberORSampleTM.docx") Then
            FileCopy cJob.DocPath.TemplateFolder2 & "WeberORSampleTM.docx", cJob.DocPath.JobDirectoryN & "WeberORSampleTM.docx"
        End If
        If Not oRoughDraft.FileExists(cJob.DocPath.JobDirectoryN & "WeberORSample2.docx") Then
            FileCopy cJob.DocPath.TemplateFolder2 & "WeberORSample2.docx", cJob.DocPath.JobDirectoryN & "WeberORSample2.docx"
        End If
    Else
    End If

    If sJurisdiction = "USBC Western Washington" Then
        If Not oRoughDraft.FileExists(cJob.DocPath.JobDirectoryN & "BankruptcyWAGuide.pdf") Then
            FileCopy cJob.DocPath.TemplateFolder1 & "BankruptcyWAGuide.pdf", cJob.DocPath.JobDirectoryN & "BankruptcyWAGuide.pdf"
        End If
    Else
    End If

    If sJurisdiction = "Food and Drug Administration" Then
        If Not oRoughDraft.FileExists(cJob.DocPath.RoughDraft) Then
            FileCopy cJob.DocPath.TemplateFolder2 & "RoughDraft-FDA.docx", cJob.DocPath.RoughDraft
        End If
    Else
    End If

    If sJurisdiction = "*Superior Court*" Then
        If Not oRoughDraft.FileExists(cJob.DocPath.JobDirectoryN & "CourtRules-WACounties.pdf") Then
            FileCopy cJob.DocPath.TemplateFolder1 & "CourtRules-WACounties.pdf", cJob.DocPath.JobDirectoryN & "CourtRules-WACounties.pdf"
        End If
    Else
    End If
    
    If sJurisdiction = "*USBC*" Then
        If Not oRoughDraft.FileExists(cJob.DocPath.JobDirectoryN & "CourtRules-Bankruptcy-TranscriptFormatGuide-1.pdf") Then
            FileCopy cJob.DocPath.TemplateFolder1 & "CourtRules-Bankruptcy-TranscriptFormatGuide-1.pdf", cJob.DocPath.JobDirectoryN & "CourtRules-Bankruptcy-TranscriptFormatGuide-1.pdf"
        End If
        If Not oRoughDraft.FileExists(cJob.DocPath.JobDirectoryN & "CourtRules-Bankruptcy-TranscriptFormatGuide-2.pdf") Then
            FileCopy cJob.DocPath.TemplateFolder1 & "CourtRules-Bankruptcy-TranscriptFormatGuide-2.pdf", cJob.DocPath.JobDirectoryN & "CourtRules-Bankruptcy-TranscriptFormatGuide-2.pdf"
        End If
    Else
    End If

    Call pfCheckFolderExistence

    DoCmd.OpenForm FormName:="PJType"            'open window with AGShortcuts, SpeakerList, and jurisdiction notes

    Shell "winword " + cJob.DocPath.RoughDraft 'open file
    Call pfClearGlobals
End Sub

Public Sub wwReplaceWeberOR()
    Call pfTCEntryReplacement
    Call FPJurors

    'Call pfTCEntryReplacement
    Call pfRoughDraftToCoverF
    'Call pfCreateIndexWeberOR
End Sub

Public Sub wwReplaceWeberNV()
    'Call .pfRoughDraftParensXEWeberNV
    Call FPJurors
    'Call pfTCEntryReplacement
    Call pfRoughDraftToCoverF
    'Call pfCreateIndexWeberNV
    
End Sub

Public Sub wwReplaceWeberBR()

    'Call pfRoughDraftParensXEWABkp
    Call FPJurors
    Call pfRoughDraftToCoverF
    Call pfTCEntryReplacement
    Call pfCreateIndexesTOAs
    'Call pfCreateIndexWeberBR

End Sub

Public Sub pfReplaceAVT()

    Call pfRoughDraftToCoverF
    Call FPJurors
    Call pfTCEntryReplacement
    
End Sub

Public Sub pfReplaceAQC()

    Call pfRoughDraftToCoverF
    Call FPJurors
    'Call pfTCEntryReplacement
    ''TODO: What is going on here?
    Call pfFindRepCitationLinks

End Sub

Public Sub pfReplaceMass()
    Call pfRoughDraftCFMass
    Call FPJurors
    Call pfTCEntryReplacement
    
End Sub

Public Sub pfRoughDraftCFMass()
    '============================================================================
    ' Name        : pfRoughDraftToCoverF
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfRoughDraftToCoverF
    ' Description : Adds rough draft to courtcover, does find/replacements of static speakers 1-17, all dynamic speakers, Q&A, : a-z, various AQC & AVT headings
    '============================================================================

    Dim sSpeakerName As String
    Dim sTextToFind As String
    Dim sReplacementText As String
    Dim wsyWordStyle As String
    
    'TODO: duplicate
    Dim oWordDoc As New Word.Document
    Dim oWordApp As New Word.Application
    
    Dim x As Long
    Dim cJob As New Job

    Dim drSpeakerName As DAO.Recordset
    Dim qdf As QueryDef
    
    Call pfCurrentCaseInfo                       'refresh transcript info

    On Error Resume Next
    Set oWordApp = GetObject(, "Word.Application")
    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    oWordApp.Visible = True

    Set oWordDoc = GetObject(cJob.DocPath.CourtCover, "Word.Document")
    With oWordDoc                                'insert rough draft at RoughBKMK bookmark

        If .bookmarks.Exists("RoughBKMK") = True Then
    
            .bookmarks("RoughBKMK").Select
            .Application.Selection.InsertFile FileName:=cJob.DocPath.RoughDraft
        
        Else
            MsgBox "Bookmark ""RoughBKMK"" does not exist!"
        End If
        .MailMerge.MainDocumentType = wdNotAMergeDocument
        .SaveAs2 FileName:=cJob.DocPath.CourtCover
        .Close
    End With
    'Documents("RoughDraft.docx").Close wdDoNotSaveChanges
    
    'Set oWordDoc = Documents.Open(cJob.DocPath.CourtCover)
    
    On Error Resume Next
    Set oWordApp = GetObject(, "Word.Application")
    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    oWordApp.Visible = True

    Set oWordDoc = GetObject(cJob.DocPath.CourtCover, "Word.Document")

    x = 18                                       '18 is number of first dynamic speaker
    
    '@Ignore UnassignedVariableUsage
    DoCmd.OpenQuery qnViewJobFormAppearancesQ, acViewNormal, acReadOnly
    
    
    'file name to do find replaces in
    
    
    '@Ignore UnassignedVariableUsage
    Set qdf = CurrentDb.QueryDefs(qnViewJobFormAppearancesQ)
    qdf.Parameters(0) = sCourtDatesID
    Set drSpeakerName = qdf.OpenRecordset
    
    If Not (drSpeakerName.EOF And drSpeakerName.BOF) Then
        drSpeakerName.MoveFirst
        Do Until drSpeakerName.EOF = True
           
            sMrMs = drSpeakerName!MrMs           'get MrMs & LastName variables
            sLastName = drSpeakerName!LastName
            sSpeakerName = UCase(sMrMs & ". " & sLastName & ":  ") 'store together in variable as a string
            
            
       
            'Do find/replaces
            sTextToFind = " snl" & x & Chr(32)
            sReplacementText = ".^p" & sSpeakerName
            Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
                
            pfDelay 3
                    
                
                
            sTextToFind = " dnl" & x & Chr(32)
            sReplacementText = " --^p" & sSpeakerName
            Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
                
            pfDelay 3
                    
                
                
            sTextToFind = " qnl" & x & Chr(32)
            sReplacementText = "?^p" & sSpeakerName
            Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
                
            pfDelay 3
                    
                
                
            sTextToFind = " sbl" & x & Chr(32)
            sReplacementText = ".^pBY " & sSpeakerName & "^pQ.  "
            Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
                
            pfDelay 3
                    
                
                
            sTextToFind = " dbl" & x & Chr(32)
            sReplacementText = " --^pBY " & sSpeakerName & "^pQ.  "
            Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
                
            pfDelay 3
                    
                
                
            sTextToFind = " qbl" & x & Chr(32)
            sReplacementText = "?^pBY " & sSpeakerName & "^pQ.  "
            Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
                
            pfDelay 3
                    
                
                
            sTextToFind = " sqnl" & x & Chr(32)
            sReplacementText = "." & Chr(34) & "^p" & sSpeakerName
            Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
                
            pfDelay 3
                    
                
                
            sTextToFind = " dqnl" & x & Chr(32)
            sReplacementText = " --" & Chr(34) & "^p" & sSpeakerName
            Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
                
            pfDelay 3
                    
                
                
                
            sTextToFind = " qqnl" & x & Chr(32)
            sReplacementText = "?" & Chr(34) & "^p" & sSpeakerName
            Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
                
            pfDelay 3
                    
                
        
            'clear variables before
            sMrMs = ""
            sLastName = ""
            sSpeakerName = ""
            sTextToFind = ""
            sReplacementText = ""
            
            x = x + 1                            'add 1 to x for next speaker name
            drSpeakerName.MoveNext               'go to next speaker name
            
            'back up to the top
            DoEvents
        Loop
    
    
    
    
        'MsgBox "Finished ing through dynamic speakers."
        
        drSpeakerName.Close                      'Close the recordset
        Set drSpeakerName = Nothing              'Clean up
        
        
        '@Ignore UnassignedVariableUsage
        DoCmd.OpenQuery qnViewJobFormAppearancesQ, acViewNormal, acReadOnly
        sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
        
        
        sTextToFind = " --"
        sReplacementText = " --"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 3
            
        
    
        sTextToFind = "  --"
        sReplacementText = " --"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 3
            
        
    
        sTextToFind = " --"
        sReplacementText = "^s--"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 3
            
        
    
        sTextToFind = "i'"
        sReplacementText = "I'"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        '**********************************Question and Answer / Q&A
        
        sTextToFind = " snlq "
        sReplacementText = ".^pQ.  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        sTextToFind = " dnlq "
        sReplacementText = "^s--^pQ.  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        sTextToFind = " qnlq "
        sReplacementText = "?^pQ.  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " snla "
        sReplacementText = ".^pA.  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
                
        sTextToFind = " dnla "
        sReplacementText = "^s--^pA.  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnla "
        sReplacementText = "?^pA.  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE COURT 1
        sTextToFind = " snl1 "
        sReplacementText = ".^pTHE COURT:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl1 "
        sReplacementText = "^s--^pTHE COURT:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl1 "
        sReplacementText = "?^pTHE COURT:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE CLERK 2
        
        sTextToFind = " dnl2 "
        sReplacementText = "^s--^pTHE CLERK:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl2 "
        sReplacementText = "?^pTHE CLERK:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " snl2 "
        sReplacementText = ".^pTHE CLERK:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE WITNESS 3
        
        sTextToFind = " snl3 "
        sReplacementText = ".^pTHE WITNESS:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl3 "
        sReplacementText = "^s--^pTHE WITNESS:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl3 "
        sReplacementText = "?^pTHE WITNESS:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE BAILIFF 4
        
        sTextToFind = " snl4 "
        sReplacementText = ".^pTHE BAILIFF:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl4 "
        sReplacementText = "^s--^pTHE BAILIFF:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
                
        sTextToFind = " qnl4 "
        sReplacementText = "?^pTHE BAILIFF:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
                
        '**********************************THE COURT REPORTER 5
        
        sTextToFind = " snl5 "
        sReplacementText = ".^pTHE COURT REPORTER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl5 "
        sReplacementText = "^s--^pTHE COURT REPORTER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl5 "
        sReplacementText = "?^pTHE COURT REPORTER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE REPORTER 6
        
        sTextToFind = " snl6 "
        sReplacementText = ".^pTHE REPORTER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl6 "
        sReplacementText = "^s--^pTHE REPORTER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl6 "
        sReplacementText = "?^pTHE REPORTER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE MONITOR 7
        
        sTextToFind = " snl7 "
        sReplacementText = ".^pTHE MONITOR:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl7 "
        sReplacementText = "^s--^pTHE MONITOR:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl7 "
        sReplacementText = "?^pTHE MONITOR:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE INTERPRETER 8
        
        sTextToFind = " snl8 "
        sReplacementText = ".^pTHE INTERPRETER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
                
        sTextToFind = " dnl8 "
        sReplacementText = "^s--^pTHE INTERPRETER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl8 "
        sReplacementText = "?^pTHE INTERPRETER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE PLAINTIFF 9
        
        sTextToFind = " snl9 "
        sReplacementText = ".^pTHE PLAINTIFF:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl9 "
        sReplacementText = "^s--^pTHE PLAINTIFF:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl9 "
        sReplacementText = "?^pTHE PLAINTIFF:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
    
        '**********************************THE DEFENDANT 10
    
        sTextToFind = " snl10 "
        sReplacementText = ".^pTHE DEFENDANT:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl10 "
        sReplacementText = "^s--^pTHE DEFENDANT:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl10 "
        sReplacementText = "?^pTHE DEFENDANT:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE PETITIONER 11
        
        sTextToFind = " snl11 "
        sReplacementText = ".^pTHE PETITIONER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl11 "
        sReplacementText = "^s--^pTHE PETITIONER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl11 "
        sReplacementText = "?^pTHE PETITIONER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
    
        '**********************************THE RESPONDENT 12
    
        sTextToFind = " snl12 "
        sReplacementText = ".^pTHE RESPONDENT:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl12 "
        sReplacementText = "^s--^pTHE RESPONDENT:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl12 "
        sReplacementText = "?^pTHE RESPONDENT:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
    
        '**********************************THE DEBTOR 13
        
        sTextToFind = " snl13 "
        sReplacementText = ".^pTHE DEBTOR:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl13 "
        sReplacementText = "^s--^pTHE DEBTOR:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl13 "
        sReplacementText = "?^pTHE DEBTOR:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
    
        '**********************************THE MOTHER 14
        
        sTextToFind = " snl14 "
        sReplacementText = ".^pTHE MOTHER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl14 "
        sReplacementText = "^s--^pTHE MOTHER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl14 "
        sReplacementText = "?^pTHE MOTHER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE JURY 15
        
        sTextToFind = " snl15 "
        sReplacementText = ".^pTHE JURY:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl15 "
        sReplacementText = "^s--^pTHE JURY:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl15 "
        sReplacementText = "?^pTHE JURY:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************THE UNIDENTIFIED SPEAKER 16
        
        sTextToFind = " snl16 "
        sReplacementText = ".^pTHE UNIDENTIFIED SPEAKER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl16 "
        sReplacementText = "^s--^pTHE UNIDENTIFIED SPEAKER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl16 "
        sReplacementText = "?^pTHE UNIDENTIFIED SPEAKER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
    
        '**********************************THE FATHER 17
        
        sTextToFind = " snl17 "
        sReplacementText = ".^pTHE FATHER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " dnl17 "
        sReplacementText = "^s--^pTHE FATHER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = " qnl17 "
        sReplacementText = "?^pTHE FATHER:  "
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        '**********************************
        'MsgBox "Finished ing through static speakers!"
    
        sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
        
    
        '********************************** :  A through Z
        sTextToFind = ":  a"
        sReplacementText = " :  A"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  b"
        sReplacementText = " :  B"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  c"
        sReplacementText = " :  C"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  d"
        sReplacementText = " :  D"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  e"
        sReplacementText = " :  E"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  f"
        sReplacementText = " :  F"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  g"
        sReplacementText = " :  G"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
                
        sTextToFind = ":  h"
        sReplacementText = " :  H"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  i"
        sReplacementText = " :  I"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
                
        sTextToFind = ":  j"
        sReplacementText = " :  J"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  k"
        sReplacementText = " :  K"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  l"
        sReplacementText = " :  L"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  m"
        sReplacementText = " :  M"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  n"
        sReplacementText = " :  N"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  o"
        sReplacementText = " :  O"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  p"
        sReplacementText = " :  P"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  q"
        sReplacementText = " :  Q"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  r"
        sReplacementText = " :  R"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  s"
        sReplacementText = " :  S"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  t"
        sReplacementText = " :  T"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  u"
        sReplacementText = " :  U"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  v"
        sReplacementText = " :  V"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  w"
        sReplacementText = " :  W"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  x"
        sReplacementText = " :  X"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  y"
        sReplacementText = " :  Y"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        sTextToFind = ":  z"
        sReplacementText = " :  Z"
        Call pfSingleFindReplace(sTextToFind, sReplacementText)
                                                     
        
        pfDelay 1
            
        
    
        
        'MsgBox "Finished looping through A: to Z:."
        
        '********************************** various style-related F/Rs
        
        sTextToFind = "Q.  "
        sReplacementText = "Q" & Chr(9)
        wsyWordStyle = "AQC-QA"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "A.  "
        sReplacementText = "A" & Chr(9)
        wsyWordStyle = "AQC-QA"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = ":  "
        sReplacementText = ":  "
        wsyWordStyle = "AQC-Colloquy"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "^p("
        sReplacementText = "^p("
        wsyWordStyle = "AQC-Parenthesis"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "admitted.)"
        sReplacementText = "admitted.)"
        wsyWordStyle = "Heading 3"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "received.)"
        sReplacementText = "received.)"
        wsyWordStyle = "Heading 3"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "marked.)"
        sReplacementText = "marked.)"
        wsyWordStyle = "Heading 3"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "SWORN"
        sReplacementText = "SWORN"
        wsyWordStyle = "Heading 1"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "DIRECT EXAMINATION"
        sReplacementText = "DIRECT EXAMINATION"
        wsyWordStyle = "Heading 2"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "CROSS-EXAMINATION"
        sReplacementText = "CROSS-EXAMINATION"
        wsyWordStyle = "Heading 2"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "REDIRECT EXAMINATION"
        sReplacementText = "REDIRECT EXAMINATION"
        wsyWordStyle = "Heading 2"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "RECROSS-EXAMINATION"
        sReplacementText = "RECROSS-EXAMINATION"
        wsyWordStyle = "Heading 2"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "FURTHER REDIRECT EXAMINATION"
        sReplacementText = "FURTHER REDIRECT EXAMINATION"
        wsyWordStyle = "Heading 2"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "FURTHER RECROSS-EXAMINATION"
        sReplacementText = "FURTHER RECROSS-EXAMINATION"
        wsyWordStyle = "Heading 2"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "OPENING STATEMENT"
        sReplacementText = "OPENING STATEMENT"
        wsyWordStyle = "Heading 1"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "CLOSING ARGUMENT"
        sReplacementText = "CLOSING ARGUMENT"
        wsyWordStyle = "Heading 1"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "VERDICT"
        sReplacementText = "VERDICT"
        wsyWordStyle = "Heading 1"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "SENTENCING"
        sReplacementText = "SENTENCING"
        wsyWordStyle = "Heading 1"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "COURT'S RULING"
        sReplacementText = "COURT'S RULING"
        wsyWordStyle = "Heading 1"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        sTextToFind = "ARGUMENT"
        sReplacementText = "ARGUMENT"
        wsyWordStyle = "Heading 1"
        Call pfSingleFindReplace(sTextToFind, sReplacementText, wsyWordStyle, bFormat:=True)
        
        x = 18                                   '18 is number of first dynamic speaker
        
        '@Ignore UnassignedVariableUsage
        DoCmd.OpenQuery qnViewJobFormAppearancesQ, acViewNormal, acReadOnly
        
        sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
        
        '@Ignore UnassignedVariableUsage
        Set qdf = CurrentDb.QueryDefs(qnViewJobFormAppearancesQ)
        qdf.Parameters(0) = sCourtDatesID
        Set drSpeakerName = qdf.OpenRecordset
        
        If Not (drSpeakerName.EOF And drSpeakerName.BOF) Then
        
            drSpeakerName.MoveFirst
            Do Until drSpeakerName.EOF = True
                
                sMrMs = drSpeakerName!MrMs
                sLastName = drSpeakerName!LastName
                sSpeakerName = UCase(sMrMs & " " & sLastName & ":  ")
            
                sTextToFind = " snl" & x & Chr(32)
                sReplacementText = ".^p" & sSpeakerName
                Call pfSingleFindReplace(sTextToFind, sReplacementText)
            
                sTextToFind = " dnl" & x & Chr(32)
                sReplacementText = "^s--^p" & sSpeakerName
                Call pfSingleFindReplace(sTextToFind, sReplacementText)
            
                sTextToFind = " qnl" & x & Chr(32)
                sReplacementText = "?^p" & sSpeakerName
                Call pfSingleFindReplace(sTextToFind, sReplacementText)
            
                sTextToFind = " sbl" & x & Chr(32)
                sReplacementText = ".^pBY " & sSpeakerName
                Call pfSingleFindReplace(sTextToFind, sReplacementText)
            
                sTextToFind = " dbl" & x & Chr(32)
                sReplacementText = "^s--^pBY " & sSpeakerName
                Call pfSingleFindReplace(sTextToFind, sReplacementText)
            
                sTextToFind = " qbl" & x & Chr(32)
                sReplacementText = "?^pBY " & sSpeakerName
                Call pfSingleFindReplace(sTextToFind, sReplacementText)
            
                sTextToFind = " sqnl" & x & Chr(32)
                sReplacementText = "." & Chr(34) & "^p" & sSpeakerName
                Call pfSingleFindReplace(sTextToFind, sReplacementText)
            
                sTextToFind = " dqnl" & x & Chr(32)
                sReplacementText = "^s--" & Chr(34) & "^p" & sSpeakerName
                Call pfSingleFindReplace(sTextToFind, sReplacementText)
            
                sTextToFind = " qqnl" & x & Chr(32)
                sReplacementText = "?" & Chr(34) & "^p" & sSpeakerName
                Call pfSingleFindReplace(sTextToFind, sReplacementText)
            
                sMrMs = ""                       'clear variables before loop
                sLastName = ""
                sSpeakerName = ""
            
                x = x + 1                        'add 1 to x for next speaker name
                drSpeakerName.MoveNext           'go to next speaker name
                
            Loop                                 'back up to the top
        Else                                     'upon completion
            MsgBox "There are no records in the recordset."
        End If
    End If
    
    drSpeakerName.Close                          'close the recordset
    Set drSpeakerName = Nothing                  'clean up


    oWordDoc.SaveAs2 FileName:=cJob.DocPath.CourtCover
    oWordDoc.Close
    oWordApp.Quit

    qdf.Close

    If sJurisdiction <> "FDA" And sJurisdiction <> "Food and Drug Administration" And sJurisdiction <> "eScribers NH" And sJurisdiction <> "eScribers Bankruptcy" And sJurisdiction <> "eScribers New Jersey" Then
        Call pfHeaders
        Call fDynamicHeaders
    End If


    Call pfClearGlobals
End Sub

