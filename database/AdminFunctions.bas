Attribute VB_Name = "AdminFunctions"

Option Compare Database
Option Explicit
'============================================================================
'class module cmAdminFunctions

'variables:
'   NONE

'functions:
    'pfUpdateCheckboxStatus:         Description:  updates Statuses checkbox field you specify
    '                                Arguments:    sStatusesField
    'pfDownloadfromFTP:              Description:  downloads files
        '                            Arguments:    NONE
    'pfDownloadFTPsite:              Description:  downloads files modified today (a.k.a. new files on FTP)
        '                            Arguments:    mySession
    'pfCheckFolderExistence:         Description:  checks for Audio, Transcripts, FTP, WorkingFiles, Notes subfolders and RoughDraft and creates if not exists
        '                            Arguments:    NONE
    'pfCommunicationHistoryAdd:      Description:  adds entry to CommunicationHistory
        '                            Arguments:    CHTopic
    'pfStripIllegalChar:             Description:  strips illegal characters from input
        '                            Arguments:    StrInput
    'pfGetFolder:                    Description:  gets folder
        '                            Arguments:    Folders, EntryID, StoreID, fld
    'pfBrowseForFolder:              Description:  browses for folder
        '                            Arguments:    StrSavePath, optional OpenAt
    'pfSingleBAScrapeSpecificBarNo:  Description:  gets one bar number's info from the WA bar website (pick any range from 1 to 55000)
        '                            Arguments:    sWebSiteBarNo
    'pfScrapingBALoop:               Description:  gets a range you specify of bar numbers' info from the WA bar website (pick any range from 1 to 55000)
        '                            Arguments:    vWebSiteBarNo, vWebSiteBarNoGoal
    'pfReformatTable:                Description:  reformats scraped Bar addresses to useable format for table
        '                            Arguments:    NONE
    'pfUpdateCheckboxStatus:         Description:  updates Statuses checkbox field you specify
        '                            Arguments:    sStatusesField
    'pfDelay:                        Description:  sleep function
        '                            Arguments:    lSeconds
    'pfPriorityPointsAlgorithm:      Description:  assigns priority points to various tasks in Tasks table and inserts it into the PriorityPoints field
        '                                          priority scale 1 to 100
        '                            Arguments:    NONE
    'pfDebugSQLStatement:            Description:  debug.prints data source query string
        '                            Arguments:    NONE
    'pfGenerateJobTasks:             Description:  generates job tasks in the Tasks table
        '                            Arguments:    NONE
    'pfDownloadfromFTP:              Description:  downloads files
        '                            Arguments:    NONE
    'pfDownloadFTPsite:              Description:  downloads files modified today (a.k.a. new files on FTP)
        '                            Arguments:    mySession
    'pfProcessFolder:                Description:  process emails in Outlook folder named AccessTest and places them in db as UnprocessedCommunication
        '                            Arguments:    oOutlookMAPIFolder
    'pfFileExists:                   Description:  check if file exists
        '                            Arguments:    path
    'pfAcrobatGetNumPages:           Description:  gets number of pages from PDF and confirms with you
        '                                          IS TOA ON SECOND PAGE?  IF YES, -2 pgs; IF NO, -1 pg
        '                            Arguments:    sCourtDatesID
    'pfReadXML:                      Description:  reads shipping XML and sends "Shipped" email to client
        '                            Arguments:    NONE
    'pfFileRenamePrompt:             Description:  renames transcript to specified name, mainly for contractors
        '                            Arguments:    NONE
    'pfWaitSeconds:                  Description:  waits for a specified number of seconds
        '                            Arguments:    iSeconds
    'pfDailyTaskAddFunction:         Description:  adds static daily tasks to Tasks table
        '                            Arguments:    NONE
    'pfAvailabilitySchedule:         Description:  opens availability calculator
        '                            Arguments:    NONE
    'pfWeeklyTaskAddFunction:        Description:  adds static weekly tasks to Tasks table
        '                            Arguments:    NONE
    'pfMonthlyTaskAddFunction:       Description:  adds static monthly tasks to Tasks table
        '                            Arguments:    NONE
    'pfMoveSelectedMessages:         Description:  move selected messages to network drive
        '                            Arguments:    NONE
    'pfEmailsExport1:                Description:  export specified fields from each mail / item in selected folder
        '                            Arguments:    NONE
    'pfCommHistoryExportSub:         Description:  exports emails to CommunicationsHistory table
        '                            Arguments:    NONE
    'pfAskforNotes:                  Description:  file dialog picker to select notes and copy them to notes folder for job
        '                            Arguments:    NONE
    'pfAskforAudio:                  Description:  file dialog picker to select audio and copy them to audio folder for job
        '                            Arguments:    NONE
    'fWunderlistGetFolders()         Description:  gets list of Wunderlist folders or folder revisions
        '                            Arguments:    NONE
    'fWunderlistGetTasksOnList()     Description:  gets tasks on Wunderlist list
        '                            Arguments:    NONE
    'fWunderlistAdd()                Description:  adds task to Wunderlist
        '                            Arguments:    NONE
    'fWLGenerateJSONInfo             Description:  get info for WL API
        '                            Arguments:    NONE
    'fWunderlistGetLists()           Description:  gets all Wunderlist lists
        '                            Arguments:    NONE
    'pfRCWRuleScraper1()             Description:  builds RCW rule links and citations
        '                            Arguments:    NONE
    'GetLevel()                      Description:  gets header level in word
        '                            Arguments:    NONE
'============================================================================

Public Function pfUpdateCheckboxStatus(sStatusesField As String)
'============================================================================
' Name        : pfUpdateCheckboxStatus
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfUpdateCheckboxStatus(sStatusesField)
' Description : updates Statuses checkbox field you specify
'============================================================================

Dim sUpdateStatusesSQL As String
Dim db As Database

Set db = CurrentDb

sUpdateStatusesSQL = "update [Statuses] set " & sStatusesField & " =(Yes)"

db.Execute sUpdateStatusesSQL

End Function

Public Function pfDebugSQLStatement()
'============================================================================
' Name        : pfDebugSQLStatement
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfDebugSQLStatement
' Description : debug.prints data source query string
'============================================================================

Dim oWordApp As New Word.Application, oWordDoc As Word.Document

oWordApp.Application.Visible = True

Set oWordApp = Nothing
Set oWordDoc = Nothing

Debug.Print oWordApp.Application.ActiveDocument.MailMerge.DataSource.QueryString

oWordApp.Quit
Set oWordApp = Nothing

End Function
Public Function pfDownloadfromFTP()
On Error Resume Next
'============================================================================
' Name        : pfDownloadfromFTP
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfDownloadfromFTP
' Description : downloads files
'============================================================================

Dim seCurrent As New Session


Call pfDownloadFTPsite(seCurrent)

If Err.Number <> 0 Then ' Query for errors
    MsgBox "Error: " & Err.Description
    Err.Clear ' Clear the error
End If
seCurrent.Dispose ' Disconnect, clean up

On Error GoTo 0 ' Restore default error handling

End Function

Public Function pfDownloadFTPsite(ByRef mySession As Session)
'============================================================================
' Name        : pfDownloadFTPsite
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfDownloadFTPsite(mySession)
' Description : downloads files modified today (a.k.a. new files on FTP)
'============================================================================

Dim seopFTPSettings As New SessionOptions
Dim sInProgressPath As String
Dim tropFTPSettings As New TransferOptions

With seopFTPSettings ' Setup session options
    .Protocol = Protocol_Ftp
    .HostName = ""
    .Username = ""
    .password = ""
End With

mySession.Open seopFTPSettings ' Connect
tropFTPSettings.TransferMode = TransferMode_Binary ' Upload files
tropFTPSettings.FileMask = "*>=1D"
sInProgressPath = "\\HUBCLOUD\evoingram\Production\1ToBeEntered\"

Dim transferResult As TransferOperationResult

Set transferResult = mySession.GetFiles("/public_html/ProjectSend/upload/files/", sInProgressPath, False, tropFTPSettings)
transferResult.Check ' Throw on any error

MsgBox "You may now find any files downloaded today in T:\Production\1ToBeEntered\."

End Function


Public Function pfProcessFolder(ByVal oOutlookPickedFolder As Outlook.MAPIFolder)
'============================================================================
' Name        : pfProcessFolder
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfProcessFolder(oOutlookPickedFolder)
' Description : process emails in Outlook folder named AccessTest and places them in db as UnprocessedCommunication
'============================================================================

Dim oOutlookNamespace As Outlook.Namespace
Dim adocOutlookExport As ADODB.Connection
Dim adorstOutlookExport As ADODB.Recordset
Dim oOutLookMAPIFolder As Outlook.MAPIFolder
Dim oOutlookMail As Outlook.MailItem
Dim dReceived As Date
Dim sReceivedTime As String, sEmailHyperlink As String, sTableHyperilnk As String

Set oOutlookNamespace = GetNamespace("MAPI")
Set oOutlookPickedFolder = oOutlookNamespace.PickFolder
Set adocOutlookExport = CreateObject("ADODB.Connection")
Set adorstOutlookExport = CreateObject("ADODB.Recordset")

For Each oOutlookMail In oOutlookPickedFolder.Items
    dReceived = oOutlookMail.ReceivedTime
    sReceivedTime = Format(dReceived, "YYYYMMDD-hhmm")
    oOutlookMail.SaveAs "T:\Database\Emails\" & sReceivedTime & "-Email.msg", 3
Next

If (oOutlookPickedFolder.Folders.Count > 0) Then
    For Each oOutLookMAPIFolder In oOutlookPickedFolder.Folders
        pfProcessFolder oOutLookMAPIFolder
    Next
End If

adocOutlookExport.Open "DSN=OutlookExportAQCP" 'DSN and target file must exist.
adorstOutlookExport.Open "SELECT * FROM Emails", adocOutlookExport, adOpenDynamic, adLockOptimistic

For Each oOutlookMail In oOutlookPickedFolder.Items
    dReceived = oOutlookMail.ReceivedTime
    sReceivedTime = Format(dReceived, "YYYYMMDD-hhmm")
    oOutlookMail.SaveAs "T:\Database\Emails\" & sReceivedTime & "-Email.msg", 3
    sEmailHyperlink = "T:\Database\Emails\" & sReceivedTime & "-Email.msg"
    sTableHyperilnk = sReceivedTime & "-Email" & "#" & sEmailHyperlink & "#"
    adorstOutlookExport.AddNew
    adorstOutlookExport("FileHyperlink") = sTableHyperilnk
    adorstOutlookExport("DateCreated") = dReceived
    adorstOutlookExport.Update
Next

If (oOutlookPickedFolder.Folders.Count > 0) Then
    For Each oOutLookMAPIFolder In oOutlookPickedFolder.Folders
        pfProcessFolder oOutLookMAPIFolder
    Next
End If

adorstOutlookExport.Close
Set adorstOutlookExport = Nothing
Set adocOutlookExport = Nothing
Set oOutlookNamespace = Nothing
Set oOutlookPickedFolder = Nothing
 
End Function
Public Function pfFileExists(ByVal path_ As String) As Boolean
'============================================================================
' Name        : pfFileExists
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfFileExists(ByVal path)
' Description : check if file exists
'============================================================================
Dim FileExists As Variant
FileExists = (Len(Dir(path_)) > 0)

End Function

Public Function pfAcrobatGetNumPages(sCourtDatesID)
'============================================================================
' Name        : pfAcrobatGetNumPages
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfAcrobatGetNumPages(sCourtDatesID)
' Description : gets number of pages from PDF and confirms with you
                'IS TOA ON SECOND PAGE?
                    'IF YES, MINUS TWO PAGES
                    'IF NO, MINUS ONE PAGE
'============================================================================

Dim dbAQC As Database
Dim qdf As QueryDef
Dim oAcrobatDoc As Object
Dim sTranscriptPDFPath As String, sActualQuantity1 As String, sActualQuantity As String
Dim sQuestion As String, sAnswer As String, sSQL As String

Set oAcrobatDoc = New AcroPDDoc

sTranscriptPDFPath = "I:\" & sCourtDatesID & "\Backups\" & sCourtDatesID & "-Transcript-FINAL.pdf"

oAcrobatDoc.Open (sTranscriptPDFPath) 'update file location

sActualQuantity = oAcrobatDoc.GetNumPages
sQuestion = "This transcript came to " & sActualQuantity & " pages.  Is the table of authorities on a separate page from the CoA?"
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

If sAnswer = vbNo Then 'IF NO THEN THIS HAPPENS

    MsgBox "Page count will be reduced by only one."
    
    sActualQuantity = sActualQuantity - 1
    sQuestion = "This transcript came to " & sActualQuantity & " billable pages.  Is that page count correct?"
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
    
    If sAnswer = vbNo Then 'IF NO THEN THIS HAPPENS
    
        sActualQuantity1 = InputBox("How many billable pages was this transcript?")
        sActualQuantity = sActualQuantity1
        
    Else 'if yes then this happens
    End If
    
Else 'if yes then this happens

    MsgBox "Page count will be reduced by two."
    
    sActualQuantity = sActualQuantity - 2
    sQuestion = "This transcript came to " & sActualQuantity & " billable pages.  Is that page count correct?"
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
    
    If sAnswer = vbNo Then 'IF NO THEN THIS HAPPENS
    
        sActualQuantity1 = InputBox("How many billable pages was this transcript?")
        sActualQuantity = sActualQuantity1
        
    Else 'if yes then this happens
    
        sActualQuantity1 = InputBox("How many billable pages was this transcript?")
        sActualQuantity = sActualQuantity1
        
    End If
    
End If

oAcrobatDoc.Close

Set dbAQC = CurrentDb



sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

sSQL = "UPDATE [CourtDates] SET [CourtDates].[ActualQuantity] = " & sActualQuantity & " WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
Set qdf = dbAQC.CreateQueryDef("", sSQL)
dbAQC.Execute sSQL

Set qdf = Nothing

DoCmd.OpenQuery "FinalUnitPriceQuery"  'PRE-QUERY FOR FINAL SUBTOTAL
dbAQC.Execute "INVUpdateFinalUnitPriceQuery" 'UPDATES FINAL SUBTOTAL
dbAQC.Close
End Function

Public Function pfReadXML()
'============================================================================
' Name        : pfReadXML
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfReadXML
' Description : reads shipping XML and sends "Shipped" email to client
'============================================================================


Dim sFullOutputDonePath As String, sTrackingNumber As String
Dim sOutputPath As String, sFullOutputPath As String
Dim dShipDate As Date, dShipDateFormatted As Date
Dim rstCurrentJob As DAO.Recordset
Dim formDOM As DOMDocument60    'Currently opened xml file
Dim ixmlRoot As IXMLDOMElement
Dim Rng As Range

sOutputPath = Dir("T:\Production\4ShippingXMLs\Output\")
Do While Len(sOutputPath) > 0
    sFullOutputPath = "T:\Production\4ShippingXMLs\Output\" & sOutputPath
    sFullOutputDonePath = "T:\Production\4ShippingXMLs\done" & sOutputPath
    
    Set formDOM = New DOMDocument60 'Open the xml file
    formDOM.resolveExternals = False 'using schema yes/no true/false
    formDOM.validateOnParse = False 'Parser validate document?  Still parses well-formed XML
    formDOM.Load (sFullOutputPath)
    
    Set ixmlRoot = formDOM.DocumentElement 'Get document reference
    
    sCourtDatesID = ixmlRoot.SelectSingleNode("//DAZzle/Package/ReferenceID").Text
    dShipDate = ixmlRoot.SelectSingleNode("//DAZzle/Package/PostmarkDate").Text
    dShipDateFormatted = DateSerial(Left(dShipDate, 4), Mid(dShipDate, 5, 2), Right(dShipDate, 2))
    sTrackingNumber = ixmlRoot.SelectSingleNode("//DAZzle/Package/PIC").Text
    
    Set rstCurrentJob = CurrentDb.OpenRecordset("SELECT * FROM CourtDates WHERE ID = " & sCourtDatesID & ";")
    
    rstCurrentJob.Edit
    rstCurrentJob.Fields("ShipDate").Value = dShipDateFormatted
    rstCurrentJob.Fields("TrackingNumber").Value = sTrackingNumber
    rstCurrentJob.Update
    
    Set rstCurrentJob = CurrentDb.OpenRecordset("SELECT * FROM [TR-Court-Q-3] WHERE [ID] = " & sCourtDatesID & ";")
    
    'global variables to use in next function
    sParty1 = rstCurrentJob.Fields("Party1").Value
    sParty2 = rstCurrentJob.Fields("Party2").Value
    sCaseNumber1 = rstCurrentJob.Fields("CaseNumber1").Value
    dHearingDate = rstCurrentJob.Fields("HearingDate").Value
    sAudioLength = rstCurrentJob.Fields("AudioLength").Value
    rstCurrentJob.Close
    
    sOutputPath = Dir
    
    Name sFullOutputPath As sFullOutputDonePath 'move file to other folder
    
    Call pfSendWordDocAsEmail("Shipped", "Transcript Shipped")
       
Loop

End Function

Public Function pfFileRenamePrompt()
'============================================================================
' Name        : pfFileRenamePrompt
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfFileRenamePrompt
' Description : renames transcript to specified name, mainly for contractors
'============================================================================

Dim db As Database

Dim sUserInput As String, sFinalTranscriptPath As String
Dim sChkBxFiledNotFiled As String, sCoverPath As String

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
sFinalTranscriptPath = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-FINAL.docx"
sCoverPath = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-CourtCover.docx"
sUserInput = InputBox("Enter the desired document name without the extension" & Chr(13) & "Weber Format:  A169195_transcript_2018-09-18_IngramEricaL" & Chr(13) & "AMOR Format: Audio Name" & Chr(13) & "eScribers format [JobNumber]_[DRAFT]_Date", "Rename your document." & Chr(13) & "Weber Format:  A169195_transcript_2018-09-18_IngramEricaL" & Chr(13) & "AMOR Format: Audio Name" & Chr(13) & "eScribers format [JobNumber]_[DRAFT]_Date", "Enter the new name for the transcript here, without the extension." & Chr(13) & "Weber Format:  A169195_transcript_2018-09-18_IngramEricaL" & Chr(13) & "AMOR Format: Audio Name" & Chr(13) & "eScribers format [JobNumber]_[DRAFT]_Date")

If sUserInput = "Enter the new name for the transcript here, without the extension." Or sUserInput = "" Then
    Exit Function
End If

sClientTranscriptName = "I:\" & sCourtDatesID & "\Transcripts\" & sUserInput & ".docx"

FileCopy sCoverPath, sFinalTranscriptPath
Name sFinalTranscriptPath As sClientTranscriptName

MsgBox "File renamed to " & sClientTranscriptName & ".  Next we will deliver the transcript."

Call pfGenericExportandMailMerge("Case", "Stage4s\ContractorTranscriptsReady")
Call pfSendWordDocAsEmail("ContractorTranscriptsReady", "Transcripts Ready", sClientTranscriptName)

sChkBxFiledNotFiled = "update [CourtDates] set FiledNotFiled =(Yes) WHERE ID=" & sCourtDatesID & ";"

CurrentDb.Execute sChkBxFiledNotFiled

MsgBox "Transcript has been delivered.  Next, let's do some admin stuff."

End Function


Public Function pfWaitSeconds(iSeconds As Integer)
On Error GoTo lERR

'============================================================================
' Name        : pfWaitSeconds
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfWaitSeconds(iSeconds)
' Description : waits for a specified number of seconds
'============================================================================
  
Dim dCurrentTime As Date

dCurrentTime = DateAdd("s", iSeconds, Now)

Do 'yield to other programs
    pfDelay 100
    DoEvents
Loop Until Now >= dCurrentTime

lEXIT:
Exit Function

lERR:
MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modDateTime.WaitSeconds"
Resume lEXIT

End Function

Public Function pfAvailabilitySchedule()
'============================================================================
' Name        : pfAvailabilitySchedule
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfAvailabilitySchedule
' Description : opens availability calculator
'============================================================================

DoCmd.OpenForm (Forms![SBFM-Availability])
End Function

Public Function pfCheckFolderExistence()
'============================================================================
' Name        : pfCheckFolderExistence
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfCheckFolderExistence
' Description : checks for Audio, Transcripts, FTP, WorkingFiles, Notes subfolders and RoughDraft and creates if not exists
'============================================================================
Dim sIPJobPath As String, sTemplatePath As String


Call pfCurrentCaseInfo  'refresh transcript info
sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
sIPJobPath = "I:\" & sCourtDatesID
sTemplatePath = "T:\Database\Templates\"

If Len(Dir(sIPJobPath, vbDirectory)) = 0 Then
   MkDir sIPJobPath
End If
If Len(Dir(sIPJobPath & "\FTP\", vbDirectory)) = 0 Then
    MkDir sIPJobPath & "\FTP\"
End If
If Len(Dir(sIPJobPath & "\WorkingFiles\", vbDirectory)) = 0 Then
    MkDir sIPJobPath & "\WorkingFiles\"
End If
If Len(Dir(sIPJobPath & "\Audio\", vbDirectory)) = 0 Then
    MkDir sIPJobPath & "\Audio\"
End If
If Len(Dir(sIPJobPath & "\Transcripts\", vbDirectory)) = 0 Then
    MkDir sIPJobPath & "\Transcripts\"
End If
If Len(Dir(sIPJobPath & "\Notes", vbDirectory)) = 0 Then
    MkDir sIPJobPath & "\Notes"
End If
If Len(Dir(sIPJobPath & "\Generated", vbDirectory)) = 0 Then
    MkDir sIPJobPath & "\Generated"
End If
If Len(Dir(sIPJobPath & "\Backups", vbDirectory)) = 0 Then
    MkDir sIPJobPath & "\Backups"
End If


If sJurisdiction Like "*Food and Drug Administration*" Then

    If Len(Dir(sIPJobPath & "\" & "RoughDraft.docx")) = 0 Then
        FileCopy sTemplatePath & "Stage2s\RoughDraft-FDA.docx", sIPJobPath & "\" & "RoughDraft.docx"
    End If
    
ElseIf sJurisdiction Like "*NonCourt*" Then

    If Len(Dir(sIPJobPath & "\" & "RoughDraft.docx")) = 0 Then
        FileCopy sTemplatePath & "Stage2s\RoughDraft.docx", sIPJobPath & "\" & "RoughDraft.docx"
    End If
    
ElseIf sJurisdiction Like "*FDA*" Then

    If Len(Dir(sIPJobPath & "\" & "RoughDraft.docx")) = 0 Then
        FileCopy sTemplatePath & "Stage2s\RoughDraft-FDA.docx", sIPJobPath & "\" & "RoughDraft.docx"
    End If
    
ElseIf sJurisdiction Like "*Weber Oregon*" Then

    If Len(Dir(sIPJobPath & "\" & "RoughDraft.docx")) = 0 Then
        FileCopy sTemplatePath & "Stage2s\RoughDraft-WeberOR.docx", sIPJobPath & "\" & "RoughDraft.docx"
    End If
    
ElseIf sJurisdiction Like "*Weber Nevada*" Then

    If Len(Dir(sIPJobPath & "\" & "RoughDraft.docx")) = 0 Then
        FileCopy sTemplatePath & "Stage2s\RoughDraft-WeberNV.docx", sIPJobPath & "\" & "RoughDraft.docx"
    End If
    
ElseIf sJurisdiction Like "*Weber Bankruptcy*" Then

    If Len(Dir(sIPJobPath & "\" & "RoughDraft.docx")) = 0 Then
        FileCopy sTemplatePath & "Stage2s\RoughDraft.docx", sIPJobPath & "\" & "RoughDraft.docx"
    End If
    
ElseIf sJurisdiction Like "*AVT*" Then

    If Len(Dir(sIPJobPath & "\" & "RoughDraft.docx")) = 0 Then
        FileCopy sTemplatePath & "Stage2s\RoughDraft.docx", sIPJobPath & "\" & "RoughDraft.docx"
    End If
    
Else

    If Len(Dir(sIPJobPath & "\" & "RoughDraft.docx")) = 0 Then
        FileCopy sTemplatePath & "Stage2s\RoughDraft.docx", sIPJobPath & "\" & "RoughDraft.docx"
    End If
    
End If
Call pfClearGlobals
End Function

Public Function pfCommunicationHistoryAdd(sCHTopic As String)
'============================================================================
' Name        : pfCommunicationHistoryAdd
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfCommunicationHistoryAdd(sCHTopic)
' Description : adds entry to CommunicationHistory
'============================================================================

Dim db As DAO.Database
Dim rstCHAdd As DAO.Recordset
Dim sCHHyperlink As String, sCurrentDocPath As String

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
sCurrentDocPath = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-" & sCHTopic & ".docx"
sCHHyperlink = sCourtDatesID & "-" & sCHTopic & "#" & sCurrentDocPath & "#"

Set db = CurrentDb
Set rstCHAdd = db.OpenRecordset("CommunicationHistory")

rstCHAdd.AddNew
rstCHAdd("FileHyperlink").Value = sCHHyperlink
rstCHAdd("DateCreated").Value = Now
rstCHAdd("CourtDatesID").Value = sCourtDatesID
rstCHAdd.Update

rstCHAdd.Close

End Function
 
Public Function pfStripIllegalChar(sInput As String)
'============================================================================
' Name        : pfStripIllegalChar
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfStripIllegalChar(StrInput)
' Description : strips illegal characters from input
'============================================================================

Dim oRegex As Object
Dim StripIllegalChar
 
Set oRegex = CreateObject("vbscript.regexp")
oRegex.Pattern = "[\" & Chr(34) & "\!\@\#\$\%\^\&\*\(\)\=\+\|\[\]\{\}\`\'\;\:\<\>\?\/\,]"
oRegex.IgnoreCase = True
oRegex.Global = True
StripIllegalChar = oRegex.Replace(sInput, "")

ExitFunction:
Set oRegex = Nothing
     
 
End Function
 
 
Public Function pfGetFolder(Folders As Collection, EntryID As Collection, StoreID As Collection, fld As MAPIFolder)
'============================================================================
' Name        : pfGetFolder
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfGetFolder(Folders, EntryID, StoreID, fld)
' Description : gets folder
'============================================================================
'
Dim SubFolder As MAPIFolder

Folders.Add fld.FolderPath
EntryID.Add fld.EntryID
StoreID.Add fld.StoreID

For Each SubFolder In fld.Folders
    pfGetFolder Folders, EntryID, StoreID, SubFolder
Next SubFolder
      
ExitSub:
Set SubFolder = Nothing

End Function
 
Public Function pfBrowseForFolder(sSavePath As String, Optional OpenAt As String) As String
'============================================================================
' Name        : pfBrowseForFolder
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfBrowseForFolder(StrSavePath, OpenAt)
                'OpenAt optional
' Description : browses for folder
'============================================================================

Dim oShell As Object, oBrowsedFolder As Object
Dim vEnvUserProfile As Variant
 
vEnvUserProfile = CStr(Environ("USERPROFILE"))
 
Set oShell = CreateObject("Shell.Application")
 
Set oBrowsedFolder = oShell.BrowseForFolder(0, "Please choose a folder", 0, vEnvUserProfile & "\My Documents\")
 
sSavePath = oBrowsedFolder.Self.Path

On Error Resume Next
On Error GoTo 0
     
ExitFunction:
Set oShell = Nothing
   
End Function
 


Public Function pfSingleBAScrapeSpecificBarNo(sWebSiteBarNo As String)
'============================================================================
' Name        : pfSingleBAScrapeSpecificBarNo
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfSingleBAScrapeSpecificBarNo(sWebSiteBarNo)
' Description:  gets one bar number's info from the WA bar website (pick any range from 1 to 55000)
'============================================================================

Dim rstBarAddresses As DAO.Recordset

Dim oCompanyName As Object, oBarName As Object, oBarNumber As Object, oEligibility As Object
Dim oActiveL As Object, oAdmitDate As Object, oAddress As Object, oEmail As Object
Dim oPhone As Object, oFax As Object, oPracticeArea As Object, oInternetE As Object

Dim sWebsiteLink As String
Dim sCompanyName As String, sBarName As String, sBarNumber As String, sEligibility As String, sActiveL As String
Dim sAdmitDate As String, sAddress As String, sEmail As String, sPhone As String, sFax As String, sPracticeArea As String

sWebsiteLink = "https://www.mywsba.org/PersonifyEbusiness/LegalDirectory/LegalProfile.aspx?Usr_ID=0000000" & sWebSiteBarNo

Set oInternetE = CreateObject("InternetExplorer.Application")
oInternetE.Visible = False
oInternetE.Navigate sWebsiteLink

While oInternetE.Busy 'Wait while oInternetE loading...
    DoEvents
Wend

' ********************************************************************** get the following info
pfDelay 5 'wait a little bit
Set oBarName = oInternetE.Document.getElementById("dnn_ctr2977_DNNWebControlContainer_ctl00_lblMemberName")
sBarName = oBarName.innerText
Set oBarNumber = oInternetE.Document.getElementById("dnn_ctr2977_DNNWebControlContainer_ctl00_lblMemberNo")
sBarNumber = oBarNumber.innerText
Set oEligibility = oInternetE.Document.getElementById("dnn_ctr2977_DNNWebControlContainer_ctl00_lblEligibleToPractice")
sEligibility = oEligibility.innerText
Set oActiveL = oInternetE.Document.getElementById("dnn_ctr2977_DNNWebControlContainer_ctl00_lblStatus")
sActiveL = oActiveL.innerText
Set oCompanyName = oInternetE.Document.getElementById("dnn_ctr2977_DNNWebControlContainer_ctl00_lblAddCompanyName")
sCompanyName = oCompanyName.innerText
Set oAddress = oInternetE.Document.getElementById("dnn_ctr2977_DNNWebControlContainer_ctl00_lblAddress")
sAddress = oAddress.innerHTML
Set oPhone = oInternetE.Document.getElementById("dnn_ctr2977_DNNWebControlContainer_ctl00_lblPhone")
sPhone = oPhone.innerText
Set oPracticeArea = oInternetE.Document.getElementById("dnn_ctr2977_DNNWebControlContainer_ctl00_lblPracticeAreas")
sPracticeArea = oPracticeArea.innerText

'print newly acquired info to debug window
Debug.Print sBarName & Chr(13) & sBarNumber & Chr(13) & sEligibility & Chr(13) & sActiveL & Chr(13) & sAdmitDate & Chr(13) & _
sCompanyName & Chr(13) & sAddress & Chr(13) & sEmail & Chr(13) & sPhone & Chr(13) & sFax & Chr(13) & sPracticeArea

'**********************************************************************

oInternetE.Quit
Set oInternetE = Nothing
'Application.ScreenUpdating = True

Set rstBarAddresses = CurrentDb.OpenRecordset("BarAddresses") 'add to a special table in access

rstBarAddresses.AddNew
rstBarAddresses.Fields("BarName").Value = sBarName
rstBarAddresses.Fields("BarNumber").Value = sBarNumber
rstBarAddresses.Fields("Eligibility").Value = sEligibility
rstBarAddresses.Fields("ActiveL").Value = sActiveL
rstBarAddresses.Fields("Company").Value = sCompanyName
rstBarAddresses.Fields("Address").Value = sAddress
rstBarAddresses.Fields("Phone").Value = sPhone
rstBarAddresses.Fields("PracticeArea").Value = sPracticeArea
rstBarAddresses.Update

rstBarAddresses.Close
Set rstBarAddresses = Nothing

End Function

Public Function pfScrapingBALoop(sWebSiteBarNo As String, sWebSiteBarNoGoal As String)
On Error Resume Next
'============================================================================
' Name        : pfScrapingBALoop
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfScrapingBALoop(sWebSiteBarNo, sWebSiteBarNoGoal)
' Description:  gets a range you specify of bar numbers' info from the WA bar website (pick any range from 1 to 55000)
'============================================================================

Dim rstBarAddresses As DAO.Recordset

Dim oCompanyName As Object, oBarName As Object, oBarNumber As Object, oEligibility As Object
Dim oActiveL As Object, oAdmitDate As Object, oAddress As Object, oEmail As Object
Dim oPhone As Object, oFax As Object, oPracticeArea As Object, oInternetE As Object
Dim sWebsiteLink As String
Dim sCompanyName As String, sBarName As String, sBarNumber As String, sEligibility As String, sActiveL As String
Dim sAdmitDate As String, sAddress As String, sEmail As String, sPhone As String, sFax As String, sPracticeArea As String



Do While sWebSiteBarNo < sWebSiteBarNoGoal

    sWebsiteLink = "https://www.mywsba.org/PersonifyEbusiness/LegalDirectory/LegalProfile.aspx?Usr_ID=0000000" & sWebSiteBarNo
    
    Set oInternetE = CreateObject("InternetExplorer.Application")
    oInternetE.Visible = False
    oInternetE.Navigate sWebsiteLink
    
    While oInternetE.Busy ' Wait while oInternetE loading.
        DoEvents
    Wend
    
    ' ********************************************************************** get the following info
    pfDelay 5 'wait a little bit
    Set oBarName = oInternetE.Document.getElementById("dnn_ctr2977_DNNWebControlContainer_ctl00_lblMemberName")
    sBarName = oBarName.innerText
    Set oBarNumber = oInternetE.Document.getElementById("dnn_ctr2977_DNNWebControlContainer_ctl00_lblMemberNo")
    sBarNumber = oBarNumber.innerText
    Set oEligibility = oInternetE.Document.getElementById("dnn_ctr2977_DNNWebControlContainer_ctl00_lblEligibleToPractice")
    sEligibility = oEligibility.innerText
    Set oActiveL = oInternetE.Document.getElementById("dnn_ctr2977_DNNWebControlContainer_ctl00_lblStatus")
    sActiveL = oActiveL.innerText
    Set oCompanyName = oInternetE.Document.getElementById("dnn_ctr2977_DNNWebControlContainer_ctl00_lblAddCompanyName")
    sCompanyName = oCompanyName.innerText
    Set oAddress = oInternetE.Document.getElementById("dnn_ctr2977_DNNWebControlContainer_ctl00_lblAddress")
    sAddress = oAddress.innerHTML
    Set oPhone = oInternetE.Document.getElementById("dnn_ctr2977_DNNWebControlContainer_ctl00_lblPhone")
    sPhone = oPhone.innerText
    Set oPracticeArea = oInternetE.Document.getElementById("dnn_ctr2977_DNNWebControlContainer_ctl00_lblPracticeAreas")
    sPracticeArea = oPracticeArea.innerText
    
    'print newly acquired info to debug window
    Debug.Print sBarName & Chr(13) & sBarNumber & Chr(13) & sEligibility & Chr(13) & sActiveL & Chr(13) & sAdmitDate & Chr(13) & _
    sCompanyName & Chr(13) & sAddress & Chr(13) & sEmail & Chr(13) & sPhone & Chr(13) & sFax & Chr(13) & sPracticeArea
    
    '**********************************************************************
    
    oInternetE.Quit
    Set oInternetE = Nothing
    'Application.ScreenUpdating = True
    
    
    Set rstBarAddresses = CurrentDb.OpenRecordset("BarAddresses") 'add to a special table in access
    
    rstBarAddresses.AddNew
    rstBarAddresses.Fields("BarName").Value = sBarName
    rstBarAddresses.Fields("BarNumber").Value = sBarNumber
    rstBarAddresses.Fields("Eligibility").Value = sEligibility
    rstBarAddresses.Fields("ActiveL").Value = sActiveL
    rstBarAddresses.Fields("Company").Value = sCompanyName
    rstBarAddresses.Fields("Address").Value = sAddress
    rstBarAddresses.Fields("Phone").Value = sPhone
    rstBarAddresses.Fields("PracticeArea").Value = sPracticeArea
    rstBarAddresses.Update
    rstBarAddresses.Close
    Set rstBarAddresses = Nothing
    
    sWebSiteBarNo = sWebSiteBarNo + 1 'move on to next bar number
    
    pfDelay 22
        
Loop
End Function


Public Function pfReformatTable()
'============================================================================
' Name        : pfReformatTable
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfReformatTable
' Description : reformats scraped Bar addresses to useable format for table
'============================================================================

'change all commas to semicolons
'export to xls
'delete columns ID, Eligibility, ActiveL, practice area
'move company column to first
'insert mrms column
'move barname column to 3rd
'after barname insert MiddleName LastName EmailAddress JobTitle columns
'move phone column to after jobtitle
'make all job titles attorney & all email addresses inquiries@aquoco.co
'insert columns MobilePhone FaxNumber after phone
'move address column to after faxnumber
'insert columns City State ZIP WebPage Notes FactorApvlID FactoringApproved after address
'change barname spaces to commas
'save as CSV
'open again in excel
'delete middlename column
'change state column space to comma
'save
'close
'delete "
'open again in excel to make sure it looks right


'ID Company MrMs LastName FirstName EmailAddress JobTitle BusinessPhone MobilePhone
'FaxNumber Address City State ZIP WebPage Notes FactorApvlID FactoringApproved



'THIS IS FOR SEPARATING OUT BAR ADDRESSES IN THE TABLE
Dim db          As DAO.Database
Dim rs          As DAO.Recordset
Dim rsADD       As DAO.Recordset

Dim strSQL      As String
Dim strField1   As String
Dim strField2   As String
Dim varData     As Variant
Dim i           As Integer

Set db = CurrentDb

' Select all eligible fields (have a comma) and unprocessed (Field2 is Null)
strSQL = "SELECT Address, Field2 FROM BarAddresses WHERE ([Address] Like ""*<br>*"") AND ([Field2] Is Null)"

Set rsADD = db.OpenRecordset("BarAddresses", dbOpenDynaset, dbAppendOnly)

Set rs = db.OpenRecordset(strSQL, dbOpenDynaset)
With rs
While Not .EOF
    strField1 = !Field1
    varData = Split(strField1, "<br>") ' Get all comma delimited fields

    ' Update First Record
    .Edit
    !Field2 = Trim(varData(0)) ' remove spaces before writing new fields
    .Update

    ' Add records with same first field
    ' and new fields for remaining data at end of string
    For i = 1 To UBound(varData)
        With rsADD
            .AddNew
            !Field1 = strField1
            !Field2 = Trim(varData(i)) ' remove spaces before writing new fields
            .Update
        End With
    Next
    .MoveNext
Wend

.Close
rsADD.Close

End With

Set rsADD = Nothing
Set rs = Nothing
db.Close
Set db = Nothing

End Function


Public Function pfDelay(lSeconds As Long)
'============================================================================
' Name        : pfDelay
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfDelay(lSeconds)
' Description : sleep function
'============================================================================

Dim dEndTime As Date

dEndTime = DateAdd("s", lSeconds, Now())
Do While Now() < dEndTime
DoEvents
Loop
End Function


Public Function pfPriorityPointsAlgorithm()
'============================================================================
' Name        : pfPriorityPointsAlgorithm
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfPriorityPointsAlgorithm
' Description:  assigns priority points to various tasks in Tasks table and inserts it into the PriorityPoints field
'               priority scale 1 to 100
'============================================================================

Dim iPriorityPoints As Integer, iTimeLength As Integer
Dim sPriority As String, sCategory As String, bCompleted As Boolean
Dim rstTasks As DAO.Recordset
Dim dDue As Date

Set rstTasks = CurrentDb.OpenRecordset("Tasks")

If Not (rstTasks.EOF And rstTasks.BOF) Then

        rstTasks.MoveFirst
        
            Do Until rstTasks.EOF = True
            
                dDue = rstTasks.Fields("Due Date").Value
                bCompleted = rstTasks.Fields("Completed").Value
                sCategory = rstTasks.Fields("Category").Value
                iTimeLength = rstTasks.Fields("TimeLength").Value
                iPriorityPoints = 0
                
                If ((DateDiff("d", Now, dDue)) < 1) Then
                    iPriorityPoints = iPriorityPoints + 70
                ElseIf ((DateDiff("d", Now, dDue)) = 1) Then
                    iPriorityPoints = iPriorityPoints + 60
                ElseIf ((DateDiff("d", Now, dDue)) < 4) And ((DateDiff("d", Now, dDue)) > 1) Then
                    iPriorityPoints = iPriorityPoints + 50
                ElseIf ((DateDiff("d", Now, dDue)) > 3) And ((DateDiff("d", Now, dDue)) < 9) Then
                    iPriorityPoints = iPriorityPoints + 40
                ElseIf ((DateDiff("d", Now, dDue)) > 8) And ((DateDiff("d", Now, dDue)) < 15) Then
                    iPriorityPoints = iPriorityPoints + 30
                ElseIf ((DateDiff("d", Now, dDue)) > 14) Then
                    iPriorityPoints = iPriorityPoints + 15
                Else
                End If
                
                If iTimeLength < 11 Then
                    iPriorityPoints = iPriorityPoints + 20
                Else
                End If
                
                If sCategory = "production" Then
                    iPriorityPoints = iPriorityPoints + 10
                ElseIf sCategory = "leisure" Then
                    iPriorityPoints = iPriorityPoints + 20
                ElseIf sCategory = "personal" Then
                    iPriorityPoints = iPriorityPoints + 20
                Else
                End If
                
                If sPriority = "(1) Stage 1" Then
                    iPriorityPoints = iPriorityPoints + 80
                ElseIf sPriority = "(2) Stage 2" Then
                    iPriorityPoints = iPriorityPoints + 55
                ElseIf sPriority = "(3) Stage 3" Then
                    iPriorityPoints = iPriorityPoints + 30
                ElseIf sPriority = "(4) Stage 4" Then
                    iPriorityPoints = iPriorityPoints + 5
            Else
            End If
            
            If sPriority = "Waiting For" Then
                iPriorityPoints = 0
            Else
            End If
            
            If bCompleted = True Then
                iPriorityPoints = 0
            End If
            
            rstTasks.Edit
            rstTasks.Fields("PriorityPoints").Value = iPriorityPoints
            rstTasks.Update
            rstTasks.MoveNext
        
        Loop
    Else
End If
    
MsgBox "Done assigning priority points to tasks."

End Function
Public Function pfGenerateJobTasks()
'============================================================================
' Name        : pfGenerateJobTasks
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfGenerateJobTasks
' Description : generates job tasks in the Tasks table
'============================================================================

Dim sTaskTitle As String, sTaskCategory As String, sPriority As String, sTaskDescription As String, sIPJobPath As String
Dim iTypingTime As Integer, iAudioProofTime As Integer, iTaskMinuteLength As Integer
Dim dStart As Date, dDue As Date
Dim qdf As QueryDef
Dim rstTasks As DAO.Recordset


Call pfCurrentCaseInfo  'refresh transcript info

sIPJobPath = "I:\" & sCourtDatesID
sTaskTitle = "(1.1) Enter job & contacts into database:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
dDue = Now + 1
dStart = Date
iTaskMinuteLength = "2"
sTaskCategory = "production"
sPriority = "(1) Stage 1"

sTaskDescription = "|Case Name:  " & sParty1 & " v. " & sParty2 & "   |" & Chr(13) & _
"|Case Nos.:  " & sCaseNumber1 & "   |   " & sCaseNumber2 & "   |" & Chr(13) & _
"|Due Date:  " & dDue & "   |   Turnaround:  " & sTurnaroundTime & " calendar days   |" & _
"|Client:   " & sCompany & "   |   Folder:   " & sIPJobPath & "   |" & Chr(13) & _
"|Exp. Advance/Deposit Date:  " & dExpectedAdvanceDate & "   |" & Chr(13) & _
"|Exp. Rebate Date:  " & dExpectedRebateDate & "   |" & Chr(13) & _
"|Estimate:  " & sSubtotal & "   |"


Set rstTasks = CurrentDb.OpenRecordset("Tasks")

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Start Date").Value = dStart
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


sTaskTitle = "(1.2) Payment:  If factored, proceed with set-up.  If not, send invoice & wait for payment :  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
dDue = Now + 1
iTaskMinuteLength = "2"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Start Date").Value = dStart
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


sTaskTitle = "(1.3) Generate documents: cover, autocorrect, AGshortcuts, Xero CSV, CD label, transcripts ready, package-enclosed letter:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
dDue = Now + 1
iTaskMinuteLength = "2"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Start Date").Value = dStart
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTypingTime = Round((((sAudioLength * 3) / 60) + 1), 0)

For i = 1 To iTypingTime
    sTaskTitle = "(2.1) Type:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
    dDue = dDueDate - 3
    iTaskMinuteLength = "60"
    sPriority = "(2) Stage 2"
    
    rstTasks.AddNew
    rstTasks.Fields("Title").Value = sTaskTitle
    rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
    rstTasks.Fields("Priority").Value = sPriority
    rstTasks.Fields("Start Date").Value = dStart
    rstTasks.Fields("Due Date").Value = dDue
    rstTasks.Fields("Category").Value = sTaskCategory
    rstTasks.Fields("Description").Value = sTaskDescription
    rstTasks.Update
Next i


sTaskTitle = "(3.1) Find/replace add to cover page:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
dDue = dDueDate - 2
iTaskMinuteLength = "3"
sPriority = "(3) Stage 3"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Start Date").Value = dStart
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


sTaskTitle = "(3.2) Hyperlink:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
dDue = dDueDate - 2
iTaskMinuteLength = "15"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Start Date").Value = dStart
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


sTaskTitle = "(3.3) Send email if more info needed and hold transcript:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
dDue = dDueDate - 2
iTaskMinuteLength = "2"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Start Date").Value = dStart
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iAudioProofTime = Round((((sAudioLength * 1.5) / 60) + 1), 0)
For i = 1 To iAudioProofTime
sTaskTitle = "(3.4) Audio-proof:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
dDue = dDueDate - 2
iTaskMinuteLength = "60"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Start Date").Value = dStart
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update
Next i


sTaskTitle = "(4.1) Make final transcript docs, pdf, zip, etc:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
dDue = dDueDate - 1
iTaskMinuteLength = "3"
sPriority = "(4) Stage 4"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Start Date").Value = dStart
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


sTaskTitle = "(4.2) Invoice if balance due or factored.  Refund if applicable:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
dDue = dDueDate - 1
iTaskMinuteLength = "1"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Start Date").Value = dStart
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


sTaskTitle = "(4.3) Deliver as necessary electronically if transcript not held:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
dDue = dDueDate - 1
iTaskMinuteLength = "1"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Start Date").Value = dStart
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


sTaskTitle = "(4.4) Send invoice to factoring:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
dDue = dDueDate - 1
iTaskMinuteLength = "1"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Start Date").Value = dStart
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


sTaskTitle = "(4.5) File transcript:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
dDue = dDueDate - 1
iTaskMinuteLength = "3"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Start Date").Value = dStart
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


sTaskTitle = "(4.6) Burn CD for mailing:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
dDue = dDueDate - 1
iTaskMinuteLength = "2"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Start Date").Value = dStart
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


sTaskTitle = "(4.7) Generate xmls for shipping:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
dDue = dDueDate - 1
iTaskMinuteLength = "1"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Start Date").Value = dStart
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


sTaskTitle = "(4.8) Produce & mail transcript:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
dDue = dDueDate - 1
iTaskMinuteLength = "15"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Start Date").Value = dStart
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


sTaskTitle = "(4.9) Add tracking number and shipping cost to DB.  Notify client:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
dDue = dDueDate - 1
iTaskMinuteLength = "2"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Start Date").Value = dStart
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


rstTasks.Close
Set rstTasks = Nothing
Call pfClearGlobals
End Function
Public Function pfDailyTaskAddFunction()
'============================================================================
' Name        : pfDailyTaskAddFunction
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfDailyTaskAddFunction
' Description : adds static daily tasks to Tasks table
'============================================================================

Dim sTaskTitle As String, sTaskCategory As String, sPriority As String, sTaskDescription As String
Dim dDue As Date
Dim rstTasks As DAO.Recordset
Dim iTaskMinuteLength As Integer

Set rstTasks = CurrentDb.OpenRecordset("Tasks")

sTaskCategory = "GTD Daily"
dDue = Now + 1
sPriority = "(1) Stage 1"
sTaskDescription = "none"
iTaskMinuteLength = "2"
sTaskTitle = "List action items, projects, waiting-fors, calendar events, someday/maybes as appropriate"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "2"
sTaskTitle = "replied to all e-mails, checked & processed all voicemails"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "2"
sTaskTitle = "export e-mails and check communication"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "10"
sTaskTitle = "review jobs sent to me"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "10"
sTaskTitle = "review tasks bin and process"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


sTaskCategory = "personal"
iTaskMinuteLength = "60"
sTaskTitle = "art time"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "60"
sTaskTitle = "yoga"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update

rstTasks.Close
Set rstTasks = Nothing

End Function


Public Function pfWeeklyTaskAddFunction()
'============================================================================
' Name        : pfWeeklyTaskAddFunction
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfWeeklyTaskAddFunction
' Description : adds static weekly tasks to Tasks table
'============================================================================

Dim sTaskTitle As String, sTaskCategory As String, sPriority As String, sTaskDescription As String
Dim vStartDate As Date, dDue As Date
Dim rstTasks As DAO.Recordset
Dim iTaskMinuteLength As Integer

Set rstTasks = CurrentDb.OpenRecordset("Tasks")


sTaskCategory = "GTD Weekly"
dDue = Now + 5
sPriority = "(2) Stage 2"
sTaskDescription = "none"
iTaskMinuteLength = "5"
sTaskTitle = "empty head about uncaptured new items"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "5"
sTaskTitle = "file material away"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "5"
sTaskTitle = "stage R/R material"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "10"
sTaskTitle = "update payment bill calendar"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "10"
sTaskTitle = "budget"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "10"
sTaskTitle = "review events coming up"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "10"
sTaskTitle = "review lists"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "10"
sTaskTitle = "review long-term projects"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "10"
sTaskTitle = "review sales reports"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


sTaskCategory = "business admin"
iTaskMinuteLength = "60"
sTaskTitle = "update AQC manual"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "60"
sTaskTitle = "do 1 hour government contracts or business/marketing plan work"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


sTaskCategory = "personal"
iTaskMinuteLength = "20"
sTaskTitle = "vacuum"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "20"
sTaskTitle = "groceries"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "30"
sTaskTitle = "laundry"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "60"
sTaskTitle = "Clean house"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


rstTasks.Close
Set rstTasks = Nothing

End Function

Public Function pfMonthlyTaskAddFunction()
'============================================================================
' Name        : pfMonthlyTaskAddFunction
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfMonthlyTaskAddFunction
' Description : adds static monthly tasks to Tasks table
'============================================================================

Dim sTaskTitle As String, sTaskCategory As String, sPriority As String, sTaskDescription As String
Dim iTaskMinuteLength As Integer
Dim dDue As Date
Dim rstTasks As DAO.Recordset

Set rstTasks = CurrentDb.OpenRecordset("Tasks")


sTaskCategory = "GTD Monthly"
dDue = Now + 20
sPriority = "(1) Stage 1"
sTaskDescription = "none"
iTaskMinuteLength = "15"
sTaskTitle = "Brainstorm Creative Ideas"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "15"
sTaskTitle = "Review 1 to 2 Year Goals"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "15"
sTaskTitle = "Review Roles and Current Responsibilities"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "15"
sTaskTitle = "Review Someday or Maybe list"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


iTaskMinuteLength = "15"
sTaskTitle = "Review Support Files"

rstTasks.AddNew
rstTasks.Fields("Title").Value = sTaskTitle
rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
rstTasks.Fields("Priority").Value = sPriority
rstTasks.Fields("Due Date").Value = dDue
rstTasks.Fields("Category").Value = sTaskCategory
rstTasks.Fields("Description").Value = sTaskDescription
rstTasks.Update


rstTasks.Close
Set rstTasks = Nothing
End Function




Public Function pfCommHistoryExportSub()
On Error Resume Next
'============================================================================
' Name        : pfCommHistoryExportSub
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfCommHistoryExportSub
' Description : exports emails to CommunicationsHistory table
'============================================================================

Dim nsOutlookNmSpc As Outlook.Namespace

Dim oOutlookAccessTestFolder As Object
Dim oOutLookMAPIFolder As Outlook.MAPIFolder
Dim oOutlookMail As Outlook.MailItem
Dim dEmailReceived As Date
Dim sEmailReceivedTime As String, sDriveHyperlink As String, sSenderName As String, sCommHistoryHyperlink As String
Dim rs As DAO.Recordset
Dim FNme As String

Set nsOutlookNmSpc = GetNamespace("MAPI")
Set oOutlookAccessTestFolder = nsOutlookNmSpc.Folders("inquiries@aquoco.co").Folders("Inbox").Folders("AccessTest")

For Each oOutlookMail In oOutlookAccessTestFolder.Items
     
    dEmailReceived = oOutlookMail.ReceivedTime 'assign received time to variable
    sSenderName = oOutlookMail.SenderName
    sEmailReceivedTime = Format(dEmailReceived, "YYYYMMDD-hhmm") 'convert time to good string value
    
    'save email on hard drive in in progress
    oOutlookMail.SaveAs "T:\Database\Emails\" & sEmailReceivedTime & "-" & sSenderName & "-Email.msg", 3
    dEmailReceived = oOutlookMail.ReceivedTime
    sEmailReceivedTime = Format(dEmailReceived, "YYYYMMDD-hhmm")
    sDriveHyperlink = "T:\Database\Emails\" & sEmailReceivedTime & "-" & sSenderName & "-Email.msg"
    
    'string for link to email in access hyperlink field
    sCommHistoryHyperlink = sEmailReceivedTime & "-" & sSenderName & "-Email" & "#" & sDriveHyperlink & "#" '
    
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM CommunicationHistory")
    
    rs.AddNew
    rs("FileHyperlink") = sCommHistoryHyperlink
    rs("FileHyperlink1") = sCommHistoryHyperlink
    rs("DateCreated") = dEmailReceived
    rs("CourtDatesID") = Null
    rs("CustomersID") = Null
    rs.Update
    
    On Error GoTo eHandler
    
Next
If (oOutlookAccessTestFolder.Folders.Count > 0) Then
    For Each oOutLookMAPIFolder In oOutlookAccessTestFolder.Folders
        pfEmailsExport1
    Next
End If

rs.Close

Set rs = Nothing
'Set adoConn = Nothing
Set nsOutlookNmSpc = Nothing
Set oOutlookAccessTestFolder = Nothing

Exit Function

eHandler:

'MsgBox ("The email " & FNme & " failed to save.")
'MsgBox Err.Description & " (" & Err.Number & ")"

Resume Next
End Function
Public Function pfEmailsExport1()
'============================================================================
' Name        : pfEmailsExport1
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfEmailsExport1
' Description : export specified fields from each mail / item in selected folder
'============================================================================

Dim nsOutlookNmSpc As Outlook.Namespace
Dim iItemCounter As Integer
Dim oOutLookMAPIFolder As Outlook.MAPIFolder
Dim rstEmails As DAO.Recordset

Set nsOutlookNmSpc = GetNamespace("MAPI")
'Set oOutLookMAPIFolder = nsOutlookNmSpc.PickFolder
Set oOutLookMAPIFolder = nsOutlookNmSpc.Folders("inquiries@aquoco.co").Folders("Inbox").Folders("AccessTest")
Set rstEmails = CurrentDb.OpenRecordset("SELECT * FROM Emails")

For iItemCounter = oOutLookMAPIFolder.Items.Count To 1 Step -1 'Cycle through selected folder.

    With oOutLookMAPIFolder.Items(iItemCounter) ' Copy property value to corresponding fields in target file.
    
        If .Class = olMail Then
            rstEmails.AddNew
            rstEmails("Subject") = .Subject
            rstEmails("Body") = .Body
            rstEmails("SenderName") = .SenderName
            rstEmails("ToName") = .To
            rstEmails("SenderEmailAddress") = .SenderEmailAddress
            rstEmails("SenderEmailType") = .SenderEmailType
            rstEmails("CCName") = .CC
            rstEmails("BCCName") = .BCC
            rstEmails("Importance") = .Importance
            rstEmails("Sensitivity") = .Sensitivity
            rstEmails.Update
        End If
        
    End With

Next

rstEmails.Close

Set rstEmails = Nothing
Set nsOutlookNmSpc = Nothing
Set oOutLookMAPIFolder = Nothing

End Function


Public Function pfMoveSelectedMessages()
'============================================================================
' Name        : pfMoveSelectedMessages
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfMoveSelectedMessages
' Description : move selected messages to network drive
'============================================================================

Dim oOutlookApp As Outlook.Application
Dim nsOutlookNmSpc As Outlook.Namespace
Dim oDestinationFolder As Outlook.MAPIFolder
Dim oSourceFolder As Outlook.Folder
Dim oCurrentExplorer As Explorer
Dim oSelection As Selection
Dim oSubSelection As Object
Dim vObjectVariant As Variant
Dim lMovedItems As Long
Dim iDateDifference As Integer
Dim sSenderName As String

Set oOutlookApp = Application
Set nsOutlookNmSpc = oOutlookApp.GetNamespace("MAPI")
Set oCurrentExplorer = oOutlookApp.ActiveExplorer
Set oSelection = oCurrentExplorer.oSelection
Set oSourceFolder = oCurrentExplorer.CurrentFolder

For Each oSubSelection In oSelection

    Set vObjectVariant = oSubSelection

    If vObjectVariant.Class = olMail Then
    
       iDateDifference = DateDiff("d", vObjectVariant.SentOn, Now)
         ' using 40 days, adjust if necessary.
       If iDateDifference > 40 Then sSenderName = vObjectVariant.SentOnBehalfOfName
       If sSenderName = ";" Then sSenderName = vObjectVariant.SenderName
       
    End If
    
    On Error Resume Next
    
    ' if the destination folder is not a subfolder of the current folder, use this:
    ' Dim oOutlookInbox  As Outlook.MAPIFolder
    ' Dim sDestinationFolder As String
    ' Set oOutlookInbox  = nsOutlookNmSpc.Folders("alias@domain.com").Folders("Inbox") ' or wherever the folder is
    ' Set oDestinationFolder = oOutlookInbox.Folders(sSenderName)
    
    Set oDestinationFolder = oSourceFolder.Folders(sSenderName)
    
    If oDestinationFolder Is Nothing Then
        Set oDestinationFolder = oSourceFolder.Folders.Add(sSenderName)
    End If
    
    vObjectVariant.Move oDestinationFolder
    'count the # of items moved
    lMovedItems = lMovedItems + 1
    Set oDestinationFolder = Nothing
    
    Err.Clear
    
Next

'Display the number of items that were moved.
MsgBox "Moved " & lMovedItems & " messages(s)."

Set oCurrentExplorer = Nothing
Set oSubSelection = Nothing
Set oSelection = Nothing
Set oOutlookApp = Nothing
Set nsOutlookNmSpc = Nothing
Set oSourceFolder = Nothing
End Function


Public Function pfAskforAudio()
'============================================================================
' Name        : pfAskforAudio
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfAskforAudio
' Description : prompts to select audio to copy to job no.'s audio folder
'============================================================================

Dim fd As FileDialog
Dim iFileChosen As Integer
Dim sFileName As String, sAudioFolder As String, sNewAudioPath As String
Dim i As Integer
Dim fs As Object

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
sAudioFolder = "I:\" & sCourtDatesID & "\Audio\"
sNewAudioPath = sAudioFolder & ""
Set fd = Application.FileDialog(msoFileDialogFilePicker)
'use the standard title and filters, but change the initial folder
fd.InitialFileName = "T:\"
fd.InitialView = msoFileDialogViewList
fd.Title = "Select the audio for this transcript."
fd.AllowMultiSelect = True 'allow multiple file selection

iFileChosen = fd.Show
If iFileChosen = -1 Then
    For i = 1 To fd.SelectedItems.Count 'open each of the files chosen
        Debug.Print "i:  " & i
        Debug.Print fd.InitialFileName
        Debug.Print fd.InitialFileName & "\" & fd.SelectedItems(i)
        Debug.Print fd.SelectedItems(i)
        sFileName = Right$(fd.SelectedItems(i), Len(fd.SelectedItems(i)) - InStrRev(fd.SelectedItems(i), "\"))
        sNewAudioPath = sAudioFolder & sFileName
        Debug.Print sNewAudioPath
        Debug.Print Len(Dir(sNewAudioPath, vbDirectory))
        
        If Len(Dir(sNewAudioPath, vbDirectory)) = 0 Then
           FileCopy fd.SelectedItems(i), sNewAudioPath
        End If
        
    Next i
    
End If

End Function
        


Public Function pfAskforNotes()
'============================================================================
' Name        : pfAskforNotes
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfAskforNotes
' Description : prompts to select notes to copy to job no.'s audio folder
'               attempts to pdf selection
'               copies pdf and original notes selected file to job no. folder
'============================================================================


Dim fd As FileDialog
Dim iFileChosen As Integer
Dim sFileName As String, sAudioFolder As String, sNewAudioPath As String
Dim sNewNotesName As String, sOriginalNotesPath As String, sNotesPath As String
Dim i As Integer
Dim fs As Object
Dim sWorkingCopyPath As String, sTranscriptWD As String, sFinalTranscriptWD As String
Dim sFinalTranscriptNoExt As String, sCourtCoverPath As String
Dim sAnswerPDFPrompt As String, sMakePDFPrompt As String
Dim oWordApp As Word.Application, oWordDoc As Word.Document, oVBComponent As Object
Dim rngStory As Range

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
sAudioFolder = "I:\" & sCourtDatesID & "\Notes\"
sNewAudioPath = sAudioFolder & ""
sNewNotesName = "I:\" & sCourtDatesID & "\Notes\" & sCourtDatesID & "-Notes.pdf"
Set fd = Application.FileDialog(msoFileDialogFilePicker)
'use the standard title and filters, but change the initial folder
fd.InitialFileName = "T:\"
fd.InitialView = msoFileDialogViewList
fd.Title = "Select your Notes."
fd.AllowMultiSelect = True 'allow multiple file selection

iFileChosen = fd.Show
If iFileChosen = -1 Then
    For i = 1 To fd.SelectedItems.Count 'copy each of the files chosen
        sFileName = Right$(fd.SelectedItems(i), Len(fd.SelectedItems(i)) - InStrRev(fd.SelectedItems(i), "\"))
        sNewAudioPath = sAudioFolder & sFileName
        
        Debug.Print sNewNotesName
        Debug.Print sNewAudioPath
        
        If Len(Dir(sNewAudioPath, vbDirectory)) = 0 Then
        
        
        'open in word and save as notes pdf
        
        sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
        sOriginalNotesPath = fd.SelectedItems(i)
        sNotesPath = "I:\" & sCourtDatesID & "\Notes\" & sCourtDatesID & "-Notes.PDF"
        
        sMakePDFPrompt = "Next we will make a PDF copy.  Click yes when ready."
        
        sAnswerPDFPrompt = MsgBox(sMakePDFPrompt, vbQuestion + vbYesNo, "???")
        
        If sAnswerPDFPrompt = vbNo Then 'Code for No
            
            MsgBox "No PDF copy of the notes will be made."
            
        Else 'Code for yes
        
            sOriginalNotesPath = fd.SelectedItems(i)
            
            
            On Error Resume Next
            
            Set oWordApp = GetObject(, "Word.Application")
            
            If Err <> 0 Then
                Set oWordApp = CreateObject("Word.Application")
            End If
            
            Set oWordDoc = GetObject(sOriginalNotesPath, "Word.Document")
            On Error GoTo 0
            
            Set oWordDoc = oWordApp.Documents.Open(FileName:=sOriginalNotesPath)
            
            oWordDoc.Application.Visible = False
            oWordDoc.Application.DisplayAlerts = False
        
            'save as pdf
            oWordDoc.ExportAsFixedFormat outputFileName:=sNotesPath, ExportFormat:=wdExportFormatPDF, CreateBookmarks:=wdExportCreateHeadingBookmarks
                        
            oWordDoc.Close SaveChanges:=False
        
            
        End If
                    
            Set oWordDoc = Nothing
            Set oWordApp = Nothing
            
           FileCopy fd.SelectedItems(i), sNewAudioPath
           FileCopy fd.SelectedItems(i), sNewNotesName
        End If
    Next i
End If
End Function

Public Function pfRCWRuleScraper()
On Error Resume Next
'============================================================================
' Name        : pfRCWRuleScraper
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfRCWRuleScraper()
' Description:  scrapes all RCWs from WA site
'============================================================================

Dim sFindCitation As String, sLongCitation As String, sRuleNumber As String
Dim sWebAddress As String, sReplaceHyperlink As String, sCurrentRule As String
Dim sChapterNumber As String, sSubchapterNumber As String, sSubtitleNumber As String
Dim sSectionNumber As String
Dim Title As String, sTitle2 As String, sTitle1 As String

Dim objHttp As Object
Dim vLettersArray1(), vLettersArray2() As Variant
Dim rstCitationHyperlinks As DAO.Recordset
Dim iErrorNum As Integer, sCHCategory As Integer

Dim i As Long, j As Long, k As Long, l As Long, m As Long
Dim w As Long, x As Long, y As Long, z As Long

For x = 1 To 91 '(RCW first portion x.###.###) '1-91


    For y = 1 To 999 '(RCW second portion ###.y.###) '1-999
    
        
        If y < 10 Then y = Str("0" & y)
                
    
        For z = 10 To 990 Step 10 '(RCW third portion ###.###.z) '10 to 990 by 10s
        
            If z < 100 Then z = Str("0" & z)
            
            'generate variables
            sCurrentRule = x & "." & y & "." & z
            sFindCitation = "RCW " & sCurrentRule
            sLongCitation = "RCW " & sCurrentRule
            sCHCategory = 2
            sRuleNumber = sCurrentRule
            sWebAddress = "https://app.leg.wa.gov/RCW/default.aspx?cite=" & sCurrentRule
            sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
            
            
            Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
            objHttp.Open "GET", sWebAddress, False
            objHttp.send ""
            
            Title = objHttp.responseText
            
            If InStr(1, UCase(Title), "<title></title>") Then
                Debug.Print ("RCW " & sCurrentRule & "is a bad RCW; moving on to try next one.")
                GoTo NextNumber1
            
            Else
            
                'add entry to citationhyperlinks
                Debug.Print ("Entering " & sFindCitation & " into CitationHyperlinks table.")
                
                'add new entry to citaitonhyperlinks table
                Set rstCitationHyperlinks = CurrentDb.OpenRecordset("CitationHyperlinks")
                rstCitationHyperlinks.AddNew
                rstCitationHyperlinks.Fields("FindCitation").Value = sFindCitation
                rstCitationHyperlinks.Fields("ReplaceHyperlink").Value = sReplaceHyperlink
                rstCitationHyperlinks.Fields("LongCitation").Value = sLongCitation
                rstCitationHyperlinks.Fields("WebAddress").Value = sWebAddress
                rstCitationHyperlinks.Fields("CHCategory").Value = sCHCategory
                rstCitationHyperlinks.Update
                
                
                
            End If


            Set objHttp = Nothing
NextNumber1:
        Next
    Next
Next


vLettersArray1 = Array("9A", "23B", "28A", "28B", "28C", "29A", "30A", "30B35A", "50A", "62A", "71A", "79A")

For w = 1 To UBound(vLettersArray1) '(RCW first portion w.###.###)

    sTitle1 = vLettersArray1(w)
    
    For x = 1 To 999 '(RCW second portion ###.x.###)
    
        If x < 10 Then x = Str("0" & x)
    
        '1-999 plus A, B, C
    
        vLettersArray2 = Array("A", "B", "C")
        
        For y = 0 To UBound(vLettersArray2) '(RCW second portion w.x[y].z)
            
            sTitle2 = x & vLettersArray2(y)
        
            If y < 10 Then sTitle2 = Str("0" & sTitle2)
    
            For z = 10 To 990 Step 10 '(RCW third portion ###.###.z)
            
                If z < 100 Then y = Str("0" & z)
                
                'generate variables
                sCurrentRule = sTitle1 & "." & "." & sTitle2 & "." & z
                sFindCitation = "RCW " & sCurrentRule
                sLongCitation = "RCW " & sCurrentRule
                sCHCategory = 2
                sRuleNumber = sCurrentRule
                sWebAddress = "https://app.leg.wa.gov/RCW/default.aspx?cite=" & sCurrentRule
                sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
                
                
                Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
                objHttp.Open "GET", sWebAddress, False
                objHttp.send ""
                
                Title = objHttp.responseText
                
                If InStr(1, UCase(Title), "<title></title>") Then
                    Debug.Print ("Bad website, moving on to try next one.")
                    GoTo NextNumber2
                
                Else
                
                    'add entry to citationhyperlinks
                    
                    Debug.Print ("Entering " & sFindCitation & " into CitationHyperlinks table.")
                    
                    'add new entry to citaitonhyperlinks table
                    Set rstCitationHyperlinks = CurrentDb.OpenRecordset("CitationHyperlinks")
                    rstCitationHyperlinks.AddNew
                    rstCitationHyperlinks.Fields("FindCitation").Value = sFindCitation
                    rstCitationHyperlinks.Fields("ReplaceHyperlink").Value = sReplaceHyperlink
                    rstCitationHyperlinks.Fields("LongCitation").Value = sLongCitation
                    rstCitationHyperlinks.Fields("WebAddress").Value = sWebAddress
                    rstCitationHyperlinks.Fields("CHCategory").Value = sCHCategory
                    rstCitationHyperlinks.Update
                    
                    
                    
                End If

    
                Set objHttp = Nothing
NextNumber2:
            Next

        Next
    Next
Next



End Function



Public Function pfUSCRuleScraper()
On Error Resume Next
'============================================================================
' Name        : pfRuleScraper
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfRuleScraper()
' Description:  gets non-section usc code from site
'============================================================================

Dim rstCitationHyperlinks As DAO.Recordset
Dim iErrorNum As Integer, sCHCategory As Integer
Dim sFindCitation As String, sLongCitation As String, sRuleNumber As String
Dim sWebAddress As String, sReplaceHyperlink As String, sCurrentRule As String
Dim vRuleNumbers() As Variant, vRules() As Variant
Dim i As Long, j As Long, k As Long, l As Long, m As Long
Dim w As Long, x As Long, y As Long, z As Long

vRules = Array("CR ", "CrR ", "RAP ", "Rule ", "RCW ", "ER ")
vRuleNumbers = Array("", "", "")


For i = 1 To 54
'Title 1-54
'http://uscode.house.gov/view.xhtml?path=/prelim@title8&edition=prelim

    For l = 1 To 3100
    'Chapter 1-3100
    'http://uscode.house.gov/view.xhtml?path=/prelim@title11/chapter1&edition=prelim
    'http://uscode.house.gov/view.xhtml?path=/prelim@title11/chapter3&edition=prelim
    
        'generate variables
        sCurrentRule = i & " U.S.C. Chapter  " & l
        sFindCitation = sCurrentRule
        sLongCitation = sCurrentRule
        sCHCategory = 2
        sRuleNumber = sCurrentRule
        sWebAddress = "https://uscode.house.gov/view.xhtml?path=/prelim@title" & i & "/chapter" & l & "&edition=prelim"
        sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
        
        'add to citationhyperlinks table
        On Error Resume Next
        iErrorNum = Err    'Save error number
        On Error GoTo 0
        If l = 850 Or l = 2000 Then
            Debug.Print ("Bad code citation for " & sCurrentRule & ", moving on to try next one.")
        ElseIf iErrorNum = 0 Then
        
            Debug.Print ("Entering " & sCurrentRule & " into CitationHyperlinks table.")
            
            'add new entry to citaitonhyperlinks table
            Set rstCitationHyperlinks = CurrentDb.OpenRecordset("CitationHyperlinks")
            rstCitationHyperlinks.AddNew
            rstCitationHyperlinks.Fields("FindCitation").Value = sFindCitation
            rstCitationHyperlinks.Fields("ReplaceHyperlink").Value = sReplaceHyperlink
            rstCitationHyperlinks.Fields("LongCitation").Value = sLongCitation
            rstCitationHyperlinks.Fields("WebAddress").Value = sWebAddress
            rstCitationHyperlinks.Fields("CHCategory").Value = sCHCategory
            rstCitationHyperlinks.Update
    
        Else
        End If

        For j = 1 To 3100
        'subchapter 1-3100
        'http://uscode.house.gov/view.xhtml?path=/prelim@title11/chapter3/subchapter1&edition=prelim
        
            'generate variables
            sCurrentRule = i & " U.S.C. Chapter  " & l & ", Subchapter " & j
            sFindCitation = sCurrentRule
            sLongCitation = sCurrentRule
            sCHCategory = 2
            sRuleNumber = sCurrentRule
            sWebAddress = "https://uscode.house.gov/view.xhtml?path=/prelim@title" & i & "/chapter" & l & "/subchapter" & j & "&edition=prelim"
            sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
            
            'add to citationhyperlinks table
            On Error Resume Next
            iErrorNum = Err    'Save error number
            On Error GoTo 0
            If j = 850 Or j = 2000 Then
                Debug.Print ("Bad code citation for " & sCurrentRule & ", moving on to try next one.")
            ElseIf iErrorNum = 0 Then
            
                Debug.Print ("Entering " & sCurrentRule & " into CitationHyperlinks table.")
                
                'add new entry to citaitonhyperlinks table
                Set rstCitationHyperlinks = CurrentDb.OpenRecordset("CitationHyperlinks")
                rstCitationHyperlinks.AddNew
                rstCitationHyperlinks.Fields("FindCitation").Value = sFindCitation
                rstCitationHyperlinks.Fields("ReplaceHyperlink").Value = sReplaceHyperlink
                rstCitationHyperlinks.Fields("LongCitation").Value = sLongCitation
                rstCitationHyperlinks.Fields("WebAddress").Value = sWebAddress
                rstCitationHyperlinks.Fields("CHCategory").Value = sCHCategory
                rstCitationHyperlinks.Update
        
            Else
            End If

        Next
    
    Next

    vRuleNumbers = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    For k = 0 To UBound(vRuleNumbers)
    'subtitle A-Z
    'http://uscode.house.gov/view.xhtml?path=/prelim@title26/subtitleG&edition=prelim
    
        'generate variables
        sCurrentRule = i & " U.S.C. Subtitle  " & vRuleNumbers(k)
        sFindCitation = sCurrentRule
        sLongCitation = sCurrentRule
        sCHCategory = 2
        sRuleNumber = sCurrentRule
        sWebAddress = "https://uscode.house.gov/view.xhtml?path=/prelim@title" & i & "/subtitle" & vRuleNumbers(k) & "&edition=prelim"
        sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
        
        'add to citationhyperlinks table
        On Error Resume Next
        iErrorNum = Err    'Save error number
        On Error GoTo 0
        If k = 850 Or k = 2000 Then
            Debug.Print ("Bad code citation for " & sCurrentRule & ", moving on to try next one.")
        ElseIf iErrorNum = 0 Then
        
            Debug.Print ("Entering " & sCurrentRule & " into CitationHyperlinks table.")
            
            'add new entry to citaitonhyperlinks table
            Set rstCitationHyperlinks = CurrentDb.OpenRecordset("CitationHyperlinks")
            rstCitationHyperlinks.AddNew
            rstCitationHyperlinks.Fields("FindCitation").Value = sFindCitation
            rstCitationHyperlinks.Fields("ReplaceHyperlink").Value = sReplaceHyperlink
            rstCitationHyperlinks.Fields("LongCitation").Value = sLongCitation
            rstCitationHyperlinks.Fields("WebAddress").Value = sWebAddress
            rstCitationHyperlinks.Fields("CHCategory").Value = sCHCategory
            rstCitationHyperlinks.Update
    
        Else
        End If

    Next

    For x = 1 To 3100
    'subtitle 1-3100
    'http://uscode.house.gov/view.xhtml?path=/prelim@title51/subtitle1&edition=prelim

    
        'generate variables
        sCurrentRule = i & " U.S.C. Subtitle  " & x
        sFindCitation = sCurrentRule
        sLongCitation = sCurrentRule
        sCHCategory = 2
        sRuleNumber = sCurrentRule
        sWebAddress = "https://uscode.house.gov/view.xhtml?path=/prelim@title" & i & "/subtitle" & x & "&edition=prelim"
        sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
        
        'add to citationhyperlinks table
        On Error Resume Next
        iErrorNum = Err    'Save error number
        On Error GoTo 0
        If l = 850 Or l = 2000 Then
            Debug.Print ("Bad code citation for " & sCurrentRule & ", moving on to try next one.")
        ElseIf iErrorNum = 0 Then
        
            Debug.Print ("Entering " & sCurrentRule & " into CitationHyperlinks table.")
            
            'add new entry to citaitonhyperlinks table
            Set rstCitationHyperlinks = CurrentDb.OpenRecordset("CitationHyperlinks")
            rstCitationHyperlinks.AddNew
            rstCitationHyperlinks.Fields("FindCitation").Value = sFindCitation
            rstCitationHyperlinks.Fields("ReplaceHyperlink").Value = sReplaceHyperlink
            rstCitationHyperlinks.Fields("LongCitation").Value = sLongCitation
            rstCitationHyperlinks.Fields("WebAddress").Value = sWebAddress
            rstCitationHyperlinks.Fields("CHCategory").Value = sCHCategory
            rstCitationHyperlinks.Update
    
        Else
        End If

    Next
Debug.Print ("Title " & i & " completed.  Next, Title " & i + 1)
Next

End Function

Function GetLevel(strItem As String) As Integer
    ' Return the heading level of a header from the
    ' array returned by Word.

    ' The number of leading spaces indicates the
    ' outline level (2 spaces per level: H1 has
    ' 0 spaces, H2 has 2 spaces, H3 has 4 spaces.

    Dim strTemp As String
    Dim strOriginal As String
    Dim intDiff As Integer

    ' Get rid of all trailing spaces.
    strOriginal = RTrim$(strItem)

    ' Trim leading spaces, and then compare with
    ' the original.
    strTemp = LTrim$(strOriginal)

    ' Subtract to find the number of
    ' leading spaces in the original string.
    intDiff = Len(strOriginal) - Len(strTemp)
    GetLevel = (intDiff / 2) + 1
End Function

Function fWLGenerateJSONInfo()
'============================================================================
' Name        : fWLGenerateJSONInfo
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fWLGenerateJSONInfo
' Description : get info for WL API
'============================================================================
'{
'  "list_id": 12345, Ingram Household = 370524335
                    'inbox = 370231796
                    '1ToBeEntered = 388499976
                    '2InProgress = 388499848
                    '3Complete = 388499951
  
'  "title": "Hallo",
'  "assignee_id": 123,
'  "completed": true,
'  "due_date": "2013-08-30",
'  "starred": false
'}

'GET a.wunderlist.com/api/v1/lists/:id
'PATCH a.wunderlist.com/api/v1/lists/:id

Dim rstTRQPlusCases As DAO.Recordset
Dim db As Database
Dim qdf As QueryDef

sWLListID = 370524335 'ingram household
'or try inbox 370231796
lAssigneeID = 88345676 'erica / 86846933 adam
bCompleted = "false"
bStarred = "false"


End Function
Function fWunderlistAdd(sTitle As String, sDueDate As String)
'============================================================================
' Name        : fWunderlistAdd
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fWunderlistAdd()
' Description : adds task to Wunderlist
'============================================================================

Dim sURL As String, sUserName As String, sPassword As String, sAuth As String, stringJSON As String, sEmail As String
Dim vInvoiceID As String, apiWaxLRS As String, vErrorIssue As String, sInvoiceTime As String
Dim oRequest As Object, Json As Object, oWebBrowser As Object
Dim vStatus As String, vTotal As String, sURL1 As String
Dim sURL2 As String
Dim rstRates As DAO.Recordset

Dim resp, response, rep, vDetails As Object
Dim sToken As String, json1 As String, json2 As String, json3 As String, json4 As String, json5 As String
Dim Parsed As Dictionary
Dim vErrorName As String, vErrorMessage As String, vErrorILink As String, vErrorDetails As String
Dim sFile1 As String, sFile2 As String, sText As String, sLine1 As String, sLine2 As String
Dim sLocal As String, sResponseText As String
'{
'  "list_id": 12345,
'  "title": "Hallo",
'  "assignee_id": 123,
'  "completed": true,
'  "due_date": "2013-08-30",
'  "starred": false
'}

Call fWLGenerateJSONInfo

Dim sFile3 As String
sFile1 = "C:\other\3.txt"
sFile2 = "C:\other\4.txt"
sFile3 = "C:\other\5.txt"

Open sFile1 For Input As #1
Line Input #1, sLine1
Close #1

Open sFile2 For Input As #2
Line Input #2, sLine2
Close #2

Dim sLine3 As String
Open sFile3 For Input As #3
Line Input #3, sLine3
Close #3


sEmail = "inquiries@aquoco.co"
sUserName = sLine1
sPassword = sLine2
sToken = sLine3

'{
'  "list_id": 12345,
'  "title": "Hallo",
'  "assignee_id": 123,
'  "completed": true,
'  "due_date": "2013-08-30",
'  "starred": false
'}
      'Public lAssigneeID As Long, sDueDate As String, bStarred As Boolean, bCompleted As Boolean, sTitle As String, sWLListID As String

    json1 = "{" & Chr(34) & _
        "list_id" & Chr(34) & ": " & sWLListID & "," & Chr(34) & _
        "title" & Chr(34) & ": " & Chr(34) & sTitle & Chr(34) & "," & Chr(34) & _
        "assignee_id" & Chr(34) & ": " & lAssigneeID & "," & Chr(34) & _
        "completed" & Chr(34) & ": " & bCompleted & "," & Chr(34) & _
        "due_date" & Chr(34) & ": " & Chr(34) & sDueDate & Chr(34) & "," & Chr(34) & _
        "starred" & Chr(34) & ": " & bStarred & _
        "}"
Debug.Print "JSON1--------------------------------------------"
Debug.Print json1

Debug.Print "RESPONSETEXT--------------------------------------------"
    sURL = "https://a.wunderlist.com/api/v1/tasks" '?completed=False" & bCompleted  '?list_id=" & sWLListID & '"&?title=" & sTitle &
        '"&?assignee_id=" & lAssigneeID & "&?completed=" & bCompleted & "&?due_date=" & sDueDate
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        '.Visible = True
        .Open "POST", sURL, False
        .setRequestHeader "X-Access-Token", sToken
        .setRequestHeader "X-Client-ID", sUserName
        .setRequestHeader "Content-Type", "application/json"
        .send json1
        apiWaxLRS = .responseText
        sToken = ""
        Debug.Print apiWaxLRS
        Debug.Print "--------------------------------------------"
        Debug.Print "Status:  " & .Status
        Debug.Print "--------------------------------------------"
        Debug.Print "StatusText:  " & .StatusText
        Debug.Print "--------------------------------------------"
        Debug.Print "ResponseBody:  " & .responseBody
        Debug.Print "--------------------------------------------"
        .abort
    End With
'Next
Debug.Print "--------------------------------------------"
Debug.Print "Error Name:  " & vErrorName
Debug.Print "Error Message:  " & vErrorMessage
Debug.Print "Error Info Link:  " & vErrorILink
'Debug.Print "Error Field:  " & vErrorIssue
'Debug.Print "Error Details:  " & vErrorDetails
Debug.Print "--------------------------------------------"
Debug.Print "Task Title:  " & sTitle & "   |   List ID:  " & sWLListID & " " & lAssigneeID
Debug.Print "Completed:  " & bCompleted & "   |   Due Date:  " & sDueDate
Debug.Print "--------------------------------------------"
'lAssigneeID As Long, sDueDate As String, bStarred As Boolean, bCompleted As Boolean, sTitle As String, sWLListID As String

End Function



Function fWunderlistGetTasksOnList()
'============================================================================
' Name        : fWunderlistGetTasksOnList
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fWunderlistGetTasksOnList()
' Description : gets tasks on Wunderlist list
'============================================================================

Dim sURL As String, sUserName As String, sPassword As String, sAuth As String, stringJSON As String, sEmail As String
Dim vInvoiceID As String, apiWaxLRS As String, vErrorIssue As String, sInvoiceTime As String
Dim oRequest As Object, Json As Object, oWebBrowser As Object
Dim vStatus As String, vTotal As String, sURL1 As String
Dim sURL2 As String
Dim rstRates As DAO.Recordset

Dim resp, response, rep, vDetails As Object
Dim sToken As String, json1 As String, json2 As String, json3 As String, json4 As String, json5 As String
Dim Parsed As Dictionary
Dim vErrorName As String, vErrorMessage As String, vErrorILink As String, vErrorDetails As String
Dim sFile1 As String, sFile2 As String, sText As String, sLine1 As String, sLine2 As String
Dim sLocal As String, sResponseText As String
'{
'  "list_id": 12345,
'  "title": "Hallo",
'  "assignee_id": 123,
'  "completed": true,
'  "due_date": "2013-08-30",
'  "starred": false
'}

Call fWLGenerateJSONInfo

Dim sFile3 As String
sFile1 = "C:\other\3.txt"
sFile2 = "C:\other\4.txt"
sFile3 = "C:\other\5.txt"

Open sFile1 For Input As #1
Line Input #1, sLine1
Close #1

Open sFile2 For Input As #2
Line Input #2, sLine2
Close #2

Dim sLine3 As String
Open sFile3 For Input As #3
Line Input #3, sLine3
Close #3

sEmail = "inquiries@aquoco.co"
sUserName = sLine1
sPassword = sLine2
sToken = sLine3
'{
'  "list_id": 12345
'}
'Public lAssigneeID As Long, sDueDate As String, bStarred As Boolean, bCompleted As Boolean, sTitle As String, sWLListID As String
    
json1 = "{" & Chr(34) & "list_id" & Chr(34) & ": " & sWLListID & "}"

Debug.Print "JSON1--------------------------------------------"
Debug.Print json1
Debug.Print "RESPONSETEXT--------------------------------------------"
    sURL = "https://a.wunderlist.com/api/v1/tasks?list_id=" & sWLListID
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        '.Visible = True
        .Open "GET", sURL, False
        .setRequestHeader "X-Access-Token", sToken
        .setRequestHeader "X-Client-ID", sUserName
        .setRequestHeader "Content-Type", "application/json"
        .send json1
        apiWaxLRS = .responseText
        sToken = ""
        Debug.Print apiWaxLRS
        Debug.Print "--------------------------------------------"
        Debug.Print .Status
        Debug.Print .StatusText
        .abort
    End With
    
apiWaxLRS = Left(apiWaxLRS, Len(apiWaxLRS) - 1)
apiWaxLRS = Right(apiWaxLRS, Len(apiWaxLRS) - 1)
apiWaxLRS = "{" & Chr(34) & "List" & Chr(34) & ":" & apiWaxLRS & "}"
'"total_amount":{"currency":"USD","value":"3.00"},
Set Parsed = JsonConverter.ParseJson(apiWaxLRS)
'sInvoiceNumber = Parsed("number") 'third level array
'vInvoiceID = Parsed("id") 'third level array
'vStatus = Parsed("status") 'third level array
'vTotal = Parsed("total_amount")("value") 'second level array
'vErrorName = Parsed("id") '("value") 'second level array
'vErrorMessage = Parsed("due_date") '("value") 'second level array
'vErrorILink = Parsed("links") '("value") 'second level array
Set vDetails = Parsed("list") 'second level array
For Each rep In vDetails ' third level objects
    vErrorIssue = rep("id")
    vErrorDetails = rep("due_date")
    Debug.Print "--------------------------------------------"
    Debug.Print "Error Name:  " & vErrorName
    Debug.Print "Error Message:  " & vErrorMessage
    'Debug.Print "Error Info Link:  " & vErrorILink
    'Debug.Print "Error Field:  " & vErrorIssue
    'Debug.Print "Error Details:  " & vErrorDetails
    Debug.Print "--------------------------------------------"
Next
Debug.Print "Task Title:  " & sTitle & "   |   List ID:  " & sWLListID & " " & lAssigneeID
Debug.Print "Completed:  " & bCompleted & "   |   Due Date:  " & sDueDate
Debug.Print "--------------------------------------------"
'lAssigneeID As Long, sDueDate As String, bStarred As Boolean, bCompleted As Boolean, sTitle As String, sWLListID As String

End Function



Function fWunderlistGetLists()
'============================================================================
' Name        : fWunderlistGetLists
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fWunderlistGetLists()
' Description : gets all Wunderlist lists
'============================================================================

Dim sURL As String, sUserName As String, sPassword As String, sAuth As String, stringJSON As String, sEmail As String
Dim vInvoiceID As String, apiWaxLRS As String, vErrorIssue As String, sInvoiceTime As String
Dim oRequest As Object, Json As Object, oWebBrowser As Object
Dim vStatus As String, vTotal As String, sURL1 As String
Dim sURL2 As String
Dim rstRates As DAO.Recordset

Dim resp, response, rep, vDetails As Object
Dim sToken As String, json1 As String, json2 As String, json3 As String, json4 As String, json5 As String
Dim Parsed As Dictionary
Dim vErrorName As String, vErrorMessage As String, vErrorILink As String, vErrorDetails As String
Dim sFile1 As String, sFile2 As String, sText As String, sLine1 As String, sLine2 As String
Dim sLocal As String, sResponseText As String
'{
'  "list_id": 12345,
'  "title": "Hallo",
'  "assignee_id": 123,
'  "completed": true,
'  "due_date": "2013-08-30",
'  "starred": false
'}

Call fWLGenerateJSONInfo

Dim sFile3 As String
sFile1 = "C:\other\3.txt"
sFile2 = "C:\other\4.txt"
sFile3 = "C:\other\5.txt"

Open sFile1 For Input As #1
Line Input #1, sLine1
Close #1

Open sFile2 For Input As #2
Line Input #2, sLine2
Close #2

Dim sLine3 As String
Open sFile3 For Input As #3
Line Input #3, sLine3
Close #3
'https://www.wunderlist.com/oauth/authorize?client_id=ID&redirect_uri=URL&state=RANDOM

sLocal = "'urn:ietf:wg:oauth:2.0:oob','oob'" '"https://localhost/"

sEmail = "inquiries@aquoco.co"
sUserName = sLine1
sPassword = sLine2
sToken = sLine3

sURL2 = "https://a.wunderlist.com/api/v1/user"
'{
'  "list_id": 12345,
'  "title": "Hallo",
'  "assignee_id": 123,
'  "completed": true,
'  "due_date": "2013-08-30",
'  "starred": false
'}
      'Public lAssigneeID As Long, sDueDate As String, bStarred As Boolean, bCompleted As Boolean, sTitle As String, sWLListID As String

    json1 = "{" & Chr(34) & "list_id" & Chr(34) & ": " & sWLListID & "}"
Debug.Print "JSON1--------------------------------------------"
Debug.Print json1
'Debug.Print "JSON2--------------------------------------------"
'Debug.Print json2
'Debug.Print "JSON3--------------------------------------------"
'Debug.Print json3
'Debug.Print "JSON4--------------------------------------------"
'Debug.Print json4
Debug.Print "RESPONSETEXT--------------------------------------------"
    sURL = "https://a.wunderlist.com/api/v1/lists"
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        '.Visible = True
        .Open "GET", sURL, False
        .setRequestHeader "X-Access-Token", sToken
        .setRequestHeader "X-Client-ID", sUserName
        .setRequestHeader "Content-Type", "application/json"
        json5 = json1
        .send json5
        apiWaxLRS = .responseText
        sToken = ""
        .abort
        Debug.Print apiWaxLRS
Debug.Print "--------------------------------------------"
    End With
Set Parsed = JsonConverter.ParseJson(apiWaxLRS)
vErrorName = Parsed("name") '("value") 'second level array
vErrorMessage = Parsed("message") '("value") 'second level array
vErrorILink = Parsed("links") '("value") 'second level array
'
'Set vDetails = Parsed("details") 'second level array
'For Each rep In vDetails ' third level objects
'    vErrorIssue = rep("field")
'    vErrorDetails = rep("issue")
'Next
Debug.Print "--------------------------------------------"
Debug.Print "Error Name:  " & vErrorName
Debug.Print "Error Message:  " & vErrorMessage
Debug.Print "Error Info Link:  " & vErrorILink
Debug.Print "--------------------------------------------"
Debug.Print "Task Title:  " & sTitle & "   |   List ID:  " & sWLListID & " " & lAssigneeID
Debug.Print "Completed:  " & bCompleted & "   |   Due Date:  " & sDueDate
Debug.Print "--------------------------------------------"
'lAssigneeID As Long, sDueDate As String, bStarred As Boolean, bCompleted As Boolean, sTitle As String, sWLListID As String

End Function


Function fWunderlistGetFolders()
'============================================================================
' Name        : fWunderlistGetFolders
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fWunderlistGetFolders()
' Description : gets list of Wunderlist folders or folder revisions
'============================================================================

Dim sURL As String, sUserName As String, sResponseText As String, json1 As String
Dim sPassword As String
Dim apiWaxLRS As String, vErrorIssue As String, sEmail As String, sToken As String
Dim Parsed As Dictionary
Dim vErrorName As String, vErrorMessage As String, vErrorILink As String, vErrorDetails As String
Dim sFile1 As String, sFile2 As String, sFile3 As String
Dim sText As String, sLine1 As String, sLine2 As String, sLine3 As String


Call fWLGenerateJSONInfo

sFile1 = "C:\other\3.txt"
sFile2 = "C:\other\4.txt"
sFile3 = "C:\other\5.txt"

Open sFile1 For Input As #1
Line Input #1, sLine1
Close #1

Open sFile2 For Input As #2
Line Input #2, sLine2
Close #2

Open sFile3 For Input As #3
Line Input #3, sLine3
Close #3

sEmail = "inquiries@aquoco.co"
sUserName = sLine1
sPassword = sLine2
sToken = sLine3

'gets list of folders or folder revisions

'GET a.wunderlist.com/api/v1/folders to get list of all folders

'Response:  Status: 200

'[
'  {
'    "id": 83526310,
'    "title": "Personal Stuff",
'    "list_ids": [1, 2, 3, 4],
'    "created_at": "2013-08-30T08:29:46.203Z",
'    "created_by_request_id": "5541b1d86e925e2dd7e5",
'    "updated_at": "2013-08-30T08:29:46.203Z",
'    "type": "folder",
'    "revision": 10
'  },
'  ...
']

Debug.Print "RESPONSETEXT--------------------------------------------"

'sURL = "https://a.wunderlist.com/api/v1/folders"
sURL = "https://a.wunderlist.com/api/v1/folder_revisions"
With CreateObject("WinHttp.WinHttpRequest.5.1")
    .Open "GET", sURL, False
    .setRequestHeader "X-Access-Token", sToken
    .setRequestHeader "X-Client-ID", sUserName
    .setRequestHeader "Content-Type", "application/json"
    .send
    apiWaxLRS = .responseText
    sToken = ""
    .abort
    Debug.Print apiWaxLRS
    Debug.Print "--------------------------------------------"
End With
    apiWaxLRS = Left(apiWaxLRS, Len(apiWaxLRS) - 1)
    apiWaxLRS = Right(apiWaxLRS, Len(apiWaxLRS) - 1)
Set Parsed = JsonConverter.ParseJson(apiWaxLRS)
vErrorName = Parsed("id") '("value") 'second level array
vErrorMessage = Parsed("title") '("value") 'second level array
vErrorILink = Parsed("list_ids") '("value") 'second level array
vErrorIssue = Parsed("revision") '("value") 'second level array

Debug.Print "--------------------------------------------"
Debug.Print "Folder ID:  " & vErrorName & "   |   " & "Folder Title:  " & vErrorMessage
Debug.Print "List IDs in Folder:  " & vErrorILink & "   |   " & "Revision No.:  " & vErrorIssue
Debug.Print "--------------------------------------------"

End Function


Public Function pfUSCRuleScraper1()
On Error Resume Next

'============================================================================
' Name        : pfUSCRuleScraper
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfUSCRuleScraper()
' Description:  validates and builds links for all us code, no front matter, no appendices
'============================================================================
Dim rstCitationHyperlinks As DAO.Recordset
Dim iErrorNum As Integer, sCHCategory As Integer
Dim sFindCitation As String, sLongCitation As String, sRuleNumber As String
Dim sWebAddress As String, sReplaceHyperlink As String, sCurrentRule As String
Dim sChapterNumber As String, sSubchapterNumber As String, sSubtitleNumber As String
Dim sSectionNumber As String
Dim Title As String
Dim objHttp As Object

Dim vRuleNumbers() As Variant, vRules() As Variant
Dim i As Long, j As Long, k As Long, l As Long, m As Long
Dim w As Long, x As Long, y As Long, z As Long
'vRules = Array("CR ", "CrR ", "RAP ", "Rule ", "RCW ", "ER ")
'vRuleNumbers = Array("", "", "")

'front matter
'http://uscode.house.gov/view.xhtml?req=granuleid:USC-prelim-title11-chapter3-front&num=0&edition=prelim


For x = 1 To 54
            
    'build title links
    
    'Title 1-54
    'http://uscode.house.gov/view.xhtml?path=/prelim@title8&edition=prelim
        
    'generate variables
    sCurrentRule = x
    sFindCitation = "Title " & sCurrentRule
    sLongCitation = "Title " & sCurrentRule
    sCHCategory = 2
    sRuleNumber = sCurrentRule
    sWebAddress = "http://uscode.house.gov/view.xhtml?path=/prelim@title" & sRuleNumber & "&edition=prelim"
    sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
    
    
    Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
    objHttp.Open "GET", sWebAddress, False
    objHttp.send ""
    
    Title = objHttp.responseText
    
    If InStr(1, UCase(Title), "<TITLE>Document Not Found") Then
        Debug.Print ("Bad website, moving on to try next one.")
        GoTo NextNumber
    
    Else
        'add entry to citationhyperlinks
        
        Debug.Print ("Entering " & sFindCitation & " into CitationHyperlinks table.")
        
        'add new entry to citaitonhyperlinks table
        Set rstCitationHyperlinks = CurrentDb.OpenRecordset("CitationHyperlinks")
        rstCitationHyperlinks.AddNew
        rstCitationHyperlinks.Fields("FindCitation").Value = sFindCitation
        rstCitationHyperlinks.Fields("ReplaceHyperlink").Value = sReplaceHyperlink
        rstCitationHyperlinks.Fields("LongCitation").Value = sLongCitation
        rstCitationHyperlinks.Fields("WebAddress").Value = sWebAddress
        rstCitationHyperlinks.Fields("CHCategory").Value = sCHCategory
        rstCitationHyperlinks.Update
        
        
        
    End If

    
            Set objHttp = Nothing
NextNumber:
    
    For y = 1 To 300
        'build related chapter links
                
        'Chapter 1-300
        'http://uscode.house.gov/view.xhtml?path=/prelim@title11/chapter1&edition=prelim
        
        'http://uscode.house.gov/view.xhtml?path=/prelim@title11/chapter3&edition=prelim
        
        
        'generate variables
        sChapterNumber = y
        sFindCitation = "Chapter " & sChapterNumber
        sLongCitation = "Chapter " & sChapterNumber
        sCHCategory = 2
        sRuleNumber = sCurrentRule
        sWebAddress = "http://uscode.house.gov/view.xhtml?path=/prelim@title" & sRuleNumber & "/chapter" & sChapterNumber & "&edition=prelim"
        sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
    
        
        Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
        objHttp.Open "GET", sWebAddress, False
        objHttp.send ""
        
        Title = objHttp.responseText
        
        If InStr(1, UCase(Title), "<TITLE>Document Not Found") Then
            Debug.Print ("Bad website, moving on to try next one.")
            GoTo NextNumber1
        
        Else
            'add entry to citationhyperlinks
            
            Debug.Print ("Entering " & sFindCitation & " into CitationHyperlinks table.")
            
            'add new entry to citaitonhyperlinks table
            Set rstCitationHyperlinks = CurrentDb.OpenRecordset("CitationHyperlinks")
            rstCitationHyperlinks.AddNew
            rstCitationHyperlinks.Fields("FindCitation").Value = sFindCitation
            rstCitationHyperlinks.Fields("ReplaceHyperlink").Value = sReplaceHyperlink
            rstCitationHyperlinks.Fields("LongCitation").Value = sLongCitation
            rstCitationHyperlinks.Fields("WebAddress").Value = sWebAddress
            rstCitationHyperlinks.Fields("CHCategory").Value = sCHCategory
            rstCitationHyperlinks.Update
            
            
            
        End If

    
            Set objHttp = Nothing
NextNumber1:
            
            
        For z = 1 To 999
            'build related subchapter links
                    
        
            'subchapter
            'http://uscode.house.gov/view.xhtml?path=/prelim@title11/chapter3/subchapter1&edition=prelim
                
                
            'generate variables
            sSubchapterNumber = z
            sFindCitation = "Subchapter " & sSubchapterNumber
            sLongCitation = "Subchapter " & sSubchapterNumber
            sCHCategory = 2
            sRuleNumber = sCurrentRule
            sWebAddress = "http://uscode.house.gov/view.xhtml?path=/prelim@title" & sRuleNumber & "/chapter" & sChapterNumber & "/subchapter" & sSubchapterNumber & "&edition=prelim"
            sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
        
            
            Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
            objHttp.Open "GET", sWebAddress, False
            objHttp.send ""
            
            Title = objHttp.responseText
            
            If InStr(1, UCase(Title), "<TITLE>Document Not Found") Then
                Debug.Print ("Bad website, moving on to try next one.")
                GoTo NextNumber2
            
            Else
                'add entry to citationhyperlinks
                
                Debug.Print ("Entering " & sFindCitation & " into CitationHyperlinks table.")
                
                'add new entry to citaitonhyperlinks table
                Set rstCitationHyperlinks = CurrentDb.OpenRecordset("CitationHyperlinks")
                rstCitationHyperlinks.AddNew
                rstCitationHyperlinks.Fields("FindCitation").Value = sFindCitation
                rstCitationHyperlinks.Fields("ReplaceHyperlink").Value = sReplaceHyperlink
                rstCitationHyperlinks.Fields("LongCitation").Value = sLongCitation
                rstCitationHyperlinks.Fields("WebAddress").Value = sWebAddress
                rstCitationHyperlinks.Fields("CHCategory").Value = sCHCategory
                rstCitationHyperlinks.Update
                
                
                
            End If

    
            Set objHttp = Nothing
NextNumber2:
    
        Next
    
            
    Next
    
    For i = 1 To 999
        'build related subtitle links
                
    
            
        'subtitle
        'http://uscode.house.gov/view.xhtml?path=/prelim@title51/subtitle1&edition=prelim
        'http://uscode.house.gov/view.xhtml?path=/prelim@title26/subtitleG&edition=prelim
            
            
        'generate variables
        sSubtitleNumber = i
        sFindCitation = "Subtitle " & sSubchapterNumber
        sLongCitation = "Subtitle " & sSubchapterNumber
        sCHCategory = 2
        sRuleNumber = sCurrentRule
        sWebAddress = "http://uscode.house.gov/view.xhtml?path=/prelim@title" & sRuleNumber & "/subtitle" & sSubtitleNumber & "&edition=prelim"
        sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
    
    
    Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
    objHttp.Open "GET", sWebAddress, False
    objHttp.send ""
    
    Title = objHttp.responseText
    
    If InStr(1, UCase(Title), "<TITLE>Document Not Found") Then
        Debug.Print ("Bad website, moving on to try next one.")
        GoTo NextNumber3
    
    Else
        'add entry to citationhyperlinks
        
        Debug.Print ("Entering " & sFindCitation & " into CitationHyperlinks table.")
        
        'add new entry to citaitonhyperlinks table
        Set rstCitationHyperlinks = CurrentDb.OpenRecordset("CitationHyperlinks")
        rstCitationHyperlinks.AddNew
        rstCitationHyperlinks.Fields("FindCitation").Value = sFindCitation
        rstCitationHyperlinks.Fields("ReplaceHyperlink").Value = sReplaceHyperlink
        rstCitationHyperlinks.Fields("LongCitation").Value = sLongCitation
        rstCitationHyperlinks.Fields("WebAddress").Value = sWebAddress
        rstCitationHyperlinks.Fields("CHCategory").Value = sCHCategory
        rstCitationHyperlinks.Update
        
        
        
    End If

    
            Set objHttp = Nothing
NextNumber3:
            
    Next

    Dim vSubtitleLetters As Variant
    vSubtitleLetters = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    For j = 1 To UBound(vSubtitleLetters)
        'build related subtitle links
                
        'subtitle
        'http://uscode.house.gov/view.xhtml?path=/prelim@title51/subtitle1&edition=prelim
        'http://uscode.house.gov/view.xhtml?path=/prelim@title26/subtitleG&edition=prelim
            
            
        'generate variables
        sSubtitleNumber = vSubtitleLetters(j)
        sFindCitation = "Subtitle " & sSubchapterNumber
        sLongCitation = "Subtitle " & sSubchapterNumber
        sCHCategory = 2
        sRuleNumber = sCurrentRule
        sWebAddress = "http://uscode.house.gov/view.xhtml?path=/prelim@title" & sRuleNumber & "/subtitle" & sSubtitleNumber & "&edition=prelim"
        sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
    
        
        Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
        objHttp.Open "GET", sWebAddress, False
        objHttp.send ""
        
        Title = objHttp.responseText
        
        If InStr(1, UCase(Title), "<TITLE>Document Not Found") Then
            Debug.Print ("Bad website, moving on to try next one.")
            GoTo NextNumber4
        
        Else
            'add entry to citationhyperlinks
            
            Debug.Print ("Entering " & sFindCitation & " into CitationHyperlinks table.")
            
            'add new entry to citaitonhyperlinks table
            Set rstCitationHyperlinks = CurrentDb.OpenRecordset("CitationHyperlinks")
            rstCitationHyperlinks.AddNew
            rstCitationHyperlinks.Fields("FindCitation").Value = sFindCitation
            rstCitationHyperlinks.Fields("ReplaceHyperlink").Value = sReplaceHyperlink
            rstCitationHyperlinks.Fields("LongCitation").Value = sLongCitation
            rstCitationHyperlinks.Fields("WebAddress").Value = sWebAddress
            rstCitationHyperlinks.Fields("CHCategory").Value = sCHCategory
            rstCitationHyperlinks.Update
            
            Set objHttp = Nothing
            
        End If
    
            Set objHttp = Nothing
NextNumber4:
        
            
    Next
    
    
    
    
    For k = 1 To 200000
        'build related section links
                
            
        'Section
        'http://uscode.house.gov/view.xhtml?req=granuleid:USC-prelim-title11-section301&num=0&edition=prelim
        
        '50 U.S.C. 1549
        'http://uscode.house.gov/view.xhtml?req=granuleid:USC-prelim-title50-section1549&num=0&edition=prelim
            
            
        'generate variables
        sSectionNumber = k
        sFindCitation = "Section " & sSectionNumber
        sLongCitation = "Section " & sSectionNumber
        sCHCategory = 2
        sRuleNumber = sCurrentRule
        sWebAddress = "http://uscode.house.gov/view.xhtml?path=/prelim@title" & sRuleNumber & "-section" & sSectionNumber & "&num=0&edition=prelim"
        sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
        
        
        Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
        objHttp.Open "GET", sWebAddress, False
        objHttp.send ""
        
        Title = objHttp.responseText
        
        If InStr(1, UCase(Title), "<TITLE>Document Not Found") Then
            Debug.Print ("Bad website, moving on to try next one.")
            GoTo NextNumber5
        
        Else
            'add entry to citationhyperlinks
            
            Debug.Print ("Entering " & sFindCitation & " into CitationHyperlinks table.")
            
            'add new entry to citaitonhyperlinks table
            Set rstCitationHyperlinks = CurrentDb.OpenRecordset("CitationHyperlinks")
            rstCitationHyperlinks.AddNew
            rstCitationHyperlinks.Fields("FindCitation").Value = sFindCitation
            rstCitationHyperlinks.Fields("ReplaceHyperlink").Value = sReplaceHyperlink
            rstCitationHyperlinks.Fields("LongCitation").Value = sLongCitation
            rstCitationHyperlinks.Fields("WebAddress").Value = sWebAddress
            rstCitationHyperlinks.Fields("CHCategory").Value = sCHCategory
            rstCitationHyperlinks.Update
            
            
            
        End If
    
    
            Set objHttp = Nothing
NextNumber5:
        
            
    Next
    

Next




End Function



Public Function pfRCWRuleScraper1()
On Error Resume Next
'============================================================================
' Name        : pfRCWRuleScraper
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfRCWRuleScraper()
' Description:  builds all RCWs and their links, checks for validation, and puts into CitationHyperlinks table
'============================================================================
Dim sFindCitation As String, sLongCitation As String, sRuleNumber As String
Dim sWebAddress As String, sReplaceHyperlink As String, sCurrentRule As String
Dim Title As String, sTitle2 As String, sTitle1 As String, sTitle As String
Dim sTitle3 As String, sCheck As String, sTitle4 As String
Dim oHTTPText As Object
Dim vLettersArray1(), vLettersArray2() As Variant
Dim rstCitationHyperlinks As DAO.Recordset
Dim iErrorNum As Integer, sCHCategory As Integer
Dim m As Long, n As Long, o As Long, p As Long
Dim i As Long, j As Long, k As Long, l As Long
Dim w As Long, x As Long, y As Long, z As Long

'i build a delay in mine by calling a separate function so it requests only once every 22 seconds
'y = 01 to 385 ; exception 28B, 43, 81
'all under 160 except 28A, 19, 18, 43, 70
'last done 29.48
y = 1
x = 32
For x = 32 To 91 '(RCW first portion x.###.###) '1-91
    
    sTitle1 = x
    
        For y = 1 To 160 '(RCW second portion ###.y.###) '1-160 ; exception 28A, 19, 18, 43, 70
        
            sTitle2 = y
            
            If y < 10 Then sTitle2 = "0" & sTitle2
                    
            For z = 1 To 999 '(RCW third portion ###.###.z) '10 to 990 by 10s
            
                sTitle3 = z
                
                If z < 100 Then sTitle3 = "0" & z
                
                'generate variables
                sCurrentRule = sTitle1 & "." & sTitle2 & "." & sTitle3
                sFindCitation = "RCW " & sCurrentRule
                sLongCitation = "RCW " & sCurrentRule
                sCHCategory = 2
                sRuleNumber = sCurrentRule
                sWebAddress = "https://app.leg.wa.gov/RCW/default.aspx?cite=" & sCurrentRule
                sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
                
                Set oHTTPText = CreateObject("MSXML2.ServerXMLHTTP")
                oHTTPText.Open "GET", sWebAddress, False
                oHTTPText.send ""
                
                Title = oHTTPText.responseText
                sCheck = Left(Title, 215)
                'Debug.Print sCheck
                sCheck = Right(sCheck, 14) 'gets full rcw ## from html title if there is one, if it's not in here it's a bad URL/RCW
                'Debug.Print sCheck
                If InStr(1, UCase(sCheck), sCurrentRule, vbTextCompare) = 0 Then
                    If z = "450" Or z = "900" Then Debug.Print ("RCW " & sCurrentRule & " is a bad RCW; moving on to try next one.")
                    GoTo NextNumber3
                
                Else
                
                    'add entry to citationhyperlinks
                    Debug.Print ("Entering " & sFindCitation & " into CitationHyperlinks table.")
                    
                    'add new entry to citationhyperlinks table
                    Set rstCitationHyperlinks = CurrentDb.OpenRecordset("CitationHyperlinks")
                    rstCitationHyperlinks.AddNew
                    rstCitationHyperlinks.Fields("FindCitation").Value = sFindCitation
                    rstCitationHyperlinks.Fields("ReplaceHyperlink").Value = sReplaceHyperlink
                    rstCitationHyperlinks.Fields("LongCitation").Value = sLongCitation
                    rstCitationHyperlinks.Fields("WebAddress").Value = sWebAddress
                    rstCitationHyperlinks.Fields("CHCategory").Value = sCHCategory
                    rstCitationHyperlinks.Update
                    
                    
                    
                End If
    
    
                Set oHTTPText = Nothing
NextNumber3:
            
            
            
            DoEvents
            Next
            
        DoEvents
        Next
    
    
    vLettersArray1 = Array("9A", "23B", "28A", "28B", "28C", "29A", "30A", "30B35A", "50A", "62A", "71A", "79A")
    
    For m = 1 To UBound(vLettersArray1) '(RCW first portion m.###.###)
    
        sTitle1 = vLettersArray1(m)
        
        For n = 1 To 160 '(RCW second portion ###.n.###) '1-160 ; exception 28A, 19, 18, 43, 70
        
            '1-999 plus A, B, C
        
            vLettersArray2 = Array("A", "B", "C")
            
            For o = 0 To UBound(vLettersArray2) '(RCW second portion m.n[o].p)
                
                sTitle2 = n & vLettersArray2(o)
                
                If n < 10 Then sTitle2 = Str("0" & sTitle2)
        
                For p = 10 To 990 Step 10 '(RCW third portion ###.###.p)
                    
                    sTitle3 = p
                    
                    If p < 100 Then sTitle3 = Str("0" & sTitle3)
                    
                    'generate variables
                    sCurrentRule = sTitle1 & "." & sTitle2 & "." & sTitle3
                    sFindCitation = "RCW " & sCurrentRule
                    sLongCitation = "RCW " & sCurrentRule
                    sCHCategory = 2
                    sRuleNumber = sCurrentRule
                    sWebAddress = "https://app.leg.wa.gov/RCW/default.aspx?cite=" & sCurrentRule
                    sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
                    
                    
                    Set oHTTPText = CreateObject("MSXML2.ServerXMLHTTP")
                    oHTTPText.Open "GET", sWebAddress, False
                    oHTTPText.send ""
                    
                    Title = oHTTPText.responseText
                    
                    If InStr(1, UCase(sCheck), sCurrentRule, vbTextCompare) = 0 Then
                
                        Debug.Print ("RCW " & sCurrentRule & " is a bad RCW; moving on to try next one.")
                        GoTo NextNumber4
                
                    Else
                    
                        'add entry to citationhyperlinks
                        Debug.Print ("Entering " & sFindCitation & " into CitationHyperlinks table.")
                            
                        'add new entry to citaitonhyperlinks table
                        Set rstCitationHyperlinks = CurrentDb.OpenRecordset("CitationHyperlinks")
                        rstCitationHyperlinks.AddNew
                        rstCitationHyperlinks.Fields("FindCitation").Value = sFindCitation
                        rstCitationHyperlinks.Fields("ReplaceHyperlink").Value = sReplaceHyperlink
                        rstCitationHyperlinks.Fields("LongCitation").Value = sLongCitation
                        rstCitationHyperlinks.Fields("WebAddress").Value = sWebAddress
                        rstCitationHyperlinks.Fields("CHCategory").Value = sCHCategory
                        rstCitationHyperlinks.Update
                        
                            
                            
                    End If
    
        
                    Set oHTTPText = Nothing
NextNumber4:
                DoEvents
                Next
            DoEvents
            Next
        DoEvents
        Next
DoEvents
Next


    
    

Next



End Function






    
Public Function pfMARuleScraper()
    
On Error Resume Next
Dim sFindCitation As String, sLongCitation As String, sRuleNumber As String
Dim sWebAddress As String, sReplaceHyperlink As String, sCurrentRule As String
Dim sCurrentPart As String, sCurrentTitle As String, sCurrentChapter As String
Dim sCurrentSection As String
Dim objHttp As Object
Dim vLetterArray(), vRomanArray() As Variant
Dim rstCitationHyperlinks As DAO.Recordset
Dim iErrorNum As Integer, sCHCategory As Integer

Dim i As Long, j As Long, k As Long, l As Long, m As Long
Dim w As Long, x As Long, y As Long, z As Long

'============================================================================
' Name        : pfMARuleScraper
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfMARuleScraper()
' Description:  scrapes all code from MA site
'============================================================================

'possibly find mass website of laws to scrape
    'https://malegislature.gov/Laws/GeneralLaws/PartI/TitleXII/Chapter71/Section37O
'Part I through V
    'Part I   Title I through XXII  Chapter   1 through 182 A through Z Section 1 through 100 A through Z
    'Part II  Title I, II, III      Chapter 183 through 210 A through Z Section 1 through 100 A through Z
    'Part III Title I through VI    Chapter 211 through 262 A through Z Section 1 through 100 A through Z
    'Part IV  Title I, II           Chapter 263 through 280 A through Z Section 1 through 100 A through Z
    'Part V   Title I               Chapter 281 and 282 A through Z     Section 1 through 100 A through Z
                                                                    '�
vRomanArray = Array("I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX", "XX", "XXI", "XXII", "XXIII", "XXIV", "XXV")
vLetterArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")

'"https://malegislature.gov/Laws/GeneralLaws/Part" & sCurrentPart & "/Title" & sCurrentTitle & "/Chapter" & sCurrentChapter & "" & sCurrentSection

For x = 1 To 5 'part I through V

    sCurrentPart = vRomanArray(x)
    
    For i = 0 To UBound(vRomanArray) 'title I through XXII
    
        sCurrentTitle = vRomanArray(x)
        
        For j = 1 To 282 'Chapter 1 through 282
        
            'Basic form:  Mass. Gen. Laws Chapter, � Section (Date).   |   Name of Act. Volume Mass. Acts Page Date.
            'Examples:    Mass. Gen. Laws ch. 71, � 1A (1966).   |  An Act Designating Certain Bridges in the Town of Middleborough. 1967 Mass. Acts 116. 8 October 1997.
            
            sCurrentChapter = j
            
            'also without letters
            
            sCurrentRule = j
            sFindCitation = "Mass. Gen. Laws ch. " & sCurrentChapter & " (####)"
            sLongCitation = "Mass. Gen. Laws ch. " & sCurrentChapter & " (####)"
            sCHCategory = 2
            sWebAddress = "https://malegislature.gov/Laws/GeneralLaws/Part" & sCurrentPart & "/Title" & sCurrentTitle & "/Chapter" & sCurrentChapter
            sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
            
            
            Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
            objHttp.Open "GET", sWebAddress, False
            objHttp.send ""
            
            If j = 150 Then 'error
            
                Debug.Print ("Completed:  " & sFindCitation & "; moving on to try next one.")
                GoTo NextNumber1
            
            Else
            
                'add entry to citationhyperlinks
                Debug.Print ("Entering " & sFindCitation & " into CitationHyperlinks table.")
                
                'add new entry to citaitonhyperlinks table
                Set rstCitationHyperlinks = CurrentDb.OpenRecordset("CitationHyperlinks")
                rstCitationHyperlinks.AddNew
                rstCitationHyperlinks.Fields("FindCitation").Value = sFindCitation
                rstCitationHyperlinks.Fields("ReplaceHyperlink").Value = sReplaceHyperlink
                rstCitationHyperlinks.Fields("LongCitation").Value = sLongCitation
                rstCitationHyperlinks.Fields("WebAddress").Value = sWebAddress
                rstCitationHyperlinks.Fields("CHCategory").Value = sCHCategory
                rstCitationHyperlinks.Update
                
            End If
        
        
            Set objHttp = Nothing
            
            For k = 0 To UBound(vLetterArray) ' A through Z
            
                'Basic form:  Mass. Gen. Laws Chapter, � Section (Date).   |   Name of Act. Volume Mass. Acts Page Date.
                'Examples:    Mass. Gen. Laws ch. 71, � 1A (1966).   |  An Act Designating Certain Bridges in the Town of Middleborough. 1967 Mass. Acts 116. 8 October 1997.
                
                sCurrentChapter = j & vLetterArray(k)
                
                sCurrentRule = sCurrentChapter
                sFindCitation = "Mass. Gen. Laws ch. " & sCurrentChapter & ", (####)"
                sLongCitation = "Mass. Gen. Laws ch. " & sCurrentChapter & ", (####)"
                sCHCategory = 2
                sWebAddress = "https://malegislature.gov/Laws/GeneralLaws/Part" & sCurrentPart & "/Title" & sCurrentTitle & "/Chapter" & sCurrentChapter
                sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
                                
                Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
                objHttp.Open "GET", sWebAddress, False
                objHttp.send ""
                                
                If j = 150 Then 'error
                
                    Debug.Print ("Completed:  " & sFindCitation & "; moving on to try next one.")
                    GoTo NextNumber1
                
                Else
                
                    'add entry to citationhyperlinks
                    Debug.Print ("Entering " & sFindCitation & " into CitationHyperlinks table.")
                    
                    'add new entry to citaitonhyperlinks table
                    Set rstCitationHyperlinks = CurrentDb.OpenRecordset("CitationHyperlinks")
                    rstCitationHyperlinks.AddNew
                    rstCitationHyperlinks.Fields("FindCitation").Value = sFindCitation
                    rstCitationHyperlinks.Fields("ReplaceHyperlink").Value = sReplaceHyperlink
                    rstCitationHyperlinks.Fields("LongCitation").Value = sLongCitation
                    rstCitationHyperlinks.Fields("WebAddress").Value = sWebAddress
                    rstCitationHyperlinks.Fields("CHCategory").Value = sCHCategory
                    rstCitationHyperlinks.Update
                    
                End If
            
                Set objHttp = Nothing
                    
                For l = 1 To 100 'Section 1 through 100
                
                    'Basic form:  Mass. Gen. Laws Chapter, � Section (Date).   |   Name of Act. Volume Mass. Acts Page Date.
                    'Examples:    Mass. Gen. Laws ch. 71, � 1A (1966).   |  An Act Designating Certain Bridges in the Town of Middleborough. 1967 Mass. Acts 116. 8 October 1997.
                        
                    sCurrentSection = l
                    
                    'also without letters
                    
                    sFindCitation = "Mass. Gen. Laws ch. " & sCurrentChapter & "� " & sCurrentSection & " (####)"
                    sLongCitation = "Mass. Gen. Laws ch. " & sCurrentChapter & "� " & sCurrentSection & " (####)"
                    sCHCategory = 2
                    sWebAddress = "https://malegislature.gov/Laws/GeneralLaws/Part" & sCurrentPart & "/Title" & sCurrentTitle & "/Chapter" & sCurrentChapter & "/Section" & sCurrentSection
                    sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
                    
                    
                    Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
                    objHttp.Open "GET", sWebAddress, False
                    objHttp.send ""
                    
                    If j = 50 Then 'error
                    
                        Debug.Print ("Completed:  " & sFindCitation & "; moving on to try next one.")
                        GoTo NextNumber1
                    
                    Else
                    
                        'add entry to citationhyperlinks
                        Debug.Print ("Entering " & sFindCitation & " into CitationHyperlinks table.")
                        
                        'add new entry to citaitonhyperlinks table
                        Set rstCitationHyperlinks = CurrentDb.OpenRecordset("CitationHyperlinks")
                        rstCitationHyperlinks.AddNew
                        rstCitationHyperlinks.Fields("FindCitation").Value = sFindCitation
                        rstCitationHyperlinks.Fields("ReplaceHyperlink").Value = sReplaceHyperlink
                        rstCitationHyperlinks.Fields("LongCitation").Value = sLongCitation
                        rstCitationHyperlinks.Fields("WebAddress").Value = sWebAddress
                        rstCitationHyperlinks.Fields("CHCategory").Value = sCHCategory
                        rstCitationHyperlinks.Update
                        
                    End If
                
                    Set objHttp = Nothing
                    
                    For m = 0 To UBound(vLetterArray) ' A through Z
                    
                        'Basic form:  Mass. Gen. Laws Chapter, � Section (Date).   |   Name of Act. Volume Mass. Acts Page Date.
                        'Examples:    Mass. Gen. Laws ch. 71, � 1A (1966).   |  An Act Designating Certain Bridges in the Town of Middleborough. 1967 Mass. Acts 116. 8 October 1997.
                            
                        sCurrentSection = l & vLetterArray(m)
                        
                        sFindCitation = "Mass. Gen. Laws ch. " & sCurrentChapter & "� " & sCurrentSection & " (####)"
                        sLongCitation = "Mass. Gen. Laws ch. " & sCurrentChapter & "� " & sCurrentSection & " (####)"
                        sCHCategory = 2
                        sWebAddress = "https://malegislature.gov/Laws/GeneralLaws/Part" & sCurrentPart & "/Title" & sCurrentTitle & "/Chapter" & sCurrentChapter & "/Section" & sCurrentSection
                        sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
                        
                        
                        Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
                        objHttp.Open "GET", sWebAddress, False
                        objHttp.send ""
                        
                        If j = 50 Then 'error
                        
                            Debug.Print ("Completed:  " & sFindCitation & "; moving on to try next one.")
                            GoTo NextNumber1
                        
                        Else
                        
                            'add entry to citationhyperlinks
                            Debug.Print ("Entering " & sFindCitation & " into CitationHyperlinks table.")
                            
                            'add new entry to citaitonhyperlinks table
                            Set rstCitationHyperlinks = CurrentDb.OpenRecordset("CitationHyperlinks")
                            rstCitationHyperlinks.AddNew
                            rstCitationHyperlinks.Fields("FindCitation").Value = sFindCitation
                            rstCitationHyperlinks.Fields("ReplaceHyperlink").Value = sReplaceHyperlink
                            rstCitationHyperlinks.Fields("LongCitation").Value = sLongCitation
                            rstCitationHyperlinks.Fields("WebAddress").Value = sWebAddress
                            rstCitationHyperlinks.Fields("CHCategory").Value = sCHCategory
                            rstCitationHyperlinks.Update
                            
                            
                            
                        End If
                    
                    
                        Set objHttp = Nothing
                        
                    
                    Next
                    
                Next
                
            Next
            
NextNumber1:
        Next
        
    Next
    
Next
    

End Function





Function fUnCompleteTimeMgmtTasks()
'============================================================================
' Name        : fUnCompleteTimeMgmtTasks
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fUnCompleteTimeMgmtTasks()
' Description:  unchecks all status boxes for a job number
'============================================================================



Dim rstCommHistory As DAO.Recordset, sQuestion As String, sAnswer As String
Dim x As Integer

sCourtDatesID = InputBox("Job Number?")

StartHere:
Set rstCommHistory = CurrentDb.OpenRecordset("SELECT * FROM Statuses WHERE CourtDatesID=" & sCourtDatesID & ";")
                                             'SELECT * FROM Tasks WHERE Title LIKE '*1945*';


If Not rstCommHistory.EOF Then
    rstCommHistory.MoveFirst
Else
    GoTo EndHere

End If

 Do Until rstCommHistory.EOF = True
    rstCommHistory.Edit
    For x = 2 To 28 Step 1
        rstCommHistory.Fields(x).Value = False
    Next
    rstCommHistory.Update
    rstCommHistory.MoveNext
Loop

rstCommHistory.Close
Set rstCommHistory = Nothing


Set rstCommHistory = CurrentDb.OpenRecordset("SELECT * FROM Tasks WHERE Title LIKE '*" & sCourtDatesID & "*';")
                                             'SELECT * FROM Tasks WHERE Priority='(1) Stage 1' AND Title LIKE '*1945*';
If Not rstCommHistory.EOF Then
    rstCommHistory.MoveFirst
Else
    GoTo EndHere

End If


 Do Until rstCommHistory.EOF = True
    rstCommHistory.Edit
        rstCommHistory.Fields("Completed").Value = False
    rstCommHistory.Update
    rstCommHistory.MoveNext
Loop

EndHere:
sQuestion = "Do you want to undo another one?"
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
    
    If sAnswer = vbNo Then 'Code for No
    
        MsgBox "Done!"
        
    Else 'Code for yes
    
        'ask for job number
        sCourtDatesID = InputBox("Job Number?")
        GoTo StartHere

End If


End Function








Function fCompleteTimeMgmtTasks()
'============================================================================
' Name        : fCompleteTimeMgmtTasks
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fCompleteTimeMgmtTasks()
' Description:  checks all tasks from tasks table for a job number
'============================================================================
'Call fWunderlistGetTasksOnList 'insert function name here


Dim rstCommHistory As DAO.Recordset, sQuestion As String, sAnswer As String

sCourtDatesID = InputBox("Job Number?")

StartHere:
Set rstCommHistory = CurrentDb.OpenRecordset("SELECT * FROM Tasks WHERE Title LIKE '*" & sCourtDatesID & "*';")
                                             'SELECT * FROM Tasks WHERE Title LIKE '*1945*';
                                             
If Not rstCommHistory.EOF Then
    rstCommHistory.MoveFirst
Else
    GoTo EndHere

End If


 Do Until rstCommHistory.EOF = True
    rstCommHistory.Edit
    rstCommHistory.Fields("Completed").Value = True
    rstCommHistory.Update
    rstCommHistory.MoveNext
Loop

EndHere:
sQuestion = "Do you want to enter another one?"
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
    
    If sAnswer = vbNo Then 'Code for No
    
        MsgBox "Done!"
        
    Else 'Code for yes
    
        'ask for job number
        sCourtDatesID = InputBox("Job Number?")
        GoTo StartHere

End If


End Function



Function fCompleteStatusBoxes()
'============================================================================
' Name        : fCompleteStatusBoxes
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fCompleteStatusBoxes()
' Description:  checks all status boxes for a job number
'============================================================================
'Call fWunderlistGetTasksOnList 'insert function name here


Dim rstCommHistory As DAO.Recordset, sQuestion As String, sAnswer As String
Dim x As Integer

sCourtDatesID = InputBox("Job Number to complete statuses of?")

StartHere:
Set rstCommHistory = CurrentDb.OpenRecordset("SELECT * FROM Statuses WHERE CourtDatesID=" & sCourtDatesID & ";")
                                             'SELECT * FROM Tasks WHERE Title LIKE '*(1.*' AND Title LIKE '*1945*';
                                             
If Not rstCommHistory.EOF Then
    rstCommHistory.MoveFirst
Else
    GoTo EndHere

End If


 Do Until rstCommHistory.EOF = True
    rstCommHistory.Edit
    For x = 2 To 28 Step 1
        rstCommHistory.Fields(x).Value = True
    Next
    rstCommHistory.Update
    rstCommHistory.MoveNext
Loop

EndHere:
sQuestion = "Do you want to complete another one?"
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
    
    If sAnswer = vbNo Then 'Code for No
    
        MsgBox "Done!"
        
    Else 'Code for yes
    
        'ask for job number
        sCourtDatesID = InputBox("Job Number?")
        GoTo StartHere

End If

rstCommHistory.Close
Set rstCommHistory = Nothing

End Function


Function fCompleteStage1Tasks()
'============================================================================
' Name        : fCompleteStage1Tasks
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fCompleteStage1Tasks()
' Description:  checks all stage 1 status boxes for a job number
'============================================================================
'Call fWunderlistGetTasksOnList 'insert function name here


Dim rstCommHistory As DAO.Recordset, sQuestion As String, sAnswer As String
Dim x As Integer

sCourtDatesID = InputBox("Job Number to complete Stage 1 of?")

StartHere:
Set rstCommHistory = CurrentDb.OpenRecordset("SELECT * FROM Statuses WHERE CourtDatesID=" & sCourtDatesID & ";")
                                             'SELECT * FROM Tasks WHERE Priority='(1) Stage 1' AND Title LIKE '*1945*';
                             'SELECT * FROM Tasks WHERE Priority='(1) Stage 1' AND Title LIKE '*1945*';
If Not rstCommHistory.EOF Then
    rstCommHistory.MoveFirst
Else
    GoTo EndHere

End If


 Do Until rstCommHistory.EOF = True
    rstCommHistory.Edit
    For x = 2 To 10 Step 1
        rstCommHistory.Fields(x).Value = True
    Next
    rstCommHistory.Update
    rstCommHistory.MoveNext
Loop


rstCommHistory.Close
Set rstCommHistory = Nothing


Set rstCommHistory = CurrentDb.OpenRecordset("SELECT * FROM Tasks WHERE Priority='(1) Stage 1' AND Title LIKE '*" & sCourtDatesID & "*';")
                                             'SELECT * FROM Tasks WHERE Priority='(1) Stage 1' AND Title LIKE '*1945*';

If Not rstCommHistory.EOF Then
    rstCommHistory.MoveFirst
Else
    GoTo EndHere

End If


 Do Until rstCommHistory.EOF = True
    rstCommHistory.Edit
        rstCommHistory.Fields("Completed").Value = True
    rstCommHistory.Update
    rstCommHistory.MoveNext
Loop

rstCommHistory.Close
Set rstCommHistory = Nothing



EndHere:
sQuestion = "Do you want to complete another one?"
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

If sAnswer = vbNo Then 'Code for No

    MsgBox "Done!"
    
Else 'Code for yes

    'ask for job number
    sCourtDatesID = InputBox("Job Number?")
    GoTo StartHere

End If

End Function

Function fCompleteStage2Tasks()
'============================================================================
' Name        : fCompleteStage2Tasks
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fCompleteStage2Tasks()
' Description:  checks all stage 2 status boxes for a job number
'============================================================================
'Call fWunderlistGetTasksOnList 'insert function name here


Dim rstCommHistory As DAO.Recordset, sQuestion As String, sAnswer As String
Dim x As Integer

sCourtDatesID = InputBox("Job Number to complete Stage 2 of?")

StartHere:
Set rstCommHistory = CurrentDb.OpenRecordset("SELECT * FROM Statuses WHERE CourtDatesID=" & sCourtDatesID & ";")
                                             'SELECT * FROM Tasks WHERE Title LIKE '*1945*';
                                             
If Not rstCommHistory.EOF Then
    rstCommHistory.MoveFirst
Else
    GoTo EndHere

End If


 Do Until rstCommHistory.EOF = True
    rstCommHistory.Edit
    For x = 11 To 15 Step 1
        rstCommHistory.Fields(x).Value = True
    Next
    rstCommHistory.Update
    rstCommHistory.MoveNext
Loop


rstCommHistory.Close
Set rstCommHistory = Nothing


Set rstCommHistory = CurrentDb.OpenRecordset("SELECT * FROM Tasks WHERE Priority='(2) Stage 2' AND Title LIKE '*" & sCourtDatesID & "*';")
                                             'SELECT * FROM Tasks WHERE Priority='(1) Stage 1' AND Title LIKE '*1945*';

                             'SELECT * FROM Tasks WHERE Priority='(1) Stage 1' AND Title LIKE '*1945*';
If Not rstCommHistory.EOF Then
    rstCommHistory.MoveFirst
Else
    GoTo EndHere

End If

 Do Until rstCommHistory.EOF = True
    rstCommHistory.Edit
        rstCommHistory.Fields("Completed").Value = True
    rstCommHistory.Update
    rstCommHistory.MoveNext
Loop

rstCommHistory.Close
Set rstCommHistory = Nothing


EndHere:
sQuestion = "Do you want to complete another one?"
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
    
    If sAnswer = vbNo Then 'Code for No
    
        MsgBox "Done!"
        
    Else 'Code for yes
    
        'ask for job number
        sCourtDatesID = InputBox("Job Number?")
        GoTo StartHere

End If


End Function


Function fCompleteStage3Tasks()
'============================================================================
' Name        : fCompleteStage3Tasks
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fCompleteStage3Tasks()
' Description:  checks all stage 3 status boxes for a job number
'============================================================================
'Call fWunderlistGetTasksOnList 'insert function name here


Dim rstCommHistory As DAO.Recordset, sQuestion As String, sAnswer As String
Dim x As Integer

sCourtDatesID = InputBox("Job Number to complete Stage 3 of?")

StartHere:
Set rstCommHistory = CurrentDb.OpenRecordset("SELECT * FROM Statuses WHERE CourtDatesID=" & sCourtDatesID & ";")
                                             'SELECT * FROM Tasks WHERE Title LIKE '*1945*';
                             'SELECT * FROM Tasks WHERE Priority='(1) Stage 1' AND Title LIKE '*1945*';
If Not rstCommHistory.EOF Then
    rstCommHistory.MoveFirst
Else
    GoTo EndHere

End If
x = 16
Do Until rstCommHistory.EOF = True
    rstCommHistory.Edit
    For x = 16 To 20 Step 1
        rstCommHistory.Fields(x).Value = True
    Next
    rstCommHistory.Update
    rstCommHistory.MoveNext
Loop


rstCommHistory.Close
Set rstCommHistory = Nothing


Set rstCommHistory = CurrentDb.OpenRecordset("SELECT * FROM Tasks WHERE Priority='(3) Stage 3' AND Title LIKE '*" & sCourtDatesID & "*';")
                                             'SELECT * FROM Tasks WHERE Priority='(1) Stage 1' AND Title LIKE '*1945*';



If Not rstCommHistory.EOF Then
    rstCommHistory.MoveFirst
Else
    GoTo EndHere

End If


 Do Until rstCommHistory.EOF = True
    rstCommHistory.Edit
        rstCommHistory.Fields("Completed").Value = True
    rstCommHistory.Update
    rstCommHistory.MoveNext
Loop


EndHere:
sQuestion = "Do you want to complete another one?"
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
    
    If sAnswer = vbNo Then 'Code for No
    
        MsgBox "Done!"
        
    Else 'Code for yes
    
        'ask for job number
        sCourtDatesID = InputBox("Job Number?")
        GoTo StartHere

End If

rstCommHistory.Close
Set rstCommHistory = Nothing

End Function

Function fCompleteStage4Tasks()
'============================================================================
' Name        : fCompleteStage4Tasks
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fCompleteStage4Tasks()
' Description:  checks all stage 4 status boxes for a job number
'============================================================================
'Call fWunderlistGetTasksOnList 'insert function name here


Dim rstCommHistory As DAO.Recordset, sQuestion As String, sAnswer As String
Dim x As Integer

sCourtDatesID = InputBox("Job Number to complete Stage 4 of?")

StartHere:
Set rstCommHistory = CurrentDb.OpenRecordset("SELECT * FROM Statuses WHERE CourtDatesID=" & sCourtDatesID & ";")
                                             'SELECT * FROM Tasks WHERE Title LIKE '*1945*';
                                             
If Not rstCommHistory.EOF Then
    rstCommHistory.MoveFirst
Else
    GoTo EndHere

End If


 Do Until rstCommHistory.EOF = True
    rstCommHistory.Edit
    For x = 21 To 28 Step 1
        rstCommHistory.Fields(x).Value = True
    Next
    rstCommHistory.Update
    rstCommHistory.MoveNext
Loop


rstCommHistory.Close
Set rstCommHistory = Nothing


Set rstCommHistory = CurrentDb.OpenRecordset("SELECT * FROM Tasks WHERE Priority='(4) Stage 4' AND Title LIKE '*" & sCourtDatesID & "*';")
                                             'SELECT * FROM Tasks WHERE Priority='(1) Stage 1' AND Title LIKE '*1945*';
                                             
If Not rstCommHistory.EOF Then
    rstCommHistory.MoveFirst
Else
    GoTo EndHere

End If


 Do Until rstCommHistory.EOF = True
    rstCommHistory.Edit
        rstCommHistory.Fields("Completed").Value = True
    rstCommHistory.Update
    rstCommHistory.MoveNext
Loop

rstCommHistory.Close
Set rstCommHistory = Nothing


EndHere:
sQuestion = "Do you want to complete another one?"
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
    
    If sAnswer = vbNo Then 'Code for No
    
        MsgBox "Done!"
        
    Else 'Code for yes
    
        'ask for job number
        sCourtDatesID = InputBox("Job Number?")
        GoTo StartHere

End If

rstCommHistory.Close
Set rstCommHistory = Nothing

End Function






Function fFixBarAddressField()
'============================================================================
' Name        : fFixBarAddressField
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fFixBarAddressField()
' Description:  fixes address field in baraddresses table
'============================================================================
Dim rstBarAddresses As DAO.Recordset, rstCustomers As DAO.Recordset
Dim sBarNameArray() As String, sAddressArray() As String, sCityArray() As String
Dim sCityArray1() As String
Dim sLastName As String, sFirstName As String, sBarName As String
Dim sCompany As String, sAddress As String, sPhone As String
Dim sCity As String, sState As String, sZIP As String
Dim sAddress1 As String, sAddress2 As String

'Customer fields: id, Company, MrMs, LastName, FirstName, EmailAddress, JobTitle, BusinessPhone
    'MobilePhone, FaxNumber, Address, City, State, ZIP, Web Page, Notes, FactoringApproved
 


    Set rstBarAddresses = CurrentDb.OpenRecordset("BarAddresses")
    rstBarAddresses.MoveFirst
    
    Debug.Print "--------------------------------"

    
    Do Until rstBarAddresses.EOF = True
        
    
        'get bar name, company name, address field value, phone
        sBarName = rstBarAddresses.Fields("BarName").Value
        sCompany = rstBarAddresses.Fields("Company").Value
        If sCompany = "" Then sCompany = "Attorney at Law"
        sAddress = rstBarAddresses.Fields("Address").Value
        sPhone = rstBarAddresses.Fields("Phone").Value
        
        'parse bar name
        sBarName = Trim(sBarName)
        sBarNameArray() = Split(sBarName, " ")
        
        If UBound(sBarNameArray) = 2 Then
            sLastName = sBarNameArray(2)
            sFirstName = sBarNameArray(0) & " " & sBarNameArray(1)
        
        ElseIf UBound(sBarNameArray) = 1 Then
            sLastName = sBarNameArray(1)
            sFirstName = sBarNameArray(0)
        
        End If
        
        'parse address
        
        sAddress = Replace(sAddress, "<br>", "|")
        sAddressArray() = Split(sAddress, "|")
        
        If UBound(sAddressArray) = 3 Then
            'address1
            sAddress1 = sAddressArray(0)
            'address2
            sAddress2 = sAddressArray(1)
            'city, state, zip
            
            sCityArray() = Split(sAddressArray(2), ",")
            sCity = sCityArray(0)
            sCityArray1() = Split(sCityArray(1), " ")
            
            sState = sCityArray1(1)
            sZIP = Left(sCityArray1(2), 5)
            
            
        ElseIf UBound(sAddressArray) = 2 Then
            'address1
            sAddress1 = sAddressArray(0)
            'city, state, zip
            sCityArray() = Split(sAddressArray(1), ",")
            
            If sCityArray(1) <> Empty Then sCityArray1() = Split(sCityArray(1), " ")
            If sCityArray1(1) <> "" Then sState = sCityArray1(1)
            If sCityArray1(0) <> "" Then sCity = sCityArray(0)
            On Error Resume Next
            If sCityArray1(2) <> "" Then sZIP = Left(sCityArray1(2), 5)
        End If
    
        Debug.Print sCompany
        Debug.Print sFirstName & " " & sLastName
        Debug.Print sAddress1
        If sAddress2 <> "" Then Debug.Print sAddress2
        Debug.Print sCity & ", " & sState & " " & sZIP
        Debug.Print sPhone
        Debug.Print "--------------------------------"
        
        
        
    
        'new customers record
        Set rstCustomers = CurrentDb.OpenRecordset("Customers")
        rstCustomers.AddNew
        rstCustomers.Fields("LastName").Value = sLastName
        rstCustomers.Fields("FirstName").Value = sFirstName
        rstCustomers.Fields("Company").Value = sCompany
        rstCustomers.Fields("BusinessPhone").Value = sPhone
        rstCustomers.Fields("Address").Value = sAddress
        rstCustomers.Fields("City").Value = sCity
        rstCustomers.Fields("State").Value = sState
        rstCustomers.Fields("ZIP").Value = sZIP
        rstCustomers.Fields("Notes").Value = "notes"
        rstCustomers.Fields("JobTitle").Value = "Attorney"
        rstCustomers.Update
        
        sAddress2 = ""

    rstBarAddresses.MoveNext
    Loop

End Function

