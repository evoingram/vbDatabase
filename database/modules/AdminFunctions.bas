Attribute VB_Name = "AdminFunctions"
'@Folder("Database.Admin.Modules")
Option Compare Database
Option Explicit
'============================================================================
'class module cmAdminFunctions

'variables:
'   NONE

'functions:
'pfUpdateCheckboxStatus:     Description:  updates Statuses checkbox field you specify
'                                Arguments:    sStatusesField
'pfDownloadfromFTP:          Description:  downloads files
'                            Arguments:    NONE
'pfDownloadFTPsite:          Description:  downloads files modified today (a.k.a. new files on FTP)
'                            Arguments:    mySession
'pfCheckFolderExistence:     Description:  checks for Audio, Transcripts, FTP, WorkingFiles, Notes subfolders and RoughDraft and creates if not exists
'                            Arguments:    NONE
'pfCommunicationHistoryAdd:  Description:  adds entry to CommunicationHistory
'                            Arguments:    CHTopic
'pfStripIllegalChar:         Description:  strips illegal characters from input
'                            Arguments:    StrInput
'pfGetFolder:                Description:  gets folder
'                            Arguments:    Folders, EntryID, StoreID, fld
'pfBrowseForFolder:          Description:  browses for folder
'                            Arguments:    StrSavePath, optional OpenAt
'pfSingleBAScrapeSpecificBarNo:  Description:  gets one bar number's info from the WA bar website (pick any range from 1 to 55000)
'                            Arguments:    sWebSiteBarNo
'pfScrapingBALoop:           Description:  gets a range you specify of bar numbers' info from the WA bar website (pick any range from 1 to 55000)
'                            Arguments:    vWebSiteBarNo, vWebSiteBarNoGoal
'pfReformatTable:            Description:  reformats scraped Bar addresses to useable format for table
'                            Arguments:    NONE
'pfUpdateCheckboxStatus:     Description:  updates Statuses checkbox field you specify
'                            Arguments:    sStatusesField
'pfDelay:                    Description:  sleep function
'                            Arguments:    lSeconds
'pfPriorityPointsAlgorithm:  Description:  assigns priority points to various tasks in Tasks table and inserts it into the PriorityPoints field
'                                          priority scale 1 to 100
'                            Arguments:    NONE
'pfDebugSQLStatement:        Description:  debug.prints data source query string
'                            Arguments:    NONE
'pfGenerateJobTasks:         Description:  generates job tasks in the Tasks table
'                            Arguments:    NONE
'pfDownloadfromFTP:          Description:  downloads files
'                            Arguments:    NONE
'pfDownloadFTPsite:          Description:  downloads files modified today (a.k.a. new files on FTP)
'                            Arguments:    mySession
'pfProcessFolder:            Description:  process emails in Outlook folder named AccessTest and places them in db as UnprocessedCommunication
'                            Arguments:    oOutlookMAPIFolder
'pfFileExists:               Description:  check if file exists
'                            Arguments:    path
'pfAcrobatGetNumPages:       Description:  gets number of pages from PDF and confirms with you
'                                          IS TOA ON SECOND PAGE?  IF YES, -2 pgs; IF NO, -1 pg
'                            Arguments:    sCourtDatesID
'pfReadXML:                  Description:  reads shipping XML and sends "Shipped" email to client
'                            Arguments:    NONE
'pfFileRenamePrompt:         Description:  renames transcript to specified name, mainly for contractors
'                            Arguments:    NONE
'pfWaitSeconds:              Description:  waits for a specified number of seconds
'                            Arguments:    iSeconds
'pfDailyTaskAddFunction:     Description:  adds static daily tasks to Tasks table
'                            Arguments:    NONE
'pfAvailabilitySchedule:     Description:  opens availability calculator
'                            Arguments:    NONE
'pfWeeklyTaskAddFunction:    Description:  adds static weekly tasks to Tasks table
'                            Arguments:    NONE
'pfMonthlyTaskAddFunction:   Description:  adds static monthly tasks to Tasks table
'                            Arguments:    NONE
'pfMoveSelectedMessages:     Description:  move selected messages to network drive
'                            Arguments:    NONE
'pfEmailsExport1:            Description:  export specified fields from each mail / item in selected folder
'                            Arguments:    NONE
'pfCommHistoryExportSub:     Description:  exports emails to CommunicationsHistory table
'                            Arguments:    NONE
'pfAskforNotes:              Description:  file dialog picker to select notes and copy them to notes folder for job
'                            Arguments:    NONE
'pfAskforAudio:              Description:  file dialog picker to select audio and copy them to audio folder for job
'                            Arguments:    NONE
'fWunderlistGetFolders()     Description:  gets list of Wunderlist folders or folder revisions
'                            Arguments:    NONE
'fWunderlistGetTasksOnList() Description:  gets tasks on Wunderlist list
'                            Arguments:    NONE
'fWunderlistAdd()            Description:  adds task to Wunderlist
'                            Arguments:    NONE
'fWLGenerateJSONInfo         Description:  get info for WL API
'                            Arguments:    NONE
'fWunderlistGetLists()       Description:  gets all Wunderlist lists
'                            Arguments:    NONE
'pfRCWRuleScraper1()         Description:  builds RCW rule links and citations
'                            Arguments:    NONE
'GetLevel()                  Description:  gets header level in word
'                            Arguments:    NONE
'============================================================================

Public Sub pfUpdateCheckboxStatus(sStatusesField As String)
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

End Sub

Public Sub pfDebugSQLStatement()
    '============================================================================
    ' Name        : pfDebugSQLStatement
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfDebugSQLStatement
    ' Description : debug.prints data source query string
    '============================================================================

    Dim oWordApp As New Word.Application
    Dim oWordDoc As Word.Document

    oWordApp.Application.Visible = True

    Set oWordApp = Nothing
    Set oWordDoc = Nothing

    Debug.Print oWordApp.Application.ActiveDocument.MailMerge.DataSource.QueryString

    oWordApp.Quit
    Set oWordApp = Nothing

End Sub

Public Sub pfDownloadfromFTP()
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

    If Err.Number <> 0 Then                      ' Query for errors
        MsgBox "Error: " & Err.Description
        Err.Clear                                ' Clear the error
    End If
    seCurrent.Dispose                            ' Disconnect, clean up

    On Error GoTo 0                              ' Restore default error handling

End Sub

Public Sub pfDownloadFTPsite(ByRef mySession As Session)
    '============================================================================
    ' Name        : pfDownloadFTPsite
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfDownloadFTPsite(mySession)
    ' Description : downloads files modified today (a.k.a. new files on FTP)
    '============================================================================
    'TODO: ftp
    Dim seopFTPSettings As New SessionOptions
    Dim tropFTPSettings As New TransferOptions
    Dim transferResult As TransferOperationResult
    
    Dim cJob As Job
    Set cJob = New Job

    With seopFTPSettings                         ' Setup session options
        .Protocol = Protocol_Ftp
        .HostName = "ftp.aquoco.co"
        .Username = Environ("ftpUserName")
        .password = Environ("ftpPassword")
    End With
    
    mySession.Open seopFTPSettings               ' Connect
    tropFTPSettings.TransferMode = TransferMode_Binary ' Upload files
    tropFTPSettings.FileMask = "*>=1D"

    Set transferResult = mySession.GetFiles("/public_html/ProjectSend/upload/files/", cJob.DocPath.UNFileInbox, False, tropFTPSettings)
    transferResult.Check                         ' Throw on any error

    MsgBox "You may now find any files downloaded today in" & cJob.DocPath.FileInbox & "."

End Sub

Public Sub pfProcessFolder(ByVal oOutlookPickedFolder As Outlook.MAPIFolder)
    '============================================================================
    ' Name        : pfProcessFolder
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfProcessFolder(oOutlookPickedFolder)
    ' Description : process emails in Outlook folder named AccessTest and places them in db as UnprocessedCommunication
    '============================================================================

    Dim dReceived As Date
    Dim sReceivedTime As String
    Dim sEmailHyperlink As String
    Dim sTableHyperilnk As String

    Dim oOutlookNamespace As Outlook.Namespace
    Dim adocOutlookExport As ADODB.Connection
    Dim adorstOutlookExport As ADODB.Recordset
    Dim oOutLookMAPIFolder As Outlook.MAPIFolder
    Dim oOutlookMail As Outlook.MailItem
    
    Dim cJob As Job
    Set cJob = New Job
    
    Set oOutlookNamespace = GetNamespace("MAPI")
    Set oOutlookPickedFolder = oOutlookNamespace.PickFolder
    Set adocOutlookExport = CreateObject("ADODB.Connection")
    Set adorstOutlookExport = CreateObject("ADODB.Recordset")

    For Each oOutlookMail In oOutlookPickedFolder.Items
        dReceived = oOutlookMail.ReceivedTime
        sReceivedTime = Format(dReceived, "YYYYMMDD-hhmm")
        oOutlookMail.SaveAs cJob.DocPath.EmailDirectory & sReceivedTime & "-Email.msg", 3
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
        oOutlookMail.SaveAs cJob.DocPath.EmailDirectory & sReceivedTime & "-Email.msg", 3
        sEmailHyperlink = cJob.DocPath.EmailDirectory & sReceivedTime & "-Email.msg"
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
 
End Sub

Public Sub pfFileExists(ByVal path_ As String)
    '============================================================================
    ' Name        : pfFileExists
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfFileExists(ByVal path)
    ' Description : check if file exists
    '============================================================================
    Dim FileExists As Variant
    '@Ignore AssignmentNotUsed
    FileExists = (Len(Dir(path_)) > 0)

End Sub

Public Sub pfAcrobatGetNumPages(sCourtDatesID As String)
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

    Dim sQuestion As String
    Dim sAnswer As String
    Dim sSQL As String

    Dim qdf As QueryDef
    
    Dim oAcrobatDoc As Object
    
    Dim cJob As Job
    Set cJob = New Job
    
    Set oAcrobatDoc = New AcroPDDoc

    oAcrobatDoc.Open (cJob.DocPath.TranscriptFPB)        'update file location

    sActualQuantity = oAcrobatDoc.GetNumPages
    sQuestion = "This transcript came to " & sActualQuantity & " pages.  Is the table of authorities on a separate page from the CoA?"
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    If sAnswer = vbNo Then                       'IF NO THEN THIS HAPPENS

        MsgBox "Page count will be reduced by only one."
    
        sActualQuantity = sActualQuantity - 1
        sQuestion = "This transcript came to " & sActualQuantity & " billable pages.  Is that page count correct?"
        sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
    
        If sAnswer = vbNo Then                   'IF NO THEN THIS HAPPENS
    
            sActualQuantity = InputBox("How many billable pages was this transcript?")
        
        Else                                     'if yes then this happens
        End If
    
    Else                                         'if yes then this happens

        MsgBox "Page count will be reduced by two."
    
        sActualQuantity = sActualQuantity - 2
        sQuestion = "This transcript came to " & sActualQuantity & " billable pages.  Is that page count correct?"
        sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
    
        If sAnswer = vbNo Then                   'IF NO THEN THIS HAPPENS
    
            sActualQuantity = InputBox("How many billable pages was this transcript?")
        
        Else                                     'if yes then this happens
    
            sActualQuantity = InputBox("How many billable pages was this transcript?")
            
        End If
    
    End If

    oAcrobatDoc.Close



    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

    sSQL = "UPDATE [CourtDates] SET [CourtDates].[ActualQuantity] = " & sActualQuantity & " WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"

    '@Ignore AssignmentNotUsed
    Set qdf = CurrentDb.CreateQueryDef("", sSQL)
    CurrentDb.Execute sSQL

    Set qdf = Nothing

    DoCmd.OpenQuery "FinalUnitPriceQuery"        'PRE-QUERY FOR FINAL SUBTOTAL
    CurrentDb.Execute "INVUpdateFinalUnitPriceQuery" 'UPDATES FINAL SUBTOTAL
    CurrentDb.Close
End Sub

Public Sub pfReadXML()
    '============================================================================
    ' Name        : pfReadXML
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfReadXML
    ' Description : reads shipping XML and sends "Shipped" email to client
    '============================================================================

    Dim sTrackingNumber As String
    Dim dShipDate As Date
    Dim dShipDateFormatted As Date
    Dim rstCurrentJob As DAO.Recordset
    Dim formDOM As DOMDocument60                 'Currently opened xml file
    Dim ixmlRoot As IXMLDOMElement
    Dim Rng As Range
    Dim cJob As Job
    Set cJob = New Job

    Do While Len(Dir(cJob.DocPath.ShippingOutputFolder)) > 0
    
        Set formDOM = New DOMDocument60          'Open the xml file
        formDOM.resolveExternals = False         'using schema yes/no true/false
        formDOM.validateOnParse = False          'Parser validate document?  Still parses well-formed XML
        formDOM.Load (cJob.DocPath.ShippingOutputFolder & Dir(cJob.DocPath.ShippingOutputFolder))
    
        Set ixmlRoot = formDOM.DocumentElement   'Get document reference
    
        sCourtDatesID = ixmlRoot.SelectSingleNode("//DAZzle/Package/ReferenceID").Text
        dShipDate = ixmlRoot.SelectSingleNode("//DAZzle/Package/PostmarkDate").Text
        dShipDateFormatted = DateSerial(Left(dShipDate, 4), Mid(dShipDate, 5, 2), Right(dShipDate, 2))
        sTrackingNumber = ixmlRoot.SelectSingleNode("//DAZzle/Package/PIC").Text
    
        Set rstCurrentJob = CurrentDb.OpenRecordset("SELECT * FROM CourtDates WHERE ID = " & sCourtDatesID & ";")
    
        rstCurrentJob.Edit
            rstCurrentJob.Fields("ShipDate").Value = dShipDateFormatted
            rstCurrentJob.Fields("TrackingNumber").Value = sTrackingNumber
        rstCurrentJob.Update
        
        Name cJob.DocPath.ShippingOutputFolder & Dir(cJob.DocPath.ShippingOutputFolder) As cJob.DocPath.ShippingFolder & "done" & Dir(cJob.DocPath.ShippingOutputFolder) 'move file to other folder
    
        Call pfSendWordDocAsEmail("Shipped", "Transcript Shipped")
       
    Loop

End Sub

Public Sub pfFileRenamePrompt()
    '============================================================================
    ' Name        : pfFileRenamePrompt
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfFileRenamePrompt
    ' Description : renames transcript to specified name, mainly for contractors
    '============================================================================

    Dim sUserInput As String
    Dim sChkBxFiledNotFiled As String
    
    Dim cJob As Job
    Set cJob = New Job

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    
    sUserInput = InputBox("Enter the desired document name without the extension" & Chr(13) & "Weber Format:  A169195_transcript_2018-09-18_IngramEricaL" & Chr(13) & _
    "AMOR Format: Audio Name" & Chr(13) & "eScribers format [JobNumber]_[DRAFT]_Date", "Rename your document." & Chr(13) & "Weber Format:  A169195_transcript_2018-09-18_IngramEricaL" _
    & Chr(13) & "AMOR Format: Audio Name" & Chr(13) & "eScribers format [JobNumber]_[DRAFT]_Date", "Enter the new name for the transcript here, without the extension." & Chr(13) & _
    "Weber Format:  A169195_transcript_2018-09-18_IngramEricaL" & Chr(13) & "AMOR Format: Audio Name" & Chr(13) & "eScribers format [JobNumber]_[DRAFT]_Date")

    If sUserInput = "Enter the new name for the transcript here, without the extension." Or sUserInput = "" Then
        Exit Sub
    End If

    FileCopy cJob.DocPath.CourtCover, cJob.DocPath.TranscriptFD
    Name cJob.DocPath.TranscriptFD As cJob.DocPath.JobDirectoryT & sUserInput & ".docx"

    MsgBox "File renamed to " & sClientTranscriptName & ".  Next we will deliver the transcript."

    Call pfGenericExportandMailMerge("Case", "Stage4s\ContractorTranscriptsReady")
    Call pfSendWordDocAsEmail("ContractorTranscriptsReady", "Transcripts Ready", sClientTranscriptName)

    sChkBxFiledNotFiled = "UPDATE [CourtDates] SET FiledNotFiled =(Yes) WHERE ID=" & sCourtDatesID & ";"

    CurrentDb.Execute sChkBxFiledNotFiled

    MsgBox "Transcript has been delivered.  Next, let's do some admin stuff."

End Sub

Public Sub pfWaitSeconds(iSeconds As Long)
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

    Do                                           'yield to other programs
        pfDelay 100
        DoEvents
    Loop Until Now >= dCurrentTime

lEXIT:
    Exit Sub

lERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modDateTime.WaitSeconds"
    Resume lEXIT

End Sub

Public Sub pfAvailabilitySchedule()
    '============================================================================
    ' Name        : pfAvailabilitySchedule
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfAvailabilitySchedule
    ' Description : opens availability calculator
    '============================================================================

    DoCmd.OpenForm (Forms![SBFM-Availability])
End Sub

Public Sub pfCheckFolderExistence()
    '============================================================================
    ' Name        : pfCheckFolderExistence
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfCheckFolderExistence
    ' Description : checks for Audio, Transcripts, FTP, WorkingFiles, Notes subfolders and RoughDraft and creates if not exists
    '============================================================================

    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID
    
    If Len(Dir(cJob.DocPath.JobDirectory, vbDirectory)) = 0 Then
        MkDir cJob.DocPath.JobDirectory
    End If
    If Len(Dir(cJob.DocPath.JobDirectoryF, vbDirectory)) = 0 Then
        MkDir cJob.DocPath.JobDirectoryF
    End If
    If Len(Dir(cJob.DocPath.JobDirectoryW, vbDirectory)) = 0 Then
        MkDir cJob.DocPath.JobDirectoryW
    End If
    If Len(Dir(cJob.DocPath.JobDirectoryA, vbDirectory)) = 0 Then
        MkDir cJob.DocPath.JobDirectoryA
    End If
    If Len(Dir(cJob.DocPath.JobDirectoryT, vbDirectory)) = 0 Then
        MkDir cJob.DocPath.JobDirectoryT
    End If
    If Len(Dir(cJob.DocPath.JobDirectoryN, vbDirectory)) = 0 Then
        MkDir cJob.DocPath.JobDirectoryN
    End If
    If Len(Dir(cJob.DocPath.JobDirectoryG, vbDirectory)) = 0 Then
        MkDir cJob.DocPath.JobDirectoryG
    End If
    If Len(Dir(cJob.DocPath.JobDirectoryB, vbDirectory)) = 0 Then
        MkDir cJob.DocPath.JobDirectoryB
    End If
    
    If sJurisdiction Like "*Food and Drug Administration*" Then

        If Len(Dir(cJob.DocPath.RoughDraft)) = 0 Then
            FileCopy cJob.DocPath.TemplateFolder2 & "RoughDraft-FDA.docx", cJob.DocPath.RoughDraft
        End If
    
    ElseIf sJurisdiction Like "*NonCourt*" Then

        If Len(Dir(cJob.DocPath.RoughDraft)) = 0 Then
            sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
            cJob.FindFirst "ID=" & sCourtDatesID
            Debug.Print cJob.DocPath.TempShipOptionsQ1XLSM
            
            FileCopy cJob.DocPath.TemplateFolder2 & "RoughDraft.docx", cJob.DocPath.RoughDraft
        End If
    
    ElseIf sJurisdiction Like "*FDA*" Then

        If Len(Dir(cJob.DocPath.RoughDraft)) = 0 Then
            FileCopy cJob.DocPath.TemplateFolder2 & "RoughDraft-FDA.docx", cJob.DocPath.RoughDraft
        End If
    
    ElseIf sJurisdiction Like "*Weber Oregon*" Then

        If Len(Dir(cJob.DocPath.RoughDraft)) = 0 Then
            FileCopy cJob.DocPath.TemplateFolder2 & "RoughDraft-WeberOR.docx", cJob.DocPath.RoughDraft
        End If
    
    ElseIf sJurisdiction Like "*Weber Nevada*" Then

        If Len(Dir(cJob.DocPath.RoughDraft)) = 0 Then
            FileCopy cJob.DocPath.TemplateFolder2 & "RoughDraft-WeberNV.docx", cJob.DocPath.RoughDraft
        End If
    
    ElseIf sJurisdiction Like "*Weber Bankruptcy*" Then

        If Len(Dir(cJob.DocPath.RoughDraft)) = 0 Then
            FileCopy cJob.DocPath.TemplateFolder2 & "RoughDraft.docx", cJob.DocPath.RoughDraft
        End If
    
    ElseIf sJurisdiction Like "*AVT*" Then

        If Len(Dir(cJob.DocPath.RoughDraft)) = 0 Then
            FileCopy cJob.DocPath.TemplateFolder2 & "RoughDraft.docx", cJob.DocPath.RoughDraft
        End If
    
    Else

        If Len(Dir(cJob.DocPath.RoughDraft)) = 0 Then
            FileCopy cJob.DocPath.TemplateFolder2 & "RoughDraft.docx", cJob.DocPath.RoughDraft
        End If
    
    End If
    Call pfClearGlobals
    Set cJob = Nothing
End Sub

Public Sub pfCommunicationHistoryAdd(sCHTopic As String)
    '============================================================================
    ' Name        : pfCommunicationHistoryAdd
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfCommunicationHistoryAdd(sCHTopic)
    ' Description : adds entry to CommunicationHistory
    '============================================================================

    Dim rstCHAdd As DAO.Recordset
    Dim sCHHyperlink As String

    Dim cJob As Job
    Set cJob = New Job
    
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    
    sCHHyperlink = sCourtDatesID & "-" & sCHTopic & "#" & cJob.DocPath.JobDirectoryGN & sCHTopic & ".docx" & "#"

    Set rstCHAdd = CurrentDb.OpenRecordset("CommunicationHistory")

    rstCHAdd.AddNew
        rstCHAdd("FileHyperlink").Value = sCHHyperlink
        rstCHAdd("DateCreated").Value = Now
        rstCHAdd("CourtDatesID").Value = sCourtDatesID
    rstCHAdd.Update

    rstCHAdd.Close

End Sub

Public Sub pfStripIllegalChar(sInput As String)
    '============================================================================
    ' Name        : pfStripIllegalChar
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfStripIllegalChar(StrInput)
    ' Description : strips illegal characters from input
    '============================================================================

    Dim oRegex As Object
    Dim StripIllegalChar As Variant
 
    Set oRegex = CreateObject("vbscript.regexp")
    oRegex.Pattern = "[\" & Chr(34) & "\!\@\#\$\%\^\&\*\(\)\=\+\|\[\]\{\}\`\'\;\:\<\>\?\/\,]"
    '@Ignore AssignmentNotUsed
    oRegex.IgnoreCase = True
    '@Ignore AssignmentNotUsed
    oRegex.Global = True
    '@Ignore AssignmentNotUsed
    StripIllegalChar = oRegex.Replace(sInput, "")

    Set oRegex = Nothing
 
End Sub

Public Sub pfGetFolder(Folders As Collection, EntryID As Collection, StoreID As Collection, fld As MAPIFolder)
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
      
    Set SubFolder = Nothing

End Sub

Public Sub pfBrowseForFolder(sSavePath As String, Optional OpenAt As String)
    '============================================================================
    ' Name        : pfBrowseForFolder
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfBrowseForFolder(StrSavePath, OpenAt)
    'OpenAt optional
    ' Description : browses for folder
    '============================================================================

    Dim oShell As Object
    Dim oBrowsedFolder As Object
    
    Dim vEnvUserProfile As Variant
 
    vEnvUserProfile = CStr(Environ("USERPROFILE"))
 
    Set oShell = CreateObject("Shell.Application")
 
    Set oBrowsedFolder = oShell.BrowseForFolder(0, "Please choose a folder", 0, vEnvUserProfile & "\My Documents\")
 
    sSavePath = oBrowsedFolder.Self.Path
     
    Set oShell = Nothing
   
End Sub

Public Sub pfSingleBAScrapeSpecificBarNo(sWebSiteBarNo As String)
    '============================================================================
    ' Name        : pfSingleBAScrapeSpecificBarNo
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfSingleBAScrapeSpecificBarNo(sWebSiteBarNo)
    ' Description:  gets one bar number's info from the WA bar website (pick any range from 1 to 55000)
    '============================================================================

    Dim rstBarAddresses As DAO.Recordset

    Dim oCompanyName As Object
    Dim oBarName As Object
    Dim oBarNumber As Object
    Dim oEligibility As Object
    Dim oActiveL As Object
    Dim oAdmitDate As Object
    Dim oAddress As Object
    Dim oEmail As Object
    Dim oPhone As Object
    Dim oFax As Object
    Dim oPracticeArea As Object
    Dim oInternetE As Object

    Dim sWebsiteLink As String
    Dim sCompanyName As String
    Dim sBarName As String
    Dim sBarNumber As String
    Dim sEligibility As String
    Dim sActiveL As String
    Dim sAdmitDate As String
    Dim sAddress As String
    Dim sEmail As String
    Dim sPhone As String
    Dim sFax As String
    Dim sPracticeArea As String

    sWebsiteLink = "https://www.mywsba.org/PersonifyEbusiness/LegalDirectory/LegalProfile.aspx?Usr_ID=0000000" & sWebSiteBarNo

    Set oInternetE = CreateObject("InternetExplorer.Application")
    oInternetE.Visible = False
    oInternetE.Navigate sWebsiteLink

    While oInternetE.Busy                        'Wait while oInternetE loading...
        DoEvents
    Wend

    ' ********************************************************************** get the following info
    pfDelay 5                                    'wait a little bit
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
    Debug.Print sBarName & Chr(13) & sBarNumber & Chr(13) & sEligibility & Chr(13) & sActiveL & Chr(13) & " " & Chr(13) & _
                                                                                                                        sCompanyName & Chr(13) & sAddress & Chr(13) & " " & Chr(13) & sPhone & Chr(13) & " " & Chr(13) & sPracticeArea

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

End Sub

Public Sub pfScrapingBALoop(sWebSiteBarNo As String, sWebSiteBarNoGoal As String)
    On Error Resume Next
    '============================================================================
    ' Name        : pfScrapingBALoop
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfScrapingBALoop(sWebSiteBarNo, sWebSiteBarNoGoal)
    ' Description:  gets a range you specify of bar numbers' info from the WA bar website (pick any range from 1 to 55000)
    '============================================================================

    Dim rstBarAddresses As DAO.Recordset

    Dim oCompanyName As Object
    Dim oBarName As Object
    Dim oBarNumber As Object
    Dim oEligibility As Object
    Dim oActiveL As Object
    Dim oAdmitDate As Object
    Dim oAddress As Object
    Dim oEmail As Object
    Dim oPhone As Object
    Dim oFax As Object
    Dim oPracticeArea As Object
    Dim oInternetE As Object
    Dim sWebsiteLink As String
    Dim sCompanyName As String
    Dim sBarName As String
    Dim sBarNumber As String
    Dim sEligibility As String
    Dim sActiveL As String
    Dim sAdmitDate As String
    Dim sAddress As String
    Dim sEmail As String
    Dim sPhone As String
    Dim sFax As String
    Dim sPracticeArea As String



    Do While sWebSiteBarNo < sWebSiteBarNoGoal

        sWebsiteLink = "https://www.mywsba.org/PersonifyEbusiness/LegalDirectory/LegalProfile.aspx?Usr_ID=0000000" & sWebSiteBarNo
    
        Set oInternetE = CreateObject("InternetExplorer.Application")
        oInternetE.Visible = False
        oInternetE.Navigate sWebsiteLink
    
        While oInternetE.Busy                    ' Wait while oInternetE loading.
            DoEvents
        Wend
    
        ' ********************************************************************** get the following info
        pfDelay 5                                'wait a little bit
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
        Debug.Print sBarName & Chr(13) & sBarNumber & Chr(13) & sEligibility & Chr(13) & sActiveL & Chr(13) & " " & Chr(13) & _
                                                                                                                            sCompanyName & Chr(13) & sAddress & Chr(13) & " " & Chr(13) & sPhone & Chr(13) & " " & Chr(13) & sPracticeArea
    
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
    
        sWebSiteBarNo = sWebSiteBarNo + 1        'move on to next bar number
    
        pfDelay 22
        
    Loop
    On Error GoTo 0
End Sub

Public Sub pfReformatTable()
    '============================================================================
    ' Name        : pfReformatTable
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfReformatTable
    ' Description : reformats scraped Bar addresses to useable format for table
    '============================================================================
    'TODO: pfReformatTable check what's going on here to finish it if necessary
    
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
    Dim i           As Long

    Set db = CurrentDb

    ' Select all eligible fields (have a comma) and unprocessed (Field2 is Null)
    strSQL = "SELECT Address, Field2 FROM BarAddresses WHERE ([Address] Like ""*<br>*"") AND ([Field2] Is Null)"

    Set rsADD = db.OpenRecordset("BarAddresses", dbOpenDynaset, dbAppendOnly)

    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset)
    With rs
        While Not .EOF
            strField1 = !Field1
            varData = Split(strField1, "<br>")   ' Get all comma delimited fields

            ' Update First Record
            .Edit
            !Field2 = Trim(varData(0))       ' remove spaces before writing new fields
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

End Sub

Public Sub pfDelay(lSeconds As Long)
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
End Sub

Public Sub pfPriorityPointsAlgorithm()
    '============================================================================
    ' Name        : pfPriorityPointsAlgorithm
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfPriorityPointsAlgorithm
    ' Description:  assigns priority points to various tasks in Tasks table and inserts it into the PriorityPoints field
    '               priority scale 1 to 100
    '============================================================================

    Dim iPriorityPoints As Long
    Dim iTimeLength As Long
    Dim sPriority As String
    Dim sCategory As String
    Dim bCompleted As Boolean
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
            sPriority = rstTasks.Fields("Priority").Value
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

End Sub


Private Sub AddTaskToTasks(sTaskTitle As String, _
                     iTaskMinuteLength As Long, _
                     sPriority As String, _
                     dDue As Date, _
                     sTaskCategory As String, _
                     sTaskDescription As String, _
                     Optional dStart As Date)
    Dim rstTasks As DAO.Recordset
    
    Set rstTasks = CurrentDb.OpenRecordset("Tasks")

    rstTasks.AddNew
    rstTasks.Fields("Title").Value = sTaskTitle
    rstTasks.Fields("TimeLength").Value = iTaskMinuteLength
    rstTasks.Fields("Priority").Value = sPriority
    rstTasks.Fields("Start Date").Value = dStart
    rstTasks.Fields("Category").Value = sTaskCategory
    rstTasks.Fields("Description").Value = sTaskDescription
        
    rstTasks.Fields("Due Date").Value = dDue
    
    rstTasks.Update
    
    rstTasks.Close

End Sub

Public Sub pfGenerateJobTasks()
    '============================================================================
    ' Name        : pfGenerateJobTasks
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfGenerateJobTasks
    ' Description : generates job tasks in the Tasks table
    '============================================================================
    
    Dim sTaskTitle As String
    Dim sTaskCategory As String
    Dim sPriority As String
    Dim sTaskDescription As String
    
    Dim iTypingTime As Long
    Dim iAudioProofTime As Long
    Dim iTaskMinuteLength As Long
    
    Dim dStart As Date
    Dim dDue As Date
    
    Dim cJob As Job
    Set cJob = New Job

    Call pfCurrentCaseInfo                       'refresh transcript info

    sTaskTitle = "(1.1) Enter job & contacts into database:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
    dDue = Date + 1
    dStart = Date
    cJob.DueDate = "12/22/2019"
    iTaskMinuteLength = "2"
    sTaskCategory = "production"
    sPriority = "(1) Stage 1"
    sTaskDescription = "|Case Name:  " & sParty1 & " v. " & sParty2 & "   |" & Chr(13) & _
                       "|Case Nos.:  " & sCaseNumber1 & "   |   " & sCaseNumber2 & "   |" & Chr(13) & _
                       "|Due Date:  " & dDue & "   |   Turnaround:  " & sTurnaroundTime & " calendar days   |" & _
                       "|Client:   " & sCompany & "   |   Folder:   " & cJob.DocPath.JobDirectory & "   |" & Chr(13) & _
                       "|Exp. Advance/Deposit Date:  " & dExpectedAdvanceDate & "   |" & Chr(13) & _
                       "|Exp. Rebate Date:  " & dExpectedRebateDate & "   |" & Chr(13) & _
                       "|Estimate:  " & sSubtotal & "   |"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription, dStart)

    sTaskTitle = "(1.2) Payment:  If factored, proceed with set-up.  If not, send invoice & wait for payment :  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
    dDue = Date + 1
    iTaskMinuteLength = "2"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription, dStart)

    sTaskTitle = "(1.3) Generate documents: cover, autocorrect, AGshortcuts, Xero CSV, CD label, transcripts ready, package-enclosed letter:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
    dDue = Date + 1
    iTaskMinuteLength = "2"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription, dStart)

    iTypingTime = Round((((sAudioLength * 3) / 60) + 1), 0)

    For i = 1 To iTypingTime
        sTaskTitle = "(2.1) Type:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
        dDue = cJob.DueDate - 3
        iTaskMinuteLength = "60"
        sPriority = "(2) Stage 2"
        Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription, dStart)
    
    Next i

    sTaskTitle = "(3.1) Find/replace add to cover page:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
    dDue = cJob.DueDate - 2
    iTaskMinuteLength = "3"
    sPriority = "(3) Stage 3"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription, dStart)

    sTaskTitle = "(3.2) Hyperlink:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
    dDue = cJob.DueDate - 2
    iTaskMinuteLength = "15"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription, dStart)

    sTaskTitle = "(3.3) Send email if more info needed and hold transcript:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
    dDue = cJob.DueDate - 2
    iTaskMinuteLength = "2"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription, dStart)

    iAudioProofTime = Round((((sAudioLength * 1.5) / 60) + 1), 0)
    
    For i = 1 To iAudioProofTime
    
        sTaskTitle = "(3.4) Audio-proof:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
        dDue = cJob.DueDate - 2
        iTaskMinuteLength = "60"
        Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription, dStart)
        
    Next i

    sTaskTitle = "(4.1) Make final transcript docs, pdf, zip, etc:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
    dDue = cJob.DueDate - 1
    iTaskMinuteLength = "3"
    sPriority = "(4) Stage 4"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription, dStart)
    
    sTaskTitle = "(4.2) Invoice if balance due or factored.  Refund if applicable:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
    dDue = cJob.DueDate - 1
    iTaskMinuteLength = "1"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription, dStart)

    sTaskTitle = "(4.3) Deliver as necessary electronically if transcript not held:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
    dDue = cJob.DueDate - 1
    iTaskMinuteLength = "1"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription, dStart)

    sTaskTitle = "(4.4) File transcript:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
    dDue = cJob.DueDate - 1
    iTaskMinuteLength = "3"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription, dStart)

    'sTaskTitle = "(4.5) Send invoice to factoring:  " & sCourtDatesID & ", Approx. " & sAudioLength & " mins"
    'dDue = cJob.DueDate - 1
    'iTaskMinuteLength = "1"
    'Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription, dStart)
    

    Call pfClearGlobals
    
End Sub

Public Sub pfDailyTaskAddFunction()
    '============================================================================
    ' Name        : pfDailyTaskAddFunction
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfDailyTaskAddFunction
    ' Description : adds static daily tasks to Tasks table
    '============================================================================
    
    Dim sTaskTitle As String
    Dim sTaskCategory As String
    Dim sPriority As String
    Dim sTaskDescription As String
    Dim dDue As Date
    Dim iTaskMinuteLength As Long

    sTaskCategory = "GTD Daily"
    dDue = Now + 1
    sPriority = "(1) Stage 1"
    sTaskDescription = "none"
    iTaskMinuteLength = "2"
    sTaskTitle = "List action items, projects, waiting-fors, calendar events, someday/maybes as appropriate"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)


    iTaskMinuteLength = "2"
    sTaskTitle = "replied to all e-mails, checked & processed all voicemails"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)
    
    iTaskMinuteLength = "2"
    sTaskTitle = "export e-mails and check communication"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)


    iTaskMinuteLength = "10"
    sTaskTitle = "review jobs sent to me"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)


    iTaskMinuteLength = "10"
    sTaskTitle = "review tasks bin and process"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)


    sTaskCategory = "personal"
    iTaskMinuteLength = "60"
    sTaskTitle = "art time"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)


    iTaskMinuteLength = "60"
    sTaskTitle = "yoga"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)

End Sub

Public Sub pfWeeklyTaskAddFunction()
    '============================================================================
    ' Name        : pfWeeklyTaskAddFunction
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfWeeklyTaskAddFunction
    ' Description : adds static weekly tasks to Tasks table
    '============================================================================
    
    Dim sTaskTitle As String
    Dim sTaskCategory As String
    Dim sPriority As String
    Dim sTaskDescription As String
    
    Dim vStartDate As Date
    Dim dDue As Date
    
    Dim iTaskMinuteLength As Long

    sTaskCategory = "GTD Weekly"
    dDue = Now + 5
    sPriority = "(2) Stage 2"
    sTaskDescription = "none"
    iTaskMinuteLength = "5"
    sTaskTitle = "empty head about uncaptured new items"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)

    iTaskMinuteLength = "5"
    sTaskTitle = "file material away"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)


    iTaskMinuteLength = "5"
    sTaskTitle = "stage R/R material"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)

    iTaskMinuteLength = "10"
    sTaskTitle = "update payment bill calendar"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)


    iTaskMinuteLength = "10"
    sTaskTitle = "budget"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)


    iTaskMinuteLength = "10"
    sTaskTitle = "review events coming up"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)

    iTaskMinuteLength = "10"
    sTaskTitle = "review lists"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)

    iTaskMinuteLength = "10"
    sTaskTitle = "review long-term projects"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)

    iTaskMinuteLength = "10"
    sTaskTitle = "review sales reports"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)

    sTaskCategory = "business admin"
    iTaskMinuteLength = "60"
    sTaskTitle = "update AQC manual"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)

    iTaskMinuteLength = "60"
    sTaskTitle = "do 1 hour government contracts or business/marketing plan work"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)

    sTaskCategory = "personal"
    iTaskMinuteLength = "20"
    sTaskTitle = "vacuum"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)

    iTaskMinuteLength = "20"
    sTaskTitle = "groceries"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)

    iTaskMinuteLength = "30"
    sTaskTitle = "laundry"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)

    iTaskMinuteLength = "60"
    sTaskTitle = "Clean house"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)

End Sub

Public Sub pfMonthlyTaskAddFunction()
    '============================================================================
    ' Name        : pfMonthlyTaskAddFunction
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfMonthlyTaskAddFunction
    ' Description : adds static monthly tasks to Tasks table
    '============================================================================
    'TODO: pfMonthlyTaskAddFunction can probably break this into separate functions
    Dim sTaskTitle As String
    Dim sTaskCategory As String
    Dim sPriority As String
    Dim sTaskDescription As String
    
    Dim iTaskMinuteLength As Long
    
    Dim dDue As Date

    sTaskCategory = "GTD Monthly"
    dDue = Now + 20
    sPriority = "(1) Stage 1"
    sTaskDescription = "none"
    iTaskMinuteLength = "15"
    sTaskTitle = "Brainstorm Creative Ideas"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)

    iTaskMinuteLength = "15"
    sTaskTitle = "Review 1 to 2 Year Goals"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)

    iTaskMinuteLength = "15"
    sTaskTitle = "Review Roles and Current Responsibilities"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)
    
    iTaskMinuteLength = "15"
    sTaskTitle = "Review Someday or Maybe list"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)

    iTaskMinuteLength = "15"
    sTaskTitle = "Review Support Files"
    Call AddTaskToTasks(sTaskTitle, iTaskMinuteLength, sPriority, dDue, sTaskCategory, sTaskDescription)

End Sub

Public Sub pfCommHistoryExportSub()
    On Error Resume Next
    '============================================================================
    ' Name        : pfCommHistoryExportSub
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfCommHistoryExportSub
    ' Description : exports emails to CommunicationsHistory table
    '============================================================================

    
    Dim sEmailReceivedTime As String
    Dim sDriveHyperlink As String
    Dim sSenderName As String
    Dim sCommHistoryHyperlink As String

    Dim nsOutlookNmSpc As Outlook.Namespace
    Dim oOutLookMAPIFolder As Outlook.MAPIFolder
    Dim oOutlookMail As Outlook.MailItem
    Dim oOutlookAccessTestFolder As Object
    
    Dim rs As DAO.Recordset
    
    Dim dEmailReceived As Date
    
    Dim cJob As Job
    Set cJob = New Job

    Set nsOutlookNmSpc = GetNamespace("MAPI")
    Set oOutlookAccessTestFolder = nsOutlookNmSpc.Folders(sCompanyEmail).Folders("Inbox").Folders("AccessTest")

    For Each oOutlookMail In oOutlookAccessTestFolder.Items
     
        dEmailReceived = oOutlookMail.ReceivedTime 'assign received time to variable
        sSenderName = oOutlookMail.SenderName
        sEmailReceivedTime = Format(dEmailReceived, "YYYYMMDD-hhmm") 'convert time to good string value
    
        'save email on hard drive in in progress
        oOutlookMail.SaveAs cJob.DocPath.EmailDirectory & sEmailReceivedTime & "-" & sSenderName & "-Email.msg", 3
        dEmailReceived = oOutlookMail.ReceivedTime
        sEmailReceivedTime = Format(dEmailReceived, "YYYYMMDD-hhmm")
        sDriveHyperlink = cJob.DocPath.EmailDirectory & sEmailReceivedTime & "-" & sSenderName & "-Email.msg"
    
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
    Set nsOutlookNmSpc = Nothing
    Set oOutlookAccessTestFolder = Nothing

    Exit Sub

eHandler:

    'MsgBox ("The email failed to save.")
    'MsgBox Err.Description & " (" & Err.Number & ")"

    Resume Next
End Sub

Public Sub pfEmailsExport1()
    '============================================================================
    ' Name        : pfEmailsExport1
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfEmailsExport1
    ' Description : export specified fields from each mail / item in selected folder
    '============================================================================

    Dim nsOutlookNmSpc As Outlook.Namespace
    Dim iItemCounter As Long
    Dim oOutLookMAPIFolder As Outlook.MAPIFolder
    Dim rstEmails As DAO.Recordset

    Set nsOutlookNmSpc = GetNamespace("MAPI")
    'Set oOutLookMAPIFolder = nsOutlookNmSpc.PickFolder
    Set oOutLookMAPIFolder = nsOutlookNmSpc.Folders(sCompanyEmail).Folders("Inbox").Folders("AccessTest")
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

End Sub

Public Sub pfMoveSelectedMessages()
    '============================================================================
    ' Name        : pfMoveSelectedMessages
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfMoveSelectedMessages
    ' Description : move selected messages to network drive
    '============================================================================

    Dim sSenderName As String
    
    Dim oOutlookApp As Outlook.Application
    Dim nsOutlookNmSpc As Outlook.Namespace
    Dim oDestinationFolder As Outlook.MAPIFolder
    Dim oSourceFolder As Outlook.Folder
    
    Dim oCurrentExplorer As Explorer
    Dim oSelection As Selection
    
    Dim oSubSelection As Object
    
    Dim vObjectVariant As Variant
    
    Dim lMovedItems As Long
    Dim iDateDifference As Long

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
    On Error GoTo 0

    'Display the number of items that were moved.
    MsgBox "Moved " & lMovedItems & " messages(s)."

    Set oCurrentExplorer = Nothing
    Set oSubSelection = Nothing
    Set oSelection = Nothing
    Set oOutlookApp = Nothing
    Set nsOutlookNmSpc = Nothing
    Set oSourceFolder = Nothing
End Sub

Public Sub pfAskforAudio()
    '============================================================================
    ' Name        : pfAskforAudio
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfAskforAudio
    ' Description : prompts to select audio to copy to job no.'s audio folder
    '============================================================================

    Dim sFileName As String
    
    Dim iFileChosen As Long
    Dim i As Long
    
    Dim fs As Object
    Dim fd As FileDialog
    
    Dim cJob As Job
    Set cJob = New Job

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    'use the standard title and filters, but change the initial folder
    'TODO: Change drive
    fd.InitialFileName = "T:\"
    fd.InitialView = msoFileDialogViewList
    fd.Title = "Select the audio for this transcript."
    fd.AllowMultiSelect = True                   'allow multiple file selection

    iFileChosen = fd.Show
    If iFileChosen = -1 Then
        For i = 1 To fd.SelectedItems.Count      'open each of the files chosen
            sFileName = Right$(fd.SelectedItems(i), Len(fd.SelectedItems(i)) - InStrRev(fd.SelectedItems(i), "\"))
        
            If Len(Dir(cJob.DocPath.JobDirectoryA & sFileName, vbDirectory)) = 0 Then
                FileCopy fd.SelectedItems(i), cJob.DocPath.JobDirectoryA & sFileName
            End If
        
        Next i
    
    End If

End Sub

Public Sub pfAskforNotes()
    '============================================================================
    ' Name        : pfAskforNotes
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfAskforNotes
    ' Description : prompts to select notes to copy to job no.'s audio folder
    '               attempts to pdf selection
    '               copies pdf and original notes selected file to job no. folder
    '============================================================================


    Dim sFileName As String
    Dim sOriginalNotesPath As String
    Dim sAnswerPDFPrompt As String
    Dim sMakePDFPrompt As String
    
    Dim i As Long
    Dim iFileChosen As Long
    
    Dim fs As Object
    Dim oVBComponent As Object
    Dim fd As FileDialog
    
    Dim oWordApp As Word.Application
    Dim oWordDoc As Word.Document
    
    Dim rngStory As Range
    
    Dim cJob As Job
    Set cJob = New Job

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    'use the standard title and filters, but change the initial folder
    fd.InitialFileName = cJob.DocPath.sDrive & ":\"
    fd.InitialView = msoFileDialogViewList
    fd.Title = "Select your Notes."
    fd.AllowMultiSelect = True                   'allow multiple file selection

    iFileChosen = fd.Show
    If iFileChosen = -1 Then
        For i = 1 To fd.SelectedItems.Count      'copy each of the files chosen
            sFileName = Right$(fd.SelectedItems(i), Len(fd.SelectedItems(i)) - InStrRev(fd.SelectedItems(i), "\"))
        
            If Len(Dir(cJob.DocPath.JobDirectoryA & sFileName, vbDirectory)) = 0 Then
        
        
                'open in word and save as notes pdf
        
                sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
                
                sOriginalNotesPath = fd.SelectedItems(i)
                sMakePDFPrompt = "Next we will make a PDF copy.  Click yes when ready."
                sAnswerPDFPrompt = MsgBox(sMakePDFPrompt, vbQuestion + vbYesNo, "???")
        
                If sAnswerPDFPrompt = vbNo Then  'Code for No
            
                    MsgBox "No PDF copy of the notes will be made."
            
                Else                             'Code for yes
        
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
                    oWordDoc.ExportAsFixedFormat outputFileName:=cJob.DocPath.Notes, ExportFormat:=wdExportFormatPDF, CreateBookmarks:=wdExportCreateHeadingBookmarks
                        
                    oWordDoc.Close SaveChanges:=False
        
            
                End If
                    
                Set oWordDoc = Nothing
                Set oWordApp = Nothing
            
                FileCopy fd.SelectedItems(i), cJob.DocPath.JobDirectoryN & sFileName
                FileCopy fd.SelectedItems(i), cJob.DocPath.Notes
            End If
        Next i
    End If
End Sub

Public Sub pfRCWRuleScraper()
    On Error Resume Next
    '============================================================================
    ' Name        : pfRCWRuleScraper
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfRCWRuleScraper()
    ' Description:  scrapes all RCWs from WA site
    '============================================================================

    Dim sFindCitation As String
    Dim sLongCitation As String
    Dim sRuleNumber As String
    Dim sWebAddress As String
    Dim sReplaceHyperlink As String
    Dim sCurrentRule As String
    Dim sChapterNumber As String
    Dim sSubchapterNumber As String
    Dim sSubtitleNumber As String
    Dim sSectionNumber As String
    Dim Title As String
    Dim sTitle2 As String
    Dim sTitle1 As String

    Dim objHttp As Object
    Dim vLettersArray1() As Variant
    Dim vLettersArray2() As Variant
    Dim rstCitationHyperlinks As DAO.Recordset
    Dim iErrorNum As Long
    Dim sCHCategory As Long

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim m As Long
    Dim w As Long
    Dim x As Long
    Dim y As Long
    Dim z As Long

    For x = 1 To 91                              '(RCW first portion x.###.###) '1-91


        For y = 1 To 999                         '(RCW second portion ###.y.###) '1-999
    
        
            If y < 10 Then y = str("0" & y)
                
    
            For z = 10 To 990 Step 10            '(RCW third portion ###.###.z) '10 to 990 by 10s
        
                If z < 100 Then z = str("0" & z)
            
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

    For w = 1 To UBound(vLettersArray1)          '(RCW first portion w.###.###)

        sTitle1 = vLettersArray1(w)
    
        For x = 1 To 999                         '(RCW second portion ###.x.###)
    
            If x < 10 Then x = str("0" & x)
    
            '1-999 plus A, B, C
    
            vLettersArray2 = Array("A", "B", "C")
        
            For y = 0 To UBound(vLettersArray2)  '(RCW second portion w.x[y].z)
            
                sTitle2 = x & vLettersArray2(y)
        
                If y < 10 Then sTitle2 = str("0" & sTitle2)
    
                For z = 10 To 990 Step 10        '(RCW third portion ###.###.z)
            
                    If z < 100 Then y = str("0" & z)
                
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

    On Error GoTo 0


End Sub

Public Sub pfUSCRuleScraper()
    On Error Resume Next
    '============================================================================
    ' Name        : pfRuleScraper
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfRuleScraper()
    ' Description:  gets non-section usc code from site
    '============================================================================

    Dim rstCitationHyperlinks As DAO.Recordset
    Dim iErrorNum As Long
    Dim sCHCategory As Long
    Dim sFindCitation As String
    Dim sLongCitation As String
    Dim sRuleNumber As String
    Dim sWebAddress As String
    Dim sReplaceHyperlink As String
    Dim sCurrentRule As String
    Dim vRuleNumbers() As Variant
    Dim vRules() As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim m As Long
    Dim w As Long
    Dim x As Long
    Dim y As Long
    Dim z As Long

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
            iErrorNum = Err                      'Save error number
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
                iErrorNum = Err                  'Save error number
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
            iErrorNum = Err                      'Save error number
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
            iErrorNum = Err                      'Save error number
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

End Sub

Public Function GetLevel(strItem As String) As Long
    ' Return the heading level of a header from the
    ' array returned by Word.

    ' The number of leading spaces indicates the
    ' outline level (2 spaces per level: H1 has
    ' 0 spaces, H2 has 2 spaces, H3 has 4 spaces.

    Dim strTemp As String
    Dim strOriginal As String
    Dim intDiff As Long

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

Public Sub fWLGenerateJSONInfo()
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

    sWLListID = 370524335                        'ingram household
    'or try inbox 370231796
    lAssigneeID = 88345676                       'erica / 86846933 adam
    bCompleted = "false"
    bStarred = "false"


End Sub

Public Sub fWunderlistAdd(sTitle As String, sDueDate As String)
    '============================================================================
    ' Name        : fWunderlistAdd
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fWunderlistAdd()
    ' Description : adds task to Wunderlist
    '============================================================================

    Dim sURL As String
    Dim vInvoiceID As String
    Dim apiWaxLRS As String
    Dim json1 As String
    
    Dim oRequest As Object
    Dim Json As Object
    Dim oWebBrowser As Object
    Dim vDetails As Object
    
    Dim rstRates As DAO.Recordset

    Dim resp As Variant
    Dim response As Variant
    Dim rep As Variant
    
    Dim parsed As Dictionary
    '{
    '  "list_id": 12345,
    '  "title": "Hallo",
    '  "assignee_id": 123,
    '  "completed": true,
    '  "due_date": "2013-08-30",
    '  "starred": false
    '}

    Call fWLGenerateJSONInfo

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
    'Debug.Print "JSON1--------------------------------------------"
    'Debug.Print json1

    'Debug.Print "RESPONSETEXT--------------------------------------------"
    sURL = "https://a.wunderlist.com/api/v1/tasks" '?completed=False" & bCompleted  '?list_id=" & sWLListID & '"&?title=" & sTitle &
    '"&?assignee_id=" & lAssigneeID & "&?completed=" & bCompleted & "&?due_date=" & sDueDate
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        '.Visible = True
        .Open "POST", sURL, False
        .setRequestHeader "X-Access-Token", Environ("apiWunderlistT")
        .setRequestHeader "X-Client-ID", Environ("apiWunderlistUN")
        .setRequestHeader "Content-Type", "application/json"
        .send json1
        apiWaxLRS = .responseText
        'Debug.Print apiWaxLRS
        'Debug.Print "--------------------------------------------"
        'Debug.Print "Status:  " & .Status
        'Debug.Print "--------------------------------------------"
        'Debug.Print "StatusText:  " & .StatusText
        'Debug.Print "--------------------------------------------"
        'Debug.Print "ResponseBody:  " & .responseBody
        'Debug.Print "--------------------------------------------"
        .abort
    End With
    'Next
    'Debug.Print "--------------------------------------------"
    'Debug.Print "Task Title:  " & sTitle & "   |   List ID:  " & sWLListID & " " & lAssigneeID
    'Debug.Print "Completed:  " & bCompleted & "   |   Due Date:  " & sDueDate
    'Debug.Print "--------------------------------------------"
    'lAssigneeID As Long, sDueDate As String, bStarred As Boolean, bCompleted As Boolean, sTitle As String, sWLListID As String

End Sub

Public Sub fWunderlistGetTasksOnList()
    '============================================================================
    ' Name        : fWunderlistGetTasksOnList
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fWunderlistGetTasksOnList()
    ' Description : gets tasks on Wunderlist list
    '============================================================================

    Dim sURL As String
    Dim vInvoiceID As String
    Dim apiWaxLRS As String
    Dim vErrorIssue As String
    Dim sInvoiceTime As String
    Dim vStatus As String
    Dim vTotal As String
    Dim json1 As String
    Dim vErrorName As String
    Dim vErrorMessage As String
    Dim vErrorILink As String
    Dim vErrorDetails As String

    Dim resp As Variant
    Dim response As Variant
    Dim rep As Variant
    
    Dim vDetails As Object
    Dim oRequest As Object
    Dim Json As Object
    Dim oWebBrowser As Object
        
    Dim parsed As Dictionary
    
    Dim rstRates As DAO.Recordset
    '{
    '  "list_id": 12345,
    '  "title": "Hallo",
    '  "assignee_id": 123,
    '  "completed": true,
    '  "due_date": "2013-08-30",
    '  "starred": false
    '}

    Call fWLGenerateJSONInfo

    '{
    '  "list_id": 12345
    '}
    'Public lAssigneeID As Long, sDueDate As String, bStarred As Boolean, bCompleted As Boolean, sTitle As String, sWLListID As String
    
    json1 = "{" & Chr(34) & "list_id" & Chr(34) & ": " & sWLListID & "}"

    'Debug.Print "JSON1--------------------------------------------"
    'Debug.Print json1
    'Debug.Print "RESPONSETEXT--------------------------------------------"
    sURL = "https://a.wunderlist.com/api/v1/tasks?list_id=" & sWLListID
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        '.Visible = True
        .Open "GET", sURL, False
        .setRequestHeader "X-Access-Token", Environ("apiWunderlistT")
        .setRequestHeader "X-Client-ID", Environ("apiWunderlistUN")
        .setRequestHeader "Content-Type", "application/json"
        .send json1
        apiWaxLRS = .responseText
        'Debug.Print apiWaxLRS
        'Debug.Print "--------------------------------------------"
        'Debug.Print .Status
        'Debug.Print .StatusText
        .abort
    End With
    
    apiWaxLRS = Left(apiWaxLRS, Len(apiWaxLRS) - 1)
    apiWaxLRS = Right(apiWaxLRS, Len(apiWaxLRS) - 1)
    apiWaxLRS = "{" & Chr(34) & "List" & Chr(34) & ":" & apiWaxLRS & "}"
    '"total_amount":{"currency":"USD","value":"3.00"},
    Set parsed = JsonConverter.ParseJson(apiWaxLRS)
    'sInvoiceNumber = Parsed("number") 'third level array
    'vInvoiceID = Parsed("id") 'third level array
    'vStatus = Parsed("status") 'third level array
    'vTotal = Parsed("total_amount")("value") 'second level array
    'vErrorName = Parsed("id") '("value") 'second level array
    'vErrorMessage = Parsed("due_date") '("value") 'second level array
    'vErrorILink = Parsed("links") '("value") 'second level array
    Set vDetails = parsed("list")                'second level array
    For Each rep In vDetails                     ' third level objects
        vErrorIssue = rep("id")
        vErrorDetails = rep("due_date")
        Debug.Print "--------------------------------------------"
        Debug.Print "Error ID:  " & vErrorIssue
        Debug.Print "Error Details:  " & vErrorDetails
        'Debug.Print "Error Info Link:  " & vErrorILink
        'Debug.Print "Error Field:  " & vErrorIssue
        'Debug.Print "Error Details:  " & vErrorDetails
        Debug.Print "--------------------------------------------"
    Next
    Debug.Print "Task Title:  " & " " & "   |   List ID:  " & sWLListID & " " & lAssigneeID
    Debug.Print "Completed:  " & bCompleted & "   |   Due Date:  " & " "
    Debug.Print "--------------------------------------------"
    'lAssigneeID As Long, sDueDate As String, bStarred As Boolean, bCompleted As Boolean, sTitle As String, sWLListID As String

End Sub

Public Sub fWunderlistGetLists()
    '============================================================================
    ' Name        : fWunderlistGetLists
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fWunderlistGetLists()
    ' Description : gets all Wunderlist lists
    '============================================================================

    Dim sURL As String
    Dim vInvoiceID As String
    Dim apiWaxLRS As String
    Dim vErrorIssue As String
    Dim sInvoiceTime As String
    Dim vStatus As String
    Dim vTotal As String
    Dim sToken As String
    Dim json1 As String
    Dim sLocal As String
    Dim vErrorName As String
    Dim vErrorMessage As String
    Dim vErrorILink As String
    Dim vErrorDetails As String
    
    Dim oRequest As Object
    Dim Json As Object
    Dim oWebBrowser As Object
    Dim vDetails As Object
    
    Dim rstRates As DAO.Recordset

    Dim resp As Variant
    Dim response As Variant
    Dim rep As Variant
    
    Dim parsed As Dictionary
    
    '{
    '  "list_id": 12345,
    '  "title": "Hallo",
    '  "assignee_id": 123,
    '  "completed": true,
    '  "due_date": "2013-08-30",
    '  "starred": false
    '}

    Call fWLGenerateJSONInfo
    'https://www.wunderlist.com/oauth/authorize?client_id=ID&redirect_uri=URL&state=RANDOM

    '@Ignore AssignmentNotUsed
    sLocal = "'urn:ietf:wg:oauth:2.0:oob','oob'" '"https://localhost/"

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
    'Debug.Print "RESPONSETEXT--------------------------------------------"
    sURL = "https://a.wunderlist.com/api/v1/lists"
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        '.Visible = True
        .Open "GET", sURL, False
        .setRequestHeader "X-Access-Token", Environ("apiWunderlistT")
        .setRequestHeader "X-Client-ID", Environ("apiWunderlistUN")
        .setRequestHeader "Content-Type", "application/json"
        .send json1
        apiWaxLRS = .responseText
        .abort
        'Debug.Print apiWaxLRS
        'Debug.Print "--------------------------------------------"
    End With
    Set parsed = JsonConverter.ParseJson(apiWaxLRS)
    vErrorName = parsed("name")                  '("value") 'second level array
    vErrorMessage = parsed("message")            '("value") 'second level array
    vErrorILink = parsed("links")                '("value") 'second level array
    '
    'Set vDetails = Parsed("details") 'second level array
    'For Each rep In vDetails ' third level objects
    '    vErrorIssue = rep("field")
    '    vErrorDetails = rep("issue")
    'Next
    'Debug.Print "--------------------------------------------"
    'Debug.Print "Error Name:  " & vErrorName
    'Debug.Print "Error Message:  " & vErrorMessage
    'Debug.Print "Error Info Link:  " & vErrorILink
    'Debug.Print "--------------------------------------------"
    'Debug.Print "Task Title:  " & " " & "   |   List ID:  " & sWLListID & " " & lAssigneeID
    'Debug.Print "Completed:  " & bCompleted & "   |   Due Date:  " & " "
    'Debug.Print "--------------------------------------------"
    'lAssigneeID As Long, sDueDate As String, bStarred As Boolean, bCompleted As Boolean, sTitle As String, sWLListID As String

End Sub

Public Sub fWunderlistGetFolders()
    '============================================================================
    ' Name        : fWunderlistGetFolders
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fWunderlistGetFolders()
    ' Description : gets list of Wunderlist folders or folder revisions
    '============================================================================

    Dim sURL As String
    Dim sResponseText As String
    Dim json1 As String
    Dim apiWaxLRS As String
    Dim vErrorIssue As String
    Dim vErrorName As String
    Dim vErrorMessage As String
    Dim vErrorILink As String
    Dim vErrorDetails As String

    Dim parsed As Dictionary

    Call fWLGenerateJSONInfo
    
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

    'Debug.Print "RESPONSETEXT--------------------------------------------"

    'sURL = "https://a.wunderlist.com/api/v1/folders"
    sURL = "https://a.wunderlist.com/api/v1/folder_revisions"
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "GET", sURL, False
        .setRequestHeader "X-Access-Token", Environ("apiWunderlistT")
        .setRequestHeader "X-Client-ID", Environ("apiWunderlistUN")
        .setRequestHeader "Content-Type", "application/json"
        .send
        apiWaxLRS = .responseText
        .abort
        'Debug.Print apiWaxLRS
        'Debug.Print "--------------------------------------------"
    End With
    apiWaxLRS = Left(apiWaxLRS, Len(apiWaxLRS) - 1)
    apiWaxLRS = Right(apiWaxLRS, Len(apiWaxLRS) - 1)
    Set parsed = JsonConverter.ParseJson(apiWaxLRS)
    vErrorName = parsed("id")                    '("value") 'second level array
    vErrorMessage = parsed("title")              '("value") 'second level array
    vErrorILink = parsed("list_ids")             '("value") 'second level array
    vErrorIssue = parsed("revision")             '("value") 'second level array

    'Debug.Print "--------------------------------------------"
    'Debug.Print "Folder ID:  " & vErrorName & "   |   " & "Folder Title:  " & vErrorMessage
    'Debug.Print "List IDs in Folder:  " & vErrorILink & "   |   " & "Revision No.:  " & vErrorIssue
    'Debug.Print "--------------------------------------------"

End Sub

Public Sub pfUSCRuleScraper1()
    On Error Resume Next

    '============================================================================
    ' Name        : pfUSCRuleScraper
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfUSCRuleScraper()
    ' Description:  validates and builds links for all us code, no front matter, no appendices
    '============================================================================
    Dim sFindCitation As String
    Dim sLongCitation As String
    Dim sRuleNumber As String
    Dim sWebAddress As String
    Dim sReplaceHyperlink As String
    Dim sCurrentRule As String
    Dim sChapterNumber As String
    Dim sSubchapterNumber As String
    Dim sSubtitleNumber As String
    Dim sSectionNumber As String
    Dim Title As String
    
    Dim objHttp As Object

    Dim vRuleNumbers() As Variant
    Dim vRules() As Variant
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim m As Long
    Dim w As Long
    Dim x As Long
    Dim y As Long
    Dim iErrorNum As Long
    Dim sCHCategory As Long
    Dim z As Long
    
    Dim rstCitationHyperlinks As DAO.Recordset
    
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

    On Error GoTo 0


End Sub

Public Sub pfRCWRuleScraper1()
    On Error Resume Next
    '============================================================================
    ' Name        : pfRCWRuleScraper
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfRCWRuleScraper()
    ' Description:  builds all RCWs and their links, checks for validation, and puts into CitationHyperlinks table
    '============================================================================
    Dim sFindCitation As String
    Dim sLongCitation As String
    Dim sRuleNumber As String
    Dim sWebAddress As String
    Dim sReplaceHyperlink As String
    Dim sCurrentRule As String
    Dim Title As String
    Dim sTitle2 As String
    Dim sTitle1 As String
    Dim sTitle As String
    Dim sTitle3 As String
    Dim sCheck As String
    Dim sTitle4 As String
    
    Dim oHTTPText As Object
    
    Dim vLettersArray1() As Variant
    Dim vLettersArray2() As Variant
    
    Dim rstCitationHyperlinks As DAO.Recordset
    
    Dim iErrorNum As Long
    Dim sCHCategory As Long
    Dim m As Long
    Dim n As Long
    Dim o As Long
    Dim p As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim w As Long
    Dim x As Long
    Dim y As Long
    Dim z As Long

    'i build a delay in mine by calling a separate function so it requests only once every 22 seconds
    'y = 01 to 385 ; exception 28B, 43, 81
    'all under 160 except 28A, 19, 18, 43, 70
    'last done 29.48
    y = 1
    x = 32
    For x = 32 To 91                             '(RCW first portion x.###.###) '1-91
    
        sTitle1 = x
    
        For y = 1 To 160                         '(RCW second portion ###.y.###) '1-160 ; exception 28A, 19, 18, 43, 70
        
            sTitle2 = y
            
            If y < 10 Then sTitle2 = "0" & sTitle2
                    
            For z = 1 To 999                     '(RCW third portion ###.###.z) '10 to 990 by 10s
            
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
                sCheck = Right(sCheck, 14)       'gets full rcw ## from html title if there is one, if it's not in here it's a bad URL/RCW
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
    
        For m = 1 To UBound(vLettersArray1)      '(RCW first portion m.###.###)
    
            sTitle1 = vLettersArray1(m)
        
            For n = 1 To 160                     '(RCW second portion ###.n.###) '1-160 ; exception 28A, 19, 18, 43, 70
        
                '1-999 plus A, B, C
        
                vLettersArray2 = Array("A", "B", "C")
            
                For o = 0 To UBound(vLettersArray2) '(RCW second portion m.n[o].p)
                
                    sTitle2 = n & vLettersArray2(o)
                
                    If n < 10 Then sTitle2 = str("0" & sTitle2)
        
                    For p = 10 To 990 Step 10    '(RCW third portion ###.###.p)
                    
                        sTitle3 = p
                    
                        If p < 100 Then sTitle3 = str("0" & sTitle3)
                    
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
    On Error GoTo 0



End Sub

Public Sub pfMARuleScraper()
    
    '============================================================================
    ' Name        : pfMARuleScraper
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfMARuleScraper()
    ' Description:  scrapes all code from MA site
    '============================================================================

    On Error Resume Next
    Dim sFindCitation As String
    Dim sLongCitation As String
    Dim sRuleNumber As String
    Dim sWebAddress As String
    Dim sReplaceHyperlink As String
    Dim sCurrentRule As String
    Dim sCurrentPart As String
    Dim sCurrentTitle As String
    Dim sCurrentChapter As String
    Dim sCurrentSection As String
    
    Dim objHttp As Object
    
    Dim vLetterArray() As Variant
    Dim vRomanArray() As Variant
    
    Dim rstCitationHyperlinks As DAO.Recordset
    
    Dim iErrorNum As Long
    Dim sCHCategory As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim m As Long
    Dim w As Long
    Dim x As Long
    Dim y As Long
    Dim z As Long

    'possibly find mass website of laws to scrape
    'https://malegislature.gov/Laws/GeneralLaws/PartI/TitleXII/Chapter71/Section37O
    'Part I through V
    'Part I   Title I through XXII  Chapter   1 through 182 A through Z Section 1 through 100 A through Z
    'Part II  Title I, II, III      Chapter 183 through 210 A through Z Section 1 through 100 A through Z
    'Part III Title I through VI    Chapter 211 through 262 A through Z Section 1 through 100 A through Z
    'Part IV  Title I, II           Chapter 263 through 280 A through Z Section 1 through 100 A through Z
    'Part V   Title I               Chapter 281 and 282 A through Z     Section 1 through 100 A through Z
    '§
    vRomanArray = Array("I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX", "XX", "XXI", "XXII", "XXIII", "XXIV", "XXV")
    vLetterArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")

    '"https://malegislature.gov/Laws/GeneralLaws/Part" & sCurrentPart & "/Title" & sCurrentTitle & "/Chapter" & sCurrentChapter & "" & sCurrentSection

    For x = 1 To 5                               'part I through V

        sCurrentPart = vRomanArray(x)
    
        For i = 0 To UBound(vRomanArray)         'title I through XXII
    
            sCurrentTitle = vRomanArray(x)
        
            For j = 1 To 282                     'Chapter 1 through 282
        
                'Basic form:  Mass. Gen. Laws Chapter, § Section (Date).   |   Name of Act. Volume Mass. Acts Page Date.
                'Examples:    Mass. Gen. Laws ch. 71, § 1A (1966).   |  An Act Designating Certain Bridges in the Town of Middleborough. 1967 Mass. Acts 116. 8 October 1997.
            
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
            
                If j = 150 Then                  'error
            
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
            
                    'Basic form:  Mass. Gen. Laws Chapter, § Section (Date).   |   Name of Act. Volume Mass. Acts Page Date.
                    'Examples:    Mass. Gen. Laws ch. 71, § 1A (1966).   |  An Act Designating Certain Bridges in the Town of Middleborough. 1967 Mass. Acts 116. 8 October 1997.
                
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
                                
                    If j = 150 Then              'error
                
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
                    
                    For l = 1 To 100             'Section 1 through 100
                
                        'Basic form:  Mass. Gen. Laws Chapter, § Section (Date).   |   Name of Act. Volume Mass. Acts Page Date.
                        'Examples:    Mass. Gen. Laws ch. 71, § 1A (1966).   |  An Act Designating Certain Bridges in the Town of Middleborough. 1967 Mass. Acts 116. 8 October 1997.
                        
                        sCurrentSection = l
                    
                        'also without letters
                    
                        sFindCitation = "Mass. Gen. Laws ch. " & sCurrentChapter & "§ " & sCurrentSection & " (####)"
                        sLongCitation = "Mass. Gen. Laws ch. " & sCurrentChapter & "§ " & sCurrentSection & " (####)"
                        sCHCategory = 2
                        sWebAddress = "https://malegislature.gov/Laws/GeneralLaws/Part" & sCurrentPart & "/Title" & sCurrentTitle & "/Chapter" & sCurrentChapter & "/Section" & sCurrentSection
                        sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
                    
                    
                        Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
                        objHttp.Open "GET", sWebAddress, False
                        objHttp.send ""
                    
                        If j = 50 Then           'error
                    
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
                    
                            'Basic form:  Mass. Gen. Laws Chapter, § Section (Date).   |   Name of Act. Volume Mass. Acts Page Date.
                            'Examples:    Mass. Gen. Laws ch. 71, § 1A (1966).   |  An Act Designating Certain Bridges in the Town of Middleborough. 1967 Mass. Acts 116. 8 October 1997.
                            
                            sCurrentSection = l & vLetterArray(m)
                        
                            sFindCitation = "Mass. Gen. Laws ch. " & sCurrentChapter & "§ " & sCurrentSection & " (####)"
                            sLongCitation = "Mass. Gen. Laws ch. " & sCurrentChapter & "§ " & sCurrentSection & " (####)"
                            sCHCategory = 2
                            sWebAddress = "https://malegislature.gov/Laws/GeneralLaws/Part" & sCurrentPart & "/Title" & sCurrentTitle & "/Chapter" & sCurrentChapter & "/Section" & sCurrentSection
                            sReplaceHyperlink = sFindCitation & "#" & sWebAddress & "#" '"test#http://www.cnn.com#"
                        
                        
                            Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
                            objHttp.Open "GET", sWebAddress, False
                            objHttp.send ""
                        
                            If j = 50 Then       'error
                        
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
    
    On Error GoTo 0

End Sub

Public Sub fUnCompleteTimeMgmtTasks()
    '============================================================================
    ' Name        : fUnCompleteTimeMgmtTasks
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fUnCompleteTimeMgmtTasks()
    ' Description:  unchecks all status boxes for a job number
    '============================================================================

    Dim rstCommHistory As DAO.Recordset
    Dim sQuestion As String
    Dim sAnswer As String
    Dim x As Long

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
        For x = 2 To 28
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
    
    If sAnswer = vbNo Then                       'Code for No
    
        MsgBox "Done!"
        
    Else                                         'Code for yes
    
        'ask for job number
        sCourtDatesID = InputBox("Job Number?")
        GoTo StartHere

    End If


End Sub

Public Sub fCompleteTimeMgmtTasks()
    '============================================================================
    ' Name        : fCompleteTimeMgmtTasks
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fCompleteTimeMgmtTasks()
    ' Description:  checks all tasks from tasks table for a job number
    '============================================================================
    'Call fWunderlistGetTasksOnList 'insert function name here

    Dim rstCommHistory As DAO.Recordset
    Dim sQuestion As String
    Dim sAnswer As String

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
    
    If sAnswer = vbNo Then                       'Code for No
    
        MsgBox "Done!"
        
    Else                                         'Code for yes
    
        'ask for job number
        sCourtDatesID = InputBox("Job Number?")
        GoTo StartHere

    End If


End Sub

Public Sub fCompleteStatusBoxes()
    '============================================================================
    ' Name        : fCompleteStatusBoxes
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fCompleteStatusBoxes()
    ' Description:  checks all status boxes for a job number
    '============================================================================
    'Call fWunderlistGetTasksOnList 'insert function name here

    Dim rstCommHistory As DAO.Recordset
    Dim sQuestion As String
    Dim sAnswer As String
    Dim x As Long

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
        For x = 2 To 28
            rstCommHistory.Fields(x).Value = True
        Next
        rstCommHistory.Update
        rstCommHistory.MoveNext
    Loop

EndHere:
    sQuestion = "Do you want to complete another one?"
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
    
    If sAnswer = vbNo Then                       'Code for No
    
        MsgBox "Done!"
        
    Else                                         'Code for yes
    
        'ask for job number
        sCourtDatesID = InputBox("Job Number?")
        GoTo StartHere

    End If

    rstCommHistory.Close
    Set rstCommHistory = Nothing

End Sub

Public Sub fCompleteStage1Tasks()
    '============================================================================
    ' Name        : fCompleteStage1Tasks
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fCompleteStage1Tasks()
    ' Description:  checks all stage 1 status boxes for a job number
    '============================================================================
    'Call fWunderlistGetTasksOnList 'insert function name here

    Dim rstCommHistory As DAO.Recordset
    Dim sQuestion As String
    Dim sAnswer As String
    Dim x As Long

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
        For x = 2 To 10
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

    If sAnswer = vbNo Then                       'Code for No

        MsgBox "Done!"
    
    Else                                         'Code for yes

        'ask for job number
        sCourtDatesID = InputBox("Job Number?")
        GoTo StartHere

    End If

End Sub

Public Sub fCompleteStage2Tasks()
    '============================================================================
    ' Name        : fCompleteStage2Tasks
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fCompleteStage2Tasks()
    ' Description:  checks all stage 2 status boxes for a job number
    '============================================================================
    'Call fWunderlistGetTasksOnList 'insert function name here

    Dim rstCommHistory As DAO.Recordset
    Dim sQuestion As String
    Dim sAnswer As String
    Dim x As Long

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
        For x = 11 To 15
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
    
    If sAnswer = vbNo Then                       'Code for No
    
        MsgBox "Done!"
        
    Else                                         'Code for yes
    
        'ask for job number
        sCourtDatesID = InputBox("Job Number?")
        GoTo StartHere

    End If

End Sub

Public Sub fCompleteStage3Tasks()
    '============================================================================
    ' Name        : fCompleteStage3Tasks
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fCompleteStage3Tasks()
    ' Description:  checks all stage 3 status boxes for a job number
    '============================================================================
    'Call fWunderlistGetTasksOnList 'insert function name here

    Dim rstCommHistory As DAO.Recordset
    Dim sQuestion As String
    Dim sAnswer As String
    Dim x As Long

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
        For x = 16 To 20
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
    
    If sAnswer = vbNo Then                       'Code for No
    
        MsgBox "Done!"
        
    Else                                         'Code for yes
    
        'ask for job number
        sCourtDatesID = InputBox("Job Number?")
        GoTo StartHere

    End If

    rstCommHistory.Close
    Set rstCommHistory = Nothing

End Sub

Public Sub fCompleteStage4Tasks()
    '============================================================================
    ' Name        : fCompleteStage4Tasks
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fCompleteStage4Tasks()
    ' Description:  checks all stage 4 status boxes for a job number
    '============================================================================
    'Call fWunderlistGetTasksOnList 'insert function name here

    Dim rstCommHistory As DAO.Recordset
    Dim sQuestion As String
    Dim sAnswer As String
    Dim x As Long

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
        For x = 21 To 28
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
    
    If sAnswer = vbNo Then                       'Code for No
    
        MsgBox "Done!"
        
    Else                                         'Code for yes
    
        'ask for job number
        sCourtDatesID = InputBox("Job Number?")
        GoTo StartHere

    End If

End Sub

Public Sub fFixBarAddressField()
    '============================================================================
    ' Name        : fFixBarAddressField
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fFixBarAddressField()
    ' Description:  fixes address field in baraddresses table
    '============================================================================
    Dim sBarNameArray() As String
    Dim sAddressArray() As String
    Dim sCityArray() As String
    Dim sCityArray1() As String
    Dim sLastName As String
    Dim sFirstName As String
    Dim sBarName As String
    Dim sCompany As String
    Dim sAddress As String
    Dim sPhone As String
    Dim sCity As String
    Dim sState As String
    Dim sZIP As String
    Dim sAddress1 As String
    Dim sAddress2 As String

    Dim rstBarAddresses As DAO.Recordset
    Dim rstCustomers As DAO.Recordset
    
    'Customer fields: id, Company, MrMs, LastName, FirstName, EmailAddress, JobTitle, BusinessPhone
    'MobilePhone, FaxNumber, Address, City, State, ZIP, Web Page, Notes, FactoringApproved
 
    Set rstBarAddresses = CurrentDb.OpenRecordset("BarAddresses")
    rstBarAddresses.MoveFirst
    
    'Debug.Print "--------------------------------"

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
    
        'Debug.Print sCompany
        'Debug.Print sFirstName & " " & sLastName
        'Debug.Print sAddress1
        'If sAddress2 <> "" Then Debug.Print sAddress2
        'Debug.Print sCity & ", " & sState & " " & sZIP
        'Debug.Print sPhone
        'Debug.Print "--------------------------------"
        
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
    On Error GoTo 0

End Sub


