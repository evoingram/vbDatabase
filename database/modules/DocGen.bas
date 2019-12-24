Attribute VB_Name = "DocGen"
'@Folder("Database.Production.Modules")
Option Compare Database
Option Explicit

'============================================================================
'class module cmDocGen

'variables:
'   NONE

'functions:
'pfGenericExportandMailMerge:  Description:  exports to specified template from \Database\Templates and saves in \Production\2InProgress\####\
'                          Arguments:    sQueryName, sExportTopic
'pfSendWordDocAsEmail:         Description:  sends Word document as an e-mail body
'                          Arguments:    vCHTopic, vSubject, Optional sAttachment1, sAttachment2, sAttachment3, sAttachment4
'pfCreateCDLabel:               Description:  makes CD label and prompts for print or no
'                          Arguments:    NONE
'pfSelectCoverTemplate:        Description:  parent function to create correct transcript cover/skeleton from template
'                          Arguments:    NONE
'pfCreateCover:                Description:  creates transcript cover/skeleton from template
'                          Arguments:    sTemplatePath
'fCreatePELLetter:             Description:  creates package enclosed letter
'                          Arguments:    NONE
'fFactorInvoicEmailF:          Description:  creates e-mail to submit invoice to factoring
'                          Arguments:    NONE
'pfInvoicesCSV:                Description:  creates CSVs used for invoicing
'                          Arguments:    NONE
'fCreateWorkingCopy:           Description:  creates "working copy" sent to client
'                          Arguments:    NONE

'============================================================================

Public Sub pfGenericExportandMailMerge(sMerge As String, sExportTopic As String)
    '============================================================================
    ' Name        : pfGenericExportandMailMerge
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfGenericExportandMailMerge(sQueryName, sExportTopic)
    ' Description:  exports to specified template from \Database\Templates and saves in \Production\2InProgress\####\
    '============================================================================

    Dim sExportInfoCSVPath As String
    Dim sArray() As String
    Dim sQueryName As String
    
    Dim iCount As Long
    Dim oWordAppDoc As Object
    
    Dim qdf As QueryDef
    Dim rstQuery As DAO.Recordset
    
    Dim xlRange As Excel.Range
    Dim oExcelWB As Excel.Workbook
    Dim oExcelApp As Excel.Application
    
    Dim cJob As Job
    Set cJob = New Job

    If sMerge = "Case" Then

        sQueryName = qnTRCourtUnionAppAddrQ
        sExportInfoCSVPath = cJob.DocPath.CaseInfo
    
    ElseIf sMerge = "Invoice" Then

        sQueryName = qnQInfobyInvNo
        sExportInfoCSVPath = cJob.DocPath.InvoiceInfo
    
    End If
    
    iCount = Len(Dir(sExportInfoCSVPath))

    If iCount = 0 Then

        DoCmd.OutputTo acOutputQuery, sQueryName, acFormatXLS, sExportInfoCSVPath, False
        
        Set oExcelApp = CreateObject("Excel.Application")
        oExcelApp.Application.Visible = False
        oExcelApp.Application.DisplayAlerts = False
    
        Set oExcelWB = oExcelApp.Workbooks.Open(sExportInfoCSVPath)
        oExcelWB.Application.DisplayAlerts = False
        oExcelWB.Application.Visible = False
    
        With oExcelWB
    
            Set xlRange = .Worksheets(1).Range("A2").CurrentRegion
            .Names.Add Name:="AAAAADataRange", RefersTo:=xlRange
            .Save
            .Saved = True
            .Close
        End With
    
        oExcelApp.Quit


    Else
        'do nothing if it exists
    
    End If

    sArray = Split(sExportTopic, "\")

    Set oWordAppDoc = GetObject(cJob.DocPath.TemplateFolder1 & sArray(1) & "-Template.docx") 'sArray(1)/export topic is folder\subject
    oWordAppDoc.Application.Visible = False

    oWordAppDoc.MailMerge.OpenDataSource _
        Name:=sExportInfoCSVPath, _
        ConfirmConversions:=False, ReadOnly:=True, LinkToSource:=True, _
        AddToRecentFiles:=False, SQLStatement:="SELECT * FROM `AAAAADataRange`" ', Revert:=False, _
                                                                                Format:=wdOpenFormatAuto _
                                                                                , SQLStatement:="SELECT * FROM `AAAAADataRange`", SQLStatement1:=""
    'SubType:=wdMergeSubTypeAccess
    ', Connection:= _
    "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cJob.DocPath.JobDirectoryGN & sArray(1) & ".docx" & ";Mode=Read;Extended Properties=" & Chr(34) & Chr(34) & "HDR=YES;IMEX=1;" _
    & Chr(34) & Chr(34) & ";Jet OLEDB:System database=" & Chr(34) & Chr(34) & Chr(34) & Chr(34) & _
    ";Jet OLEDB:Engine Type=34;Jet OLEDB"
        
    oWordAppDoc.MailMerge.Execute
    oWordAppDoc.MailMerge.MainDocumentType = wdNotAMergeDocument
    'oWordAppDoc.Application.ActiveDocument.ExportAsFixedFormat outputFileName:=cJob.DocPath.JobDirectoryGN & sArray(1) & ".pdf", ExportFormat:=wdExportFormatPDF, CreateBookmarks:=wdExportCreateHeadingBookmarks
    oWordAppDoc.Application.ActiveDocument.SaveAs FileName:=cJob.DocPath.JobDirectoryGN & sArray(1) & ".docx"
    oWordAppDoc.Application.ActiveDocument.Close
    Set oWordAppDoc = Nothing

End Sub

Public Sub pfSendWordDocAsEmail(vCHTopic As String, vSubject As String, _
                                Optional sAttachment1 As String, Optional sAttachment2 As String, _
                                Optional sAttachment3 As String, Optional sAttachment4 As String)
    On Error Resume Next
    '============================================================================
    ' Name        : pfSendWordDocAsEmail
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfSendWordDocAsEmail(vCHTopic, vSubject, sAttachment1, sAttachment2, sAttachment3, sAttachment4)
    'attachments optional
    ' Description:  sends Word document as an e-mail body
    '============================================================================
    
    Dim oOutlookApp As Outlook.Application
    Dim oOutlookMail As Outlook.MailItem
    
    Dim oWordApp As New Word.Application
    Dim oWordEditor As Word.editor
    Dim oWordDoc As New Word.Document
    
    Dim cJob As Job
    Set cJob = New Job
    
    Set oOutlookApp = CreateObject("Outlook.Application")
    Set oOutlookMail = oOutlookApp.CreateItem(0)
    Set oWordApp = CreateObject("Word.Application")
    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.JobDirectoryGN & vCHTopic & ".docx")
    oWordDoc.Content.Copy
    
    With oOutlookMail
        .To = ""
        .CC = ""
        .BCC = ""
        .Subject = vSubject
        .BodyFormat = olFormatRichText
        Set oWordEditor = .GetInspector.WordEditor
        .GetInspector.WordEditor.Content.Paste
        .Display
        If sAttachment1 = "" And sAttachment2 = "" And sAttachment3 = "" And sAttachment4 = "" Then GoTo LoopExit
        If Not sAttachment1 = "" Then GoTo At1
        If Not sAttachment1 = "" And sAttachment2 = "" Then GoTo At2
        If Not sAttachment1 = "" And sAttachment2 = "" And sAttachment3 = "" Then GoTo At3
        If Not sAttachment1 = "" And sAttachment2 = "" And sAttachment3 = "" And sAttachment4 = "" Then GoTo At4
At4:
        .Attachments.Add (sAttachment4)
At3:
        .Attachments.Add (sAttachment3)
At2:
        .Attachments.Add (sAttachment2)
At1:
        .Attachments.Add (sAttachment1)
LoopExit:
    End With
    
    On Error GoTo 0
    Set oWordApp = Nothing
    oWordDoc.Close
    oWordApp.Quit
End Sub

Public Sub pfCreateCDLabel()
    '============================================================================
    ' Name        : pfCreateCDLabel
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfCreateCDLabel
    ' Description : makes CD label and prompts for print or no
    '============================================================================

    Dim sCommHistoryHyperlink As String
    Dim sAnswer As String
    Dim sQuestion As String
    
    Dim oPubDoc As Publisher.Document
    Dim oPubApp As Publisher.Application
    
    Dim rstCommHistory As DAO.Recordset
    
    Dim cJob As Job
    Set cJob = New Job

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID
    
    Call pfCurrentCaseInfo 'refresh transcript info

    DoCmd.TransferSpreadsheet TransferType:=acExport, TableName:=qnTRCourtUnionAppAddrQ, FileName:=cJob.DocPath.CaseInfo

    Set oPubApp = New Publisher.Application
    Set oPubDoc = oPubApp.Open(cJob.DocPath.TemplateFolder & "Stage1s/CD-Label.pub")

    sCommHistoryHyperlink = sCourtDatesID & "-CD-Label" & "#" & cJob.DocPath.CDLabel
    

    oPubDoc.MailMerge.OpenDataSource bstrDataSource:=cJob.DocPath.CaseInfo, bstrTable:="", fOpenExclusive:=True, fneverprompt:=1 'performs mail merge
    oPubDoc.MailMerge.Execute Pause:=True, Destination:=pbMergeToNewPublication
    oPubDoc.SaveAs FileName:=cJob.DocPath.CDLabel         'saves file in job number folder in in progress
    oPubDoc.Close
    oPubApp.Quit
    Set rstCommHistory = CurrentDb.OpenRecordset("CommunicationHistory")

    'Adds record to CommHistoryTable and link to document on hard drive
    rstCommHistory.AddNew
        rstCommHistory("FileHyperlink").Value = sCommHistoryHyperlink
        rstCommHistory("DateCreated").Value = Now
        rstCommHistory("CourtDatesID").Value = sCourtDatesID
        rstCommHistory("CustomersID").Value = cJob.App0.ID
    rstCommHistory.Update

    'sQuestion = "Want to burn the CD?"
    'sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    'If sAnswer = vbNo Then 'Code for No
    '    MsgBox "No CD will be burned at this time.  You're done!"
    
    'Else 'Code for yes

    'Call pfBurnCD
    
    'End If

    'sQuestion = "Print CD Label? (MAKE SURE PAPER IS CORRECT IN PRINTER)"
    'sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???") '
    
    'If sAnswer = vbNo Then 'Code for No

    '    MsgBox "CD label will not print."
    
    'Else 'Code for yes

    '    Call fEmailtoPrint(sPubDocPDFName)
    '    Set oPubDoc = Nothing
    
    'End If

    Call fPrintKCIEnvelope
    Call pfClearGlobals
    
End Sub

Public Sub pfSelectCoverTemplate()
    '============================================================================
    ' Name        : pfSelectCoverTemplate
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfSelectCoverTemplate
    ' Description : parent function to create correct transcript cover/skeleton from template
    '============================================================================
    
    Dim sFDAQuery As String

    sFDAQuery = "Food" & "*" & "and" & "*" & "Drug" & "*" & "Administration"

    Call pfCurrentCaseInfo                       'refresh transcript info

    If ((sJurisdiction) Like ("*" & sFDAQuery & "*")) Then
        pfCreateCover ("Stage2s\TR-JEW-FDA.doc")
    ElseIf sJurisdiction = "AVT USBC" Then pfCreateCover ("Stage2s\TR-AVT-Bankruptcy.dotm")
    ElseIf sJurisdiction = "Weber Oregon" Then pfCreateCover ("Stage2s\TR-WeberOregon-Template.dotm")
    ElseIf sJurisdiction = "Weber Nevada" Then pfCreateCover ("Stage2s\TR-WeberNevada-Template.dotm")
    ElseIf sJurisdiction = "Weber Bankruptcy" Then pfCreateCover ("Stage2s\TR-WeberBankruptcy-Template.dotm")
    ElseIf sJurisdiction = "Non-Court" Then pfCreateCover ("Stage2s\TR-noncourt.docx")
    ElseIf sJurisdiction Like "District" & " " & "of" & " " Then pfCreateCover ("Stage2s\TR-Bankruptcy.dotm")
    ElseIf sJurisdiction = "JJ BK" Then pfCreateCover ("Stage2s\TR-JJ-Bankruptcy.docx")
    ElseIf sJurisdiction = "JJ NJ" Then pfCreateCover ("Stage2s\TR-JJ-NJ.docx")
    ElseIf sJurisdiction = "AVT NH" Then pfCreateCover ("Stage2s\TR-AVT-NH.dotm")
    ElseIf sJurisdiction = "AVTOC" Then pfCreateCover ("Stage2s\TR-AVT-OC-CA.dotm")
    ElseIf sJurisdiction = "eScribers NH" Then pfCreateCover ("Stage2s\TR-AVT-NH.dotm")
    ElseIf sJurisdiction = "eScribers AVT NH" Then pfCreateCover ("Stage2s\TR-AVT-NH.dotm") '1.3.2
    ElseIf sJurisdiction = "eScribers Bankruptcy" Then pfCreateCover ("Stage2s\TR-AVT-Bankruptcy.dotm")
    ElseIf sJurisdiction = "eScribers bankruptcy" Then pfCreateCover ("Stage2s\TR-AVT-Bankruptcy.dotm")
    ElseIf sJurisdiction Like "Massachusetts" Then pfCreateCover ("Stage2s\TR-Mass.dotm")
    Else: pfCreateCover ("Stage2s\TR-WACounties.dotm")
    End If

    Call pfCommunicationHistoryAdd("CourtCover")
    Call pfClearGlobals
End Sub

Public Sub pfCreateCover(sTemplatePath As String)
    '============================================================================
    ' Name        : pfCreateCover
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfCreateCover(sTemplatePath)
    ' Description : creates transcript cover/skeleton from template
    '============================================================================

    Dim sCommHistoryAddSQL As String
    Dim sSource As String
    
    Dim oExcelApp As New Excel.Application
    Dim oExcelWB As New Excel.Workbook
    Dim oWordApp As New Word.Application
    Dim oWordDoc As New Word.Document
    Dim xlRange As Excel.Range
    
    Dim oDocuments As Object
    
    Dim x As Long
    
    Dim rstCommHistory As DAO.Recordset
    
    Dim cJob As Job
    Set cJob = New Job

    If Len(Dir(cJob.DocPath.CaseInfo)) = 0 Then

        DoCmd.OutputTo acOutputQuery, qnTRCourtUnionAppAddrQ, acFormatXLS, cJob.DocPath.CaseInfo, False
    
        Set oExcelApp = CreateObject("Excel.Application")
        oExcelApp.Application.Visible = False
        oExcelApp.Application.DisplayAlerts = False
    
        Set oExcelWB = oExcelApp.Application.Workbooks.Open(cJob.DocPath.CaseInfo)
        oExcelWB.Application.DisplayAlerts = False
        oExcelWB.Application.Visible = False
    
        With oExcelWB
            Set xlRange = .Worksheets(1).Range("A2").CurrentRegion
            .Names.Add Name:="AAAAADataRange", RefersTo:=xlRange
            .SaveAs FileName:=cJob.DocPath.CaseInfo
            .Saved = True
            .Close
        End With
    
        oExcelApp.Application.Quit
        
    Else
        'do nothing if it exists
    
    End If

    On Error Resume Next
    Set oWordApp = GetObject(, "Word.Application")

    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If

    With oWordApp
        .Application.DisplayAlerts = False
        .Application.Visible = True
    End With

    Set oWordDoc = GetObject(cJob.DocPath.TemplateFolder & sTemplatePath, "Word.Document") 'sTemplatePath is folder\subject

    On Error GoTo 0

    With oWordDoc
        .MailMerge.OpenDataSource _
        Name:=cJob.DocPath.CaseInfo, _
        LinkToSource:=True, _
 _
        Format:=wdOpenFormatAuto, Connection:= _
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cJob.DocPath.CourtCover & ";Mode=Read;Extended Properties=" & Chr(34) & Chr(34) & "HDR=YES;IMEX=1;" _
            & Chr(34) & Chr(34) & ";Jet OLEDB:System database=" & Chr(34) & Chr(34) & Chr(34) & Chr(34) & ";Jet OLEDB:Registry Path=" & Chr(34) & Chr(34) & Chr(34) & Chr(34) & _
            ";Jet OLEDB:Engine Type=34;Jet OLEDB;" _
            , SQLStatement:="SELECT * FROM `AAAAADataRange`", SQLStatement1:="", _
            SubType:=wdMergeSubTypeAccess
        .MailMerge.DataSource.FirstRecord = wdDefaultFirstRecord
        .MailMerge.DataSource.LastRecord = wdDefaultLastRecord
        .MailMerge.Execute
        .MailMerge.MainDocumentType = wdNotAMergeDocument
    
    End With

    Set oDocuments = Documents
    For x = oDocuments.Count To 1 Step -1
        Debug.Print x
        sSource = ActiveWindow.Caption
    
        If sSource <> "Form Letters1" Then
    
            If sSource <> sCourtDatesID & "-CourtCover.docx" Then
                sSource = Left(sSource, Len(sSource) - 22)
                sSource = Trim(sSource)
            End If
        End If
    
        'Debug.Print sSource
        If sSource = "Form Letters1" Then
            Documents("Form Letters1").Activate
            Documents("Form Letters1").SaveAs FileName:=cJob.DocPath.CourtCover
        Else
            Documents(sSource).Activate
            Documents(sSource).Close SaveChanges:=wdDoNotSaveChanges
        End If
    
    Next x

    'ActiveDocument.Close
    Set oExcelWB = Nothing
    Set oWordApp = Nothing
    Set oExcelApp = Nothing

    Call pfCreateBookmarks
    Call pfCreateIndexesTOAs

    Set rstCommHistory = CurrentDb.OpenRecordset("CommunicationHistory")
    rstCommHistory.AddNew
        rstCommHistory.Fields("FileHyperlink").Value = sCourtDatesID & "#" & cJob.DocPath.CourtCover
        rstCommHistory.Fields("CourtDatesID").Value = sCourtDatesID
        rstCommHistory.Fields("DateCreated").Value = Now
    rstCommHistory.Update
    
    rstCommHistory.Close
    Set rstCommHistory = Nothing
    
    Call pfClearGlobals
End Sub

Public Sub fCreatePELLetter()
    '============================================================================
    ' Name        : fCreatePELLetter
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fCreatePELLetter
    ' Description : creates package enclosed letter
    '============================================================================

    Dim sQuestion As String
    Dim sAnswer As String
    
    Dim cJob As Job
    Set cJob = New Job

    Call pfGenericExportandMailMerge("Case", "Stage1s\PackageEnclosedLetter")

    sQuestion = "Print letter to enclose with transcript?"
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    If sAnswer = vbNo Then                       'Code for No
        MsgBox "Package-enclosed letter will not print."
    Else                                         'Code for yes
        Call fEmailtoPrint(cJob.DocPath.PELP)
    End If

    Call pfCommunicationHistoryAdd("PackageEnclosedLetter")
    Call pfClearGlobals
End Sub

Public Sub fFactorInvoicEmailF()
    '============================================================================
    ' Name        : fFactorInvoicEmailF
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fFactorInvoicEmailF
    ' Description : creates e-mail to submit invoice to factoring
    '============================================================================

    Dim sQuestion As String
    Dim sAnswer As String
    Dim sContactName As String
    Dim sPONumber As String
    Dim sUnitPriceSQL As String
    
    Dim rstUPCourtDates As DAO.Recordset
    Dim rstUnitPrice As DAO.Recordset
    Dim rstFactoringCSV As DAO.Recordset
    
    Dim qdf As QueryDef
    
    Dim oExcelApp As New Excel.Application
    Dim oExcelWB As New Excel.Workbook
    
    Dim cJob As Job
    Set cJob = New Job

    Call pfCurrentCaseInfo                       'refresh transcript info

    Call pfGenericExportandMailMerge("Case", "Stage4s\FactorInvoiceEmail")

    MsgBox "Factoring Invoice Email Created!"

    Call pfCommunicationHistoryAdd("FactorInvoiceEmail")
    
    Set qdf = CurrentDb.QueryDefs(qFCSVQ)
        qdf.Parameters(0) = cJob.ActualQuantity
        qdf.Parameters(1) = sCourtDatesID
    Set rstFactoringCSV = qdf.OpenRecordset
    rstFactoringCSV.MoveFirst
    
        sContactName = rstFactoringCSV.Fields("ContactName").Value
        sPONumber = rstFactoringCSV.Fields("PO Number").Value
        
    Set oExcelApp = CreateObject("Excel.Application")
    Set oExcelWB = oExcelApp.Workbooks.Open(cJob.DocPath.TemplateFolder4 & "Client_Basic_Schedule.xls")

    With oExcelWB
        .Application.DisplayAlerts = False
        .Application.Visible = False
        .Activate
        .Application.ActiveCell.Select
        MsgBox .Application.ActiveCell.Address
        .Application.ActiveCell.Offset(1, 2).Value = sContactName
        .Application.ActiveCell.Offset(2, 2).Value = cJob.InvoiceNo
        .Application.ActiveCell.Offset(3, 2).Value = sPONumber
        .Application.ActiveCell.Offset(4, 2).Value = cJob.InvoiceDate
        .Application.ActiveCell.Offset(5, 2).Value = 28 'net terms
        .Application.ActiveCell.Offset(6, 2).Value = (cJob.ActualQuantity * cJob.PageRate)
        .SaveAs FileName:=cJob.DocPath.JobDirectoryW & "Client_Basic_Schedule.xls", FileFormat:=xlExcel8
        .Close
    End With

    Set oExcelApp = Nothing
    rstFactoringCSV.Close

    sQuestion = "Click yes after you have created your final invoice at " & cJob.DocPath.InvoiceP
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    If sAnswer = vbNo Then                       'IF NO THEN THIS HAPPENS
        MsgBox "No invoice will be sent to factoring."
    Else                                         'if yes then this happens
        Call pfSendWordDocAsEmail("FactorInvoiceEmail", "Invoice to Factor", cJob.DocPath.InvoiceP) 'send email and add attachment yourself (from xero)
    End If

    Application.FollowHyperlink ("https://cirrus.factorfox.net/")
    
    Call pfUpdateCheckboxStatus("InvoicetoFactorEmail")

    qdf.Close
    CurrentDb.Close
    Call pfClearGlobals
    
End Sub


Public Sub pfInvoicesCSV()
    '============================================================================
    ' Name        : pfInvoicesCSV
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfInvoicesCSV
    ' Description : creates CSVs used for invoicing
    '============================================================================

    Dim cJob As Job
    Set cJob = New Job

    'DoCmd.OpenQuery qXeroCSVQ, acViewNormal, acAdd
    DoCmd.TransferText acExportDelim, , qSelectXero, cJob.DocPath.XeroCSV, True

    'real factoring csv plus invoice generated FactoringCSVQuery
    If cJob.App0.FactoringApproved = True Then

        DoCmd.OpenQuery qFCSVQ, acViewNormal, acAdd
        DoCmd.TransferText acExportDelim, , qFCSVQ, cJob.DocPath.XeroCSV, True
        DoCmd.Close (qFCSVQ)
    
    Else
    End If

    Application.FollowHyperlink ("https://go.xero.com/Import/Import.aspx?type=IMPORTTYPE/ARINVOICES")


    Call pfUpdateCheckboxStatus("InvoiceCompleted")
    Call pfClearGlobals
    
End Sub

Public Sub fCreateWorkingCopy()
    '============================================================================
    ' Name        : fCreateWorkingCopy
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fCreateWorkingCopy
    ' Description : creates "working copy" sent to client
    '============================================================================

    Dim sAnswer As String
    Dim sQuestion As String
    
    Dim vbComp As Object
    Dim x As Variant
    Dim oRng As Range
    
    Dim oWordApp As New Word.Application
    Dim oWordDoc As New Word.Document
    Dim wsSections As Word.Sections
    Dim wsSection As Word.Section
    
    Dim cJob As Job
    Set cJob = New Job

    Call pfWordIndexer

    sQuestion = "Next we will make a working copy.  Click yes when ready."
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    If sAnswer = vbNo Then                       'Code for No
        MsgBox "No working copy will be made."
    
    Else                                         'Code for yes
        On Error Resume Next
        Set oWordApp = GetObject(, "Word.Application")


        If Err <> 0 Then
            Set oWordApp = CreateObject("Word.Application")
        End If
        On Error GoTo 0
        oWordApp.Application.DisplayAlerts = False
        oWordApp.Application.Visible = False

        Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.CourtCover)


        SendKeys "+{Home}"
        With oWordDoc
    
            'add continuous section break at top
            .Paragraphs(1).Range.InsertBreak Type:=wdSectionBreakContinuous
        
            If .ProtectionType <> wdNoProtection Then .Unprotect password:="wrts0419"
            'removes doc info and macro code within document
            .RemoveDocumentInformation (wdRDIAll)
            For Each vbComp In .VBProject.VBComponents
                With vbComp
                    If .Type = 100 Then
                        .CodeModule.DeleteLines 1, .CodeModule.CountOfLines
                    Else
                        .VBProject.VBComponents.Remove vbComp
                    End If
                End With
            Next vbComp
            SendKeys "+{Home}"
        
            'delete cert section
            Set wsSections = .Sections
            Set oRng = .Range(Start:=.bookmarks("CertBMK").Range.End, End:=.bookmarks("ToABMK").Range.Start)
            oRng.delete
        
            For x = wsSections.Last.Index To 1 Step -1
                Set wsSection = wsSections.item(x)
                If x = 1 Then
                    wsSection.ProtectedForForms = True
                Else
                    wsSection.ProtectedForForms = False
                End If
            
            Next x
        
            'lock up header, leave all other sections unlocked
            .Protect Type:=wdAllowOnlyFormFields, noReset:=True, password:="wrts0419"
            .SaveAs FileName:=cJob.DocPath.TranscriptWC, FileFormat:=wdFormatXMLDocument 'save as file
        
        End With
    End If

    oWordDoc.Close
    oWordApp.Quit
    Set oWordDoc = Nothing
    Set oWordApp = Nothing

    FileCopy cJob.DocPath.TranscriptWC, cJob.DocPath.TranscriptWCB

End Sub

Public Sub pfPrepareCover()
    '============================================================================
    ' Name        : pfPrepareCover
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfPrepareCover
    ' Description : prepares volume cover for transcript compendium
    '============================================================================

    Dim sOriginalCurrentTranscriptPath As String
    Dim sNewCurrentTranscriptPath As String
    Dim sVolumeText As String
    Dim sBookmarkName As String
    Dim sCurrentJobID As String
    
    Dim x As Long
    Dim y As Long
    
    Dim oWordApp As Word.Application
    Dim oWordDoc As Word.Document
    
    Dim rstJobsByCase As DAO.Recordset
    
    Dim cJob As Job
    Set cJob = New Job
    
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    
    'query for all dates & job numbers for a case
    Set rstJobsByCase = CurrentDb.OpenRecordset("SELECT * FROM Courtdates WHERE CasesID = " & cJob.CaseInfo.ID & ";")
        x = rstJobsByCase.RecordCount
    rstJobsByCase.MoveFirst
    y = 1
    'loop through each one

    Do While rstJobsByCase.EOF = False

        sCurrentJobID = rstJobsByCase.Fields("ID").Value

        'copy other transcript pdfs for same volume into \transcripts\ folder
        sOriginalCurrentTranscriptPath = cJob.DocPath.InProgressFolder & sCurrentJobID & "\Transcripts\" & sCurrentJobID & "-Transcript-FINAL.pdf"
        sNewCurrentTranscriptPath = cJob.DocPath.InProgressFolder & sCourtDatesID & "\Transcripts\" & sCurrentJobID & "-Transcript-FINAL.pdf"
    
        If Not sOriginalCurrentTranscriptPath = sNewCurrentTranscriptPath Then
            FileCopy sOriginalCurrentTranscriptPath, sNewCurrentTranscriptPath
        End If
        
        On Error Resume Next
        Set oWordApp = GetObject(, "Word.Application")
    
        If Err <> 0 Then
            Set oWordApp = CreateObject("Word.Application")
        End If
        
        Set oWordDoc = GetObject(cJob.DocPath.WACoverD, "Word.Document")
    
    
        Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.WACoverD)
        On Error GoTo 0
    
        oWordDoc.Application.Visible = True
        oWordDoc.Application.DisplayAlerts = False
    
        sBookmarkName = "VolumeBMK0" & y
        
        'list current date on volumes cover
        sVolumeText = "Volume " & y & ":  " & Format(cJob.HearingDate, "dddd, mmmm d, yyyy")
        'Debug.Print sVolumeText
        With oWordDoc
            .bookmarks(sBookmarkName).Select
            .Application.Selection.TypeText Text:=sVolumeText
        End With
        y = y + 1
        rstJobsByCase.MoveNext
    
    Loop

    oWordDoc.Save

    'make volume cover pdf
    oWordDoc.ExportAsFixedFormat outputFileName:=cJob.DocPath.WACoverP, ExportFormat:=wdExportFormatPDF

    oWordDoc.Close
    oWordApp.Quit
    rstJobsByCase.Close
    
End Sub


