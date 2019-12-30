Attribute VB_Name = "~AScratchPad"
'@Folder("Database.General.Modules")

'TODO: get permissions and hook into various case mgmt software to get orders

'*****Medium Priority*****

'*****Low Priority*****

'============================================================================

'============================================================================
' Name        : Name
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call Name(Argument, Argument)
'Argument optional
' Description : comment
'============================================================================

'============================================================================
'class module cmClassName

'variables:
'   Sleep(Milliseconds)

'functions:
'Name:  Description:  comment
'   Arguments:    NONE

'Name:  Description:  comment
'   Arguments:    NONE

'Name:  Description:  comment
'   Arguments:    NONE

'Name:  Description:  comment
'   Arguments:    NONE

'Name:  Description:  comment
'   Arguments:    NONE

'Name:  Description:  comment
'   Arguments:    NONE
'============================================================================

Option Explicit

'Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As Long, _
ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long

'Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

'Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
(ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long

'Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

'Private Declare PtrSafe Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" _
(ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

'Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long

'@Ignore ProcedureNotUsed
Private Sub testClassesInfo()
    Dim cJob As Job
    Set cJob = New Job
    'On Error Resume Next
    sCourtDatesID = 1874
    cJob.FindFirst "ID=" & sCourtDatesID
    Debug.Print cJob.ID
    Debug.Print cJob.Status.AddRDtoCover
    Debug.Print cJob.AudioLength
    Debug.Print cJob.TurnaroundTime
    Debug.Print cJob.CaseInfo.Party1
    Debug.Print cJob.CaseInfo.Party1Name
    Debug.Print Format$(cJob.HearingStartTime, "h:mm AM/PM")
    Debug.Print Format$(cJob.HearingEndTime, "h:mm AM/PM")
    Debug.Print Format$(cJob.HearingDate, "mm-dd-yyyy")
    Debug.Print cJob.App1.ID
    Debug.Print cJob.App1.Company
    Debug.Print cJob.App0.ID
    Debug.Print cJob.App0.Company
    Debug.Print cJob.App0.FactoringApproved
    Debug.Print cJob.CaseInfo.Party1
    Debug.Print "Page Rate = " & cJob.UnitPrice
    Debug.Print cJob.PageRate
    Debug.Print cJob.InventoryRateCode
    cJob.Status.AddRDtoCover = True
    Debug.Print cJob.Status.AddRDtoCover
    Debug.Print "Template Folder = " & cJob.DocPath.TemplateFolder2
    Debug.Print "Template Folder = " & cJob.DocPath.OrderConfirmationD
    Debug.Print "Page Rate = " & cJob.PageRate
    'cJob.Update
    'On Error GoTo 0
End Sub


Private Sub pfWashingtonTranscriptCompiler()
    '============================================================================
    ' Name        : pfWashingtonTranscriptCompiler
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfWashingtonTranscriptCompiler()
    ' Description:  compiles transcripts to file with COA
    '============================================================================
    
    'TODO: meet 2020 Washington transcript requirements
    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID
    
    Dim sAllTranscriptDates As String
    Dim sSource As String
    Dim sCurrentOtherCourtDatesID As String
    Dim sFinalTranscripts As String
    Dim OK As String
    
    Dim oDocuments As Object
    
    Dim x As Integer
    Dim primaryPageCount As Integer
    Dim sourcePageCount As Integer
    
    Dim rstAllTranscriptsInCase As DAO.Recordset
    Dim rstCommHistory As DAO.Recordset
    
    Dim oWordDoc As New Word.Document
    Dim oWordApp As New Word.Application
    Dim oWordDoc1 As New Word.Document
    
    Dim xlRange As Excel.Range
    Dim oExcelWB As Excel.Workbook
    Dim oExcelApp As Excel.Application
    
    Dim aeAcroExchange As Acrobat.CAcroApp
    Dim aePrimaryDoc As Acrobat.CAcroPDDoc
    Dim aeSourceDoc As Acrobat.CAcroPDDoc
    
    'To Fix Bookmarks:
    'Final Transcript Outline:
        'Cover All
        'Cover Date
            'General Index
                'normal
            'Witness Index
                'normal
            'Exhibit Index
                'normal
            'Transcript Body
            'Certificate
            'Table of Authorities
        'Cover Date
            'General Index
                'normal
            'Witness Index
                'normal
            'Exhibit Index
                'normal
            'Transcript Body
            'Certificate
            'Table of Authorities
        
        Forms![NewMainMenu].Form!lblFlash.Caption = "Compiling Washington COA transcript."
        
        'does this transcript need to be compiled
            'if not, do like normal
            'if yes, do following
                'select query to get all transcript dates of current caseID
                'export to csv in /workingfiles/
                sAllTranscriptDates = "SELECT * FROM CourtDates WHERE [CasesID]=" & cJob.CaseID & ";"
                If IsNull(DLookup("name", "msysobjects", "name='qAllTranscriptDates'")) Then
                    CurrentDb.CreateQueryDef "qAllTranscriptDates", sAllTranscriptDates
                Else
                    CurrentDb.QueryDefs("qAllTranscriptDates").Sql = sAllTranscriptDates
                End If
                
                DoCmd.OutputTo acOutputQuery, "qAllTranscriptDates", acFormatXLS, cJob.DocPath.JobDirectoryW & "CompiledTranscripts.xls", False
                
                'get xls content into range
                Set oExcelApp = CreateObject("Excel.Application")
                oExcelApp.Application.Visible = False
                oExcelApp.Application.DisplayAlerts = False
            
                Set oExcelWB = oExcelApp.Workbooks.Open(cJob.DocPath.JobDirectoryW & "CompiledTranscripts.xls")
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

                'generate cover all
                Set oWordDoc = GetObject(cJob.DocPath.TemplateFolder1 & "TR-AllCover.dotm", "Word.Document")
                oWordDoc.Application.Visible = False
        
                On Error GoTo 0
                             
                 With oWordDoc
                     .MailMerge.OpenDataSource _
                     Name:=cJob.DocPath.JobDirectoryW & "CompiledTranscripts.xls", _
                     LinkToSource:=True, _
 _
                     Format:=wdOpenFormatAuto, Connection:= _
                     "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cJob.DocPath.WAConsolidatedD & ";Mode=Read;Extended Properties=" & Chr(34) & Chr(34) & "HDR=YES;IMEX=1;" _
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
                     'Debug.Print x
                     sSource = ActiveWindow.Caption
                 
                     If sSource <> "Form Letters1" Then
                 
                         If sSource <> sCourtDatesID & "-Cover.docx" Then
                             sSource = Left(sSource, Len(sSource) - 27)
                             sSource = Trim(sSource)
                         End If
                     End If
                 
                     'Debug.Print sSource
                     If sSource = "Form Letters1" Then
                         Documents("Form Letters1").Activate
                         Documents("Form Letters1").SaveAs FileName:=cJob.DocPath.WAConsolidatedD
                'make pdf
                         Documents(cJob.DocPath.WAConsolidatedD).SaveAs cJob.DocPath.WAConsolidatedP
                     Else
                         Documents(sSource).Activate
                         Documents(sSource).Close SaveChanges:=wdDoNotSaveChanges
                     End If
                 
                 Next x
             
                 Set oExcelWB = Nothing
                 Set oWordApp = Nothing
                 Set oExcelApp = Nothing
             
                 Set rstCommHistory = CurrentDb.OpenRecordset("CommunicationHistory")
                 rstCommHistory.AddNew
                     rstCommHistory.Fields("FileHyperlink").Value = sCourtDatesID & "#" & cJob.DocPath.WAConsolidatedD
                     rstCommHistory.Fields("CourtDatesID").Value = sCourtDatesID
                     rstCommHistory.Fields("DateCreated").Value = Now
                 rstCommHistory.Update
                 
                 rstCommHistory.Close
                 Set rstCommHistory = Nothing
                 
                 Call pfClearGlobals
                                
                'select query to get all transcript dates of current caseID
                Set rstAllTranscriptsInCase = CurrentDb.OpenRecordset(sAllTranscriptDates)
                'FileCopy cJob.DocPath.TranscriptFP, cJob.DocPath.TranscriptFPB
                
                If Not (rstAllTranscriptsInCase.EOF And rstAllTranscriptsInCase.BOF) Then
            
                    rstAllTranscriptsInCase.MoveFirst
                
                    Do Until rstAllTranscriptsInCase.EOF = True
                    
                'copy in other transcripts for same case
                        sCurrentOtherCourtDatesID = rstAllTranscriptsInCase.Fields("ID").Value
                        FileCopy cJob.DocPath.WAConsolidatedD, cJob.DocPath.InProgressFolder & sCurrentOtherCourtDatesID & "/Generated/" & sCurrentOtherCourtDatesID & "-Cover.docx"
                        FileCopy cJob.DocPath.WAConsolidatedP, cJob.DocPath.InProgressFolder & sCurrentOtherCourtDatesID & "/Generated/" & sCurrentOtherCourtDatesID & "-Cover.pdf"
                        rstAllTranscriptsInCase.MoveNext
                    
                    Loop
                
                Else
            
                    MsgBox "There are no records in the recordset."
                
                End If

                'branch to complete other transcript dates if not done already
                
                Set rstAllTranscriptsInCase = CurrentDb.OpenRecordset(sAllTranscriptDates)
                
                If Not (rstAllTranscriptsInCase.EOF And rstAllTranscriptsInCase.BOF) Then
            
                    rstAllTranscriptsInCase.MoveFirst
                
                    Do Until rstAllTranscriptsInCase.EOF = True
                    
                'copy into first compiled ID# the other transcripts for same case
                        sCurrentOtherCourtDatesID = rstAllTranscriptsInCase.Fields("ID").Value
                        FileCopy cJob.DocPath.InProgressFolder & sCurrentOtherCourtDatesID & "/Transcripts/" & sCurrentOtherCourtDatesID & "-Transcript-FINAL.pdf", _
                                 cJob.DocPath.InProgressFolder & sCourtDatesID & "/Transcripts/" & sCurrentOtherCourtDatesID & "-Transcript-FINAL.pdf"
                        rstAllTranscriptsInCase.MoveNext
                    
                    Loop
                
                Else
            
                    Debug.Print "There are no records in the recordset."
                
                End If

            
                'get all to-be-compiled files
                sFinalTranscripts = Dir(cJob.DocPath.JobDirectoryG & "\*" & "-Transcript-FINAL.PDF")
                
                Set aeAcroExchange = CreateObject("Acroexch.app")
                Set aePrimaryDoc = CreateObject("AcroExch.PDDoc")
                OK = aePrimaryDoc.Open("CoverName.pdf")
                Debug.Print "PRIMARY DOC OPENED & PDDOC SET: " & OK
                
                'compile all final transcript PDFs for a case
                Do While Len(sFinalTranscripts) > 0
                    'add current pdf to end of original pdf
                    primaryPageCount = aePrimaryDoc.GetNumPages() - 1
            
                    Set aeSourceDoc = CreateObject("AcroExch.PDDoc")
                    OK = aeSourceDoc.Open("ToBeInsertedTranscriptName.pdf")
                    Debug.Print "SOURCE DOC OPENED & PDDOC SET: " & OK
            
                    sourcePageCount = aeSourceDoc.GetNumPages
            
                    OK = aePrimaryDoc.InsertPages(primaryPageCount, aeSourceDoc, 0, sourcePageCount, False)
                    Debug.Print "PAGES INSERTED SUCCESSFULLY: " & OK
            
                    OK = aePrimaryDoc.Save(PDSaveFull, "FinalCompiledPDFName")
                    Debug.Print "PRIMARYDOC SAVED PROPERLY: " & OK
            
                    Set aeSourceDoc = Nothing
                    
                Loop
                
                Set aeSourceDoc = Nothing
                Set aePrimaryDoc = Nothing
                aeAcroExchange.Exit
                Set aeAcroExchange = Nothing
                Forms![NewMainMenu].Form!lblFlash.Caption = "Ready to process."
                MsgBox "Compilation complete.  Make sure your COA transcript looks fine, including bookmarks."

    sCourtDatesID = ""
End Sub



'@Ignore EmptyMethod
'@Ignore ProcedureNotUsed
Private Sub emptyFunction()
        
        
    '============================================================================
    ' Name        : pfWashingtonTranscriptCompiler
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfWashingtonTranscriptCompiler()
    ' Description:  compiles transcripts to file with COA
    '============================================================================
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID
    
    
    
    
    Dim oWordDoc As New Word.Document
    Dim oWordApp As New Word.Application
    Dim oWordDoc1 As New Word.Document
        
        
        
        
        
    'Debug.Print cJob.DocPath.CaseInfo
    'Debug.Print "test"
    'Call pfSendWordDocAsEmail("PP-FactoredInvoiceEmail", "Transcript Delivery & Invoice", cJob.DocPath.InvoiceP)
    'On Error Resume Next
    
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID
    'Debug.Print ("---------------------------------------------")
    
    Set oWordApp = CreateObject("Word.Application")
    oWordApp.Visible = False
        
    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPButton) 'open button in word
    oWordDoc.Content.Copy
    Set oWordApp = CreateObject("Word.Application")
    oWordApp.Visible = False
        
    Set oWordDoc1 = oWordApp.Documents.Open(cJob.DocPath.InvoiceD) 'open invoice docx

    With oWordDoc1.Application

        .Selection.Find.ClearFormatting
        .Selection.Find.Replacement.ClearFormatting
    
        With .Selection.Find
        
            .Text = "#PPB1#"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute
        
            'enter code here
            If .Forward = True Then
                .Application.Selection.Collapse Direction:=wdCollapseStart
            Else
                .Application.Selection.Collapse Direction:=wdCollapseEnd
            End If
        
            .Execute Replace:=wdReplaceOne
            If .Forward = True Then
                .Application.Selection.Collapse Direction:=wdCollapseEnd
            Else
                .Application.Selection.Collapse Direction:=wdCollapseStart
            End If
            .Execute
            .Application.Selection.PasteAndFormat (wdFormatOriginalFormatting) 'paste button
        
        End With
        .Selection.Find.ClearFormatting
        .Selection.Find.Replacement.ClearFormatting
    
        With .Selection.Find
            .Text = "#PPB2#"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute
            If .Forward = True Then
                .Application.Selection.Collapse Direction:=wdCollapseStart
            Else
                .Application.Selection.Collapse Direction:=wdCollapseEnd
            End If
            .Execute Replace:=wdReplaceOne
            If .Forward = True Then
                .Application.Selection.Collapse Direction:=wdCollapseEnd
            Else
                .Application.Selection.Collapse Direction:=wdCollapseStart
            End If
            .Execute
            .Application.Selection.PasteAndFormat (wdFormatOriginalFormatting) 'paste button
        End With
    
        'save invoice
        oWordDoc1.Save
    
    End With

    oWordDoc1.Close
    oWordDoc.Close
    oWordApp.Quit
    
    Set oWordApp = Nothing
    Set oWordDoc = Nothing
    Set oWordDoc1 = Nothing
    sCourtDatesID = ""
End Sub



