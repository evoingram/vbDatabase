Attribute VB_Name = "Stage3"
'@Folder("Database.Production.Modules")
Option Compare Database
Option Explicit

'============================================================================
'class module cmStage3

'variables:
'   NONE

'functions:

    'pfStage3Ppwk:        Description:  completes all stage 3 tasks
        '                 Arguments:    NONE
    'pfBurnCD:            Description:  burns CD to D drive
        '                 Arguments:    NONE
    'pfCreateRegularPDF:  Description:  creates final PDF of transcript and saves to main/transcripts folders
        '                 Arguments:    NONE
    'pfHeaders            Description : add sections and headers programmatically
        '                 Arguments:    NONE

'============================================================================


Public Sub pfStage3Ppwk()
'============================================================================
' Name        : pfStage3Ppwk
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfStage3Ppwk
' Description : completes all stage 3 tasks
'============================================================================

Dim sContTrReadyPath As String
Dim oWordApp As New Word.Application, oWordEditor As Word.editor, oWordDoc As New Word.Document
Dim oOutlookApp As Outlook.Application, oOutlookMail As Outlook.MailItem
Dim sDeliveryURL As String


Call pfGetOrderingAttorneyInfo
Call pfCheckFolderExistence 'checks for job folder and creates it if not exists

Call pfUpdateCheckboxStatus("AudioProof")

Call pfCurrentCaseInfo  'refresh transcript info
If sJurisdiction Like "*AVT*" Then
    sDeliveryURL = "http://tabula.escribers.net/"
    
    Call pfFileRenamePrompt
    
    Call fTranscriptDeliveryF
    
    Call pfCommunicationHistoryAdd("FileTranscriptEmail") 'LOG FACTORED CLIENT INVOICE
    
    MsgBox "Next, upload this job to eScribers."
    
    Application.FollowHyperlink (sDeliveryURL) 'FILE WITH ESCRIBERS
    
ElseIf sJurisdiction = "eScribers NH" Then
    sDeliveryURL = "http://tabula.escribers.net/"
    Call pfFileRenamePrompt
    Call fTranscriptDeliveryF
    Call pfCommunicationHistoryAdd("FileTranscriptEmail") 'LOG FACTORED CLIENT INVOICE
    
    MsgBox "Next, upload this job to eScribers."
    
    Application.FollowHyperlink (sDeliveryURL) 'FILE WITH ESCRIBERS
   
ElseIf sJurisdiction Like "*FDA*" Then
    Call pfCurrentCaseInfo  'refresh transcript info
    Call fTranscriptDeliveryF '(only the filing)
    Call pfGenericExportandMailMerge("Case", "Stage4s\ContractorTranscriptsReady")
    Call pfFileRenamePrompt
    
    Set oOutlookApp = CreateObject("Outlook.Application")
    Set oOutlookMail = oOutlookApp.CreateItem(0)
    
    Set oWordApp = CreateObject("Word.Application")
    
    sContTrReadyPath = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-ContractorTranscriptsReady" & ".docx"
    
    Set oWordDoc = oWordApp.Documents.Open(sContTrReadyPath)
    
    oWordDoc.Content.Copy
    
    With oOutlookMail
        .To = ""
        .CC = ""
        .BCC = ""
    
        .Attachments.Add (sClientTranscriptName)
        .Subject = sJurisdiction & " " & dHearingDate & " Transcript Ready " & sCourtDatesID
        .BodyFormat = olFormatRichText
        Set oWordEditor = .GetInspector.WordEditor
        .GetInspector.WordEditor.Content.Paste
        .Display
    End With
    oWordDoc.Close
    oWordApp.Quit
    Set oWordApp = Nothing
    
    Call pfCommunicationHistoryAdd("FileTranscriptEmail") 'LOG FACTORED CLIENT INVOICE
    
ElseIf sJurisdiction Like "Food and Drug Administration" Then
    Call pfCurrentCaseInfo  'refresh transcript info
    Call fTranscriptDeliveryF '(only the filing)
    Call pfGenericExportandMailMerge("Case", "Stage4s\ContractorTranscriptsReady")
    Call pfFileRenamePrompt
    
    Set oOutlookApp = CreateObject("Outlook.Application")
    Set oOutlookMail = oOutlookApp.CreateItem(0)
    Set oWordApp = CreateObject("Word.Application")
    
    sContTrReadyPath = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-ContractorTranscriptsReady" & ".docx"
    
    Set oWordDoc = oWordApp.Documents.Open(sContTrReadyPath)
    oWordDoc.Content.Copy
    
    With oOutlookMail
        .To = ""
        .CC = ""
        .BCC = ""
        .Attachments.Add (sClientTranscriptName)
        .Subject = sJurisdiction & " " & dHearingDate & " Transcript Ready " & sCourtDatesID
        .BodyFormat = olFormatRichText
        
        Set oWordEditor = .GetInspector.WordEditor
        .GetInspector.WordEditor.Content.Paste

        .Display
    End With
    oWordDoc.Close
    oWordApp.Quit
    Set oWordApp = Nothing
    
    Call pfCommunicationHistoryAdd("FileTranscriptEmail") 'LOG FACTORED CLIENT INVOICE
    
ElseIf sJurisdiction Like "Weber" Then
    Call pfCurrentCaseInfo  'refresh transcript info
    Call fTranscriptDeliveryF '(only the filing)
    Call pfGenericExportandMailMerge("Case", "Stage4s\ContractorTranscriptsReady")
    Call pfFileRenamePrompt
    
    Set oOutlookApp = CreateObject("Outlook.Application")
    Set oOutlookMail = oOutlookApp.CreateItem(0)
    
    Set oWordApp = CreateObject("Word.Application")
    sContTrReadyPath = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-ContractorTranscriptsReady" & ".docx"
    
    Set oWordDoc = oWordApp.Documents.Open(sContTrReadyPath)
    oWordDoc.Content.Copy
    
    With oOutlookMail
        .To = ""
        .CC = ""
        .BCC = ""
        .Attachments.Add (sClientTranscriptName)
        .Subject = sJurisdiction & " " & dHearingDate & " Transcript Ready " & sCourtDatesID
        .BodyFormat = olFormatRichText
        
        Set oWordEditor = .GetInspector.WordEditor
        .GetInspector.WordEditor.Content.Paste
        
        .Display
    End With
    
    oWordDoc.Close
    oWordApp.Quit
    Set oWordApp = Nothing
    
    Call pfCommunicationHistoryAdd("FileTranscriptEmail") 'LOG FACTORED CLIENT INVOICE
    
Else '2-up, 4-up, zips

    Call fGenerateZIPsF
    Call pfUpdateCheckboxStatus("GenerateZIPs")
    
End If

    MsgBox "Stage 3 complete."
    Call pfClearGlobals
End Sub

Public Sub pfBurnCD()
'============================================================================
' Name        : pfBurnCD
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfBurnCD
' Description : burns CD to D drive
'============================================================================

Dim sAnswer As String, sQuestion As String, qnViewJobFormAppearancesQ As String
Dim sDrive As String, sSource As String, sCDVolumeName As String, sBurnDir As String
Dim oWSHShell As Object, oShell As Object, oSource As Object


sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

Call pfCheckFolderExistence 'checks for job folder and creates it if not exists

sQuestion = "Is there a blank CD in the D drive?"
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

If sAnswer = vbNo Then 'Code for No

    MsgBox "CD will not burn."
    
Else 'Code for yes
 
    sDrive = InputBox("Driveletter : ", "Driveletter", "D") & ":\" 'CD burner drive letter
    sSource = "I:\" & sCourtDatesID & "\FTP\" 'source directory
    sCDVolumeName = sCourtDatesID 'CD volume name (16-char max)
    Const MY_COMPUTER = &H11
    
    Set oWSHShell = CreateObject("WScript.Shell")
    Set oShell = CreateObject("Shell.Application")
    sBurnDir = oWSHShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\" & "Explorer\Shell Folders\CD Burning")
    'come back
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    sSource = "I:\" & sCourtDatesID & "\FTP\" 'source directory
    
    
    Set oSource = oShell.Namespace(sSource)
    oShell.Namespace(sBurnDir).CopyHere oSource.Items
    oShell.Namespace(&H11&).ParseName(sDrive).InvokeVerbEx ("Write these files to CD")
    
End If
    
End Sub

Public Sub pfCreateRegularPDF()
'============================================================================
' Name        : pfCreateRegularPDF
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfCreateRegularPDF
' Description : creates final PDF of transcript and saves to main/transcripts folders
'============================================================================

Dim sWorkingCopyPath As String, sTranscriptWD As String, sFinalTranscriptWD As String
Dim sFinalTranscriptNoExt As String, sCourtCoverPath As String
Dim sAnswerPDFPrompt As String, sMakePDFPrompt As String

Dim oWordApp As Word.Application, oWordDoc As Word.Document, oVBComponent As Object
Dim rngStory As Range

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
sWorkingCopyPath = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-WorkingCopy.docx"
sTranscriptWD = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript.docx"
sFinalTranscriptWD = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-FINAL.docx"
sFinalTranscriptNoExt = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-FINAL"
sMakePDFPrompt = "Next we will make a PDF copy.  Click yes when ready."
sAnswerPDFPrompt = MsgBox(sMakePDFPrompt, vbQuestion + vbYesNo, "???")

If sAnswerPDFPrompt = vbNo Then 'Code for No
    MsgBox "No PDF copy will be made."
    
Else 'Code for yes

    sCourtCoverPath = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-CourtCover.docx"
    
    'Set oWordApp = GetObject(, "Word.Application")
    
    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If
    
    Set oWordDoc = GetObject(sCourtCoverPath, "Word.Document")
    oWordDoc.Application.Visible = False
    oWordDoc.Application.DisplayAlerts = False

    'Set oWordDoc = oWordApp.Documents.Open(FileName:=sCourtCoverPath)

    
    With oWordDoc
        If .ProtectionType <> wdNoProtection Then .Unprotect password:="wrts0419"
        .Activate
        .Application.Selection.HomeKey Unit:=wdStory
        For Each rngStory In .StoryRanges
            Do
                With rngStory.Find
                        .Replacement.ClearFormatting
                        .Text = "***WORKING COPY***^p"
                        .Replacement.Text = ""
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                        .Execute Replace:=wdReplaceAll
                End With
                
                    .Application.Selection.Find.ClearFormatting
                    .Application.Selection.Find.Replacement.ClearFormatting
                With rngStory.Find
                        .Text = "***WORKING COPY***"
                        .Replacement.Text = ""
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                        .Execute Replace:=wdReplaceAll
                End With
                
                Set rngStory = rngStory.NextStoryRange
                
            Loop Until rngStory Is Nothing
            
        Next rngStory
        
        .RemoveDocumentInformation (wdRDIAll) 'remove vba and document info
        
        'For Each oVBComponent In oWordApp.VBProject.oVBComponents
           ' With oVBComponent
                'If .Type = 100 Then
                    '.CodeModule.DeleteLines 1, .CodeModule.CountOfLines
                'Else
                    '.VBProject.oVBComponentonents.Remove oVBComponent
               ' End If
            'End With
        'Next oVBComponent
        
    End With
End If

'lock document in whole and save as final
oWordDoc.Protect Type:=wdAllowOnlyReading, noReset:=True, password:="wrts0419"
oWordDoc.SaveAs FileName:=sFinalTranscriptWD 'sFinalTranscriptNoExt
oWordDoc.ExportAsFixedFormat outputFileName:=sFinalTranscriptNoExt, ExportFormat:=wdExportFormatPDF, CreateBookmarks:=wdExportCreateHeadingBookmarks
oWordDoc.Close SaveChanges:=False
Set oWordDoc = Nothing
Set oWordApp = Nothing

Call pfTopOfTranscriptBookmark



MsgBox "Created Regular PDF Copy!"

FileCopy sFinalTranscriptWD, "I:\" & sCourtDatesID & "\Backups\" & sCourtDatesID & "-Transcript-FINAL.docx"
FileCopy sFinalTranscriptNoExt & ".pdf", "I:\" & sCourtDatesID & "\Backups\" & sCourtDatesID & "-Transcript-FINAL.pdf"

End Sub



Sub fDynamicHeaders()
'============================================================================
' Name        : fDynamicHeaders
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fDynamicHeaders()
' Description : adds the dynamic headers to transcript automatically
'============================================================================
'add headers to transcript
    'go to RoughBKMK
    'find each heading 1/2
    'hit home key
    'insert continuing page break
    'rinse and repeat
    'go to beginning
    'go to beginning of each section
    'view header
    'insert header "heading -- witness name"
    'stop at CertBMK
'save & close

Dim oWordApp As Word.Application, oWordDoc As Word.Document
Dim oWordField As Field, oWordSection As Section
Dim x As Integer, y As Integer, z As Integer
Dim sFileName As String, sBookmarkName As String, sHeading As String
Dim sHeadings() As String
Dim oRange As Range
sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
'Call pfCurrentCaseInfo

sFileName = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-CourtCover.docx"

On Error Resume Next
Set oWordApp = GetObject(, "Word.Application")
If Err <> 0 Then
    Set oWordApp = CreateObject("Word.Application")
End If

Set oWordDoc = oWordApp.Documents.Open(sFileName)
oWordDoc.Application.Visible = True

'go to roughbkmk, beginning of body
oWordDoc.Application.Selection.Goto What:=wdGoToBookmark, Name:="RoughBKMK"
oWordDoc.Application.Selection.Find.ClearFormatting

sHeadings = oWordDoc.GetCrossReferenceItems(wdRefTypeHeading)


'set find style, not necessary for now
'oWordDoc.Application.Selection.Find.Style = ActiveDocument.Styles("Heading 1")
Debug.Print oWordDoc.Fields.Count & " headings"
x = 1
y = 1

' Loop through fields/cross-references in transcript
Debug.Print UBound(sHeadings) & " headings in document"
sHeading = sHeadings(x)
Debug.Print "Current Heading: " & sHeading


'go to first heading
oWordDoc.Application.Selection.Goto What:=wdGoToHeading, which:=wdGoToFirst, Count:=1, Name:=""
        
        
    For x = 1 To UBound(sHeadings)
    
        'go to beginning of its page
        If oWordDoc.Application.Selection.Range.Information(wdActiveEndPageNumber) = 1 Then
            oWordDoc.Application.Selection.Goto wdGoToPage, wdGoToNext
            oWordDoc.Application.Selection.Goto wdGoToPage, wdGoToPrevious
        Else
            oWordDoc.Application.Selection.Goto wdGoToPage, wdGoToPrevious
            oWordDoc.Application.Selection.Goto wdGoToPage, wdGoToNext
        End If
        
        'insert continuous break
        oWordDoc.Application.Selection.InsertBreak Type:=wdSectionBreakContinuous
        
        
        'edit header for this section
        oWordDoc.Application.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        
        'unlink to previous section
        If oWordDoc.Application.Selection.HeaderFooter.LinkToPrevious = Selection.HeaderFooter. _
            LinkToPrevious Then
            oWordDoc.Application.Selection.HeaderFooter.LinkToPrevious = Not Selection.HeaderFooter. _
            LinkToPrevious
        End If
        
        'exit header
        oWordDoc.Application.Selection.EscapeKey
        
        oWordDoc.Application.Selection.Goto What:=wdGoToHeading, which:=wdGoToNext, Count:=1, Name:=""
        
        oWordDoc.Application.Selection.Goto What:=wdGoToHeading, which:=wdGoToNext, Count:=1, Name:=""
    Next
    
    oWordDoc.Application.Selection.HomeKey Unit:=wdStory
x = 1
y = 1
' Loop through fields/cross-references in transcript
Debug.Print UBound(sHeadings) & " headings in document"
Debug.Print "next field:  " & x & sHeadings(x)

'go to first heading
oWordDoc.Application.Selection.Goto What:=wdGoToHeading, which:=wdGoToFirst, Count:=1, Name:=""
        
    For x = 1 To UBound(sHeadings)
        sHeading = sHeadings(x)
    
        Debug.Print "Current Heading: " & sHeading
        
        
        'edit header for this section
        oWordDoc.Application.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    
        'keep header single spaced, centered, and NOT underlined
        oWordDoc.Application.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        oWordDoc.Application.Selection.Font.Underline = wdUnderlineNone
        oWordDoc.Application.Selection.ParagraphFormat.LineSpacing = LinesToPoints(32888)
        
        'go to first heading
        '***WORKING COPY***
        oWordDoc.Application.Selection.TypeText Text:="***WORKING COPY***" & Chr(10)
        
        'insert heading and " -- "
        oWordDoc.Application.Selection.InsertCrossReference ReferenceType:="Heading", ReferenceKind:= _
                wdContentText, ReferenceItem:=x, InsertAsHyperlink:=True, _
                IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
        oWordDoc.Application.Selection.TypeText Text:=" -- WITNESSNAME"
        
        'exit header
        oWordDoc.Application.Selection.EscapeKey

        'go to next heading starting from next page
        
        oWordDoc.Application.Selection.Goto What:=wdGoToHeading, which:=wdGoToNext, Count:=1, Name:=""
        
        Debug.Print "next field:  " & x + 1 & sHeadings(x + 1)
Next

x = 1
For Each oWordSection In oWordDoc.Sections
    Set oRange = oWordSection
    oWordDoc.Sections(x).Headers(wdHeaderFooterPrimary).Application.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    oWordDoc.Sections(x).Headers(wdHeaderFooterPrimary).Application.Selection.ParagraphFormat.LineSpacing = LinesToPoints(32888)
    oWordDoc.Sections(x).Headers(wdHeaderFooterPrimary).Application.Selection.Font.Underline = wdUnderlineNone
    x = x + 1
Next
    
    
    
HeadersDone:
'save & close
oWordDoc.SaveAs2 FileName:=sFileName
oWordDoc.Close
oWordApp.Quit
Set oWordDoc = Nothing
Set oWordApp = Nothing
            
    
End Sub


Sub pfHeaders()
'============================================================================
' Name        : pfHeaders
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfHeaders
' Description : add sections and headers programmatically
'============================================================================

Dim astrHeadings As Variant
Dim rCurrentSection As Range, HdrRange As Range
Dim bFound As Boolean
Dim oWordDoc As New Word.Document, oWordApp As New Word.Application
Dim sFileName As String, sCurrentSection As String, sCurrentHeading As String
Dim intItem As Integer, iCurrentSectionNo As Integer, intLevel As Integer
Dim aStyleList() As String, sStyleName As String
Dim StyleName As Variant, Header As Variant
Dim sListStyle As String
bFound = True
sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
Debug.Print ("---------------------------------------------")
ReDim aStyleList(1 To 1) As String
Dim x As Integer, z As Integer, k As Integer
Dim oHF As HeaderFooter
Dim iMaxHeadingsCount As Integer
Dim sec As Word.Section

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

sFileName = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-CourtCover.docx"

On Error Resume Next
Set oWordApp = GetObject(, "Word.Application")

If Err <> 0 Then
    Set oWordApp = CreateObject("Word.Application")
End If

Set oWordDoc = GetObject(sFileName, "Word.Document")

oWordDoc.Application.Visible = True


With oWordDoc.Application


    astrHeadings = oWordDoc.GetCrossReferenceItems(wdRefTypeHeading)
    
    For intItem = LBound(astrHeadings) To UBound(astrHeadings)
    
        sCurrentHeading = astrHeadings(intItem)
        intLevel = GetLevel(CStr(astrHeadings(intItem)))
        
        sStyleName = "Heading " & intLevel
        
        Debug.Print ("Heading Level:  " & intLevel & ", " & sCurrentHeading)
        
        .Selection.Find.ClearFormatting
        .Selection.Find.Replacement.ClearFormatting
        
        .Selection.Goto What:=wdGoToHeading, which:=wdGoToNext, Count:=1
        sStyleName = "Heading " & intLevel
        'aStyleList(intLevel) = sStyleName
    
        aStyleList(intItem) = sStyleName
        sStyleName = "Heading " & intLevel
        For Each StyleName In aStyleList
            Debug.Print StyleName
        Next
        
        ReDim Preserve aStyleList(1 To UBound(aStyleList) + 1) As String
        
        With .Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
    
        'Ctrl Page up once
        .Selection.GoToPrevious wdGoToPage
        
        'page down once
        .browser.Next
        
        'press home once
        .Selection.HomeKey Unit:=wdLine
        
        'insert continuous section break
        'Selection.InsertBreak Type:=wdSectionBreakContinuous
        .Selection.Paragraphs(1).Range.InsertBreak Type:=wdSectionBreakContinuous
        
        '.Selection.HeaderFooter.LinkToPrevious = False
    
        .Selection.Find.ClearFormatting
        .Selection.Find.Replacement.ClearFormatting
        
        .Selection.Goto What:=wdGoToHeading, which:=wdGoToNext, Count:=1
        sStyleName = "Heading " & intLevel
        
        With .Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Debug.Print ("---------------------------------------------")
    Next
    
    intItem = 1
    
    For intItem = 1 To oWordDoc.Sections.Count
        'add headers to each section
        SendKeys "^{HOME}"
        
        
        For Each StyleName In aStyleList
                
            For Each sec In oWordDoc.Sections
                                
                With sec.Headers(wdHeaderFooterPrimary)
                        
                    'header formatting
                    
                    '.Selection.HeaderFooter.LinkToPrevious = False
                                       
                    .LinkToPrevious = False
                    
                    
                End With
                             
            Next sec
            
        Next StyleName
                
    Next
            
            
    oWordDoc.Application.Selection.HomeKey Unit:=wdStory
    
    
            
    intItem = 2
    
    oWordDoc.Application.Selection.Goto What:=wdGoToHeading, which:=wdGoToNext, Count:=1
    
    'For intItem = 2 To oWordDoc.Sections.Count + 1
    
                       
        Dim iHeadingsNumber As Integer, iSectionNumber As Integer, iSectionIndex As Integer
        astrHeadings = oWordDoc.GetCrossReferenceItems(wdRefTypeHeading)
        
    
            For Each sec In oWordDoc.Sections
            
                iSectionIndex = sec.Index
                iHeadingsNumber = iSectionIndex - 1
                iSectionNumber = iSectionIndex
                        
                If iHeadingsNumber > UBound(astrHeadings) Then GoTo NextItem
                
                If iHeadingsNumber = 0 Then iHeadingsNumber = 1
                
                If UBound(astrHeadings) = 0 Then GoTo NextItem
                
                sCurrentHeading = astrHeadings(iHeadingsNumber)
                intLevel = GetLevel(CStr(astrHeadings(iHeadingsNumber)))
                sStyleName = "Heading " & intLevel
                
                iMaxHeadingsCount = UBound(astrHeadings)
                
                'add headers to each section
                        
                If iSectionNumber <= iMaxHeadingsCount + 1 Then
                        
                    sCurrentHeading = astrHeadings(iHeadingsNumber)
                    intLevel = GetLevel(CStr(astrHeadings(iHeadingsNumber)))
                    
                    sStyleName = "Heading " & intLevel
                                        
                    iSectionIndex = sec.Index
                    Debug.Print ("Section Number:  " & iSectionIndex & "   |   " & "Headings Number:  " & iHeadingsNumber)
                    If iSectionNumber = 1 Then GoTo SkipFrontPage
                                                                 
                    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
                        ActiveWindow.Panes(2).Close
                    End If
                    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
                        ActivePane.View.Type = wdOutlineView Then
                        ActiveWindow.ActivePane.View.Type = wdPrintView
                    End If
                    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
                    
                    With oWordDoc.Application
                        Selection.TypeText Text:="***WORKING COPY***"
                        Selection.Collapse Direction:=wdCollapseEnd
                        Selection.TypeParagraph
                        Selection.InsertCrossReference ReferenceType:="Heading", ReferenceKind:= _
                            wdContentText, ReferenceItem:=iHeadingsNumber, InsertAsHyperlink:=True, _
                            IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
                        
                        If sStyleName = "Heading 2" Then Selection.TypeText Text:=" -- WITNESSNAME"
                        
                        Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
                        Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
                        Selection.Find.ClearFormatting
                        With Selection.Find
                            .Text = ""
                            .Replacement.Text = ""
                            .Forward = True
                            .Wrap = wdFindContinue
                            .Format = False
                            .MatchCase = False
                            .MatchWholeWord = False
                            .MatchWildcards = False
                            .MatchSoundsLike = False
                            .MatchAllWordForms = False
                        End With
                        Selection.Style = ActiveDocument.Styles("AQC-Header")
                    
                    End With
                    
                    oWordDoc.Application.Selection.Goto What:=wdGoToHeading, which:=wdGoToNext, Count:=2
                    
                End If
SkipFrontPage:
                With sec
                
                .Footers(wdHeaderFooterPrimary).Range.Text = "www.aquoco.co   |   inquiries@aquoco.co"
                .Footers(wdHeaderFooterPrimary).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                End With
                
            Next sec
NextItem:
    'Next intItem
            
End With

oWordDoc.SaveAs2 FileName:=sFileName
oWordDoc.Close
Set oWordApp = Nothing
Set oWordDoc = Nothing
Set rCurrentSection = Nothing

    
End Sub

Sub pfTopOfTranscriptBookmark()

Dim AcroApp As Acrobat.CAcroApp
Dim PDoc As Acrobat.CAcroPDDoc
Dim PDocAll As Acrobat.CAcroPDDoc
Dim PDocCover As Acrobat.CAcroPDDoc
Dim ADoc As AcroAVDoc
Dim PDBookmark As AcroPDBookmark
Dim PDFPageView As AcroAVPageView
Dim bTitle, n, sTranscriptsFolderFinalPDF As String
Dim sTranscriptVolumesALLPath As String, sVolumesCoverPath  As String
Dim oPDFBookmarks As Object, parentBookmark As AcroPDBookmark
Dim jso As Object, BookMarkRoot As Object
Dim numpages


Set AcroApp = CreateObject("AcroExch.App")

Set PDoc = CreateObject("AcroExch.PDDoc")
Set PDocAll = CreateObject("AcroExch.PDDoc")
Set PDocCover = CreateObject("AcroExch.PDDoc")
Set ADoc = CreateObject("AcroExch.AVDoc")
Set PDBookmark = CreateObject("AcroExch.PDBookmark", "")

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
sTranscriptsFolderFinalPDF = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-FINAL.pdf"
sTranscriptVolumesALLPath = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcripts-All.pdf"
sVolumesCoverPath = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Cover.pdf"

'Table of Contents Bookmark



PDocCover.Open (sVolumesCoverPath)

Set ADoc = PDocCover.OpenAVDoc(sVolumesCoverPath)
Set PDFPageView = ADoc.GetAVPageView()
Call PDFPageView.Goto(0)
AcroApp.MenuItemExecute ("NewBookmark")

bTitle = PDBookmark.GetByTitle(PDocCover, "Untitled")
bTitle = PDBookmark.SetTitle("Table of Contents")

n = PDocCover.Save(PDSaveFull, sVolumesCoverPath)

' Insert the pages of Part2 after the end of Part1
    numpages = PDocCover.GetNumPages()
    
PDoc.Open (sTranscriptsFolderFinalPDF)

Set ADoc = PDoc.OpenAVDoc(sTranscriptsFolderFinalPDF)
SendKeys ("^{HOME}")
Set PDFPageView = ADoc.GetAVPageView()

AppActivate "Adobe Acrobat Pro Extended"

'Top of Transcript Bookmark
Call PDFPageView.Goto(0)
AcroApp.MenuItemExecute ("NewBookmark")

bTitle = PDBookmark.GetByTitle(PDoc, "Untitled")
bTitle = PDBookmark.SetTitle("TOP OF TRANSCRIPT")

'Index Bookmark
Call PDFPageView.Goto(1)
AcroApp.MenuItemExecute ("NewBookmark")

bTitle = PDBookmark.GetByTitle(PDoc, "Untitled")
bTitle = PDBookmark.SetTitle("TRANSCRIPT INDEXES")

'General Index Bookmark
Call PDFPageView.Goto(0)
AcroApp.MenuItemExecute ("NewBookmark")

bTitle = PDBookmark.GetByTitle(PDoc, "Untitled")
bTitle = PDBookmark.SetTitle("General")

'Witnesses Index Bookmark
Call PDFPageView.Goto(0)
AcroApp.MenuItemExecute ("NewBookmark")

bTitle = PDBookmark.GetByTitle(PDoc, "Untitled")
bTitle = PDBookmark.SetTitle("Witnesses")

'Exhibits Index Bookmark
Call PDFPageView.Goto(0)
AcroApp.MenuItemExecute ("NewBookmark")

bTitle = PDBookmark.GetByTitle(PDoc, "Untitled")
bTitle = PDBookmark.SetTitle("Exhibits")

'Cases Bookmark
Call PDFPageView.Goto(0)
AcroApp.MenuItemExecute ("NewBookmark")

bTitle = PDBookmark.GetByTitle(PDoc, "Untitled")
bTitle = PDBookmark.SetTitle("Cases")

'Rules, Regulation, Code, Statutes Bookmark
Call PDFPageView.Goto(0)
AcroApp.MenuItemExecute ("NewBookmark")

bTitle = PDBookmark.GetByTitle(PDoc, "Untitled")
bTitle = PDBookmark.SetTitle("Rules, Regulation, Code, Statutes")

'Other Authorities Bookmark
Call PDFPageView.Goto(0)
AcroApp.MenuItemExecute ("NewBookmark")

bTitle = PDBookmark.GetByTitle(PDoc, "Untitled")
bTitle = PDBookmark.SetTitle("Other Authorities")

MsgBox ("Hit okay after you've moved all bookmarks to their corresponding headings; General, Witnesses, or Exhibits; and also created the Authorities bookmark structure.")

n = PDoc.Save(PDSaveFull, sTranscriptsFolderFinalPDF)



'for each -Transcript-FINAL.pdf in \Transcripts\ do the following

    If PDocCover.InsertPages(numpages - 1, PDoc, 0, PDoc.GetNumPages(), True) = False Then
        MsgBox "Cannot insert pages"
    End If

    If PDocCover.Save(PDSaveFull, sTranscriptVolumesALLPath) = False Then
        MsgBox "Cannot save the modified document"
    End If
    
    
    
    
PDoc.Close
PDocCover.Close
AcroApp.Exit
Set AcroApp = Nothing
Set PDoc = Nothing
Set ADoc = Nothing

End Sub

                
Sub fPDFBookmarks()

On Error GoTo eHandler
'============================================================================
' Name        : fPrint2upPDF
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fPrint2upPDF
' Description : prints 2-up transcript PDF
'============================================================================


Dim sTranscriptsFolderFinalPDF As String

Dim sTranscriptsFolder2upPDF As String

Dim sTranscript2upPSPath As String
Dim sJavascriptPrint As String
Dim jobsettings As String
Dim sLogFilePath As String

Dim aaAcroApp As Acrobat.AcroApp
Dim aaAcroAVDoc As Acrobat.AcroAVDoc
Dim aaAcroPDDoc As Acrobat.AcroPDDoc
Dim bret
Dim pp As Object

Dim pdTranscriptFinalDistiller As PdfDistiller
Dim aaAFormApp As AFORMAUTLib.AFormApp

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
sTranscriptsFolderFinalPDF = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-FINAL.pdf"

Set aaAcroApp = New AcroApp
Set aaAcroAVDoc = CreateObject("AcroExch.AVDoc")

If aaAcroAVDoc.Open(sTranscriptsFolderFinalPDF, "") Then
    aaAcroAVDoc.Maximize (1)
    
    Set aaAcroPDDoc = aaAcroAVDoc.GetPDDoc()
    Set aaAFormApp = CreateObject("AFormAut.App")
    
      sJavascriptPrint = "function MakeBkMks(oBkMkParent, aBkMks) {" & _
            "var aBkMkNames = [ " & Chr(34) & "General" & Chr(34) & ", " & Chr(34) & "Witnesses" & Chr(34) & _
            ", " & Chr(34) & "Exhibits" & Chr(34) & ", " & Chr(34) & "Authorities" & Chr(34) & _
            ", [" & Chr(34) & "Case Law" & Chr(34) & "," & _
            Chr(34) & "Rules, Regulation, Code, Statutes" & Chr(34) & "," & _
            Chr(34) & "Other Authority" & Chr(34) & "] ];" & _
            "for(var index=0;index<aBkMks.length;index++) {" & _
            "if(typeof(aBkMks[index]) == " & Chr(34) & "string" & Chr(34) & ") oBkMkParent.createChild({cName:aBkMks[i], nIndex:index });" & _
            "else {" & _
            "// Assume this is a sub Array" & _
            "oBkMkParent.createChild({cName:aBkMks[index][0], nIndex:index});" & _
            "MakeBkMks(oBkMkParent.children[index], aBkMks[index].slice(1) );}}}}" & _
            "MakeBkMks(this.bookmarkRoot, aBkMkNames);"

    aaAFormApp.Fields.ExecuteThisJavascript sJavascriptPrint
    
    aaAcroPDDoc.Save PDSaveFull, sTranscriptsFolderFinalPDF
    'aaAcroPDDoc.Close
    'aaAcroApp.CloseAllDocs
    
End If

eHandlerX:
Set aaAcroPDDoc = Nothing
Set aaAcroAVDoc = Nothing
Set aaAcroApp = Nothing
MsgBox "PDF Bookmarks Created"
Exit Sub

eHandler:
MsgBox Err.Number & ": " & Err.Description, vbCritical, "Error Detail"
GoTo eHandlerX
Resume

End Sub





