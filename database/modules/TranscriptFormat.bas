Attribute VB_Name = "TranscriptFormat"
'@Folder("Database.Production.Modules")
Option Compare Database
'============================================================================
'class module cmTranscriptFormat

'variables:
'   Private sFileName As String
'   Private oWordApp As Object, oWordDoc As Object
'   Private qdf As QueryDef
'   Private sQueryName As String
'   Private db As Database
'   Public sBookmarkName As String

'functions:

'pfCreateBookmarks:             Description:  replaces #TOC_# notations in transcript with bookmarks and then places index at bookmarks
'                           Arguments:    NONE
'pfReplaceBMKWwithBookmark:     Description:  replaces #__# notations with bookmarks
'                           Arguments:    NONE
'pfApplyStyle:                  Description:  finds specific phrases in activedocument(transcript) and applies a specific style
'                           Arguments:    sStyleName, sTextToFind, sReplacementText
'pfFindRepCitationLinks:        Description:  adds citations and hyperlinks from CitationHyperlinks table in transcript
'                           Arguments:    NONE
'pfCreateIndexesTOAs:           Description:  creates indexes and indexes certain things
'                           Arguments:    NONE
'pfSingleFindReplace:           Description:  find and replace all of one item
'                           Arguments:    sTextToFind, sReplacementText
'                                         Optional wsyWordStyle = "", bForward = True, bWrap = "wdFindContinue"
'                                         Optional bFormat = False, bMatchCase = True, bMatchWholeWord = False
'                                         Optional bMatchWildcards = False, bMatchSoundsLike = False, bMatchAllWordForms = False
'pfReplaceFDA:                  Description:  doctor speaker name find/replaces for FDA transcripts
'                           Arguments:    NONE
'pfDynamicSpeakersFindReplace:  Description:  gets speaker names from ViewJobFormAppearancesQ query and find/replaces in transcript as appropriate
'                           Arguments:    NONE
'pfSingleTCReplaceAll:          Description:  one replace TC entry function for ones with no field entry
'                           Arguments:    sTexttoSearch, sReplacementText
'pfFieldTCReplaceAll:           Description:  one replace TC entry function for ones with field entry
'                           Arguments:    sTexttoSearch, sReplacementText, sFieldText
'pfWordIndexer:                 Description:  builds word index in separate PDF from transcript
'                           Arguments:    NONE
'FPJurors:                      Description:  does find/replacements of prospective jurors in transcript
'                           Arguments:    NONE
'pfTCEntryReplacement:          Description:  parent function that finds certain entries within a transcript and assigns TC entries to them for indexing purposes
'                           Arguments:    NONE
'pfFindRepCitationLinks:        Description:  'originally named fEfficientCiteSearch 'old one now named pfFindRepCitationLinks1
'find citation markings like phonetic in transcript
'list separately so marking doesn't take so long
'                           Arguments:    NONE
        
'============================================================================

Private sFileName As String
Private oWordApp As New Word.Application
Private oWordDoc As New Word.Document
Private qdf As QueryDef
Private sQueryName As String
Private db As Database
Public sBookmarkName As String

Public Sub test1()
    '============================================================================
    ' Name        : pfTCEntryReplacement
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfTCEntryReplacement
    ' Description : parent function that finds certain entries within a transcript and assigns TC entries to them for indexing purposes
    '============================================================================
    
    Dim sMrMs2 As String
    Dim sLastName2 As String
    Dim vSpeakerName As String
    
    Dim oWordApp As New Word.Application
    Dim oCourtCoverWD As New Word.Document
    
    Dim rstTRCourtQ As DAO.Recordset
    Dim rstViewJFAppQ As DAO.Recordset
    Dim qdf As QueryDef

    Dim cJob As New Job
    
    DoCmd.OpenQuery qnViewJobFormAppearancesQ, acViewNormal, acReadOnly 'open query
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField] 'job number
    '@Ignore AssignmentNotUsed
    Set oWordApp = CreateObject("Word.Application")

    On Error Resume Next
    Set oWordApp = GetObject(, "Word.Application")
    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If
    oWordApp.Visible = False
    
    Set oCourtCoverWD = oWordApp.Documents.Open(cJob.DocPath.CourtCover)
    Call pfFieldTCReplaceAll("(ee1)", "^p" & "DIRECT EXAMINATION" & "^p", Chr(34) & "Direct Examination by " & Chr(34) & " \l 3")
    Call pfFieldTCReplaceAll("(ee2)", "^p" & "CROSS-EXAMINATION" & "^p", Chr(34) & "Cross-Examination by " & Chr(34) & " \l 3")
    Call pfFieldTCReplaceAll("(ee3)", "^p" & "REDIRECT EXAMINATION" & "^p", Chr(34) & "Redirect Examination by " & Chr(34) & " \l 3")
    Call pfFieldTCReplaceAll("(ee4)", "^p" & "RECROSS-EXAMINATION" & "^p", Chr(34) & "Recross-Examination by " & Chr(34) & " \l 3")
    Call pfFieldTCReplaceAll("(ee5)", "^p" & "FURTHER REDIRECT EXAMINATION" & "^p", Chr(34) & "Further Redirect Examination by " & Chr(34) & " \l 3")
    Call pfFieldTCReplaceAll("(ee6)", "^p" & "FURTHER RECROSS-EXAMINATION" & "^p", Chr(34) & "Further Recross-Examination by " & Chr(34) & " \l 3")
    Call pfFieldTCReplaceAll("(e1c)", "^p" & "DIRECT EXAMINATION CONTINUED" & "^p", Chr(34) & "Direct Examination Continued by " & Chr(34) & " \l 3")
    Call pfFieldTCReplaceAll("(e2c)", "^p" & "CROSS-EXAMINATION CONTINUED" & "^p", Chr(34) & "Cross-Examination Continued by " & Chr(34) & " \l 3")
    Call pfFieldTCReplaceAll("(e3c)", "^p" & "REDIRECT EXAMINATION CONTINUED" & "^p", Chr(34) & "Redirect Examination Continued by " & Chr(34) & " \l 3")
    Call pfFieldTCReplaceAll("(e4c)", "^p" & "RECROSS-EXAMINATION CONTINUED" & "^p", Chr(34) & "Recross-Examination Continued by " & Chr(34) & " \l 3")
    Call pfFieldTCReplaceAll("(e5c)", "^p" & "FURTHER REDIRECT EXAMINATION CONTINUED" & "^p", Chr(34) & "Further Redirect Examination Continued by " & Chr(34) & " \l 3")
    Call pfFieldTCReplaceAll("(e6c)", "^p" & "FURTHER RECROSS-EXAMINATION CONTINUED" & "^p", Chr(34) & "Further Recross-Examination Continued by " & Chr(34) & " \l 3")
    'Public Function pfFieldTCReplaceAll(sTexttoSearch As String, sReplacementText As String, sFieldText As String)
    '.Fields.Add Type:=wdFieldTOCEntry, Text:=sFieldText, PreserveFormatting:=False, Range:=.Range 'sFieldText sample = "TC ""WitnessName"" \l 2-3"
        
    rstViewJFAppQ.Close
    Set rstViewJFAppQ = Nothing
    oCourtCoverWD.SaveAs2 FileName:=cJob.DocPath.CourtCover
    oCourtCoverWD.Close
    oWordApp.Quit
    On Error GoTo 0
    Set oCourtCoverWD = Nothing
    Set oWordApp = Nothing
            
End Sub

Public Sub pfReplaceBMKWwithBookmark()
    '============================================================================
    ' Name        : pfReplaceBMKWwithBookmark
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfReplaceBMKWwithBookmark
    ' Description : replaces #__# notations with bookmarks
    '============================================================================

    Dim sBookmarkName As String

    ActiveDocument.Application.Selection.Find.ClearFormatting
    ActiveDocument.Application.Selection.Find.Replacement.ClearFormatting

    With ActiveDocument.Application.Selection.Find
        sBookmarkName = "RoughBKMK"
        .Text = "#RBMK#"
        ActiveDocument.bookmarks.Add Name:=sBookmarkName
        sBookmarkName = "CertBMK"
        .Text = "#CBMK#"
        ActiveDocument.bookmarks.Add Name:=sBookmarkName
        sBookmarkName = "ToABMK"
        .Text = "#TBMK#"
        ActiveDocument.bookmarks.Add Name:=sBookmarkName
        sBookmarkName = "TopLine"
        .Text = "#TOPL#"
        ActiveDocument.bookmarks.Add Name:=sBookmarkName
        sBookmarkName = "EndTime"
        .Text = "#ENDT#"
        ActiveDocument.bookmarks.Add Name:=sBookmarkName
    End With
    With ActiveDocument                          'insert topline at TopLine bookmark

        If .bookmarks.Exists("TopLine") = True Then
    
            'sHearingLocation, sStartTime, sEndTime
            .bookmarks("TopLine").Select
            .Application.Selection.TypeText Text:=UCase(sHearingLocation) & ", " & _
                                                                          FormatDateTime(dHearingDate, vbLongDate) & ", " & UCase(sStartTime)
        Else
            MsgBox "Bookmark ""TopLine"" does not exist!"
        End If
        .MailMerge.MainDocumentType = wdNotAMergeDocument

        If .bookmarks.Exists("EndTime") = True Then
    
            If Right(sEndTime, 2) = "AM" Then
                sEndTime = Replace(sEndTime, "AM", "a.m.")
        
            ElseIf Right(sEndTime, 2) = "PM" Then
                sEndTime = Replace(sEndTime, "PM", "p.m.")
        
            End If
    
    
            'sHearingLocation, sStartTime, sEndTime
            .bookmarks("EndTime").Select
            .Application.Selection.TypeText Text:=UCase(sEndTime)
        Else
            MsgBox "Bookmark ""EndTime"" does not exist!"
        End If
        .MailMerge.MainDocumentType = wdNotAMergeDocument

    End With

End Sub

Public Sub pfCreateBookmarks()
    '============================================================================
    ' Name        : pfCreateBookmarks
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfCreateBookmarks
    ' Description : replaces #TOC_# notations in transcript with bookmarks and then places index at bookmarks
    '============================================================================

    Dim sBookmarkName As String
    Dim vBookmarkName As String
            Dim sTopLine As String
    
    Dim cJob As New Job

    'oWordDoc.Activate
    On Error Resume Next
    Set oWordApp = GetObject(, "Word.Application")
    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If

    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.CourtCover)
    On Error GoTo 0


    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting

    With oWordDoc.Application.Selection.Find
        .Text = "#TOPOFT#"
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
    End With
    sBookmarkName = "TopOfT"
    oWordDoc.bookmarks.Add Name:=sBookmarkName


    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting

    With oWordDoc.Application.Selection.Find
        .Text = "#RBMK#"
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
    End With
    sBookmarkName = "RoughBKMK"
    oWordDoc.bookmarks.Add Name:=sBookmarkName



    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting

    With oWordDoc.Application.Selection.Find
        .Text = "#TOPL#"
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
    End With
    sBookmarkName = "TopLine"
    oWordDoc.bookmarks.Add Name:=sBookmarkName

    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting

    With oWordDoc.Application.Selection.Find
        .Text = "#ENDT#"
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
    End With
    sBookmarkName = "EndTime"
    oWordDoc.bookmarks.Add Name:=sBookmarkName

    oWordApp.Application.Selection.Find.ClearFormatting
    oWordApp.Application.Selection.Find.Replacement.ClearFormatting
    '

    With oWordDoc.Application.Selection.Find
        .Text = "#CBMK#"
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
    End With
    sBookmarkName = "CertBMK"
    oWordDoc.bookmarks.Add Name:=sBookmarkName


    oWordApp.Application.Selection.Find.ClearFormatting
    oWordApp.Application.Selection.Find.Replacement.ClearFormatting

    With oWordDoc.Application.Selection.Find
        .Text = "#TBMK#"
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
    End With
    sBookmarkName = "ToABMK"
    oWordDoc.bookmarks.Add Name:=sBookmarkName

    oWordApp.Application.Selection.Find.ClearFormatting
    oWordApp.Application.Selection.Find.Replacement.ClearFormatting

    With oWordDoc.Application.Selection.Find
        .Text = "#TOCA#"
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
    End With
    sBookmarkName = "IndexA"
    oWordDoc.bookmarks.Add Name:=sBookmarkName

    sBookmarkName = "IndexB"

    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting
    With oWordDoc.Application.Selection.Find
        .Text = "#TOCB#"
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
    End With

    oWordDoc.bookmarks.Add Name:=sBookmarkName

    sBookmarkName = "IndexC"

    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting
    With oWordDoc.Application.Selection.Find
        .Text = "#TOCC#"
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
    End With

    oWordDoc.bookmarks.Add Name:=sBookmarkName

    sBookmarkName = "IndexD"

    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting
    With oWordDoc.Application.Selection.Find
        .Text = "#TOCD#"
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
    End With

    oWordDoc.bookmarks.Add Name:=sBookmarkName

    sBookmarkName = "IndexE"

    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting
    With oWordDoc.Application.Selection.Find
        .Text = "#TOCE#"
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
    End With

    oWordDoc.bookmarks.Add Name:=sBookmarkName


    vBookmarkName = "TOAC"

    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting

    With oWordDoc.Application.Selection.Find

        .Text = "#TOAC#"
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
    End With

    oWordDoc.bookmarks.Add Name:=vBookmarkName



    vBookmarkName = "TOAR"
    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting

    With oWordDoc.Application.Selection.Find

        .Text = "#TOAR#"
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
    End With

    oWordDoc.bookmarks.Add Name:=vBookmarkName

    vBookmarkName = "TOAO"
    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting

    With oWordDoc.Application.Selection.Find
        .Text = "#TOAO#"
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
    End With

    oWordDoc.bookmarks.Add Name:=vBookmarkName

    With oWordDoc                                'insert topline at TopLine bookmark

        If .bookmarks.Exists("TopLine") = True Then
    
            If Right(sStartTime, 6) = ":00 AM" Then
                sStartTime = Replace(sStartTime, ":00 AM", " a.m.")
        
            ElseIf Right(sStartTime, 6) = ":00 PM" Then
                sStartTime = Replace(sStartTime, ":00 PM", " p.m.")
        
            End If
            
            sTopLine = UCase(sHearingLocation) & ", " & UCase(FormatDateTime(dHearingDate, vbLongDate)) & ", " & UCase(sStartTime)
            'sHearingLocation, sStartTime, sEndTime
            .bookmarks("TopLine").Select
            .Application.Selection.Font.Underline = wdUnderlineSingle
            .Application.Selection.TypeText Text:=sTopLine
            .Application.Selection.Font.Underline = wdUnderlineNone
        Else
            MsgBox "Bookmark ""TopLine"" does not exist!"
        End If
        .MailMerge.MainDocumentType = wdNotAMergeDocument

        If .bookmarks.Exists("EndTime") = True Then
    
            If Right(sEndTime, 6) = ":00 AM" Then
                sEndTime = Replace(sEndTime, ":00 AM", " a.m.")
        
            ElseIf Right(sEndTime, 6) = ":00 PM" Then
                sEndTime = Replace(sEndTime, ":00 PM", " p.m.")
        
            End If
    
            'sHearingLocation, sStartTime, sEndTime
            .bookmarks("EndTime").Select
            .Application.Selection.TypeText Text:=sEndTime
        Else
            MsgBox "Bookmark ""EndTime"" does not exist!"
        End If
        .MailMerge.MainDocumentType = wdNotAMergeDocument

    End With

    oWordDoc.SaveAs2 FileName:=cJob.DocPath.CourtCover
    oWordDoc.Close (wdSaveChanges)




End Sub

Public Sub pfApplyStyle(sStyleName As String, sTextToFind As String, sReplacementText As String)
    '============================================================================
    ' Name        : pfApplyStyle
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfApplyStyle(sStyleName, sTextToFind, sReplacementText)
    ' Description : finds specific phrases in oWordDoc(transcript) and applies a specific style
    '============================================================================
    Dim oWordApp As New Word.Application
    Dim oWordDoc As New Word.Document
    
    'Set oWordApp = GetObject(, "Word.Application")

    'If Err <> 0 Then
    '    Set oWordApp = CreateObject("Word.Application")
    'End If

    'oWordApp.Activate
    'Set oWordDoc = oWordApp.Documents.Add(sCourtDatesID & "-CourtCover.docx")
    Set oWordDoc = GetObject(sCourtDatesID & "-CourtCover.docx", "Word.Document")

    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.Style = oWordDoc.Styles(sStyleName)

    With oWordDoc.Application.Selection.Find
        .Text = sTextToFind
        .Replacement.Text = sReplacementText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

    oWordDoc.Application.Selection.Find.Execute Replace:=wdReplaceAll


    oWordDoc.SaveAs2 FileName:=sFileName

End Sub

Public Sub pfCreateIndexesTOAs()
    '============================================================================
    ' Name        : pfCreateIndexesTOAs
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfCreateIndexesTOAs
    ' Description : creates indexes and indexes certain things
    '============================================================================
        
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

    Dim oWordApp As New Word.Application
    Dim oWordDoc As New Word.Document
    
    Dim cJob As New Job

    On Error Resume Next
    Set oWordApp = GetObject(, "Word.Application")

    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0

    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.CourtCover)

    With oWordDoc
        .Application.Selection.Goto What:=wdGoToBookmark, Name:="IndexA"
        With oWordDoc.bookmarks
            .DefaultSorting = wdSortByName
            .ShowHidden = False
        End With
        .TablesOfContents.Add Range:=.Application.Selection.Range, UseHeadingStyles:=False, UseFields:=True, TableID:="a", RightAlignPageNumbers:=True, IncludePageNumbers:=True, UseHyperlinks:=True
        '.indexes.Add Range:=.Application.Selection.Range, HeadingSeparator:= _
        'wdHeadingSeparatorNone, Type:=wdIndexRunin, NumberOfColumns:= _
        'wdAutoPosition, IndexLanguage:=wdEnglishUS
        '.indexes(1).TabLeader = wdTabLeaderDots
        .Application.Selection.Goto What:=wdGoToBookmark, Name:="IndexB"
        With oWordDoc.bookmarks
            .DefaultSorting = wdSortByName
            .ShowHidden = False
        End With
        With oWordDoc
            .TablesOfContents.Add Range:=.Application.Selection.Range, UseHeadingStyles:=False, UseFields:=True, TableID:="b", RightAlignPageNumbers:=True, IncludePageNumbers:=True, UseHyperlinks:=True
            '.indexes.Add Range:=.Application.Selection.Range, HeadingSeparator:= _
            'wdHeadingSeparatorNone, Type:=wdIndexRunin, NumberOfColumns:= _
            'wdAutoPosition, IndexLanguage:=wdEnglishUS
            '.indexes(1).TabLeader = wdTabLeaderDots
        End With
        .Application.Selection.Goto What:=wdGoToBookmark, Name:="IndexC"
        With oWordDoc.bookmarks
            .DefaultSorting = wdSortByName
            .ShowHidden = False
        End With
        With oWordDoc
            .TablesOfContents.Add Range:=.Application.Selection.Range, UseHeadingStyles:=False, UseFields:=True, TableID:="c", RightAlignPageNumbers:=True, IncludePageNumbers:=True, UseHyperlinks:=True
            '.indexes.Add Range:=.Application.Selection.Range, HeadingSeparator:= _
            'wdHeadingSeparatorNone, Type:=wdIndexRunin, NumberOfColumns:= _
            'wdAutoPosition, IndexLanguage:=wdEnglishUS
            '.indexes(1).TabLeader = wdTabLeaderDots
        End With
        .Application.Selection.Goto What:=wdGoToBookmark, Name:="IndexD"
        With oWordDoc.bookmarks
            .DefaultSorting = wdSortByName
            .ShowHidden = False
        End With
        With oWordDoc
            .TablesOfContents.Add Range:=.Application.Selection.Range, UseHeadingStyles:=False, UseFields:=True, TableID:="d", RightAlignPageNumbers:=True, IncludePageNumbers:=True, UseHyperlinks:=True
            '.indexes.Add Range:=.Application.Selection.Range, HeadingSeparator:= _
            'wdHeadingSeparatorNone, Type:=wdIndexRunin, NumberOfColumns:= _
            'wdAutoPosition, IndexLanguage:=wdEnglishUS
            '.indexes(1).TabLeader = wdTabLeaderDots
        End With
        .Application.Selection.Goto What:=wdGoToBookmark, Name:="IndexE"
        With oWordDoc.bookmarks
            .DefaultSorting = wdSortByName
            .ShowHidden = False
        End With
        .TablesOfContents.Add Range:=.Application.Selection.Range, UseHeadingStyles:=False, UseFields:=True, TableID:="e", RightAlignPageNumbers:=True, IncludePageNumbers:=True, UseHyperlinks:=True
        '.indexes.Add Range:=.Application.Selection.Range, HeadingSeparator:= _
        'wdHeadingSeparatorNone, Type:=wdIndexRunin, NumberOfColumns:= _
        'wdAutoPosition, IndexLanguage:=wdEnglishUS
        '.indexes(1).TabLeader = wdTabLeaderDots
    
     
        If InStr(sJurisdiction, "AVT") = 0 And InStr(sJurisdiction, "eScribers") = 0 And InStr(sJurisdiction, "FDA") = 0 And InStr(sJurisdiction, "Food and Drug Administration") = 0 And InStr(sJurisdiction, "Weber") = 0 Then
                
            
            oWordDoc.Application.Selection.Goto What:=wdGoToBookmark, Name:="TOAC"
            With oWordDoc.bookmarks
                .DefaultSorting = wdSortByName
                .ShowHidden = False
            End With
            .TablesOfAuthorities.Add Range:=oWordDoc.Application.Selection.Range, Category:=1, Passim _
                                     :=False, KeepEntryFormatting:=True
            .TablesOfAuthorities(1).TabLeader = wdTabLeaderDots
            .TablesOfAuthorities.Format = wdIndexIndent
             
            
             
            oWordDoc.Application.Selection.Goto What:=wdGoToBookmark, Name:="TOAR"
            With oWordDoc.bookmarks
                .DefaultSorting = wdSortByName
                .ShowHidden = False
            End With
            .TablesOfAuthorities.Add Range:=oWordDoc.Application.Selection.Range, Category:=2, Passim _
                                     :=False, KeepEntryFormatting:=True
            .TablesOfAuthorities(1).TabLeader = wdTabLeaderDots
            .TablesOfAuthorities.Format = wdIndexIndent
             
             
            oWordDoc.Application.Selection.Goto What:=wdGoToBookmark, Name:="TOAO"
            With oWordDoc.bookmarks
                .DefaultSorting = wdSortByName
                .ShowHidden = False
            End With
             
            .TablesOfAuthorities.Add Range:=.Application.Selection.Range, Category:=3, Passim _
                                     :=False, KeepEntryFormatting:=True
            .TablesOfAuthorities(1).TabLeader = wdTabLeaderDots
            .TablesOfAuthorities.Format = wdIndexIndent
             
        Else
        End If
        
        With oWordDoc
            With .Styles("TOA Heading")
                .AutomaticallyUpdate = False
                .BaseStyle = "Normal"
                .NextParagraphStyle = "Normal"
            End With
            With .Styles("TOA Heading").Font
                .Name = "Courier"
                .Size = 12
                .Bold = True
                .Italic = False
                .Underline = wdUnderlineSingle
                .UnderlineColor = wdColorAutomatic
            End With
            .TablesOfAuthoritiesCategories(2).Name = "Rules, Regulation, Code, Statutes"
        End With


        'TOC \f e
        With oWordDoc.Application.Selection.Find
            .Text = "TOC \f e"
            .Replacement.Text = "TOC \l 2-3"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        oWordDoc.Application.Selection.Find.Execute Replace:=wdReplaceAll
    
        With oWordDoc.Application.Selection.Find
            .Text = "For the :" & "^p"
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
        oWordDoc.Application.Selection.Find.Execute Replace:=wdReplaceAll
    
    
        With oWordDoc.Application.Selection.Find
            .Text = "By:   , ESQ." & "^p"
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
        oWordDoc.Application.Selection.Find.Execute Replace:=wdReplaceAll
    
        oWordDoc.SaveAs2 FileName:=cJob.DocPath.CourtCover
        
        
    End With
End Sub

Public Sub pfReplaceFDA()
    '============================================================================
    ' Name        : pfReplaceFDA
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfReplaceFDA
    ' Description : doctor speaker name find/replaces for FDA transcripts
    '============================================================================

    Dim ReplaceWithName As String
    Dim FindFDA1 As String
    Dim FindFDA3 As String
    Dim QueryName As String
    Dim FindFDA2 As String
    Dim FindFDA4 As String
    
    Dim oWordApp As New Word.Application
    Dim oWordDoc As New Word.Document
    
    Dim rs As DAO.Recordset
    Dim rs1 As DAO.Recordset
    Dim qdf As QueryDef
    
    Dim cJob As New Job

    Call pfCurrentCaseInfo                       'refresh transcript info

    'Set oWordApp = CreateObject("Word.Application")
    oWordApp.Visible = False
    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.CourtCover)
    If sJurisdiction = "Food and Drug Administration" Then
        'run a query to pull doctors FDA find
        QueryName = "Q-Doctors"
                
        Set qdf = CurrentDb.QueryDefs(QueryName)
        qdf.Parameters(0) = sCourtDatesID
        Set rs1 = qdf.OpenRecordset
        'open Word document
        'run another query to pull up FDA finds
        QueryName = "QFDA1FindReplaceShortcuts"
        Set rs = CurrentDb.OpenRecordset(QueryName)
                
        If Not (rs.EOF And rs.BOF) Then
        
            rs.Move (1)
        
            Do Until rs.EOF = True
        
                FindFDA1 = rs!ID
                FindFDA2 = "L" & FindFDA1 - 1
                FindFDA3 = FindFDA2 & ": "
            
                If FindFDA1 = 72 Then
                    Exit Do
                Else
                End If
            
                FindFDA4 = rs1.Fields(FindFDA2).Value
                ReplaceWithName = "DR" & Chr(46) & Chr(32) & FindFDA4 & Chr(58) & Chr(32) & Chr(32)
                With oWordDoc.Application.Selection.Find
                    .Text = FindFDA3
                    .Replacement.Text = ReplaceWithName
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                oWordDoc.Application.Selection.Find.Execute Replace:=wdReplaceAll

                'cleaning up
                FindFDA1 = ""
                FindFDA2 = ""
                FindFDA3 = ""
                FindFDA4 = ""
                ReplaceWithName = ""
            
                rs.MoveNext
            Loop
        Else
            MsgBox "There are no records in the recordset."
        End If
    Else
    End If
    
    oWordDoc.Save
    oWordDoc.Close
    rs.Close                                     'Close the recordset
    rs1.Close
    oWordApp.Quit
    Set oWordDoc = Nothing
    Set rs = Nothing                             'Clean up
    Set rs1 = Nothing                            'Clean up
    Set oWordApp = Nothing
    Call pfClearGlobals
End Sub

Public Sub pfDynamicSpeakersFindReplace()
    '============================================================================
    ' Name        : pfDynamicSpeakersFindReplace
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfDynamicSpeakersFindReplace
    ' Description : gets speaker names from ViewJobFormAppearancesQ query and find/replaces in transcript as appropriate
    '============================================================================

    Dim sMrMs As String
    Dim sLastName As String
    Dim vSpeakerName As String
    
    Dim oWordApp As New Word.Application
    Dim oWordDoc As New Word.Document
    
    Dim x As Long
    
    Dim qdf As QueryDef
    Dim rs As DAO.Recordset
    
    Dim cJob As New Job
    
    x = 18                                       '18 is number of first dynamic speaker

    DoCmd.OpenQuery qnViewJobFormAppearancesQ, acViewNormal, acReadOnly 'open query

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField] 'job number

    Set qdf = CurrentDb.QueryDefs(qnViewJobFormAppearancesQ) 'open query
    qdf.Parameters(0) = sCourtDatesID
    Set rs = qdf.OpenRecordset

    Set oWordApp = CreateObject("Word.Application")
    If oWordApp Is Nothing Then
        Set oWordApp = CreateObject("Word.Application")
    End If
    oWordApp.Visible = False

    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.CourtCover)

    With oWordDoc

        If Not (rs.EOF And rs.BOF) Then
    
            rs.MoveFirst
    
            Do Until rs.EOF = True
        
                sMrMs = rs!MrMs                  'get MrMs & LastName variables
                sLastName = rs!LastName
                vSpeakerName = UCase(sMrMs & ". " & sLastName & ":  ") 'store together in variable as a string
            
                
                
                .Application.Selection.Find.ClearFormatting 'Do find/replaces
                .Application.Selection.Find.Replacement.ClearFormatting
                With .Application.Selection.Find
                    .Text = " snl" & x & Chr(32)
                    .Replacement.Text = ".^p" & vSpeakerName
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                .Application.Selection.Find.Execute Replace:=wdReplaceAll
            
                .Application.Selection.Find.ClearFormatting
                .Application.Selection.Find.Replacement.ClearFormatting
                With .Application.Selection.Find
                    .Text = " dnl" & x & Chr(32)
                    .Replacement.Text = " --^p" & vSpeakerName
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                .Application.Selection.Find.Execute Replace:=wdReplaceAll
            
                .Application.Selection.Find.ClearFormatting
                .Application.Selection.Find.Replacement.ClearFormatting
                With .Application.Selection.Find
                    .Text = " qnl" & x & Chr(32)
                    .Replacement.Text = "?^p" & vSpeakerName
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                .Application.Selection.Find.Execute Replace:=wdReplaceAll
                .Application.Selection.Find.ClearFormatting
                .Application.Selection.Find.Replacement.ClearFormatting
                With .Application.Selection.Find
                    .Text = " sbl" & x & Chr(32)
                    .Replacement.Text = ".^pBY " & vSpeakerName
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
            
                .Application.Selection.Find.Execute Replace:=wdReplaceAll
                .Application.Selection.Find.ClearFormatting
                .Application.Selection.Find.Replacement.ClearFormatting
                With .Application.Selection.Find
                    .Text = " dbl" & x & Chr(32)
                    .Replacement.Text = " --^pBY " & vSpeakerName
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                .Application.Selection.Find.Execute Replace:=wdReplaceAll
            
                .Application.Selection.Find.ClearFormatting
                .Application.Selection.Find.Replacement.ClearFormatting
                With .Application.Selection.Find
                    .Text = " qbl" & x & Chr(32)
                    .Replacement.Text = "?^pBY " & vSpeakerName
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                .Application.Selection.Find.Execute Replace:=wdReplaceAll
            
                .Application.Selection.Find.ClearFormatting
                .Application.Selection.Find.Replacement.ClearFormatting
                With .Application.Selection.Find
                    .Text = " sqnl" & x & Chr(32)
                    .Replacement.Text = "." & Chr(34) & "^p" & vSpeakerName
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                .Application.Selection.Find.Execute Replace:=wdReplaceAll
            
                .Application.Selection.Find.ClearFormatting
                .Application.Selection.Find.Replacement.ClearFormatting
                With .Application.Selection.Find
                    .Text = " dqnl" & x & Chr(32)
                    .Replacement.Text = " --" & Chr(34) & "^p" & vSpeakerName
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                .Application.Selection.Find.Execute Replace:=wdReplaceAll
            
                .Application.Selection.Find.ClearFormatting
                .Application.Selection.Find.Replacement.ClearFormatting
                With .Application.Selection.Find
                    .Text = " qqnl" & x & Chr(32)
                    .Replacement.Text = "?" & Chr(34) & "^p" & vSpeakerName
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                'clear variables before loop
                sMrMs = ""
                sLastName = ""
                vSpeakerName = ""
            
                x = x + 1                        'add 1 to x for next speaker name
                rs.MoveNext                      'go to next speaker name
                
            Loop                                 'back up to the top
        Else

            MsgBox "There are no dynamic speakers." 'msg upon completion
        End If
    
        'MsgBox "Finished looping through dynamic speakers."
    
        rs.Close                                 'Close the recordset
        Set rs = Nothing                         'Clean up
        oWordDoc.SaveAs2 FileName:=cJob.DocPath.CourtCover
        oWordDoc.Close
        oWordApp.Quit
    

        Set oWordDoc = Nothing
        Set oWordApp = Nothing

    End With
    
End Sub

Public Sub pfSingleFindReplace(ByVal sTextToFind As String, ByVal sReplacementText As String, Optional ByVal wsyWordStyle As String = "", Optional bForward As Boolean = True, _
                               Optional bWrap As String = "wdFindContinue", Optional bFormat As Boolean = False, Optional bMatchCase As Boolean = True, _
                               Optional bMatchWholeWord As Boolean = False, Optional bMatchWildcards As Boolean = False, _
                               Optional bMatchSoundsLike As Boolean = False, Optional bMatchAllWordForms As Boolean = False)
    '============================================================================
    ' Name        : pfSingleFindReplace
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfSingleFindReplace
    ' Description : find and replace one item
    '============================================================================
    Dim cJob As New Job

    'Set oWordDoc = Documents.Open(cJob.DocPath.CourtCover)

    On Error Resume Next
    Set oWordApp = GetObject(, "Word.Application")
    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0

    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.CourtCover)

    oWordApp.Visible = True
    oWordDoc.Application.Selection.HomeKey Unit:=wdStory

    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting
   
    With oWordDoc.Application.Selection.Find
        .Text = sTextToFind
        .Replacement.Text = sReplacementText
        If wsyWordStyle <> "" Then
            .Replacement.Style = oWordDoc.Styles(wsyWordStyle)
        Else
        End If
        .Forward = bForward
        '.Wrap = bWrap
        .Format = bFormat
        .MatchCase = bMatchCase
        .MatchWholeWord = bMatchWholeWord
        .MatchWildcards = bMatchWildcards
        .MatchSoundsLike = bMatchSoundsLike
        .MatchAllWordForms = bMatchAllWordForms
    End With
    oWordDoc.Application.Selection.Find.Execute
        
    With oWordDoc.Application.Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceAll
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
        .Find.Execute
    End With
    oWordDoc.Save
    'Debug.Print "no"
End Sub

Public Sub pfSingleTCReplaceAll(ByVal sTextToFind As String, ByVal sReplacementText As String, Optional ByVal wsyWordStyle As String = "", Optional bForward As Boolean = True, _
                                Optional bWrap As String = "wdFindContinue", Optional bFormat As Boolean = False, Optional bMatchCase As Boolean = True, _
                                Optional bMatchWholeWord As Boolean = False, Optional bMatchWildcards As Boolean = False, _
                                Optional bMatchSoundsLike As Boolean = False, Optional bMatchAllWordForms As Boolean = False)
    '============================================================================
    ' Name        : pfSingleReplaceAll
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfSingleReplaceAll(sTexttoSearch, sReplacementText)
    ' Description : one replace TC entry function for ones with no field entry
    '============================================================================
    
    Dim cJob As New Job

    On Error Resume Next

    Set oWordApp = GetObject(, "Word.Application")

    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    
    End If

    On Error GoTo 0

    oWordApp.Visible = True
    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.CourtCover)



    With oWordDoc.Application

        .Selection.Find.ClearFormatting
    
        With .Selection.Find
            .Text = sTextToFind
            .Replacement.Text = sReplacementText
            If wsyWordStyle <> "" Then
                .Replacement.Style = oWordDoc.Styles(wsyWordStyle)
            Else
            End If
            .Forward = True
            .Wrap = wdFindContinue
            .Format = bFormat
            .MatchCase = False
            If bMatchCase <> Empty Then
                .MatchCase = bMatchCase
            Else
            End If
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
    
        .Selection.Find.Execute Replace:=wdReplaceAll
    
    End With
    oWordDoc.Save

End Sub

Public Sub pfFieldTCReplaceAll(sTexttoSearch As String, sReplacementText As String, sFieldText As String)
    '============================================================================
    ' Name        : pfFieldTCReplaceAll
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfFieldTCReplaceAll(sTexttoSearch, sReplacementText, sFieldText)
    ' Description : one replace TC entry function for ones with field entry
    '============================================================================

    Dim cJob As New Job

    'wdFieldTOCEntry
    On Error Resume Next
    Set oWordApp = GetObject(, "Word.Application")
    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    oWordApp.Visible = True

    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.CourtCover)


    With oWordDoc.Application
        .Selection.Find.ClearFormatting
    
        With .Selection.Find
            .Text = sTexttoSearch
            .Replacement.Text = sReplacementText
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        .Selection.Find.Execute
        Do While .Selection.Find.Found
            .Selection.Find.Execute Replace:=wdReplaceOne
            .Selection.Range.Text = sReplacementText
            .Selection.Fields.Add Type:=wdFieldTOCEntry, Text:=sFieldText, PreserveFormatting:=False, Range:=.Selection.Range 'sFieldText sample = "TC ""WitnessName"" \l 2-3"
            .Selection.Find.Execute
        Loop
        
        With .Selection
            If .Find.Forward = True Then
                .Collapse Direction:=wdCollapseStart
            Else
                .Collapse Direction:=wdCollapseEnd
            End If
            '.Find.Execute 'Replace:=wdReplaceOne 'wdReplaceAll
        
        
            'Do While .Find.Found
            '    .Find.Execute
            '   .Fields.Add Type:=wdFieldTOCEntry, Text:=sFieldText, PreserveFormatting:=False, Range:=.Range 'sFieldText sample = "TC ""WitnessName"" \l 2-3"
            '    .Find.Execute
            'Loop
        
            If .Find.Forward = True Then
                .Collapse Direction:=wdCollapseEnd
            Else
                .Collapse Direction:=wdCollapseStart
            End If
        
    
        End With
    
    End With

    oWordDoc.Save

End Sub

Public Sub pfWordIndexer()
    '============================================================================
    ' Name        : pfWordIndexer
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfWordIndexer
    ' Description : builds word index in separate PDF from transcript
    '============================================================================

    Dim sInput As String
    Dim sCurrentIndexEntry As String
    Dim sCurrentEntryOriginal As String
    Dim sExclusions As String
    Dim sCurrentEntry1 As String
    Dim sCurrentEntry2 As String
    Dim sCurrentEntry3 As String
    Dim sCurrentEntry4 As String
    Dim sCurrentEntry5 As String
    Dim vBookmarkName As String
    
    Dim oWordApp As New Word.Application
    Dim oWordDoc As New Word.Document
    Dim oWordApp1 As New Word.Application
    Dim oWordDoc1 As New Word.Document
    
    Dim w As Long
    Dim x As Long
    Dim y As Long
    Dim z As Long
    
    Dim Rng As Variant
    
    Dim cJob As New Job
    
    'TODO: Take out duplicate page ##s
    
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

    'Set oWordApp1 = CreateObject(Class:="Word.Application")
    oWordApp1.Visible = True

    oWordApp1.AutomationSecurity = msoAutomationSecurityLow
    Set oWordDoc1 = oWordApp1.Application.Documents.Open(cJob.DocPath.CourtCover)

    sExclusions = "a,am,an,and,are,as,at,b,be,but,by,c,can,cm,d,did,case,cases,about,cause,ask,asks,asked,asking," & _
                  "do,does,e,eg,en,eq,etc,f,for,g,get,go,got,h,has,have,correct,conduct,examination,direct,cross," & _
                  "he,her,him,how,i,ie,if,in,into,is,it,its,j,k,l,m,me,don't,didn't,county,court,motion,look,looking,looked," & _
                  "mi,mm,my,n,na,nb,no,not,o,of,off,ok,on,one,or,our,out,had,going,first,knew,know,under,thing,things,took," & _
                  "p,q,r,re,s,she,so,t,the,their,them,they,this,t,to,u,v,his,her,honor,here,objection," & _
                  "like,let,law,other,order,last,know,judge,petitioner's,respondent's,plaintiff's,defendant's,court's," & _
                  "from,then,than,court,there's,that,that's,order,indiscernible,who,what,when,where,why,yes,yeah,i've,I'm,just,right,order,all,because,it's,aquoco.co,no,that,that's,I've,there,petitioner,respondent,plaintiff,defendant,right,um,uh,huh," & _
                  "via,vs,w,was,we,were,who,will,with,would,x,y,yd,you,your,you're,yours,he's,she's,she,z," & _
                  "well,since,sorry,there,there'stook,too,such,than,times,1,2,3,4,5,6,7,8,9,0,98119,again,after,address,actually,a.m,p.m,anyway,anything," & _
                  "anyone,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,able,another,anyone,anything,anywhere,anytime," & _
                  "being,before,asked,asking,around,ask,away,ave,bac,bad,been,before,being,beings,between,boys,c.d,call,called,calling,cannot,can't,won't,don't,aren't,isn't," & _
                  "clerk,clear,child,children,children's,course,closee,come,coming,contact,correct,could,couldn't,wouldn't,shouldn't,didn't,doesn't,current,day,doing,done," & _
                  "even,every,excuse,evidence,evidencing,exactly,factors,factor,fear,feel,feet,fifth,female,sixth,seventh,eighth,ninth,first,second,third,fourth,front,hard," & _
                  "soft,gone,given,hear,hearing,have,having,have,folks,jury,jurors,venire,herself,himself,her,hers,his,help,handle,happy,guys,guy,group,gotten,good,full,form," & _
                  "forth,family,excuse,guilty,he's,she's,high,his,hold,huh,uh,i.d,i'd,i'm,i'll,i.m,however,hyperlinked,include,included,including,indeed,index,information,indiscernible," & _
                  "job,king,judge,law,know,knew,knows,last,lasted,later,interest,interested,issue,issues,issued,let,leave,hours,court," & _
                  "live,might,lives,lived,living,long,longer,look,looked,looking,looks,love,made,mail,make,makes,making,man,march,matter,mean,meaning,means,meant,meet,meets,might,mind," & _
                  "met,more,most,mount,names,name,need,needed,needs,never,new,news,next,nor,notice,number,numbers,numbered,old,only,open,original,other,own,owned,page,parent,parents,parties,party," & _
                  "pattern,period,periods,petition,petitioner,response,responses,respondent,problem,problems,point,please,put,read,purpose,record,records,prior,report,restraining,service,sorry,sort,kind,statute," & _
                  "six,school,under,through,think,thought,things,thing,they're,these,there's,there,tell,telling,table,take,such,stattues,still,temporary,thrown,took,too,though,through,sure," & _
                  "wi,try,trying,tried,tries,see,seeing,saw,sees,self,person,persons,people," & _
                  "you've,you're,well,we'll,went,we're,why,what,who,will,way,wanted,want,very,us,until,week,weeks,yesterday,talk,talking,use,which,wherever,some,question,questions"
          
    With oWordDoc1
        .Application.DisplayAlerts = False
        .Application.Visible = False
        sInput = .Content.Text
        
        For w = 1 To 255                         'hyphens & single quotes kept; strip unwanted chars
            Select Case w
            Case 1 To 35, 37 To 38, 40 To 43, 45, 47, 58 To 64, 91 To 96, 123 To 127, 129 To 144, 147 To 149, 152 To 162, 164, 166 To 171, 174 To 191, 247
                sInput = Replace(sInput, Chr(w), " ")
            End Select
        Next
    
        sInput = Replace(Replace(Replace(Replace(sInput, Chr(44) & Chr(32), " "), Chr(44) & vbCr, " "), Chr(46) & Chr(32), " "), Chr(46) & vbCr, " ")
        sInput = Replace(Replace(Replace(Replace(sInput, Chr(145), "'"), Chr(146), "'"), "' ", " "), " '", " ")
        sInput = " " & LCase(Trim(sInput)) & " "
    
        For w = 0 To UBound(Split(sExclusions, ",")) 'loop through sExclusions
            While InStr(sInput, " " & Split(sExclusions, ",")(w) & " ") > 0
                sInput = Replace(sInput, " " & Split(sExclusions, ",")(w) & " ", " ")
            Wend
        Next
    
        While InStr(sInput, "  ") > 0
            sInput = Replace(sInput, "  ", " ")
        Wend
    
        sInput = " " & Trim(sInput) & " "
        x = UBound(Split(sInput, " "))
        z = x
    
        For w = 1 To x
            sCurrentEntryOriginal = Split(sInput, " ")(1) 'get word count
            While InStr(sInput, " " & sCurrentEntryOriginal & " ") > 0
                sInput = Replace(sInput, " " & sCurrentEntryOriginal & " ", " ")
            Wend
            y = z - UBound(Split(sInput, " "))   'calculate replaced count
            sCurrentIndexEntry = sCurrentIndexEntry & sCurrentEntryOriginal & vbTab & y & vbCr 'update current index entry
            z = UBound(Split(sInput, " "))
            If z = 1 Then Exit For
            DoEvents
        Next
    
        sInput = sCurrentIndexEntry
        sCurrentIndexEntry = ""
        sCurrentEntry5 = UBound(Split(sInput, vbCr)) - 1
    
        For w = 0 To sCurrentEntry5
            sCurrentEntryOriginal = ""
            With .Range
                With .Find
                    .ClearFormatting
                    sCurrentEntry4 = Split(Split(sInput, vbCr)(w), vbTab)(0)
                    sCurrentEntry1 = " " & Split(Split(sInput, vbCr)(w), vbTab)(1)
                    .Text = sCurrentEntry4
                    .Replacement.Text = ""
                    .Wrap = wdFindStop
                    .Forward = True
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = True
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                    .Execute
                End With
                Do While .Find.Found
                    If sCurrentEntryOriginal = "" Then sCurrentEntryOriginal = sCurrentEntryOriginal & " " & .Information(wdActiveEndPageNumber)
                    sCurrentEntry1 = Right(sCurrentEntryOriginal, 2)
                    sCurrentEntry2 = " " & .Information(wdActiveEndPageNumber)
                    If sCurrentEntry1 = sCurrentEntry2 Then sCurrentEntryOriginal = sCurrentEntryOriginal
                    If sCurrentEntry1 <> sCurrentEntry2 Then sCurrentEntryOriginal = sCurrentEntryOriginal & " " & .Information(wdActiveEndPageNumber)
                    .Collapse (wdCollapseEnd)
                    .Find.Execute
                
                    If sCurrentEntry1 = "" Then GoTo ExitLoop1
                Loop
ExitLoop1:
            End With
            sCurrentEntryOriginal = Replace(Trim(sCurrentEntryOriginal), " ", ",")
            sCurrentIndexEntry = sCurrentIndexEntry & Split(sInput, vbCr)(w) & vbTab & sCurrentEntryOriginal & vbCr
            If sCurrentEntryOriginal = "" Then GoTo ExitLoop2
        Next
    End With
    oWordApp1.Quit
ExitLoop2:

    'Set oWordApp = CreateObject("Word.Application")
    oWordApp.Visible = False
    
    Set oWordDoc = oWordApp.Documents.Add(cJob.DocPath.WordIndexT) 'template

    With oWordDoc
        Set Rng = .Range.Characters.Last

        'Create the word index
        With Rng
            .InsertAfter vbCr & Chr(12) & sCurrentIndexEntry
            .Start = .Start
            .ConvertToTable Separator:=vbTab, NumColumns:=3
            .Tables(1).Sort Excludeheader:=False, FieldNumber:=1, _
                            SortFieldType:=wdSortFieldAlphanumeric, _
                            SortOrder:=wdSortOrderAscending, CaseSensitive:=False
            .Tables.item(1).Columns(2).delete
            '.Tables.item(1).Columns(1).Width = InchesToPoints(1.1)
            '.Tables.item(1).Columns(2).Width = InchesToPoints(0.8)
        End With
    
        With Rng
            .Tables(1).Columns(1).Select
            .Application.Selection.Font.Bold = wdToggle
        End With
    
        vBookmarkName = "WordIndex"
    
        .Application.Selection.Find.ClearFormatting
        .Application.Selection.Find.Replacement.ClearFormatting
        With .Application.Selection.Find
            .Text = "#WI#"
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
        End With
    
        .bookmarks.Add Name:=vBookmarkName
        .Application.Selection.EndKey Unit:=wdStory, Extend:=wdExtend
    
        With .Application.Selection.PageSetup.TextColumns
            .SetCount NumColumns:=3
            .EvenlySpaced = True
            '.Width = InchesToPoints(1)
            .LineBetween = False
        End With
    
        .Application.Selection.Goto What:=wdGoToBookmark, Name:="WordIndex"
    
        If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
            .ActiveWindow.ActivePane.View.Type = wdNormalView
        Else
            .ActiveWindow.View.Type = wdNormalView
        End If
    
        .Application.Selection.MoveDown Unit:=wdLine, Count:=4
        .Application.Selection.delete Count:=3
    
        If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
            .ActiveWindow.ActivePane.View.Type = wdPrintView
        Else
            .ActiveWindow.View.Type = wdPrintView
        End If
    
        .Application.Selection.HomeKey Unit:=wdLine
        .Application.Selection.HomeKey Unit:=wdStory
        .Application.Selection.EndKey Unit:=wdLine
        .Application.Selection.MoveRight Unit:=wdCharacter, Count:=1
        .Application.Selection.EndKey Unit:=wdStory, Extend:=wdExtend
        .Application.Selection.Font.Size = 10
    
        .SaveAs cJob.DocPath.WordIndexDB
        .SaveAs cJob.DocPath.WordIndexD
        .SaveAs cJob.DocPath.WordIndexP
        .Close
    End With

    oWordApp.Quit

End Sub

Public Sub FPJurors()
    '============================================================================
    ' Name        : FPJurors
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call FPJurors()
    ' Description : does find/replacements of prospective jurors in transcript
    '============================================================================

    Dim sSpeakerName As String
    Dim ssSpeakerFind As String
    Dim sdSpeakerFind As String
    Dim sqSpeakerFind As String
    Dim ssSpeakerName As String
    Dim sdSpeakerName As String
    Dim sqSpeakerName As String
    
    Dim oCourtCoverWD As New Word.Document
    Dim oWordApp As New Word.Application
    
    Dim x As Long
    Dim y As Long

    Dim cJob As New Job

    Call pfCurrentCaseInfo                       'refresh transcript info

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    x = 101                                      '101 is number of first PROSPECTIVE JUROR
    On Error Resume Next

    Set oWordApp = GetObject(, "Word.Application")
    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If

    oWordApp.Visible = False

    Set oCourtCoverWD = oWordApp.Documents.Open(cJob.DocPath.CourtCover)


    With oCourtCoverWD

        ssSpeakerFind = " snl100 "
        sdSpeakerFind = " dnl100 "
        sqSpeakerFind = " qnl100 "
        ssSpeakerName = ".^p" & UCase("PROSPECTIVE JUROR") & ":  "
        sdSpeakerName = "^s--^p" & UCase("PROSPECTIVE JUROR") & ":  "
        sqSpeakerName = "?^p" & UCase("PROSPECTIVE JUROR") & ":  "
    
        'Do find/replaces
            
        .Application.Selection.Find.ClearFormatting
        .Application.Selection.Find.Replacement.ClearFormatting
        With .Application.Selection.Find
            .Text = ssSpeakerFind
            .Replacement.Text = ssSpeakerName
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        .Application.Selection.Find.Execute Replace:=wdReplaceAll
    
        .Application.Selection.Find.ClearFormatting
        .Application.Selection.Find.Replacement.ClearFormatting
        With .Application.Selection.Find
            .Text = sdSpeakerFind
            .Replacement.Text = sdSpeakerName
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        .Application.Selection.Find.Execute Replace:=wdReplaceAll
    
        .Application.Selection.Find.ClearFormatting
        .Application.Selection.Find.Replacement.ClearFormatting
        With .Application.Selection.Find
            .Text = sqSpeakerFind
            .Replacement.Text = sqSpeakerName
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        .Application.Selection.Find.Execute Replace:=wdReplaceAll

        If Not (x = 100 And x = 200) Then
            Do Until x = 200
                ssSpeakerFind = " snl" & x & " "
                sdSpeakerFind = " dnl" & x & " "
                sqSpeakerFind = " qnl" & x & " "
                y = x - 100
                ssSpeakerName = ".^p" & UCase("PROSPECTIVE JUROR NO. ") & y & ":  "
                sdSpeakerName = "^s--^p" & UCase("PROSPECTIVE JUROR NO. ") & y & ":  "
                sqSpeakerName = "?^p" & UCase("PROSPECTIVE JUROR NO. ") & y & ":  "
    
                'Do find/replaces
            
                .Application.Selection.Find.ClearFormatting
                .Application.Selection.Find.Replacement.ClearFormatting
                With .Application.Selection.Find
                    .Text = ssSpeakerFind
                    .Replacement.Text = ssSpeakerName
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                .Application.Selection.Find.Execute Replace:=wdReplaceAll
    
                .Application.Selection.Find.ClearFormatting
                .Application.Selection.Find.Replacement.ClearFormatting
                With .Application.Selection.Find
                    .Text = sdSpeakerFind
                    .Replacement.Text = sdSpeakerName
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                .Application.Selection.Find.Execute Replace:=wdReplaceAll
    
                .Application.Selection.Find.ClearFormatting
                .Application.Selection.Find.Replacement.ClearFormatting
                With .Application.Selection.Find
                    .Text = sqSpeakerFind
                    .Replacement.Text = sqSpeakerName
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                .Application.Selection.Find.Execute Replace:=wdReplaceAll
        
                x = x + 1                        'add 1 to x for next speaker name
            Loop
        Else
            'upon completion
            MsgBox "There are no records in the recordset."
        End If
    End With
    oCourtCoverWD.Save
    oCourtCoverWD.Close
    On Error GoTo 0
    oWordApp.Quit
    Call pfClearGlobals
End Sub

Public Sub pfTCEntryReplacement()
    '============================================================================
    ' Name        : pfTCEntryReplacement
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfTCEntryReplacement
    ' Description : parent function that finds certain entries within a transcript and assigns TC entries to them for indexing purposes
    '============================================================================
    
    Dim sMrMs2 As String
    Dim sLastName2 As String
    Dim vSpeakerName As String
    
    Dim rstTRCourtQ As DAO.Recordset
    Dim rstViewJFAppQ As DAO.Recordset
    Dim qdf As QueryDef
    
    Dim oWordApp As New Word.Application
    Dim oCourtCoverWD As New Word.Document
    
    Dim cJob As New Job
    
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField] 'job number
    
    On Error Resume Next
    Set oWordApp = GetObject(, "Word.Application")
    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    oWordApp.Visible = False
    
    Set oCourtCoverWD = oWordApp.Documents.Open(cJob.DocPath.CourtCover)
    Set qdf = CurrentDb.QueryDefs(qnTRCourtQ)    'open query
    qdf.Parameters(0) = sCourtDatesID
    Set rstTRCourtQ = qdf.OpenRecordset
    sJurisdiction = rstTRCourtQ!Jurisdiction
    sParty1Name = rstTRCourtQ!Party1Name
    sParty2Name = rstTRCourtQ!Party2Name

    qdf.Close
    rstTRCourtQ.Close

    Set qdf = CurrentDb.QueryDefs(qnViewJobFormAppearancesQ) 'open query
    qdf.Parameters(0) = sCourtDatesID
    Set rstViewJFAppQ = qdf.OpenRecordset

    rstViewJFAppQ.MoveFirst
    sMrMs2 = rstViewJFAppQ!MrMs
    sLastName2 = rstViewJFAppQ!LastName

    If Not (rstViewJFAppQ.EOF And rstViewJFAppQ.BOF) Then
        rstViewJFAppQ.MoveFirst
        With oCourtCoverWD.Application           'beginning of file do these replacements
            .Selection.Find.ClearFormatting
            Call pfFieldTCReplaceAll("(nnn)", "^p ", Chr(34) & "TC" & Chr(34) & " " & Chr(34) & "WitnessName" & Chr(34) & " " & "\l 2")
            Call pfFieldTCReplaceAll("(ema)", "^p(Exhibit ## marked and admitted.)^p", Chr(34) & "TC" & Chr(34) & " " & Chr(34) & "Exhibit  marked and admitted" & Chr(34) & " " & "\f cd")
            Call pfFieldTCReplaceAll("(emm)", "^p(Exhibit ## marked.)^p", Chr(34) & "TC" & Chr(34) & " " & Chr(34) & "Exhibit  marked" & Chr(34) & " " & "\f cd")
            Call pfFieldTCReplaceAll("(eaa)", "^p(Exhibit ## admitted.)^p", Chr(34) & "TC" & Chr(34) & " " & Chr(34) & "Exhibit  admitted" & Chr(34) & " " & "\f cd")
            Call pfFieldTCReplaceAll("(exa)", "^p(Exhibit ## admitted.)^p", Chr(34) & "TC" & Chr(34) & " " & Chr(34) & "Exhibit  admitted" & Chr(34) & " " & "\f cd")
        
            Call pfFieldTCReplaceAll("(ee1)", "^pDIRECT EXAMINATION^p", Chr(34) & "Direct Examination by " & Chr(34) & " \l 3")
            Call pfFieldTCReplaceAll("(ee2)", "^pCROSS-EXAMINATION^p", Chr(34) & "Cross-Examination by " & Chr(34) & " \l 3")
            Call pfFieldTCReplaceAll("(ee3)", "^pREDIRECT EXAMINATION^p", Chr(34) & "Redirect Examination by " & Chr(34) & " \l 3")
            Call pfFieldTCReplaceAll("(ee4)", "^pRECROSS-EXAMINATION^p", Chr(34) & "Recross-Examination by " & Chr(34) & " \l 3")
            Call pfFieldTCReplaceAll("(ee5)", "^pFURTHER REDIRECT EXAMINATION^p", Chr(34) & "Further Redirect Examination by " & Chr(34) & " \l 3")
            Call pfFieldTCReplaceAll("(ee6)", "^pFURTHER RECROSS-EXAMINATION^p", Chr(34) & "Further Recross-Examination by " & Chr(34) & " \l 3")
            Call pfFieldTCReplaceAll("(e1c)", "^pDIRECT EXAMINATION CONTINUED^p", Chr(34) & "Direct Examination Continued by " & Chr(34) & " \l 3")
            Call pfFieldTCReplaceAll("(e2c)", "^pCROSS-EXAMINATION CONTINUED^p", Chr(34) & "Cross-Examination Continued by " & Chr(34) & " \l 3")
            Call pfFieldTCReplaceAll("(e3c)", "^pREDIRECT EXAMINATION CONTINUED^p", Chr(34) & "Redirect Examination Continued by " & Chr(34) & " \l 3")
            Call pfFieldTCReplaceAll("(e4c)", "^pRECROSS-EXAMINATION CONTINUED^p", Chr(34) & "Recross-Examination Continued by " & Chr(34) & " \l 3")
            Call pfFieldTCReplaceAll("(e5c)", "^pFURTHER REDIRECT EXAMINATION CONTINUED^p", Chr(34) & "Further Redirect Examination Continued by " & Chr(34) & " \l 3")
            Call pfFieldTCReplaceAll("(e6c)", "^pFURTHER RECROSS-EXAMINATION CONTINUED^p", Chr(34) & "Further Recross-Examination Continued by " & Chr(34) & " \l 3")
            Call pfFieldTCReplaceAll("(crr)", "^pCOURT'S RULING" & "^p", Chr(34) & "TC" & Chr(34) & " " & Chr(34) & "Court's Ruling" & Chr(34) & " " & "\f e")
            Call pfFieldTCReplaceAll("(aa1)", "^pARGUMENT FOR THE " & UCase(sParty1Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Argument for the " & sParty1Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
            Call pfFieldTCReplaceAll("(ar1)", "^pREBUTTAL ARGUMENT FOR THE " & UCase(sParty1Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Rebuttal Argument for the " & sParty1Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
            Call pfFieldTCReplaceAll("(ao1)", "^pOPENING STATEMENT FOR THE " & UCase(sParty1Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Opening Statement for the " & sParty1Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
            Call pfFieldTCReplaceAll("(ac1)", "^pCLOSING ARGUMENT FOR THE " & UCase(sParty1Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Closing Argument for the " & sParty1Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
            Call pfSingleTCReplaceAll("(sbb)", "^p(Sidebar begins at ##:## ap.m.)^p")
            Call pfSingleTCReplaceAll("(sbe)", "^p(Sidebar ends at ##:## ap.m.)^p")
            Call pfSingleTCReplaceAll("(rrr)", "^p(Recess taken from ##:## ap.m. to ##:## ap.m.)^p")
            Call pfSingleTCReplaceAll("(sbn)", "^p(Sidebar taken from ##:## ap.m. to ##:## ap.m.)^p")
            Call pfSingleTCReplaceAll("(jen)", "^p(Jury panel enters at ##:## ap.m.)^p")
            Call pfSingleTCReplaceAll("(jex)", "^p(Jury panel exits at ##:## ap.m.)^p")
            Call pfSingleTCReplaceAll("(pjn)", "^p(Prospective jury panel enters at ##:## ap.m.)^p")
            Call pfSingleTCReplaceAll("(pjx)", "^p(Prospective jury panel exits at ##:## ap.m.)^p")
            Call pfSingleTCReplaceAll("(wsu)", "^p(Witness summoned.)^p")
            Call pfSingleTCReplaceAll("(wsw)", "^p(The witness was sworn.)^p")
            Call pfSingleTCReplaceAll("(vub)", "^p(Video played at ##:## ap.m.)^p")
            Call pfSingleTCReplaceAll("(vue)", "^p(Video ends at ##:## ap.m.)^p")
            Call pfSingleTCReplaceAll("(vup)", "^p(Video played from ##:## ap.m. to ##:## ap.m.)^p")
            Call pfSingleTCReplaceAll("(aup)", "^p(Audio played from ##:## ap.m. to ##:## ap.m.)^p")
            Call pfSingleTCReplaceAll("(aue)", "^p(Audio ends at ##:## ap.m.)^p")
            Call pfSingleTCReplaceAll("(aub)", "^p(Audio begins at ##:## ap.m.)^p")
            Call pfSingleTCReplaceAll("(ccc)", "^p(Counsel confer.)^p")
            Call pfSingleTCReplaceAll("(pcc)", "^p(Parties confer.)^p")
            Call pfSingleTCReplaceAll("(ppr)", "^p(The witness paused to review the document.)^p")
            Call pfSingleTCReplaceAll("(nrp)", "^p(No response.)^p")
            Call pfSingleTCReplaceAll("(rrr)", "^p(Whereupon, at ##:## ap.m., a recess was taken.)^p")
            Call pfSingleTCReplaceAll("(rrl)", "^p(Whereupon, at ##:## ap.m., a luncheon recess was taken.)^p")
            Call pfSingleTCReplaceAll("(ppp)", "^p(Pause.)^p")
            Call pfSingleTCReplaceAll("(otr)", "^p(Off the record.)^p")
            Call pfSingleTCReplaceAll("(dtr)", "^p(Discussion held off the record.)^p")
            Call pfSingleTCReplaceAll("(wxu)", "^p(Witness excused.)^p")
            Call pfSingleTCReplaceAll("(cco)", "^p(Whereupon, the following proceedings were held in open court outside the presence of the jury:)^p")
            Call pfSingleTCReplaceAll("(cci)", "^p(Whereupon, the following proceedings were held in open court in the presence of the jury:)^p")
            Call pfSingleTCReplaceAll("Uh-huh.", "Uh-huh.")
            Call pfSingleTCReplaceAll("Huh-uh.", "Huh-uh.")
            'Call pfFieldTCReplaceAll(, , )
        
            If Not rstViewJFAppQ.EOF Then rstViewJFAppQ.MoveNext
        
            If Not rstViewJFAppQ.EOF Then
                sMrMs2 = rstViewJFAppQ!MrMs      'get MrMs & LastName variables
                sLastName2 = rstViewJFAppQ!LastName
                Call pfFieldTCReplaceAll("(aa2)", "^pARGUMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Argument for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ar2)", "^pREBUTTAL ARGUMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Rebuttal Argument for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ao2)", "^pOPENING STATEMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Opening Statement for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ac2)", "^pCLOSING ARGUMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Closing Argument for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
            End If
        
            If Not rstViewJFAppQ.EOF Then rstViewJFAppQ.MoveNext
        
            If Not rstViewJFAppQ.EOF Then
                sMrMs2 = rstViewJFAppQ!MrMs      'get MrMs & LastName variables
                sLastName2 = rstViewJFAppQ!LastName
                Call pfFieldTCReplaceAll("(aa3)", "^pARGUMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Argument for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ar3)", "^pREBUTTAL ARGUMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Rebuttal Argument for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ao3)", "^pOPENING STATEMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Opening Statement for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ac3)", "^pCLOSING ARGUMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Closing Argument for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
            End If
        
            If Not rstViewJFAppQ.EOF Then rstViewJFAppQ.MoveNext
        
            If Not rstViewJFAppQ.EOF Then
                sMrMs2 = rstViewJFAppQ!MrMs      'get MrMs & LastName variables
                sLastName2 = rstViewJFAppQ!LastName
                Call pfFieldTCReplaceAll("(aa4)", "^pARGUMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Argument for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ar4)", "^pREBUTTAL ARGUMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Rebuttal Argument for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ao4)", "^pOPENING STATEMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Opening Statement for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ac4)", "^pCLOSING ARGUMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Closing Argument for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
            End If
        
            If Not rstViewJFAppQ.EOF Then rstViewJFAppQ.MoveNext
        
            If Not rstViewJFAppQ.EOF Then
                sMrMs2 = rstViewJFAppQ!MrMs      'get MrMs & LastName variables
                sLastName2 = rstViewJFAppQ!LastName
                Call pfFieldTCReplaceAll("(aa5)", "^pARGUMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Argument for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ar5)", "^pREBUTTAL ARGUMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Rebuttal Argument for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ao5)", "^pOPENING STATEMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Opening Statement for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ac5)", "^pCLOSING ARGUMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Closing Argument for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
            End If
        
            If Not rstViewJFAppQ.EOF Then rstViewJFAppQ.MoveNext
        
            If Not rstViewJFAppQ.EOF Then
                sMrMs2 = rstViewJFAppQ!MrMs      'get MrMs & LastName variables
                sLastName2 = rstViewJFAppQ!LastName
                Call pfFieldTCReplaceAll("(aa6)", "^pARGUMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Argument for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ar6)", "^pREBUTTAL ARGUMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Rebuttal Argument for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ao6)", "^pOPENING STATEMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Opening Statement for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ac6)", "^pCLOSING ARGUMENT FOR THE " & UCase(sParty2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Closing Argument for the " & sParty2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
            End If
        
            GoTo ParenDone
        End With
    End If
ParenDone:
    MsgBox "Finished looping through TC entries for the various parties."

    rstViewJFAppQ.Close
    Set rstViewJFAppQ = Nothing
    oCourtCoverWD.SaveAs2 FileName:=cJob.DocPath.CourtCover
    oCourtCoverWD.Close
    oWordApp.Quit
    Set oCourtCoverWD = Nothing
    Set oWordApp = Nothing
End Sub


Public Sub pfFindRepCitationLinks()
    '============================================================================
    ' Name        : pfFindRepCitationLinks
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfFindRepCitationLinks
    ' Description : find and link citation markings like phonetic in transcript
    '============================================================================


    'for each result from Word doc query

    'look up in database
    'if yes
    'use current code
    'if not
    'look up on courtlistener
    'for each courtlistener result, list in input box
    'prompt input box choice (enter 1-2-3-4-5 etc)
    'enter choice into database
    'use choice in transcript or goto correct code place
            
            
    
    'Rule ##
    'Rule ##.##
    'Rule ###(#)
    'Rule ###
    'RCW ##.###.###
    'usc ## U.S.C. ####
    'Rule ER ###
    'CrR ###, CrR ##
    'CR ##.##
    'RAP ##.##
    'Section ### or ####
    'Title # or ##
    'case law
        
    'sCitationList(x)will be: insert search key after syntax?
    '(cht) <syntax from above> | garbage to let drop off
        
    Dim sAbsoluteURL As String
    Dim sURL As String
    Dim apiWaxLRS As String
    Dim sCaseName As String
    Dim sInputState As String
    Dim sToken As String
    Dim sInput1 As String
    Dim sInput2 As String
    Dim sInputCourt As String
    Dim sCourt As String
    Dim sFile1 As String
    Dim qReplaceHyperlink As String
    Dim sQLongCitation As String
    Dim sQCHCategory As String
    Dim sQWebAddress As String
    Dim sCurrentCitation As String
    Dim sCitationList() As String
    Dim sHyperlinkList() As String
    Dim sCurrentLinkSQL As String
    Dim sCurrentLinkFC As String
    Dim sCurrentLinkRH As String
    Dim sCurrentLinkLC As String
    Dim sCurrentLinkCHC As String
    Dim sCurrentLinkWeb As String
    Dim sBeginCHT As String
    Dim sEndCHT As String
    Dim sCurrentSearch As String
    Dim sCurrentTerm As String
    Dim sCLChoiceList As String
    Dim sQFindCitation As String
    Dim sInitialSearchSQL As String
    Dim sOriginalSearchTerm As String
    Dim vBookmarkName As String
    
    Dim Parsed As Dictionary
    
    Dim x As Integer
    Dim y As Integer
    Dim z As Integer
    Dim j As Integer
    Dim iLongCitationLength As Integer
    Dim iStartPos As Integer
    Dim iStopPos As Integer
    Dim i As Integer
    
    Dim letter As Variant
    Dim sSearchTermArray() As Variant
    Dim rep As Variant
    Dim resp As Variant
    Dim sCitation As Variant
    Dim oEntry As Variant
    
    Dim sID As Object
    Dim oCitations As Object
    Dim oRequest As Object
    Dim vDetails As Object
    
    Dim rCurrentCitation As Range
    Dim rCurrentSearch As Range
    Dim oWordApp As Word.Application
    Dim oWordDoc As Word.Document
    
    Dim rstCurrentHyperlink As DAO.Recordset
    Dim rstCurrentSearchMatching As DAO.Recordset
    
    Dim cJob As New Job
    
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField] 'job number
    
    x = 1
    
    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If
    
    
    Set oWordApp = CreateObject("Word.Application")
    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.CourtCover) 'open word document
    oWordApp.Visible = True
    
    y = 1
    sBeginCHT = "(cht) "
    sEndCHT = " |"
    
    ReDim sSearchTermArray(0 To 0)
    'Get all the document text and store it in a variable. 'TODO: What is going on here?
    Set rCurrentSearch = oWordDoc.Range
     
    On Error Resume Next
     
    Set oWordApp = GetObject(, "Word.Application")
    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.CourtCover) 'open word document
    oWordApp.Visible = True
    
    y = 1
    sBeginCHT = "(cht) "
    sEndCHT = " |"
    
    ReDim sSearchTermArray(0 To 0)
    
    
    Set rCurrentSearch = oWordDoc.Range
    sCurrentSearch = rCurrentSearch.Text
    sCurrentLinkSQL = "SELECT * FROM CitationHyperlinks WHERE [FindCitation]=" & Chr(34) 'TODO: add usc table
    'Loop sCurrentSearch till you can't find any more matching "terms"
    x = UBound(sSearchTermArray) - LBound(sSearchTermArray) + 1
    Debug.Print x
    Do Until x = 0
        If y > 1 Then
            sCurrentLinkSQL = sCurrentLinkSQL & " OR [FindCitation]=" & Chr(34)
        End If
        If x = 0 Then GoTo Done
        iStartPos = InStr(x, sCurrentSearch, sBeginCHT, vbTextCompare)
        If iStartPos = 0 Then GoTo ExitLoop
        iStopPos = InStr(iStartPos, sCurrentSearch, sEndCHT, vbTextCompare)
        If iStopPos = 0 Then GoTo ExitLoop
        sCurrentTerm = Mid$(sCurrentSearch, iStartPos + Len(sBeginCHT), iStopPos - iStartPos - Len(sEndCHT))
        x = InStr(iStopPos, sCurrentSearch, sBeginCHT, vbTextCompare)
        sCurrentTerm = Left(sCurrentTerm, Len(sCurrentTerm) - 4)
        
        'add term to array which we will use to search document again later
        ReDim Preserve sSearchTermArray(UBound(sSearchTermArray) + 1)
        sSearchTermArray(UBound(sSearchTermArray)) = sCurrentTerm
        'construct sql statement from this
        sCurrentLinkSQL = sCurrentLinkSQL & sCurrentTerm & Chr(34)
        Debug.Print "Current Search Term:  " & sCurrentTerm
        Debug.Print "Current Search Array:  " & Join(sSearchTermArray, ", ")
        Debug.Print "SQL Statement:  " & sCurrentLinkSQL
        Debug.Print "----------------------------------------------------------"
        
        sOriginalSearchTerm = ""
        
        y = y + 1
    
    
    Loop
    
    
ExitLoop:
    sCurrentLinkSQL = sCurrentLinkSQL & ";"
    Debug.Print "Final Search Array:  " & Join(sSearchTermArray, ", ")
    Debug.Print "Final SQL Statement:  " & sCurrentLinkSQL
    Debug.Print "----------------------------------------------------------"
    'MsgBox "I'm done"
    
        
    'query those from citationhyperlinks and get hyperlink info back
    x = 1
    z = 0
            
    'sSearchTermArray Join(sSearchTermArray, ", ")
            
    sInputState = InputBox("Enter name of state here.  Will also search federal and special court jurisdictions.")
    For x = 1 To (UBound(sSearchTermArray) - 1)
        sInitialSearchSQL = "SELECT * FROM CitationHyperlinks WHERE [FindCitation] = " & Chr(34) & "*" & sSearchTermArray(x) & "*" & Chr(34) & ";"
        'look each one up in CitationHyperlinks
        Debug.Print "Initial Search SQL" & sInitialSearchSQL
        'TODO: What is going on here?
        'GoTo NextSearchTerm
        Set rstCurrentSearchMatching = CurrentDb.OpenRecordset(sInitialSearchSQL)
                
        On Error Resume Next
        rstCurrentSearchMatching.MoveFirst
        On Error GoTo 0
                
        'if result is NOT in CitationHyperlinks do this
        If rstCurrentSearchMatching.EOF = True Then
                                    
            'look up on courtlistener
                                        
                    
            sInput1 = sSearchTermArray(x)        'search term 1
            sInput2 = ""                         'search term 2     'enter name of state here, 'federal', 'special'
            sInputCourt = "scotus+ca1+ca2+ca3+ca4+ca5+ca6+ca7+ca8+ca9+ca10+ca11+cadc+cafc+ag+afcca+asbca+armfor+acca+uscfc+tax+mc+mspb+nmcca+cavc+bva+fiscr+fisc+cit+usjc+jpml+sttex+stp+cc+com+ccpa+cusc+eca+tecoa+reglrailreorgct+kingsbench"
            'which courts go with which states
            sOriginalSearchTerm = sInput1
            If sInputState = "Alabama" Then
                sInputCourt = "almd+alnd+alsd+almb+alnb+alsb+ala+alactapp+alacrimapp+alacivapp"
            ElseIf sInputState = "special" Then sInputCourt = "ag+afcca+asbca+armfor+acca+uscfc+tax+mc+mspb+nmcca+cavc+bva+fiscr+fisc+cit+usjc+jpml+sttex+stp+cc+com+ccpa+cusc+eca+tecoa+reglrailreorgct+kingsbench"
            ElseIf sInputState = "federal" Then sInputCourt = "scotus+ca1+ca2+ca3+ca4+ca5+ca6+ca7+ca8+ca9+ca10+ca11+cadc+cafc"
            ElseIf sInputState = "Alaska" Then sInputCourt = "akd+akb+alaska+alaskactapp+" & sInputCourt
            ElseIf sInputState = "Arizona" Then sInputCourt = "azd+arb+ariz+arizctapp+ariztaxct+" & sInputCourt
            ElseIf sInputState = "Arkansas" Then sInputCourt = "ared+arwd+areb+arwb+ark+arkctapp+arkworkcompcom+arkag+" & sInputCourt
            ElseIf sInputState = "California" Then sInputCourt = "cacd+caed+cand+casd+californiad+caca+cacb+caeb+canb+casb+cal+calctapp+calappdeptsuper+calag+" & sInputCourt
            ElseIf sInputState = "Colorado" Then sInputCourt = "cod+cob+colo+coloctapp+coloworkcompcom+coloag+" & sInputCourt
            ElseIf sInputState = "Connecticut" Then sInputCourt = "ctd+ctb+conn+connappct+connsuperct+connworkcompcom+" & sInputCourt
            ElseIf sInputState = "Delaware" Then sInputCourt = "ded+circtdel+deb+del+delch+delsuperct+delctcompl+delfamct+deljudct+" & sInputCourt
            ElseIf sInputState = "Florida" Then sInputCourt = "flmd+flnd+flsd+flmb+flnb+flsb+fla+fladistctapp+flaag+" & sInputCourt
            ElseIf sInputState = "Georgia" Then sInputCourt = "gamd+gand+gasd+gamb+ganb+gasb+ga+gactapp+" & sInputCourt
            ElseIf sInputState = "Hawaii" Then sInputCourt = "hid+hib+haw+hawapp+" & sInputCourt
            ElseIf sInputState = "Idaho" Then sInputCourt = "idd+idb+idaho+idahoctapp+" & sInputCourt
            ElseIf sInputState = "Illinois" Then sInputCourt = "ilcd+ilnd+ilsd+illinoised+illinoisd+ilcb+ilnb+ilsb+ill+illappct+" & sInputCourt
            ElseIf sInputState = "Indiana" Then sInputCourt = "innd+insd+indianad+innb+insb+ind+indctapp+indtc+" & sInputCourt
            ElseIf sInputState = "Iowa" Then sInputCourt = "iand+iasd+ianb+iasb+iowa+iowactapp+" & sInputCourt
            ElseIf sInputState = "Kansas" Then sInputCourt = "ksd+ksb+kan+kanctapp+kanag+" & sInputCourt
            ElseIf sInputState = "Kentucky" Then sInputCourt = "kyed+kywd+kyeb+kywb+ky+kyctapp+kyctapphigh+" & sInputCourt
            ElseIf sInputState = "Louisiana" Then sInputCourt = "laed+lamd+lawd+laeb+lamb+lawb+la+lactapp+laag+" & sInputCourt
            ElseIf sInputState = "Maine" Then sInputCourt = "med+bapme+meb+me+" & sInputCourt
            ElseIf sInputState = "Maryland" Then sInputCourt = "mdd+mdb+md+mdctspecapp+mdag+" & sInputCourt
            ElseIf sInputState = "Massachusetts" Then sInputCourt = "mad+bapma+mab+mass+massappct+masssuperct+massdistct+maworkcompcom+" & sInputCourt
            ElseIf sInputState = "Michigan" Then sInputCourt = "mied+miwd+mieb+miwb+mich+michctapp+" & sInputCourt
            ElseIf sInputState = "Minnesota" Then sInputCourt = "mnd+mnb+minn+minnctapp+minnag+" & sInputCourt
            ElseIf sInputState = "Mississippi" Then sInputCourt = "msnd+mssd+msnb+mssb+miss+missctapp+" & sInputCourt
            ElseIf sInputState = "Missouri" Then sInputCourt = "moed+mowd+moeb+mowb+mo+moctapp+moag+" & sInputCourt
            ElseIf sInputState = "Montana" Then sInputCourt = "mtd+mtb+mont+monttc+montag+" & sInputCourt
            ElseIf sInputState = "Nebraska" Then sInputCourt = "ned+nebraskab+neb+nebctapp+nebag+" & sInputCourt
            ElseIf sInputState = "Nevada" Then sInputCourt = "nvd+nvb+nev+" & sInputCourt
            ElseIf sInputState = "New Hampshire" Then sInputCourt = "nhd+nhb+nh+" & sInputCourt
            ElseIf sInputState = "New Jersey" Then sInputCourt = "njd+njb+nj+njsuperctappdiv+njtaxct+njch+" & sInputCourt
            ElseIf sInputState = "New Mexico" Then sInputCourt = "nmd+nmb+nm+nmctapp+" & sInputCourt
            ElseIf sInputState = "New York" Then sInputCourt = "nyed+nynd+nysd+nywd+nyeb+nynb+nysb+nywb+ny+nyappdiv+nyappterm+nysupct+nyfamct+nysurct+nycivct+nycrimct+nyag+" & sInputCourt
            ElseIf sInputState = "North Carolina" Then sInputCourt = "nced+ncmd+ncwd+circtnc+nceb+ncmb+ncwb+nc+ncctapp+ncsuperct+ncworkcompcom+" & sInputCourt
            ElseIf sInputState = "North Dakota" Then sInputCourt = "ndd+ndb+nd+ndctapp+" & sInputCourt
            ElseIf sInputState = "Ohio" Then sInputCourt = "ohnd+ohsd+ohiod+ohnb+ohsb+ohio+ohioctapp+ohioctcl+" & sInputCourt
            ElseIf sInputState = "Oklahoma" Then sInputCourt = "oked+oknd+okwd+okeb+oknb+okwb+okla+oklacivapp+oklacrimapp+oklajeap+oklacoj+oklaag+" & sInputCourt
            ElseIf sInputState = "Oregon" Then sInputCourt = "ord+orb+or+orctapp+ortc+" & sInputCourt
            ElseIf sInputState = "Pennsylvania" Then sInputCourt = "paed+pamd+pawd+pennsylvaniad+paeb+pamb+pawb+pa+pasuperct+pacommwct+cjdpa+stp+" & sInputCourt
            ElseIf sInputState = "Rhode Island" Then sInputCourt = "rid+rib+ri+risuperct+" & sInputCourt
            ElseIf sInputState = "South Carolina" Then sInputCourt = "scd+southcarolinaed+southcarolinawd+scb+sc+scctapp+" & sInputCourt
            ElseIf sInputState = "South Dakota" Then sInputCourt = "sdd+sdb+sd+" & sInputCourt
            ElseIf sInputState = "Tennessee" Then sInputCourt = "tned+tnmd+tnwd+tennessed+circttenn+tneb+tnmb+tnwb+tennesseeb+tenn+tennctapp+tenncrimapp+tennsuperct+" & sInputCourt
            ElseIf sInputState = "Texas" Then sInputCourt = "txed+txnd+txsd+txwd+txeb+txnb+txsb+txwb+tex+texapp+texcrimapp+texreview+texjpml+texag+sttex+" & sInputCourt
            ElseIf sInputState = "Utah" Then sInputCourt = "utd+utb+utah+utahctapp+" & sInputCourt
            ElseIf sInputState = "Vermont" Then sInputCourt = "vtd+vtb+vt+vtsuperct+" & sInputCourt
            ElseIf sInputState = "Virginia" Then sInputCourt = "vaed+vawd+vaeb+vawb+va+vactapp+" & sInputCourt
            ElseIf sInputState = "Washington" Then sInputCourt = "waed+wawd+waeb+wawb+wash+washctapp+washag+" & sInputCourt
            ElseIf sInputState = "West Virginia" Then sInputCourt = "wvnd+wvsd+wvnb+wvsb+wva+" & sInputCourt
            ElseIf sInputState = "Wisconsin" Then sInputCourt = "wied+wiwd+wieb+wiwb+wis+wisctapp+wisag+" & sInputCourt
            ElseIf sInputState = "Wyoming" Then sInputCourt = "wyd+wyb+wyo+" & sInputCourt
            End If
            If sInput2 = "" Then
                'only input1
                sURL = "https://www.courtlistener.com/api/rest/v3/search/" & "?q=" & sInput1 & "&court=" & sInputCourt & "&order_by=score+desc&stat_Precedential=on" & "&fields=caseName" '
            Else
                'with input2
                'https://www.courtlistener.com/?type=o&q=westview&type=o&order_by=score+desc&stat_Precedential=on&court=waed+wawd+waeb+wawb+wash+washctapp+washag
                sURL = "https://www.courtlistener.com/api/rest/v3/search/" & "?q=" & sInput1 & "&q=" & sInput2 & "&court=" & sInputCourt & "&order_by=score+desc&stat_Precedential=on" & "&fields=caseName" '
            End If
                    
            With CreateObject("WinHttp.WinHttpRequest.5.1")
                .Open "GET", sURL, False         'options or head
                .setRequestHeader "Accept", "application/json"
                .setRequestHeader "content-type", "application/x-www-form-urlencoded"
                'make sure this is token and that it works; wasn't assigned correctly prior to change
                .setRequestHeader "Authorization", "Bearer " & Environ("apiCourtListener")
                .send
                apiWaxLRS = .responseText
                .abort
                'Debug.Print apiWaxLRS
                'Debug.Print "--------------------------------------------"
            End With
            x = 1
            y = 1
            Set Parsed = JsonConverter.ParseJson(apiWaxLRS)
            Set sID = Parsed.item("results")
            
            'create new table TempCitations 'TODO: What is going on here?
                                        
            On Error Resume Next
            CurrentDb.Execute "DROP TABLE TempCitations"
            On Error GoTo 0
            CurrentDb.Execute "CREATE TABLE TempCitations (ID COUNTER(1, 1) PRIMARY KEY, " & _
                              "j NUMBER, " & _
                              "OriginalTerm TEXT, " & _
                              "FindCitation TEXT, " & _
                              "ReplaceHyperlink TEXT, " & _
                              "LongCitation TEXT, " & _
                              "WebAddress TEXT );"
            j = 1
            sCLChoiceList = "Current Transcript Term:  " & sOriginalSearchTerm & Chr(10) & Chr(10) & "Choices:  " & Chr(10)
            'list first 15 results only
            For j = 1 To 15
                    
                For Each rep In sID
                        
                    If Not IsNull(rep.item("citation")) Then
                                
                        Set oCitations = rep.item("citation")
                                    
                        For Each oEntry In oCitations
                            If oEntry = "null" Or oEntry = "" Or oEntry = Null Or oEntry = "Null" Then
                                sCitation = sCitation & ", " & "null"
                                'Debug.Print sCitation
                                            
                            ElseIf oEntry <> "" Then
                                sCitation = sCitation & ", " & oEntry
                                            
                                'Debug.Print sCitation
                                            
                            Else
                                        
                                Set sCitation = oEntry
                                For Each resp In sCitation
                                    sCitation = sCitation & ", " & oEntry
                                    y = x + 1
                                    'Debug.Print sCitation
                                                
                                Next
                                            
                                y = 1
                                            
                            End If
                                        
                        Next
                                    
                    Else
                                
                        sCitation = sCitation & ", " & "Null"
                                    
                    End If
                    sCitation = Right(sCitation, Len(sCitation) - 2)
                    'Debug.Print "Citation Number:  " & sCitation
                    sCaseName = rep.item("caseName")
                    'Debug.Print "Case Name:  " & sCaseName
                    sAbsoluteURL = rep.item("absolute_url")
                    'Debug.Print "URL:  https://www.courtlistener.com" & sAbsoluteURL
                    sCourt = rep.item("court")
                    'Debug.Print "Court:  " & sCourt
                            
                    'prepare to add citations to new temporary table TempCitations
                    sQFindCitation = Left(sCaseName, 255)
                    qReplaceHyperlink = "https://www.courtlistener.com/" & sAbsoluteURL 'format is test#http://www.cnn.com#, can't add hyperlink field to a table in vba
                    iLongCitationLength = 253 - Len(sCitation)
                            
                    'TODO: What is going on here?
                    'GoTo NextSearchTerm
                    sQFindCitation = Left(sQFindCitation, iLongCitationLength)
                    sQLongCitation = sQFindCitation & ", " & sCitation
                            
                    'Debug.Print "Short Citation:  " & sQFindCitation
                    'Debug.Print "Long Citation:  " & Len(sQLongCitation) & " letters, " & sQLongCitation
                    x = x + 1
                    'Debug.Print "--------------------------------------------"
                            
                    'string for input box
                    sCLChoiceList = sCLChoiceList & "(" & j & ")" & sQLongCitation & Chr(10) & sAbsoluteURL & Chr(10) & "-------------------" & Chr(10)
                            
                            
                    'add citations to new temporary table TempCitations with field j
                            
                    'TODO: What is going on here?
                    'GoTo NextSearchTerm
                    Set rstCurrentHyperlink = CurrentDb.OpenRecordset("TempCitations")
                    rstCurrentHyperlink.AddNew
                    rstCurrentHyperlink.Fields("OriginalTerm").Value = sOriginalSearchTerm
                    rstCurrentHyperlink.Fields("j").Value = j
                    rstCurrentHyperlink.Fields("FindCitation").Value = sQFindCitation
                            
                    rstCurrentHyperlink.Fields("ReplaceHyperlink").Value = qReplaceHyperlink
                            
                    rstCurrentHyperlink.Fields("LongCitation").Value = sQLongCitation
                    rstCurrentHyperlink.Fields("WebAddress").Value = sAbsoluteURL
                    rstCurrentHyperlink.Update
                            
                    y = 1
                            
                    sCitation = ""
                    j = j + 1
                Next
                        
            Next
            rstCurrentHyperlink.Close
                    
            MsgBox ("On the next screen, you will enter a number to choose one of the following authorities:" & Chr(10) & Chr(10) & "(0)" & " None." & _
                    sCLChoiceList)
                    
            'prompt input box choice (enter 0-1-2-3-4-5 etc)
            z = InputBox("Enter your choice from 0 to 15 for which result to use, " & _
                         "0 being none and 1-15 being a choice of result from CourtListener.")
                    
                    
                    
            If z = 0 Then GoTo NextSearchTerm
            'put choice into proper variables
            Set rstCurrentHyperlink = CurrentDb.OpenRecordset("SELECT * FROM TempCitations WHERE j =" & z & ";")
            sQFindCitation = rstCurrentHyperlink.Fields("FindCitation").Value
            'test#http://www.cnn.com#
            qReplaceHyperlink = sQFindCitation & "#" & rstCurrentHyperlink.Fields("WebAddress").Value & "#"
            sQLongCitation = rstCurrentHyperlink.Fields("LongCitation").Value
            sAbsoluteURL = "https://www.courtlistener.com" & rstCurrentHyperlink.Fields("WebAddress").Value
            rstCurrentHyperlink.Close
                    
            'after choice made and stored, delete temporary table TempCitations
            DoCmd.SetWarnings False
            DoCmd.DeleteObject acTable = acDefault, "TempCitations"
            DoCmd.SetWarnings True
                            
            'do something with your choice
            Select Case z
                    
                Case 0
                    'go to next term to search
                    GoTo NextSearchTerm
                                
                Case Else
                    'enter choice into database
                    Set rstCurrentHyperlink = CurrentDb.OpenRecordset("CitationHyperlinks")
                    rstCurrentHyperlink.AddNew
                    rstCurrentHyperlink.Fields("FindCitation").Value = sOriginalSearchTerm 'sQFindCitation
                                
                    rstCurrentHyperlink.Fields("ReplaceHyperlink").Value = qReplaceHyperlink
                                
                    rstCurrentHyperlink.Fields("LongCitation").Value = sQLongCitation
                    rstCurrentHyperlink.Fields("ChCategory").Value = 1
                    rstCurrentHyperlink.Fields("WebAddress").Value = sAbsoluteURL
                    rstCurrentHyperlink.Update
                    'Debug.Print "Citation:  " & sQLongCitation & " | Web Address:  " & sAbsoluteURL
                    'Debug.Print "z:  " & z
                    'Debug.Print "----------------------------------------------------------"
                                
                            
            End Select
                    
        Else
            'if it is in the database already do this stuff
                    
            'Set rstCurrentSearchMatching = CurrentDb.OpenRecordset(sInitialSearchSQL)
            'rstCurrentSearchMatching.MoveFirst
                    
        End If
                
NextSearchTerm:
    Next
            
    'run refreshed sql statement and proceed as normal
            
    x = 1
            
    'this is the original sql statement constructed, which is what we want
    'format sCurrentLinkSQL = "SELECT * FROM CitationHyperlinks WHERE [FindCitation] = " & Chr(34) & "*" & sCitationList(x) & "*" & Chr(34)
    Debug.Print "Current SQL Statement:  " & sCurrentLinkSQL
    If iStartPos = 0 Then GoTo ExitLoop1
            
    'TODO: What is going on here?
    'GoTo NextSearchTerm
    'GoTo NextSearchTerm
    Set rstCurrentHyperlink = CurrentDb.OpenRecordset(sCurrentLinkSQL)
                
    Do While rstCurrentHyperlink.EOF = False
                
        sQFindCitation = rstCurrentHyperlink.Fields("FindCitation").Value
                
        'test#http://www.cnn.com#
        qReplaceHyperlink = rstCurrentHyperlink.Fields("ReplaceHyperlink").Value
                
        sQLongCitation = rstCurrentHyperlink.Fields("LongCitation").Value
        sQCHCategory = rstCurrentHyperlink.Fields("ChCategory").Value
        sQWebAddress = rstCurrentHyperlink.Fields("WebAddress").Value
                
        'Debug.Print "Citation:  " & sQFindCitation & " | Web Address:  " & sQWebAddress
        'Debug.Print "x:  " & x
        'If x > 0 Then Debug.Print "Current Search Term x-1:  " & sSearchTermArray(UBound(sSearchTermArray) - 1)
        'Debug.Print "Current Search Term x-2:  " & sSearchTermArray(x - 2)
        'Debug.Print "Current Search Term x-3:  " & sSearchTermArray(x - 3)
        'Debug.Print "Current Search Term:  " & sSearchTermArray(x - 1)
        'Debug.Print "Current Search Term x-4:  " & sSearchTermArray(x - 4)
        'Debug.Print "Current Search Term x-5:  " & sSearchTermArray(x - 5)
        'Debug.Print "Current Search Term x-6:  " & sSearchTermArray(x - 6)
        'Debug.Print "Current Search Term x-7:  " & sSearchTermArray(x - 7)
        'Debug.Print "Current Search Term x-8:  " & sSearchTermArray(x - 8)
        'Debug.Print UBound(sSearchTermArray())
        'Debug.Print UBound(sSearchTermArray())
                
        Debug.Print "----------------------------------------------------------"
            
        'use that info on word document
                
        'TODO: What is going on here?
        'GoTo NextSearchTerm
        oWordDoc.Application.Selection.Find.ClearFormatting
        oWordDoc.Application.Selection.Find.Replacement.ClearFormatting
                
        oWordDoc.Content.Select
                                 
        With oWordDoc.Application.Selection
                    
            .Find.Text = sQFindCitation
            .Find.Replacement.Text = sQFindCitation
            .Find.Forward = True
            .Find.Wrap = wdFindStop
            .Find.Format = False
            .Find.MatchCase = True
                    
        End With
                    
        If oWordDoc.Application.Selection.Find.Execute = True Then
                    
            oWordDoc.Application.Selection.Text = sQFindCitation
            oWordDoc.TablesOfAuthorities.MarkAllCitations ShortCitation:=sQFindCitation, _
                                                          LongCitation:=sQLongCitation, LongCitationAutoText:=sQLongCitation, Category:=sQCHCategory
                         
            oWordDoc.Application.Selection.Find.Execute Replace:=wdReplaceAll
        Else
        End If
                    
        oWordDoc.Application.Selection.HomeKey Unit:=wdStory
                       
        Do While oWordDoc.Application.Selection.Find.Execute
                    
            oWordDoc.Application.Selection.Text = sQFindCitation
                         
            oWordDoc.Hyperlinks.Add Anchor:=oWordDoc.Application.Selection.Range, _
                                    Address:=sQWebAddress, ScreenTip:=sQLongCitation & ":" & Chr(13) & sQWebAddress, _
                                    TextToDisplay:=sQFindCitation
                             
        Loop
                                 
        'End With
               
        If sQCHCategory = "1" Then
                
            oWordDoc.Application.ActiveWindow.ActivePane.View.ShowAll = Not oWordDoc.Application.ActiveWindow.ActivePane.View.ShowAll
                                    
                    
                    
            With oWordDoc.Application.Selection.Find
                .Text = "\s " & Chr(34) & sQFindCitation & Chr(34)
                .Replacement.Text = "\s " & Chr(34) & sQLongCitation & Chr(34)
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            oWordDoc.Application.Selection.Find.Execute Replace:=wdReplaceAll
                    
            Debug.Print "\s fixed for cases"
                
            'find "\s " & Chr(34) & sQFindCitation & Chr(34)
            'replace "s " & Chr(34) & sQLongCitation & Chr(34)
                                
        End If
        sQFindCitation = ""
        qReplaceHyperlink = ""
        sQLongCitation = ""
        sQCHCategory = ""
        sQWebAddress = ""
        
        oWordDoc.Application.Selection.HomeKey Unit:=wdStory
        rstCurrentHyperlink.MoveNext
        x = x + 1
    Loop
    rstCurrentHyperlink.Close
    
ExitLoop1:
    With oWordDoc.Application.Selection.Find
    
        .ClearFormatting
        .Replacement.ClearFormatting
    
        .Text = " |"
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
    
    
        .Text = "(cht) "
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
        
        .ClearFormatting
        .Replacement.ClearFormatting
    
        .Text = ", Wa "
        .Replacement.Text = ", Washington "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
    
        .Text = "^p" & "." & "^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
    
        .Text = "^p" & " ." & "^p"
        .Replacement.Text = "^p"
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
    
Done:
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField] 'TODO: job number
    oWordDoc.SaveAs2 FileName:=cJob.DocPath.CourtCover         'save and close word doc
    oWordDoc.Close wdDoNotSaveChanges
    oWordApp.Quit
    
    
    Set oWordDoc = Nothing
    Set oWordApp = Nothing
    
End Sub

Public Sub pfTopOfTranscriptBookmark()

    Dim bTitle As String
    Dim n As String
    
    Dim AcroApp As Acrobat.CAcroApp
    Dim PDoc As Acrobat.CAcroPDDoc
    Dim PDocCover As Acrobat.CAcroPDDoc
    Dim ADoc As AcroAVDoc
    Dim PDocAll As Acrobat.CAcroPDDoc
    Dim parentBookmark As AcroPDBookmark
    Dim PDBookmark As AcroPDBookmark
    Dim PDFPageView As AcroAVPageView
    
    Dim jso As Object
    Dim BookMarkRoot As Object
    Dim oPDFBookmarks As Object
    
    Dim numpages As Variant

    Dim cJob As New Job
    
    Set AcroApp = CreateObject("AcroExch.App")
    '@Ignore AssignmentNotUsed
    Set PDoc = CreateObject("AcroExch.PDDoc")
    Set PDocAll = CreateObject("AcroExch.PDDoc")
    Set PDocCover = CreateObject("AcroExch.PDDoc")
    Set ADoc = CreateObject("AcroExch.AVDoc")
    Set PDBookmark = CreateObject("AcroExch.PDBookmark", "")

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

    PDocCover.Open (cJob.DocPath.WACoverP)

    Set ADoc = PDocCover.OpenAVDoc(cJob.DocPath.WACoverP)
    
    'Table of Contents Bookmark
    '@Ignore AssignmentNotUsed
    Set PDFPageView = ADoc.GetAVPageView()
    Call PDFPageView.Goto(0)
    AcroApp.MenuItemExecute ("NewBookmark")

    'TODO: pfTopOfTranscriptBookmarks make sure this is correc
    '@Ignore AssignmentNotUsed
    bTitle = PDBookmark.GetByTitle(PDocCover, "Untitled")
    '@Ignore AssignmentNotUsed, AssignmentNotUsed
    bTitle = PDBookmark.SetTitle("Table of Contents")

    n = PDocCover.Save(PDSaveFull, cJob.DocPath.WACoverP)

    ' Insert the pages of Part2 after the end of Part1
    numpages = PDocCover.GetNumPages()
    
    PDoc.Open (cJob.DocPath.TranscriptFP)

    Set ADoc = PDoc.OpenAVDoc(cJob.DocPath.TranscriptFP)
    SendKeys ("^{HOME}")
    Set PDFPageView = ADoc.GetAVPageView()

    AppActivate "Adobe Acrobat Pro Extended"

    'Top of Transcript Bookmark
    Call PDFPageView.Goto(0)
    AcroApp.MenuItemExecute ("NewBookmark")

    'TODO: pfTopOfTranscriptBookmark Make sure this is correct
    '@Ignore AssignmentNotUsed
    bTitle = PDBookmark.GetByTitle(PDoc, "Untitled")
    '@Ignore AssignmentNotUsed
    bTitle = PDBookmark.SetTitle("TOP OF TRANSCRIPT")

    'Index Bookmark
    Call PDFPageView.Goto(1)
    AcroApp.MenuItemExecute ("NewBookmark")

    'TODO: pfTopOfTranscriptBookmark Make sure this is correct
    '@Ignore AssignmentNotUsed
    bTitle = PDBookmark.GetByTitle(PDoc, "Untitled")
    '@Ignore AssignmentNotUsed
    bTitle = PDBookmark.SetTitle("TRANSCRIPT INDEXES")

    'General Index Bookmark
    Call PDFPageView.Goto(0)
    AcroApp.MenuItemExecute ("NewBookmark")

    'TODO: pfTopOfTranscriptBookmark Make sure this is correct
    '@Ignore AssignmentNotUsed
    bTitle = PDBookmark.GetByTitle(PDoc, "Untitled")
    '@Ignore AssignmentNotUsed
    bTitle = PDBookmark.SetTitle("General")

    'Witnesses Index Bookmark
    Call PDFPageView.Goto(0)
    AcroApp.MenuItemExecute ("NewBookmark")

    'TODO: pfTopOfTranscriptBookmark Make sure this is correct
    '@Ignore AssignmentNotUsed
    bTitle = PDBookmark.GetByTitle(PDoc, "Untitled")
    '@Ignore AssignmentNotUsed
    bTitle = PDBookmark.SetTitle("Witnesses")

    'Exhibits Index Bookmark
    Call PDFPageView.Goto(0)
    AcroApp.MenuItemExecute ("NewBookmark")

    'TODO: pfTopOfTranscriptBookmark Make sure this is correct
    '@Ignore AssignmentNotUsed
    bTitle = PDBookmark.GetByTitle(PDoc, "Untitled")
    '@Ignore AssignmentNotUsed
    bTitle = PDBookmark.SetTitle("Exhibits")

    'Cases Bookmark
    Call PDFPageView.Goto(0)
    AcroApp.MenuItemExecute ("NewBookmark")

    'TODO: pfTopOfTranscriptBookmark Make sure this is correct
    '@Ignore AssignmentNotUsed
    bTitle = PDBookmark.GetByTitle(PDoc, "Untitled")
    '@Ignore AssignmentNotUsed
    bTitle = PDBookmark.SetTitle("Cases")

    'Rules, Regulation, Code, Statutes Bookmark
    Call PDFPageView.Goto(0)
    AcroApp.MenuItemExecute ("NewBookmark")

    'TODO: pfTopOfTranscriptBookmark Make sure this is correct
    '@Ignore AssignmentNotUsed
    bTitle = PDBookmark.GetByTitle(PDoc, "Untitled")
    '@Ignore AssignmentNotUsed
    bTitle = PDBookmark.SetTitle("Rules, Regulation, Code, Statutes")

    'Other Authorities Bookmark
    Call PDFPageView.Goto(0)
    AcroApp.MenuItemExecute ("NewBookmark")

    'TODO: pfTopOfTranscriptBookmark Make sure this is correct
    '@Ignore AssignmentNotUsed
    bTitle = PDBookmark.GetByTitle(PDoc, "Untitled")
    '@Ignore AssignmentNotUsed
    bTitle = PDBookmark.SetTitle("Other Authorities")

    MsgBox ("Hit okay after you've moved all bookmarks to their corresponding headings; General, Witnesses, or Exhibits; and also created the Authorities bookmark structure.")

    n = PDoc.Save(PDSaveFull, cJob.DocPath.TranscriptFP)
    
    'for each -Transcript-FINAL.pdf in \Transcripts\ do the following

    If PDocCover.InsertPages(numpages - 1, PDoc, 0, PDoc.GetNumPages(), True) = False Then
        MsgBox "Cannot insert pages"
    End If

    If PDocCover.Save(PDSaveFull, cJob.DocPath.WAConsolidatedP) = False Then
        MsgBox "Cannot save the modified document"
    End If
    
    PDoc.Close
    PDocCover.Close
    AcroApp.Exit
    Set AcroApp = Nothing
    Set PDoc = Nothing
    Set ADoc = Nothing

End Sub


