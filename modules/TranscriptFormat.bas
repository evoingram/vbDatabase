Attribute VB_Name = "TranscriptFormat"
'@Ignore OptionExplicit
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
'                                         Optional wsyWordStyle = vbNullString, bForward = True, bWrap = "wdFindContinue"
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

Public Sub test1()
    '============================================================================
    ' Name        : pfTCEntryReplacement
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfTCEntryReplacement
    ' Description : parent function that finds certain entries within a transcript and assigns TC entries to them for indexing purposes
    '============================================================================
    
    Dim oWordApp As New Word.Application
    Dim oCourtCoverWD As New Word.Document
    
    Dim rstViewJFAppQ As DAO.Recordset

    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID
    '@Ignore AssignmentNotUsed
    
    On Error Resume Next
    Set oWordApp = GetObject(cJob.DocPath.CourtCover)

    Set oWordApp = GetObject(, "Word.Application")
    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    Set oCourtCoverWD = oWordApp.Documents.Open(cJob.DocPath.CourtCover)
    oWordApp.Application.Visible = False
    
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

    sCourtDatesID = vbNullString
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
    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID

    ActiveDocument.Application.Selection.Find.ClearFormatting
    ActiveDocument.Application.Selection.Find.Replacement.ClearFormatting

    With ActiveDocument.Application.Selection.Find
        sBookmarkName = "RoughBKMK"
        .Text = "#RBMK#"
        ActiveDocument.Bookmarks.Add Name:=sBookmarkName
        sBookmarkName = "CertBMK"
        .Text = "#CBMK#"
        ActiveDocument.Bookmarks.Add Name:=sBookmarkName
        sBookmarkName = "ToABMK"
        .Text = "#TBMK#"
        ActiveDocument.Bookmarks.Add Name:=sBookmarkName
        sBookmarkName = "TopLine"
        .Text = "#TOPL#"
        ActiveDocument.Bookmarks.Add Name:=sBookmarkName
        sBookmarkName = "EndTime"
        .Text = "#ENDT#"
        ActiveDocument.Bookmarks.Add Name:=sBookmarkName
    End With
    With ActiveDocument                          'insert topline at TopLine bookmark

        If .Bookmarks.Exists("TopLine") = True Then
    
            .Bookmarks("TopLine").Select
            .Application.Selection.TypeText Text:=UCase(cJob.Location) & ", " & _
                                                                          FormatDateTime(Format(cJob.HearingDate, "mm-dd-yyyy"), vbLongDate) & ", " & UCase(Format(cJob.HearingStartTime, "h:mm AM/PM"))
        Else
            MsgBox "Bookmark ""TopLine"" does not exist!"
        End If
        .MailMerge.MainDocumentType = wdNotAMergeDocument

        If .Bookmarks.Exists("EndTime") = True Then
    
            If Right(Format(cJob.HearingEndTime, "h:mm AM/PM"), 2) = "AM" Then
                Format(cJob.HearingEndTime, "h:mm AM/PM") = Replace(Format(cJob.HearingEndTime, "h:mm AM/PM"), "AM", "a.m.")
        
            ElseIf Right(Format(cJob.HearingEndTime, "h:mm AM/PM"), 2) = "PM" Then
                Format(cJob.HearingEndTime, "h:mm AM/PM") = Replace(Format(cJob.HearingEndTime, "h:mm AM/PM"), "PM", "p.m.")
        
            End If
    
            .Bookmarks("EndTime").Select
            .Application.Selection.TypeText Text:=UCase(Format(cJob.HearingEndTime, "h:mm AM/PM"))
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
    'TODO: rewrite this function
    Dim sBookmarkName As String
    Dim vBookmarkName As String
    Dim sTopLine As String
    
    Dim oWordDoc As Word.Document
    Dim oWordApp As Word.Application
    
    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID

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
        .Replacement.Text = vbNullString
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
    oWordDoc.Bookmarks.Add Name:=sBookmarkName


    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting

    With oWordDoc.Application.Selection.Find
        .Text = "#RBMK#"
        .Replacement.Text = vbNullString
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
    oWordDoc.Bookmarks.Add Name:=sBookmarkName


    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting

    With oWordDoc.Application.Selection.Find
        .Text = "#TOPL#"
        .Replacement.Text = vbNullString
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
    oWordDoc.Bookmarks.Add Name:=sBookmarkName

    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting

    With oWordDoc.Application.Selection.Find
        .Text = "#ENDT#"
        .Replacement.Text = vbNullString
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
    oWordDoc.Bookmarks.Add Name:=sBookmarkName

    oWordApp.Application.Selection.Find.ClearFormatting
    oWordApp.Application.Selection.Find.Replacement.ClearFormatting
    '

    With oWordDoc.Application.Selection.Find
        .Text = "#CBMK#"
        .Replacement.Text = vbNullString
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
    oWordDoc.Bookmarks.Add Name:=sBookmarkName


    oWordApp.Application.Selection.Find.ClearFormatting
    oWordApp.Application.Selection.Find.Replacement.ClearFormatting

    With oWordDoc.Application.Selection.Find
        .Text = "#TBMK#"
        .Replacement.Text = vbNullString
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
    oWordDoc.Bookmarks.Add Name:=sBookmarkName

    oWordApp.Application.Selection.Find.ClearFormatting
    oWordApp.Application.Selection.Find.Replacement.ClearFormatting

    With oWordDoc.Application.Selection.Find
        .Text = "#TOCA#"
        .Replacement.Text = vbNullString
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
    oWordDoc.Bookmarks.Add Name:=sBookmarkName

    sBookmarkName = "IndexB"

    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting
    With oWordDoc.Application.Selection.Find
        .Text = "#TOCB#"
        .Replacement.Text = vbNullString
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

    oWordDoc.Bookmarks.Add Name:=sBookmarkName

    sBookmarkName = "IndexC"

    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting
    With oWordDoc.Application.Selection.Find
        .Text = "#TOCC#"
        .Replacement.Text = vbNullString
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

    oWordDoc.Bookmarks.Add Name:=sBookmarkName

    sBookmarkName = "IndexD"

    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting
    With oWordDoc.Application.Selection.Find
        .Text = "#TOCD#"
        .Replacement.Text = vbNullString
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

    oWordDoc.Bookmarks.Add Name:=sBookmarkName

    sBookmarkName = "IndexE"

    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting
    With oWordDoc.Application.Selection.Find
        .Text = "#TOCE#"
        .Replacement.Text = vbNullString
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

    oWordDoc.Bookmarks.Add Name:=sBookmarkName


    vBookmarkName = "TOAC"

    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting

    With oWordDoc.Application.Selection.Find

        .Text = "#TOAC#"
        .Replacement.Text = vbNullString
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

    oWordDoc.Bookmarks.Add Name:=vBookmarkName



    vBookmarkName = "TOAR"
    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting

    With oWordDoc.Application.Selection.Find

        .Text = "#TOAR#"
        .Replacement.Text = vbNullString
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

    oWordDoc.Bookmarks.Add Name:=vBookmarkName

    vBookmarkName = "TOAO"
    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting

    With oWordDoc.Application.Selection.Find
        .Text = "#TOAO#"
        .Replacement.Text = vbNullString
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

    oWordDoc.Bookmarks.Add Name:=vBookmarkName

    With oWordDoc                                'insert topline at TopLine bookmark

        If .Bookmarks.Exists("TopLine") = True Then
    
            If Right(Format(cJob.HearingStartTime, "h:mm AM/PM"), 6) = ":00 AM" Then
                Format(cJob.HearingStartTime, "h:mm AM/PM") = Replace(Format(cJob.HearingStartTime, "h:mm AM/PM"), ":00 AM", " a.m.")
        
            ElseIf Right(Format(cJob.HearingStartTime, "h:mm AM/PM"), 6) = ":00 PM" Then
                Format(cJob.HearingStartTime, "h:mm AM/PM") = Replace(Format(cJob.HearingStartTime, "h:mm AM/PM"), ":00 PM", " p.m.")
        
            End If
            
            sTopLine = UCase(cJob.Location) & ", " & UCase(FormatDateTime(Format(cJob.HearingDate, "mm-dd-yyyy"), vbLongDate)) & ", " & UCase(Format(cJob.HearingStartTime, "h:mm AM/PM"))

            .Bookmarks("TopLine").Select
            oWordDoc.Application.Selection.Font.Underline = wdUnderlineSingle
            oWordDoc.Application.Selection.TypeText Text:=sTopLine
            oWordDoc.Application.Selection.Font.Underline = wdUnderlineNone
        Else
            MsgBox "Bookmark ""TopLine"" does not exist!"
        End If
        .MailMerge.MainDocumentType = wdNotAMergeDocument

        If .Bookmarks.Exists("EndTime") = True Then
    
            If Right(Format(cJob.HearingEndTime, "h:mm AM/PM"), 6) = ":00 AM" Then
                Format(cJob.HearingEndTime, "h:mm AM/PM") = Replace(Format(cJob.HearingEndTime, "h:mm AM/PM"), ":00 AM", " a.m.")
        
            ElseIf Right(Format(cJob.HearingEndTime, "h:mm AM/PM"), 6) = ":00 PM" Then
                Format(cJob.HearingEndTime, "h:mm AM/PM") = Replace(Format(cJob.HearingEndTime, "h:mm AM/PM"), ":00 PM", " p.m.")
        
            End If
    
            .Bookmarks("EndTime").Select
            .Application.Selection.TypeText Text:=Format(cJob.HearingEndTime, "h:mm AM/PM")
        Else
            MsgBox "Bookmark ""EndTime"" does not exist!"
        End If
        .MailMerge.MainDocumentType = wdNotAMergeDocument

    End With

    oWordDoc.Save
    oWordDoc.Close (wdSaveChanges)
    oWordApp.Quit



    sCourtDatesID = vbNullString
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
    
    On Error Resume Next
    Set oWordApp = GetObject(, "Word.Application")

    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If

    'oWordApp.Activate
    Set oWordDoc = oWordApp.Documents.Open(sCourtDatesID & "-CourtCover.docx")
    On Error GoTo 0
    
    Set oWordDoc = GetObject(sCourtDatesID & "-CourtCover.docx")

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

    oWordDoc.Save
    'oWordDoc.SaveAs2 FileName:=sCourtDatesID & "-CourtCover.docx"

    sCourtDatesID = vbNullString
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
    
    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID

    On Error Resume Next
    Set oWordApp = GetObject(, "Word.Application")

    If Err <> 0 Then
        Set oWordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0

    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.CourtCover)

    With oWordDoc
        .Application.Selection.Goto What:=wdGoToBookmark, Name:="IndexA"
        With oWordDoc.Bookmarks
            .DefaultSorting = wdSortByName
            .ShowHidden = False
        End With
        .TablesOfContents.Add Range:=.Application.Selection.Range, UseHeadingStyles:=False, UseFields:=True, TableID:="a", RightAlignPageNumbers:=True, IncludePageNumbers:=True, UseHyperlinks:=True
        '.indexes.Add Range:=.Application.Selection.Range, HeadingSeparator:= _
        'wdHeadingSeparatorNone, Type:=wdIndexRunin, NumberOfColumns:= _
        'wdAutoPosition, IndexLanguage:=wdEnglishUS
        '.indexes(1).TabLeader = wdTabLeaderDots
        .Application.Selection.Goto What:=wdGoToBookmark, Name:="IndexB"
        With oWordDoc.Bookmarks
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
        With oWordDoc.Bookmarks
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
        With oWordDoc.Bookmarks
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
        With oWordDoc.Bookmarks
            .DefaultSorting = wdSortByName
            .ShowHidden = False
        End With
        .TablesOfContents.Add Range:=.Application.Selection.Range, UseHeadingStyles:=False, UseFields:=True, TableID:="e", RightAlignPageNumbers:=True, IncludePageNumbers:=True, UseHyperlinks:=True
        '.indexes.Add Range:=.Application.Selection.Range, HeadingSeparator:= _
        'wdHeadingSeparatorNone, Type:=wdIndexRunin, NumberOfColumns:= _
        'wdAutoPosition, IndexLanguage:=wdEnglishUS
        '.indexes(1).TabLeader = wdTabLeaderDots
    
     
        If InStr(cJob.CaseInfo.Jurisdiction, "AVT") = 0 And InStr(cJob.CaseInfo.Jurisdiction, "eScribers") = 0 And InStr(cJob.CaseInfo.Jurisdiction, "FDA") = 0 And InStr(cJob.CaseInfo.Jurisdiction, "Food and Drug Administration") = 0 And InStr(cJob.CaseInfo.Jurisdiction, "Weber") = 0 Then
                
            
            oWordDoc.Application.Selection.Goto What:=wdGoToBookmark, Name:="TOAC"
            With oWordDoc.Bookmarks
                .DefaultSorting = wdSortByName
                .ShowHidden = False
            End With
            .TablesOfAuthorities.Add Range:=oWordDoc.Application.Selection.Range, Category:=1, Passim _
                                     :=False, KeepEntryFormatting:=True
            .TablesOfAuthorities(1).TabLeader = wdTabLeaderDots
            .TablesOfAuthorities.Format = wdIndexIndent
             
            
             
            oWordDoc.Application.Selection.Goto What:=wdGoToBookmark, Name:="TOAR"
            With oWordDoc.Bookmarks
                .DefaultSorting = wdSortByName
                .ShowHidden = False
            End With
            .TablesOfAuthorities.Add Range:=oWordDoc.Application.Selection.Range, Category:=2, Passim _
                                     :=False, KeepEntryFormatting:=True
            .TablesOfAuthorities(1).TabLeader = wdTabLeaderDots
            .TablesOfAuthorities.Format = wdIndexIndent
             
             
            oWordDoc.Application.Selection.Goto What:=wdGoToBookmark, Name:="TOAO"
            With oWordDoc.Bookmarks
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
            .Replacement.Text = vbNullString
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
            .Replacement.Text = vbNullString
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
    
        oWordDoc.Save
        oWordDoc.Close
        oWordApp.Quit
        

    sCourtDatesID = vbNullString
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
    
    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID

    'Set oWordApp = CreateObject("Word.Application")
    oWordApp.Visible = False
    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.CourtCover)
    If cJob.CaseInfo.Jurisdiction = "Food and Drug Administration" Then
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
                FindFDA1 = vbNullString
                FindFDA2 = vbNullString
                FindFDA3 = vbNullString
                FindFDA4 = vbNullString
                ReplaceWithName = vbNullString
            
                rs.MoveNext
            Loop
        Else
            MsgBox "There are no records in the recordset."
        End If
    Else
    End If
    
    oWordDoc.Save
    oWordDoc.Close
    rs.Close
    rs1.Close
    oWordApp.Quit
    Set oWordDoc = Nothing
    Set rs = Nothing
    Set rs1 = Nothing
    Set oWordApp = Nothing
    
    sCourtDatesID = vbNullString
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
    
    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID
    
    x = 18                                       '18 is number of first dynamic speaker

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
                sMrMs = vbNullString
                sLastName = vbNullString
                vSpeakerName = vbNullString
            
                x = x + 1                        'add 1 to x for next speaker name
                rs.MoveNext                      'go to next speaker name
                
            Loop 'back up to the top
        Else

            MsgBox "There are no dynamic speakers."
            
        End If
    
        'MsgBox "Finished looping through dynamic speakers."
    
        rs.Close
        Set rs = Nothing
        oWordDoc.SaveAs2 FileName:=cJob.DocPath.CourtCover
        oWordDoc.Close
        oWordApp.Quit
    

        Set oWordDoc = Nothing
        Set oWordApp = Nothing

    End With

    sCourtDatesID = vbNullString
End Sub

Public Sub pfSingleFindReplace(ByVal sTextToFind As String, ByVal sReplacementText As String, Optional ByVal wsyWordStyle As String = vbNullString, Optional bForward As Boolean = True, _
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
    
    Dim oWordDoc As Word.Document
    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID
    
    Set oWordDoc = GetObject(cJob.DocPath.CourtCover)
    oWordDoc.Application.Visible = False
    oWordDoc.Application.Selection.HomeKey Unit:=wdStory
    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting
   
    With oWordDoc.Application.Selection.Find
        .Text = sTextToFind
        .Replacement.Text = sReplacementText
        If wsyWordStyle <> vbNullString Then
            .Replacement.Style = oWordDoc.Styles(wsyWordStyle)
        Else
        End If
        .Forward = bForward
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
End Sub

Public Sub pfSingleTCReplaceAll(ByVal sTextToFind As String, ByVal sReplacementText As String, Optional ByVal wsyWordStyle As String = vbNullString, Optional bForward As Boolean = True, _
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
    
    Dim oWordDoc As Word.Document
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID

    On Error Resume Next
    
    Set oWordDoc = GetObject(cJob.DocPath.CourtCover)
    oWordDoc.Application.Visible = False
    oWordDoc.Application.Selection.HomeKey Unit:=wdStory

    With oWordDoc.Application

        .Selection.Find.ClearFormatting
    
        With .Selection.Find
            .Text = sTextToFind
            .Replacement.Text = sReplacementText
            If wsyWordStyle <> vbNullString Then
                .Replacement.Style = oWordDoc.Styles(wsyWordStyle)
            Else
            End If
            If bForward <> Empty Then
                .Forward = bForward
            Else
                .Forward = True
            End If
            If bFormat <> Empty Then
                .Format = bFormat
            Else
                .Format = False
            End If
            If bMatchCase <> Empty Then
                .MatchCase = bMatchCase
            Else
                .MatchCase = False
            End If
            If bMatchWholeWord <> Empty Then
                .MatchWholeWord = bMatchWholeWord
            Else
                .MatchWholeWord = False
            End If
            If bMatchWildcards <> Empty Then
                .MatchWildcards = bMatchWildcards
            Else
                .MatchWildcards = False
            End If
            If bMatchSoundsLike <> Empty Then
                .MatchSoundsLike = bMatchSoundsLike
            Else
                .MatchSoundsLike = False
            End If
            If bMatchAllWordForms <> Empty Then
                .MatchAllWordForms = bMatchAllWordForms
            Else
                .MatchAllWordForms = False
            End If
        End With
    
        .Selection.Find.Execute Replace:=wdReplaceAll
    
    End With
    oWordDoc.Save
    On Error GoTo 0
End Sub

Public Sub pfFieldTCReplaceAll(sTexttoSearch As String, sReplacementText As String, sFieldText As String)
    '============================================================================
    ' Name        : pfFieldTCReplaceAll
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfFieldTCReplaceAll(sTexttoSearch, sReplacementText, sFieldText)
    ' Description : one replace TC entry function for ones with field entry
    '============================================================================

    
    Dim oWordDoc As Word.Document
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID

    'wdFieldTOCEntry
    

    Set oWordDoc = GetObject(cJob.DocPath.CourtCover)
    oWordDoc.Application.Visible = False
    
    oWordDoc.Application.Selection.HomeKey Unit:=wdStory
    oWordDoc.Application.Selection.Find.ClearFormatting
    oWordDoc.Application.Selection.Find.Replacement.ClearFormatting


    With oWordDoc.Application
        .Selection.Find.ClearFormatting
    
        With .Selection.Find
            .Text = sTexttoSearch
            .Replacement.Text = sReplacementText
            .Forward = True
            '.Wrap = wdFindContinue
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
    
    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID

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
                  "six,school,under,through,think,thought,things,thing,they're,the,these,there's,there,tell,telling,table,take,such,stattues,still,temporary,thrown,took,too,though,through,sure," & _
                  "wi,try,trying,tried,tries,see,seeing,saw,sees,self,person,persons,people," & _
                  "you've,you're,well,we'll,went,we're,why,what,who,will,way,wanted,want,very,us,until,week,weeks,yesterday,talk,talking,use,which,wherever,some,question,questions"
          
    With oWordDoc1
        .Application.DisplayAlerts = False
        '.Application.Visible = False
        sInput = .Content.Text
        
        For w = 1 To 255                         'hyphens & single quotes kept; strip unwanted chars
            Select Case w
            Case 1 To 35, 37 To 38, 40 To 43, 45, 47, 58 To 64, 91 To 96, 123 To 127, 129 To 144, 147 To 149, 152 To 162, 164, 166 To 171, 174 To 191, 247
                sInput = Replace(sInput, Chr(w), " ")
            End Select
        Next
        'TODO: word index -- insert somewhere an if statement:  if current entry to add is identical to previous, go to next page number
        sInput = Replace(Replace(Replace(Replace(sInput, Chr(44) & Chr(32), " "), Chr(44) & vbCr, " "), Chr(46) & Chr(32), " "), Chr(46) & vbCr, " ")
        sInput = Replace(Replace(Replace(Replace(sInput, Chr(145), "'"), Chr(146), "'"), "' ", " "), " '", " ")
        sInput = " " & LCase(Trim(sInput)) & " "
    
        For w = 0 To UBound(Split(sExclusions, ",")) 'loop through sExclusions
            Do While InStr(sInput, " " & Split(sExclusions, ",")(w) & " ") > 0
                sInput = Replace(sInput, " " & Split(sExclusions, ",")(w) & " ", " ")
            Loop
        Next
    
        Do While InStr(sInput, "  ") > 0
            sInput = Replace(sInput, "  ", " ")
        Loop
    
        sInput = " " & Trim(sInput) & " "
        x = UBound(Split(sInput, " "))
        z = x
    
        For w = 1 To x
            sCurrentEntryOriginal = Split(sInput, " ")(1) 'get word count
            Do While InStr(sInput, " " & sCurrentEntryOriginal & " ") > 0
                sInput = Replace(sInput, " " & sCurrentEntryOriginal & " ", " ")
            Loop
            y = z - UBound(Split(sInput, " "))   'calculate replaced count
            sCurrentIndexEntry = sCurrentIndexEntry & sCurrentEntryOriginal & vbTab & y & vbCr 'update current index entry
            z = UBound(Split(sInput, " "))
            If z = 1 Then Exit For
            DoEvents
        Next
    
        sInput = sCurrentIndexEntry
        sCurrentIndexEntry = vbNullString
        sCurrentEntry5 = UBound(Split(sInput, vbCr)) - 1
    
        For w = 0 To sCurrentEntry5
            sCurrentEntryOriginal = vbNullString
            With .Range
                With .Find
                    .ClearFormatting
                    sCurrentEntry4 = Split(Split(sInput, vbCr)(w), vbTab)(0)
                    sCurrentEntry1 = " " & Split(Split(sInput, vbCr)(w), vbTab)(1)
                    .Text = sCurrentEntry4
                    .Replacement.Text = vbNullString
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
                    If sCurrentEntryOriginal = vbNullString Then sCurrentEntryOriginal = sCurrentEntryOriginal & " " & .Information(wdActiveEndPageNumber)
                    sCurrentEntry1 = Right(sCurrentEntryOriginal, 2)
                    sCurrentEntry2 = " " & .Information(wdActiveEndPageNumber)
                    If sCurrentEntry1 = sCurrentEntry2 Then sCurrentEntryOriginal = sCurrentEntryOriginal
                    If sCurrentEntry1 <> sCurrentEntry2 Then sCurrentEntryOriginal = sCurrentEntryOriginal & " " & .Information(wdActiveEndPageNumber)
                    .Collapse (wdCollapseEnd)
                    .Find.Execute
                
                    If sCurrentEntry1 = vbNullString Then GoTo ExitLoop1
                Loop
ExitLoop1:
            End With
            sCurrentEntryOriginal = Replace(Trim(sCurrentEntryOriginal), " ", ",")
            sCurrentIndexEntry = sCurrentIndexEntry & Split(sInput, vbCr)(w) & vbTab & sCurrentEntryOriginal & vbCr
            If sCurrentEntryOriginal = vbNullString Then GoTo ExitLoop2
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
            .Replacement.Text = vbNullString
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
    
        .Bookmarks.Add Name:=vBookmarkName
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

    sCourtDatesID = vbNullString
End Sub

Public Sub FPJurors()
    '============================================================================
    ' Name        : FPJurors
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call FPJurors()
    ' Description : does find/replacements of prospective jurors in transcript
    '============================================================================

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

    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    x = 101                                      '101 is number of first PROSPECTIVE JUROR
    On Error Resume Next
    Forms![NewMainMenu].Form!lblFlash.Caption = "Step 9 of 10:  Processing prospective juror replacements..."
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
    
    sCourtDatesID = vbNullString
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
    
    Dim rstTRCourtQ As DAO.Recordset
    Dim rstViewJFAppQ As DAO.Recordset
    Dim qdf As QueryDef
    
    Dim oWordApp As New Word.Application
    Dim oCourtCoverWD As New Word.Document
    
    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID
    
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
            pfDelay 1
            Call pfFieldTCReplaceAll("(ema)", "^p(Exhibit ## marked and admitted.)^p", Chr(34) & "TC" & Chr(34) & " " & Chr(34) & "Exhibit  marked and admitted" & Chr(34) & " " & "\f cd")
            pfDelay 1
            Call pfFieldTCReplaceAll("(emm)", "^p(Exhibit ## marked.)^p", Chr(34) & "TC" & Chr(34) & " " & Chr(34) & "Exhibit  marked" & Chr(34) & " " & "\f cd")
            pfDelay 1
            Call pfFieldTCReplaceAll("(eaa)", "^p(Exhibit ## admitted.)^p", Chr(34) & "TC" & Chr(34) & " " & Chr(34) & "Exhibit  admitted" & Chr(34) & " " & "\f cd")
            pfDelay 1
            Call pfFieldTCReplaceAll("(exa)", "^p(Exhibit ## admitted.)^p", Chr(34) & "TC" & Chr(34) & " " & Chr(34) & "Exhibit  admitted" & Chr(34) & " " & "\f cd")
            pfDelay 1
        
            Call pfFieldTCReplaceAll("(ee1)", "^pDIRECT EXAMINATION^p", Chr(34) & "Direct Examination by " & Chr(34) & " \l 3")
            pfDelay 1
            Call pfFieldTCReplaceAll("(ee2)", "^pCROSS-EXAMINATION^p", Chr(34) & "Cross-Examination by " & Chr(34) & " \l 3")
            pfDelay 1
            Call pfFieldTCReplaceAll("(ee3)", "^pREDIRECT EXAMINATION^p", Chr(34) & "Redirect Examination by " & Chr(34) & " \l 3")
            pfDelay 1
            Call pfFieldTCReplaceAll("(ee4)", "^pRECROSS-EXAMINATION^p", Chr(34) & "Recross-Examination by " & Chr(34) & " \l 3")
            pfDelay 1
            Call pfFieldTCReplaceAll("(ee5)", "^pFURTHER REDIRECT EXAMINATION^p", Chr(34) & "Further Redirect Examination by " & Chr(34) & " \l 3")
            pfDelay 1
            Call pfFieldTCReplaceAll("(ee6)", "^pFURTHER RECROSS-EXAMINATION^p", Chr(34) & "Further Recross-Examination by " & Chr(34) & " \l 3")
            pfDelay 1
            Call pfFieldTCReplaceAll("(e1c)", "^pDIRECT EXAMINATION CONTINUED^p", Chr(34) & "Direct Examination Continued by " & Chr(34) & " \l 3")
            pfDelay 1
            Call pfFieldTCReplaceAll("(e2c)", "^pCROSS-EXAMINATION CONTINUED^p", Chr(34) & "Cross-Examination Continued by " & Chr(34) & " \l 3")
            pfDelay 1
            Call pfFieldTCReplaceAll("(e3c)", "^pREDIRECT EXAMINATION CONTINUED^p", Chr(34) & "Redirect Examination Continued by " & Chr(34) & " \l 3")
            pfDelay 1
            Call pfFieldTCReplaceAll("(e4c)", "^pRECROSS-EXAMINATION CONTINUED^p", Chr(34) & "Recross-Examination Continued by " & Chr(34) & " \l 3")
            pfDelay 1
            Call pfFieldTCReplaceAll("(e5c)", "^pFURTHER REDIRECT EXAMINATION CONTINUED^p", Chr(34) & "Further Redirect Examination Continued by " & Chr(34) & " \l 3")
            pfDelay 1
            Call pfFieldTCReplaceAll("(e6c)", "^pFURTHER RECROSS-EXAMINATION CONTINUED^p", Chr(34) & "Further Recross-Examination Continued by " & Chr(34) & " \l 3")
            pfDelay 1
            Call pfFieldTCReplaceAll("(crr)", "^pCOURT'S RULING" & "^p", Chr(34) & "TC" & Chr(34) & " " & Chr(34) & "Court's Ruling" & Chr(34) & " " & "\f e")
            pfDelay 1
            Call pfFieldTCReplaceAll("(aa1)", "^pARGUMENT FOR THE " & UCase(cJob.CaseInfo.Party1Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Argument for the " & cJob.CaseInfo.Party1Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
            pfDelay 1
            Call pfFieldTCReplaceAll("(ar1)", "^pREBUTTAL ARGUMENT FOR THE " & UCase(cJob.CaseInfo.Party1Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Rebuttal Argument for the " & cJob.CaseInfo.Party1Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
            pfDelay 1
            Call pfFieldTCReplaceAll("(ao1)", "^pOPENING STATEMENT FOR THE " & UCase(cJob.CaseInfo.Party1Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Opening Statement for the " & cJob.CaseInfo.Party1Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
            pfDelay 1
            Call pfFieldTCReplaceAll("(ac1)", "^pCLOSING ARGUMENT FOR THE " & UCase(cJob.CaseInfo.Party1Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Closing Argument for the " & cJob.CaseInfo.Party1Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
            pfDelay 1
            Call pfSingleTCReplaceAll("(sbb)", "^p(Sidebar begins at ##:## ap.m.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(sbe)", "^p(Sidebar ends at ##:## ap.m.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(rrr)", "^p(Recess taken from ##:## ap.m. to ##:## ap.m.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(sbn)", "^p(Sidebar taken from ##:## ap.m. to ##:## ap.m.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(jen)", "^p(Jury panel enters at ##:## ap.m.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(jex)", "^p(Jury panel exits at ##:## ap.m.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(pjn)", "^p(Prospective jury panel enters at ##:## ap.m.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(pjx)", "^p(Prospective jury panel exits at ##:## ap.m.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(wsu)", "^p(Witness summoned.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(wsw)", "^p(The witness was sworn.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(vub)", "^p(Video played at ##:## ap.m.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(vue)", "^p(Video ends at ##:## ap.m.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(vup)", "^p(Video played from ##:## ap.m. to ##:## ap.m.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(aup)", "^p(Audio played from ##:## ap.m. to ##:## ap.m.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(aue)", "^p(Audio ends at ##:## ap.m.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(aub)", "^p(Audio begins at ##:## ap.m.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(ccc)", "^p(Counsel confer.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(pcc)", "^p(Parties confer.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(ppr)", "^p(The witness paused to review the document.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(nrp)", "^p(No response.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(rrr)", "^p(Whereupon, at ##:## ap.m., a recess was taken.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(rrl)", "^p(Whereupon, at ##:## ap.m., a luncheon recess was taken.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(ppp)", "^p(Pause.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(otr)", "^p(Off the record.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(dtr)", "^p(Discussion held off the record.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(wxu)", "^p(Witness excused.)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(cco)", "^p(Whereupon, the following proceedings were held in open court outside the presence of the jury:)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("(cci)", "^p(Whereupon, the following proceedings were held in open court in the presence of the jury:)^p")
            pfDelay 1
            Call pfSingleTCReplaceAll("Uh-huh.", "Uh-huh.")
            pfDelay 1
            Call pfSingleTCReplaceAll("Huh-uh.", "Huh-uh.")
            pfDelay 1
            'Call pfFieldTCReplaceAll(, , )
        
            If Not rstViewJFAppQ.EOF Then rstViewJFAppQ.MoveNext
        
            If Not rstViewJFAppQ.EOF Then
                sMrMs2 = rstViewJFAppQ!MrMs      'get MrMs & LastName variables
                sLastName2 = rstViewJFAppQ!LastName
                Call pfFieldTCReplaceAll("(aa2)", "^pARGUMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Argument for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ar2)", "^pREBUTTAL ARGUMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Rebuttal Argument for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ao2)", "^pOPENING STATEMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Opening Statement for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ac2)", "^pCLOSING ARGUMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Closing Argument for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
            End If
        
            If Not rstViewJFAppQ.EOF Then rstViewJFAppQ.MoveNext
        
            If Not rstViewJFAppQ.EOF Then
                sMrMs2 = rstViewJFAppQ!MrMs      'get MrMs & LastName variables
                sLastName2 = rstViewJFAppQ!LastName
                Call pfFieldTCReplaceAll("(aa3)", "^pARGUMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Argument for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ar3)", "^pREBUTTAL ARGUMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Rebuttal Argument for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ao3)", "^pOPENING STATEMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Opening Statement for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ac3)", "^pCLOSING ARGUMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Closing Argument for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
            End If
        
            If Not rstViewJFAppQ.EOF Then rstViewJFAppQ.MoveNext
        
            If Not rstViewJFAppQ.EOF Then
                sMrMs2 = rstViewJFAppQ!MrMs      'get MrMs & LastName variables
                sLastName2 = rstViewJFAppQ!LastName
                Call pfFieldTCReplaceAll("(aa4)", "^pARGUMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Argument for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ar4)", "^pREBUTTAL ARGUMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Rebuttal Argument for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ao4)", "^pOPENING STATEMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Opening Statement for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ac4)", "^pCLOSING ARGUMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Closing Argument for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
            End If
        
            If Not rstViewJFAppQ.EOF Then rstViewJFAppQ.MoveNext
        
            If Not rstViewJFAppQ.EOF Then
                sMrMs2 = rstViewJFAppQ!MrMs      'get MrMs & LastName variables
                sLastName2 = rstViewJFAppQ!LastName
                Call pfFieldTCReplaceAll("(aa5)", "^pARGUMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Argument for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ar5)", "^pREBUTTAL ARGUMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Rebuttal Argument for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ao5)", "^pOPENING STATEMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Opening Statement for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ac5)", "^pCLOSING ARGUMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Closing Argument for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
            End If
        
            If Not rstViewJFAppQ.EOF Then rstViewJFAppQ.MoveNext
        
            If Not rstViewJFAppQ.EOF Then
                sMrMs2 = rstViewJFAppQ!MrMs      'get MrMs & LastName variables
                sLastName2 = rstViewJFAppQ!LastName
                Call pfFieldTCReplaceAll("(aa6)", "^pARGUMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Argument for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ar6)", "^pREBUTTAL ARGUMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Rebuttal Argument for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ao6)", "^pOPENING STATEMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Opening Statement for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
                Call pfFieldTCReplaceAll("(ac6)", "^pCLOSING ARGUMENT FOR THE " & UCase(cJob.CaseInfo.Party2Name) & " BY " & UCase(sMrMs2) & ". " & UCase(sLastName2) & "^p", "TC ""Closing Argument for the " & cJob.CaseInfo.Party2Name & " by " & sMrMs2 & ". " & sLastName2 & """ \f a")
            End If
        
            GoTo ParenDone
        End With
    End If
ParenDone:
    'Debug.Print "Finished looping through TC entries for the various parties."

    rstViewJFAppQ.Close
    Set rstViewJFAppQ = Nothing
    Set oCourtCoverWD = Nothing
    Set oWordApp = Nothing
    sCourtDatesID = vbNullString
End Sub


