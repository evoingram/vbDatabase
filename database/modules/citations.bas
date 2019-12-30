Attribute VB_Name = "citations"
'@Ignore OptionExplicit
'@Folder("Database.Production.Modules")
Option Compare Database

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
    Dim sInput1 As String
    Dim sInput2 As String
    Dim sInputCourt As String
    Dim sCourt As String
    Dim qReplaceHyperlink As String
    Dim sQLongCitation As String
    Dim sQCHCategory As String
    Dim sQWebAddress As String
    Dim sCitationList() As String
    Dim sHyperlinkList() As String
    Dim sCurrentLinkSQL As String
    Dim sBeginCHT As String
    Dim sEndCHT As String
    Dim sCurrentSearch As String
    Dim sCurrentTerm As String
    Dim sCLChoiceList As String
    Dim sQFindCitation As String
    Dim sInitialSearchSQL As String
    Dim sOriginalSearchTerm As String
    Dim sCHSQLstm1 As String
    Dim sUSCSQLstm2 As String
    Dim sCH1SQLstm3 As String
    
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim j As Long
    Dim iLongCitationLength As Long
    Dim iStartPos As Long
    Dim iStopPos As Long
    
    Dim sSearchTermArray() As Variant
    Dim rep As Variant
    Dim resp As Variant
    Dim sCitation As Variant
    Dim oEntry As Variant
    
    Dim sID As Object
    Dim oCitations As Object
    
    Dim rCurrentCitation As Range
    Dim rCurrentSearch As Range
    Dim oWordApp As Word.Application
    Dim oWordDoc As Word.Document
    
    Dim rstCurrentHyperlink As DAO.Recordset
    Dim rstCurrentSearchMatching As DAO.Recordset
    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID
    Forms![NewMainMenu].Form!lblFlash.Caption = "Step 10 of 10:  Processing citations found..."
    
    x = 1
    
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
    'Get all the document text and store it in a variable.
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
    sCurrentLinkSQL = "SELECT * FROM CitationHyperlinks WHERE [FindCitation]=" & Chr(34)
    sCHSQLstm1 = "SELECT * FROM CitationHyperlinks WHERE [FindCitation]=" & Chr(34)
    sUSCSQLstm2 = " UNION " & _
                  "SELECT * FROM USC WHERE [USC].[FindCitation]=" & Chr(34)
    sCH1SQLstm3 = " UNION " & _
                  "SELECT * FROM CitationHyperlinks1 WHERE [CitationHyperlinks1].[FindCitation]=" & Chr(34)
    'Loop sCurrentSearch till you can't find any more matching "terms"
    'x = UBound(sSearchTermArray) - LBound(sSearchTermArray) + 1
    Debug.Print x
    
    'this loop adds to the search array everything marked found in the document
    Do Until x = 0
        If y > 1 Then
            sCurrentLinkSQL = sCurrentLinkSQL & " OR [CitationHyperlinks1].[FindCitation]=" & Chr(34)
            sCHSQLstm1 = sCHSQLstm1 & " OR [CitationHyperlinks].[FindCitation]=" & Chr(34)
            sUSCSQLstm2 = sUSCSQLstm2 & " OR [USC].[FindCitation]=" & Chr(34)
            sCH1SQLstm3 = sCH1SQLstm3 & " OR [CitationHyperlinks1].[FindCitation]=" & Chr(34)
        End If
        If x = 0 Then GoTo Done
        iStartPos = InStr(x, sCurrentSearch, sBeginCHT, vbTextCompare)
        If iStartPos = 0 Then GoTo ExitLoop
        iStopPos = InStr(iStartPos, sCurrentSearch, sEndCHT, vbTextCompare)
        If iStopPos = 0 Then GoTo ExitLoop
        sCurrentTerm = Mid$(sCurrentSearch, iStartPos + Len(sBeginCHT), iStopPos - iStartPos - Len(sEndCHT))
        x = InStr(iStopPos, sCurrentSearch, sBeginCHT, vbTextCompare)
        'Debug.Print x
        sCurrentTerm = Left(sCurrentTerm, Len(sCurrentTerm) - 4)
        
        'add term to array which we will use to search document again later
        ReDim Preserve sSearchTermArray(UBound(sSearchTermArray) + 1)
        sSearchTermArray(UBound(sSearchTermArray)) = sCurrentTerm
        'construct sql statement from this
        sCurrentLinkSQL = sCurrentLinkSQL & sCurrentTerm & Chr(34) & _
                            " UNION " & _
                            "SELECT * FROM USC WHERE [USC].[FindCitation]=" & Chr(34) & "*" & sCurrentTerm & "*" & Chr(34) & _
                            " UNION " & _
                            "SELECT * FROM CitationHyperlinks1 WHERE [CitationHyperlinks1].[FindCitation]=" & Chr(34) & sCurrentTerm & Chr(34)
        sCHSQLstm1 = sCHSQLstm1 & sCurrentTerm & Chr(34)
        sUSCSQLstm2 = sUSCSQLstm2 & sCurrentTerm & Chr(34)
        sCH1SQLstm3 = sCH1SQLstm3 & sCurrentTerm & Chr(34)
        
        'Debug.Print "Current Search Term:  " & sCurrentTerm
        'Debug.Print "Current Search Array:  " & Join(sSearchTermArray, ", ")
        'Debug.Print "SQL Statement:  " & sCurrentLinkSQL
        'Debug.Print "----------------------------------------------------------"
        
        sOriginalSearchTerm = ""
        
        y = y + 1
    
    Loop
    
    
ExitLoop:
    sCurrentLinkSQL = sCurrentLinkSQL & ";"
    
    sInitialSearchSQL = sCHSQLstm1 & _
                        sUSCSQLstm2 & _
                        sCH1SQLstm3 & ";"
                            
    'Debug.Print "Final Search Array:  " & Join(sSearchTermArray, ", ")
    Debug.Print "Original Final SQL Statement:  " & sCurrentLinkSQL
    Debug.Print "New Final SQL Statement:  " & sInitialSearchSQL
    'Debug.Print "----------------------------------------------------------"
    'MsgBox "I'm done"
    
        
    'query those from citationhyperlinks and get hyperlink info back
    x = 1
    z = 0
            
    'sSearchTermArray Join(sSearchTermArray, ", ")
            
    sInputState = InputBox("Enter name of state here.  Will also search federal and special court jurisdictions.")
    For x = 1 To (UBound(sSearchTermArray) - 1)
        'Debug.Print UBound(sSearchTermArray) - 1
        'Debug.Print x
        sCHSQLstm1 = "SELECT * FROM CitationHyperlinks WHERE [FindCitation]=" & Chr(34) & "*" & sSearchTermArray(x) & "*" & Chr(34)
        sUSCSQLstm2 = " UNION " & _
                      "SELECT * FROM USC WHERE [USC].[FindCitation]=" & Chr(34) & "*" & sSearchTermArray(x) & "*" & Chr(34)
        sCH1SQLstm3 = " UNION " & _
                      "SELECT * FROM CitationHyperlinks1 WHERE [CitationHyperlinks1].[FindCitation]=" & Chr(34) & sCurrentTerm & Chr(34) & ";"
                      
        sInitialSearchSQL = sCHSQLstm1 & _
                            sUSCSQLstm2 & _
                            sCH1SQLstm3
        'look each one up in CitationHyperlinks
        'Debug.Print "Initial Search SQL = " & sInitialSearchSQL
        Set rstCurrentSearchMatching = CurrentDb.OpenRecordset(sInitialSearchSQL)
                
        On Error Resume Next
        rstCurrentSearchMatching.MoveFirst
        On Error GoTo 0
                
        'if result is NOT in CitationHyperlinks do this
        If rstCurrentSearchMatching.EOF = True Then
                                    
            'look up on courtlistener
                    
            sInput1 = sSearchTermArray(x)        'search term 1
            sInput2 = ""                         'search term 2     'enter name of state here, 'federal', 'special''which courts go with which states
            sOriginalSearchTerm = sInput1
            
            Set sID = apiCourtListener(sInputState, sOriginalSearchTerm)
            
            
            CurrentDb.Execute "DELETE FROM TempCitations"
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
                            
                    sQFindCitation = Left(sQFindCitation, iLongCitationLength)
                    sQLongCitation = sQFindCitation & ", " & sCitation
                            
                    Debug.Print "Short Citation:  " & sQFindCitation
                    Debug.Print "Long Citation:  " & Len(sQLongCitation) & " letters, " & sQLongCitation
                    Debug.Print "--------------------------------------------"
                            
                    'string for input box
                    sCLChoiceList = sCLChoiceList & "(" & j & ")" & sQLongCitation & Chr(10) & sAbsoluteURL & Chr(10) & "-------------------" & Chr(10)
                    
                            
                    'add citations to new temporary table TempCitations with field j
                            
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
                    
            'after choice made and stored, delete contents of TempCitations
            CurrentDb.Execute "DELETE FROM TempCitations"
                            
            'do something with your choice
            Select Case z
                    
                Case 0
                    'go to next term to search
                    GoTo NextSearchTerm
                                
                Case Else
                    'enter choice into database
                    Set rstCurrentHyperlink = CurrentDb.OpenRecordset("CitationHyperlinks1")
                    rstCurrentHyperlink.AddNew
                    rstCurrentHyperlink.Fields("FindCitation").Value = sOriginalSearchTerm 'sQFindCitation
                                
                    rstCurrentHyperlink.Fields("ReplaceHyperlink").Value = qReplaceHyperlink
                                
                    rstCurrentHyperlink.Fields("LongCitation").Value = sQLongCitation
                    rstCurrentHyperlink.Fields("ChCategory").Value = 1
                    rstCurrentHyperlink.Fields("WebAddress").Value = sAbsoluteURL
                    rstCurrentHyperlink.Update
                    Debug.Print "Citation:  " & sQLongCitation & " | Web Address:  " & sAbsoluteURL
                    Debug.Print "z:  " & z
                    Debug.Print "----------------------------------------------------------"
                                
                            
            End Select
                    
        Else
            'if it is in the database already do this stuff
                    
            'Set rstCurrentSearchMatching = CurrentDb.OpenRecordset(sInitialSearchSQL)
            'rstCurrentSearchMatching.MoveFirst
            
        End If
                
NextSearchTerm:
    'x = x + 1
    Debug.Print x
    
    Next
            
            
    'run refreshed sql statement and proceed as normal
            
    x = 1
            
    'this is the original sql statement constructed, which is what we want
    'format sCurrentLinkSQL = "SELECT * FROM CitationHyperlinks WHERE [FindCitation] = " & Chr(34) & "*" & sCitationList(x) & "*" & Chr(34)
    Debug.Print "Current SQL Statement:  " & sInitialSearchSQL
    If iStartPos = 0 Then GoTo ExitLoop1
            
        
    Set rstCurrentHyperlink = CurrentDb.OpenRecordset(sInitialSearchSQL)
                    
                
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
    
    CurrentDb.Execute "DELETE FROM TempCitations"
    
    
Done:
    oWordDoc.Save 'save and close word doc
    oWordDoc.Close wdDoNotSaveChanges
    oWordApp.Quit
    
    
    Set oWordDoc = Nothing
    Set oWordApp = Nothing

    sCourtDatesID = ""
End Sub

Public Function apiCourtListener(sInputState, sInput1, Optional sInput2 As String)

Dim sInputCourt As String
Dim sURL As String
Dim apiWaxLRS As String

Dim parsed As Dictionary
            sInputCourt = "scotus+ca1+ca2+ca3+ca4+ca5+ca6+ca7+ca8+ca9+ca10+ca11+cadc+cafc+ag+afcca+asbca+armfor+acca+uscfc+tax+mc+mspb+nmcca+cavc+bva+fiscr+fisc+cit+usjc+jpml+sttex+stp+cc+com+ccpa+cusc+eca+tecoa+reglrailreorgct+kingsbench"

            If sInputState = "Alabama" Then
                sInputCourt = "almd+alnd+alsd+almb+alnb+alsb+ala+alactapp+alacrimapp+alacivapp+" & sInputCourt
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
                .setRequestHeader "Authorization", "Bearer " & Environ("apiCourtListener")
                .send
                apiWaxLRS = .responseText
                .abort
                'Debug.Print apiWaxLRS
                'Debug.Print "--------------------------------------------"
            End With
            Set parsed = JsonConverter.ParseJson(apiWaxLRS)
            Set apiCourtListener = parsed.item("results")
    
End Function
