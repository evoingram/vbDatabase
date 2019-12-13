Attribute VB_Name = "Speech"
'@Folder("Database.Speech.Modules")
Option Compare Database
Option Explicit

'============================================================================
'class module cmSpeech

'variables:
    'Private URLDownloadToFile as long
    
'functions:

    'pfGetFile:                                Description: read binary file as a string value
        '                                      Arguments:  sFileName
    'pfDoFolder:                               Description: cycles through subfolders of /Prepared/####/ and makes find/replaces listed below
        '                                      Arguments:  Folder
    'pfIEPostStringRequest:                    Description: sends URL encoded form data To the URL using IE
        '                                      Arguments:  sURL, sFormData, sBoundary
    'pfUploadFile:                             Description: uploads corpus file to Sphinx site to get LM/DIC files back
        '                                                   upload file using input type=file
        '                                      Arguments:  sDestinationURL, sCorpusPath, sFieldName
        '                                                  Optional sFieldName = "corpus"
    'pfCorpusUpload:                           Description: uploads corpus to lmtool at Sphinx to get compatible LM & DIC files back via download
        '                                      Arguments:  NONE
    'pfDownloadFile:                           Description: downloads provided file
        '                                      Arguments:  sURL, sSaveAs
    'pfAddSubfolder:                           Description: add concatenatedaudio folder to /UnprocessedAudio/####
        '                                      Arguments:  NONE
    'pfTrainEngine:                            Description: HOW TO PREPARE FILES FOR TRAINING
        '                                      Arguments:  NONE
    'pfPrepareAudio:                           Description: runs batch file audioprep / audioprep1.bat in Prepared folder
        '                                      Arguments:  NONE
    'pfSplitAudio:                             Description: runs batch file splitaudio / audioprep1.bat in Prepared folder
        '                                      Arguments:  NONE
    'pfRenameBaseFiles:                        Description: runs batch file FileRename / audioprep1.bat in Prepared folder
        '                                      Arguments:  NONE
    'pfSRTranscribe:                           Description: runs batch file SRTranscribe
        '                                      Arguments:  NONE
    'pfTrainAudio:                             Description: HOW TO TRAIN AUDIO / run batch file audiotrain.bat
        '                                      Arguments:  NONE
    'pfCopyTranscriptFromCompletedToPrepared:  Description: copies transcripts from /completed/ folder in T drive to corresponding /prepared/ folder in S drive
        '                                      Arguments:  NONE
    'pfMultipleAudioOneTranscript:             Description: HOW TO RUN MULTIPLE AUDIO FILES LIST
        '                                      Arguments:  NONE
    'pfPrepareTranscript:                      Description: makes changes to each transcript so it fits into speech recognition requirements
        '                                      Arguments:  NONE
    'pfRunCopyTranscTextBAT:                   Description: runs batch file CopyTranscriptTXT
        '                                      Arguments:  NONE

'============================================================================

Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal _
    szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    
Public Sub pfCopyTranscriptFromCompletedToPrepared()
'============================================================================
' Name        : pfCopyTranscriptFromCompletedToPrepared
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfCopyTranscriptFromCompletedToPrepared
' Description : copies transcripts from /completed/ folder in T drive to corresponding /prepared/ folder in S drive
'============================================================================

Dim sCompletedPath As String, sPreparedPath As String, sFullCompletedDocPath As String
Dim oCurrentFileString As String, sFileExtension As String, sFolderName As String
Dim oFolderObject As Object, oRootFolder As Object, oCurrentFile As Object, oSubfolder As Object

Set oFolderObject = CreateObject("Scripting.FileSystemObject")
sPreparedPath = "S:\UnprocessedAudio\Prepared\"   'Change to identify your main folder
sCompletedPath = "T:\Production\3Complete"
Set oRootFolder = oFolderObject.GetFolder(sPreparedPath)
For Each oSubfolder In oRootFolder.oSubfolders
    Set oCurrentFile = oSubfolder.Files
    sFolderName = Right(oSubfolder, 3)
    
    For Each oCurrentFile In oCurrentFile
    
        oCurrentFileString = oFolderObject.GetFileName(oCurrentFile)
        sFileExtension = oFolderObject.GetExtensionName(oCurrentFile)
        
        If sFileExtension = "docx" Or sFileExtension = "doc" Then
        
            If oCurrentFile Like "*.doc" Or oCurrentFile Like "*.docx" Then
            
                sFullCompletedDocPath = sCompletedPath & "\" & sFolderName & "\" & oCurrentFileString
                
                FileCopy sFullCompletedDocPath, oCurrentFile
                
                Debug.Print "Original Transcript:  " & sFullCompletedDocPath; ""
                Debug.Print "Copied to:  " & oCurrentFileString
                
            End If
            
        End If
        
    Next oCurrentFile
        
Next oSubfolder

End Sub

'four separate functions on separate schedules

    'preparing audio:  call PrepareAudio (loop)
    
    'preparing transcripts: call PrepareTranscript (loop for each docx)
    
'======>still need to match transcripts:audio files

'======>online lms

    'train engine on audio:  Call TrainAudio (loop)
    
    'transcribe audio:  call SRTranscribe
    
'change 'uNPrepared' to 'prepared' before you run these

Public Sub pfMultipleAudioOneTranscript()
'============================================================================
' Name        : pfMultipleAudioOneTranscript
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfMultipleAudioOneTranscript
' Description : HOW TO RUN MULTIPLE AUDIO FILES LIST
'============================================================================


'HOW TO RUN MULTIPLE AUDIO FILES LIST
'When running multiple audio files, PRIOR TO ABOVE STEPS:
    'Append transcripts of desired audio together into one file in order you are processing audio.
        'Use the following format for transcripts:
        '<s> the sentence goes here without any punctuation </s> (AudioFileNameinParenthesisNoExtension)
        'Within *.transcription, place audio file name only at last sentence of that file's transcription
'*.FileIDs should be in the following format
    'arctic_001
    'arctic_002
'Follow rest of steps above starting at #1.

'SO BASICALLY
    'match audio to transcripts with wavfilename in transcription
    'replace wavfilename at correct break points
    'call PrepareTranscript and it will create a new fileids
    'run as normal
    
End Sub


Public Sub pfPrepareTranscript()
'============================================================================
' Name        : pfPrepareTranscript
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfPrepareTranscript
' Description : makes changes to each transcript so it fits into speech recognition requirements
'============================================================================

Dim oFolderObject As Scripting.FileSystemObject
Dim oRootFolder As Object, oSubfolder As Object
Dim oPreparedWD As New Word.Document, oWordApp As New Word.Application, oCurrentFile As Object
Dim sPreparedPath As String, sFileExtension As String, sCurrentFile As String, sFolderName As String
Dim sTextToFind As String, sReplacementText As String
Dim wsStyle As Word.Style


sPreparedPath = "S:\UnprocessedAudio\Prepared\"
Set oFolderObject = CreateObject("Scripting.FileSystemObject")
Set oRootFolder = oFolderObject.GetFolder(sPreparedPath)

For Each oSubfolder In oRootFolder.oSubfolders
    sCourtDatesID = oSubfolder.ParentFolder.Name
    Debug.Print oSubfolder.Path
    Set oCurrentFile = oSubfolder.Files
    
    For Each oCurrentFile In oCurrentFile
        sFileExtension = oFolderObject.GetExtensionName(oCurrentFile)
        sCurrentFile = oFolderObject.GetFileName(oCurrentFile)
        If sFileExtension = "docx" Or sFileExtension = "doc" Then
            If Not oCurrentFile Like "*" & Chr(36) & "*" Then
            
                Debug.Print oSubfolder.Path
                
                Set oWordApp = CreateObject("Word.Application")
                Set oPreparedWD = oWordApp.Documents.Open(FileName:=oSubfolder.Path & "\" & sCurrentFile)
                oPreparedWD.Activate
                'preparing transcripts (for each docx)
                'loop through files to place bookmark at beginning of proceedings so system knows where transcript starts

                With oPreparedWD
                
                    If .ProtectionType <> wdNoProtection Then
                        .Unprotect password:="wrts0419"
                        Else
                    End If
                    
                    For Each wsStyle In ActiveDocument.Styles
                        
                        If wsStyle = "Heading 1" Or wsStyle = "Heading 2" Or wsStyle = "Heading 3" Or wsStyle = "AQC-Working" Or wsStyle = "ESSworn" Or wsStyle = "ESBYLawyer" Or wsStyle = "ESHeading" Then
                            
                            With .Application.Selection.Find
                            
                                .ClearFormatting
                                .Replacement.ClearFormatting
                                .Style = wsStyle
                                .Text = "*"
                                .Replacement.Text = ""
                               .Forward = True
                                .Wrap = wdFindStop
                                .Format = True
                                .MatchCase = False
                                .MatchWholeWord = False
                                .MatchWildcards = True
                                .MatchSoundsLike = False
                                .MatchAllWordForms = False
                                .Execute Replace:=wdReplaceAll
                                
                            End With
                            
                        Else
                        
                            Exit For
                            
                        End If
                        
                    Next wsStyle
                    
                '   Find/replace so that every sentence begins with "<s> " and ends with "</s> (wavfilename)"
                '       ALL SENTENCES CAPITALIZED but lowercase <s>, </s> and (wavfilename).
                '       One sentence per line
                
    
                    sTextToFind = "^13^t^t(*[!^13]{1,})\: {2,2}"
                    sReplacementText = "^13<s> "
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText, bMatchWildcards:=True)
                    sTextToFind = "^13^t{2,2}"
                    sReplacementText = "^13<s> "
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText, bMatchWildcards:=True)
                    
                    sTextToFind = "<s> ^p"
                    sReplacementText = "<s> "
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText, bMatchWildcards:=False)
                    
                
                    sTextToFind = "-- ^p"
                    sReplacementText = " </s> (wavfilename)^p"
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText)
                
                    sTextToFind = " --^p"
                    sReplacementText = " </s> (wavfilename)^p"
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText)
                
                
                    sTextToFind = "?  ^p"
                    sReplacementText = " </s> (wavfilename)^p"
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText)
                                    
                    sTextToFind = "? ^p"
                    sReplacementText = " </s> (wavfilename)^p"
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText)
                    
                    sTextToFind = "?^p"
                    sReplacementText = " </s> (wavfilename)^p"
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText)
                
                    sTextToFind = ".  ^p"
                    sReplacementText = " </s> (wavfilename)^p"
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText)
                
                    sTextToFind = ". ^p"
                    sReplacementText = " </s> (wavfilename)^p"
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText)
                    
                    sTextToFind = ".^p"
                    sReplacementText = " </s> (wavfilename)^p"
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText)
                
                    sTextToFind = "?  "
                    sReplacementText = " </s> (wavfilename)^p<s> "
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText)
                    
                    sTextToFind = ".  "
                    sReplacementText = " </s> (wavfilename)^p<s> "
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText)
                
                
                    sTextToFind = " -- "
                    sReplacementText = " "
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText)
                    
                    sTextToFind = "."
                    sReplacementText = ""
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText)
                
                    sTextToFind = ","
                    sReplacementText = ""
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText)
                    
                    sTextToFind = "'"
                    sReplacementText = ""
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText)
                
                    sTextToFind = "?"
                    sReplacementText = ""
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText)
                
                    sTextToFind = "^13\((*)\)^13"
                    sReplacementText = ""
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText, bWrap:="wdFindStop", bMatchCase:=False)
                
                    sTextToFind = "CERTIFICATE OF TRANSCRIBER"
                    sReplacementText = ""
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText, bWrap:="wdFindContinue", bFormat:=False)
                    
                    .Application.Selection.EndKey Unit:=wdStory, Extend:=wdExtend
                
                
                    sTextToFind = "^13\<*[!^13]\:  "
                    sReplacementText = "^13<s> "
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText, bForward:=False, bWrap:="wdFindContinue", _
                        bFormat:=False, bMatchWildcards:=True)
                        
                    Selection.WholeStory
                    Selection.Range.Case = wdLowerCase
                    
                    .Application.Selection.EndKey Unit:=wdStory, Extend:=wdExtend
                    .Application.Selection.EndKey Unit:=wdStory, Extend:=wdExtend
                    .Application.Selection.delete Unit:=wdCharacter, Count:=1
                    .Save
                    .SaveAs2 FileName:="S:\UnprocessedAudio\Prepared\" & sCourtDatesID & "\WorkingFiles\Transcript.txt", FileFormat:=wdFormatText
                    
                    sTextToFind = "</s> (wavfilename)"
                    sReplacementText = ""
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText, bMatchWildcards:=False, bFormat:=False)
                    
                    sTextToFind = "<s> "
                    sReplacementText = ""
                    Call pfSingleFindReplace(sTextToFind:=sTextToFind, sReplacementText:=sReplacementText, bMatchWildcards:=False, bForward:=True)
                    
                    .SaveAs2 FileName:=oSubfolder.Path & "\WorkingFiles\" & "Base-corpus.txt", FileFormat:=wdFormatText
                    .Close
                End With
            End If
        Else
        End If
        
                Call pfRunCopyTranscTextBAT
                    'Run batch CopyTranscriptTXT.bat (makes *.transcription & *.fileids file)

                Set oPreparedWD = Nothing
        'Else
        'End If
    Next oCurrentFile
Next oSubfolder

End Sub
Public Sub pfRunCopyTranscTextBAT()
'============================================================================
' Name        : pfRunCopyTranscTextBAT
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfRunCopyTranscTextBAT
' Description : runs batch file CopyTranscriptTXT
'============================================================================

  Dim PathCrnt As String

  PathCrnt = "S:\UnprocessedAudio\Unprepared"
  Call Shell(PathCrnt & "\CopyTranscriptTXT.bat " & PathCrnt)
End Sub

    
Public Sub pfSRTranscribe()
'============================================================================
' Name        : pfSRTranscribe
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfSRTranscribe
' Description : runs batch file SRTranscribe
'============================================================================

'runs batch file SRTranscribe
'HOW TO TRANSCRIBE FILES AFTER TRAINING
'Run the following commands in a command window with administrator:
    'cd /D S:\UnprocessedAudio\3 (or number you're using)
    'S:\pocketsphinx\bin\Release\Win32\pocketsphinx_continuous.exe -infile wavfilename.wav -hmm en-us-adapt -lm wavfilename.lm -dict wavfilename.dic >> full-output.txt
    'Check output in full-output.txt in S:\UnprocessedAudio\##

Dim PathCrnt As String

PathCrnt = "S:\UnprocessedAudio\Prepared"

Call Shell(PathCrnt & "\SRTranscribe.bat " & PathCrnt)

End Sub


Public Sub pfTrainAudio()
'============================================================================
' Name        : pfTrainAudio
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfTrainAudio
' Description : HOW TO TRAIN AUDIO / run batch file audiotrain.bat
'============================================================================

'---------------------------------------------------------
'
'HOW TO TRAIN AUDIO
'NOTE: "S:\training BATs\G-runWAlignforAccuracyReport.bat"
'
'Change "wavfilename" to your wav file name.
'Run the following commands in a command window with administrator:
'
'1. cd /D S:\UnprocessedAudio\3
'
'2. S:\sphinxtrain\bin\Release\Win32\sphinx_fe.exe -argfile en-us/feat.params -samprate 16000 -c wavfilename.fileids -di . -do . -ei wav -eo mfc -mswav yes -nfft 2048
'
'3. S:\sphinxtrain\bin\Release\Win32\pocketsphinx_mdef_convert.exe -text en-us/mdef en-us/mdef.txt
'
'4. S:\sphinxtrain\bin\Release\Win32\bw.exe -hmmdir en-us -moddeffn en-us/mdef.txt -ts2cbfn .ptm. -svspec 0-12/13-25/26-38 -feat 1s_c_d_dd -cmn current -agc none -dictfn wavfilename.dic -ctlfn wavfilename.fileids -lsnfn wavfilename.transcription -accumdir .
'
'#### S:\sphinxtrain\bin\Release\Win32\bw.exe -hmmdir en-us -moddeffn en-us/mdef.txt -ts2cbfn .ptm. -svspec 0-12/13-25/26-38 -feat 1s_c_d_dd -cmn current -agc none -dictfn cmudict-en-us.dict -ctlfn wavfilename.fileids -lsnfn wavfilename.transcription -accumdir .
'
'5. S:\sphinxtrain\bin\Release\Win32\mllr_solve.exe -meanfn en-us/means -varfn en-us/variances -outmllrfn mllr_matrix -accumdir .
'
'6. S:\sphinxtrain\bin\Release\Win32\map_adapt.exe -moddeffn en-us/mdef.txt -ts2cbfn .ptm. -meanfn en-us/means -varfn en-us/variances -mixwfn en-us/mixture_weights -tmatfn en-us/transition_matrices -accumdir . -mapmeanfn en-us-adapt/means -mapvarfn en-us-adapt/variances -mapmixwfn en-us-adapt/mixture_weights -maptmatfn en-us-adapt/transition_matrices
'
'7. S:\sphinxtrain\bin\Release\Win32\pocketsphinx_batch.exe -adcin yes -cepdir wav -cepext .wav -ctl ryan-weed-sample.fileids -lm wavfilename.lm -dict wavfilename.dic -hmm en-us-adapt -hyp wavfilename.hyp
'
'### S:\sphinxtrain\bin\Release\Win32\pocketsphinx_batch.exe -adcin yes -cepdir wav -cepext .wav -ctl wavfilename.fileids -lm wavfilename.lm -dict wavfilename.dic -hmm en-us-adapt -hyp wavfilename.hyp
'
'8. S:\sphinxtrain\scripts\decode\word_align.pl wavfilename.transcription wavfilename.hyp >> word_align_output.txt
'
'
'---------------------------------------------------------
'
Dim PathCrnt As String

PathCrnt = "S:\UnprocessedAudio\Prepared"
Call Shell(PathCrnt & "\audiotrain.bat " & PathCrnt)

End Sub
    

Public Sub pfRenameBaseFiles()
'============================================================================
' Name        : pfRenameBaseFiles
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfRenameBaseFiles
' Description : runs batch file FileRename / audioprep1.bat in Prepared folder
'============================================================================

Dim PathCrnt As String

PathCrnt = "S:\UnprocessedAudio\Prepared"
Call Shell(PathCrnt & "\FileRename.bat " & PathCrnt)
End Sub


Public Sub pfSplitAudio()
'============================================================================
' Name        : pfSplitAudio
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfSplitAudio
' Description : runs batch file splitaudio / audioprep1.bat in Prepared folder
'============================================================================
Dim PathCrnt As String

PathCrnt = "S:\UnprocessedAudio\Prepared"
Call Shell(PathCrnt & "\splitaudio.bat " & PathCrnt)
End Sub

Public Sub pfPrepareAudio()
'============================================================================
' Name        : pfPrepareAudio
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfPrepareAudio
' Description : runs batch file audioprep / audioprep1.bat in Prepared folder
'============================================================================

Dim PathCrnt As String

PathCrnt = "S:\UnprocessedAudio\Prepared"
Call Shell(PathCrnt & "\audioprep.bat " & PathCrnt)
End Sub


'---------------------------------------------------------
'
Public Sub pfTrainEngine()
'============================================================================
' Name        : pfTrainEngine
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfTrainEngine
' Description : HOW TO PREPARE FILES FOR TRAINING
'============================================================================


'HOW TO PREPARE FILES FOR TRAINING
'
'You may use other numbers in \\HUBCLOUD\evoingram\speech\UnprocessedAudio\ for samples of other working documents you may want to look at.
'
'PREPARING AUDIO:
'
'   Find a completed job from \\HUBCLOUD\evoingram\3Complete\.
'       I have done 453 and 3 dates from 446, 2/8/18, 2/12/18, & 3/9/18.  Everything else up for grabs.
'   ADAM: If you want to skip the rest of this section, please just copy the direct audio + transcript into a new folder in
'       \UnprocessedAudio\ and go to "PREPARING ALL OTHER FILES"
'   Copy its corresponding final transcript docx into \UnprocessedAudio\ OR this folder you just made within \UnprocessedAudio\.
'   Open audio player (FTR, liberty, javs, wmp, etc) & load corresponding audio from folder you selected.
'   Open Audacity.
'   Make sure project rate is 16000 ONLY and record moderately loudly, not quietly.
'   Record corresponding audio into audacity file (hit record in audacity and then hit play to play the file).
'   When complete, zoom out so you can see entire audio file in half-hour marks across the ruler on one screen.
'   Select first half-hour.
'   Go to file --> export-selected audio -->save as short name with "-##of##" at end 16-bit pcm something.
'       See other folders for examples.
'   Save in \UnprocessedAudio\ directory or next number folder in that directory.
'   Repeat for every half-hour until you are all the way through the audio.
'
'PREPARING ALL OTHER FILES:
'   Create a folder that is one higher than the highest number in \UnprocessedAudio\
'   OR rename one you just put audio+transcript in to one number higher than the highest number in \UnprocessedAudio\
'   Move your audio and transcript from \UnprocessedAudio\ into that folder.
'   Create the following 'new text documents': (make * the same as your wav name without the ##of## part)
'       *.fileids
'       *.transcription
'       *.dic
'       *.lm
'       *.docx
'       *-corpus.txt
'   Open final transcript in number folder in \\HUBCLOUD\evoingram\speech\UnprocessedAudio\
'   Copy everything from the SECOND line of the transcript body, so everything below the "CITY, STATE, DATE, TIME"
'       line, down to and including one line BEFORE (Hearing/Proceedings concluded at time.) at the end.
'   Paste as text only into wavfilename.docx.
'   Close final transcript, don't save.  You're done with this file.
'   Save wavfilename.docx, but do not close it.  This is the one we are mainly working with.
'   Find/replace so that every sentence begins with "<s> " and ends with "</s> (wavfilename)"
'       ALL SENTENCES CAPITALIZED but lowercase <s>, </s> and (wavfilename).
'       One sentence per line
'   I suggest replacements in the following order:
'       ++ Turn on Wildcards ++
'       " {8,8}(*{1,})\: {2,2}" to "<s> "
'       "^t{2,2}(*{1,})\: {2,2}" to "<s> "
'       -- Turn off Wildcads --
'       "^p {8,8}" to nothing
'       " -- ^p" to " </s> (wavfilename)^p"
'       " --^p" to " </s> (wavfilename)^p"
'       "?  ^p" to " </s> (wavfilename)^p"
'       "? ^p" to " </s> (wavfilename)^p"
'       "?p" to " </s> (wavfilename)^p"
'       ".  ^p" to " </s> (wavfilename)^p"
'       ". ^p" to " </s> (wavfilename)^p"
'       ".^p" to " </s> (wavfilename)^p"
'       "?  " to " </s> (wavfilename)^p<s> "
'       ".  " to " </s> (wavfilename)^p<s> "
'       " -- " to " "
'       "." "," "'" "?" to nothing
'   Delete all lines with:
'       Headings
'       SHORT PHRASES IN ALL CAPS LIKE 'COURT'S RULING' 'ARGUMENT BY THE...' 'DIRECT EXAMINATION' etc
'       Lines with parentheses'd comments
'       Any extra blank lines or anything that doesn't fit the <s> text here </s> (wavfilename) format
'   Select all, lowercase all, and copy.  Do not close.
'   Open wavfilename-corpus.txt file and paste the lowercase stuff in there.
'   Delete "<s> " and "</s> (wavfilename)" in corpus file (find/replace with nothing)
'   Save, close, and go back to previous word document you were making all the replacements on.
'   Select all, capitalize all.
'   Find/replace, match case, <S>, </S> and (WAVFILENAME) to make lowercase <s>, </s> and (wavfilename).
'   Select all and copy.
'   Open wavfilename.transcription file and paste in there.  Save.
'   Copy again and open Visual studio or other programming software that shows you lines
'       Visual basic component in office also does this.
'   Paste your transcription file in there.
'   Note number of lines.
'   Save and close.
'   Open wavfilename.fileids from your folder.
'   Put wavfilename on each line for the EXACT SAME number of lines in your wavfilename.transcription.
'       i copy it enough times to where i think it's close and then verify line count in programming software.
'   Save and close.
'   Go to http://www.speech.cs.cmu.edu/tools/lmtool-new.html
'   Upload wavfilename-corpus.txt file.  When it completes its task:
'   "save link as" *.dic as wavfilename.dic
'   "save link as" *.lm as wavfilename.lm
'   Now, open up the word document, wavfilename.docx, again and load up each wav file in a player.
'   Note first and last line of words of the wav file.
'   Search for them in the wavfilename.docx so that you select everything that's on the audio you just played.
'   In each selected line, ensure the 'wavfilename' matches each of the wav files it belongs in.
'       You want to make sure the wavfilename includes the "-##of##" part.
'       So each selection, you will find/replace "(wavfilename)" with your new file name, "(wavfilename-##of##)"
'       Repeat matching audio to lines and find/replace within selection for each half-hour audio file you have.
'   When complete, save, select all, and copy.
'   Then open wavfilename.transcription.
'   Select all and paste.  Make sure there are no blank lines at the bottom.
'   Save and close wavfilename.transcription.
'   Go back to wavfilename.docx one more time.
'   Find and replace, check "use wildcards", find "\)(*{1,})\(", and replace with ")^p("
'   That should leave only the wavfilename a bunch of times in a list.
'   Select all and copy.
'   Open wavfilename.fileids.
'   Select all and paste.
'   Save and close wavfilename.fileids.
'   Without saving, close wavfilename.docx.
'

End Sub

Public Sub pfAddSubfolder()
'============================================================================
' Name        : pfAddSubfolder
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfAddSubfolder
' Description : add concatenatedaudio folder to /UnprocessedAudio/####
'============================================================================

Dim oFolderObject As Scripting.FileSystemObject
Dim oRootFolder As Object, oSubfolder As Object
Dim sUnprocessedAudioPath As String, sConcatenatedAudioPath As String

Set oFolderObject = CreateObject("Scripting.FileSystemObject")
sUnprocessedAudioPath = "S:\UnprocessedAudio\"
Set oRootFolder = oFolderObject.GetFolder(sUnprocessedAudioPath)

For Each oSubfolder In oRootFolder.SubFolders

    Debug.Print oSubfolder.Path
    
    sConcatenatedAudioPath = oSubfolder.Path & "\ConcatenatedAudio"
    
    If Not oFolderObject.FolderExists(sConcatenatedAudioPath) Then
        MkDir (sConcatenatedAudioPath)
    End If
    
Next oSubfolder

End Sub

Public Sub pfDownloadFile(sURL As String, sSaveAs As String)
'============================================================================
' Name        : pfDownloadFile
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfDownloadFile(sURL, sSaveAs)
' Description : downloads provided file
'============================================================================

Dim x As Long

x = URLDownloadToFile(0, sURL, sSaveAs, 0, 0)

If x = 0 Then
    Debug.Print "Download has completed!"
Else
    Debug.Print "Error!"
End If

End Sub

Public Sub pfCorpusUpload()
'============================================================================
' Name        : pfCorpusUpload
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfCorpusUpload
' Description : uploads corpus to lmtool at Sphinx to get compatible LM & DIC files back via download
'============================================================================

Dim sDestinationURL As String, sFieldName As String, sMainURL As String
Dim oFolderObject As Scripting.FileSystemObject
Dim sPreparedPath As String

sMainURL = "http://www.speech.cs.cmu.edu/tools/lmtool-new.html"
sDestinationURL = "http://www.speech.cs.cmu.edu/cgi-bin/tools/lmtool/run"
sFieldName = "corpus"
sPreparedPath = "S:\UnprocessedAudio\Prepared\"

Set oFolderObject = CreateObject("Scripting.FileSystemObject")
pfDoFolder oFolderObject.GetFolder(sPreparedPath)

End Sub

Public Sub pfUploadFile(sDestinationURL As String, sCorpusPath As String, _
  Optional ByVal sFieldName As String = "corpus")
'============================================================================
' Name        : pfUploadFile
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfUploadFile(sDestinationURL, sCorpusPath, sFieldName)
                'Optional sFieldName = "corpus"
' Description : uploads corpus file to Sphinx site to get LM/DIC files back
'               upload file using input type=file
'============================================================================

Dim sFormData As String, d As String

'Boundary of fields.
'Be sure this string is Not In the source file
Const Boundary As String = "---------------------------0123456789012"

'Get source file As a string.
sFormData = pfGetFile(sCorpusPath)

'Build source form with file contents
d = "--" + Boundary + vbCrLf
d = d + "Content-Disposition: form-data; name=""" + sFieldName + """;"
d = d + " sCorpusPath=""" + sCorpusPath + """" + vbCrLf
d = d + "Content-Type: application/upload" + vbCrLf + vbCrLf
d = d + sFormData
d = d + vbCrLf + "--" + Boundary + "--" + vbCrLf

'Post the data To the destination URL
pfIEPostStringRequest sDestinationURL, d, Boundary
End Sub

Public Sub pfIEPostStringRequest(sURL As String, sFormData As String, sBoundary As String)
'============================================================================
' Name        : pfIEPostStringRequest
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfIEPostStringRequest(sURL, sFormData, sBoundary)
' Description : sends URL encoded form data To the URL using IE
'============================================================================

Dim sFileNameURL As String, sFilePathDIC As String
Dim sFilePathLM As String, sFileURL As String, sCourtDatesID As String, sActionURL As String, sMainURL As String
Dim sResponseText As String, sResponseIDNumber As String, sResponseURL As String, sDestinationURL As String, sCorpusPath As String
Dim sDownloadID As String, sDIC As String, sLM As String, sOnlineDIC As String, sOnlineLM As String
Dim bFormData() As Byte
Dim oWebBrowser As Object, oWebBrowser01 As Object, oStream As Object
Dim m_isRedirected As Boolean

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
sMainURL = "http://www.speech.cs.cmu.edu/tools/lmtool-new.html"
sDestinationURL = "http://www.speech.cs.cmu.edu/cgi-bin/tools/lmtool/run"
sCorpusPath = "S:\UnprocessedAudio\Prepared\" & sCourtDatesID & "\WorkingFiles\Base-corpus.txt"
sFilePathDIC = "S:\UnprocessedAudio\Prepared\" & sCourtDatesID & "\WorkingFiles\Base.dic"
sFilePathLM = "S:\UnprocessedAudio\Prepared\" & sCourtDatesID & "\WorkingFiles\Base.lm"

Set oWebBrowser = CreateObject("InternetExplorer.Application")
oWebBrowser.Visible = True

'Send the form data To URL As POST request
ReDim bFormData(Len(sFormData) - 1)
bFormData = StrConv(sFormData, vbFromUnicode)

oWebBrowser.Navigate sURL, , , bFormData, _
  "Content-Type: multipart/form-data; boundary=" + sBoundary + vbCrLf

Do While oWebBrowser.Busy
    'Sleep 100
    DoEvents
Loop

sResponseText = oWebBrowser.Document.Body.innerHTML
'Debug.Print "Response Text:  " & sResponseText
Debug.Print "-----------------------"
sResponseIDNumber = Right(sResponseText, 18)
sResponseIDNumber = Left(sResponseIDNumber, 16)
Debug.Print "Response URL:  " & sResponseIDNumber
sResponseURL = "http://www.speech.cs.cmu.edu/tools/product/" & sResponseIDNumber
Debug.Print "Response URL:  " & sResponseURL
Debug.Print "-----------------------"
Set oWebBrowser01 = CreateObject("InternetExplorer.Application")
oWebBrowser01.Navigate sResponseURL
Do While oWebBrowser01.Busy
    '    Sleep 100
    DoEvents
Loop

sResponseText = oWebBrowser01.Document.Body.innerHTML
sResponseIDNumber = Right(sResponseText, 152)

Debug.Print "-----------------------"
sDownloadID = Left(sResponseIDNumber, 4)
Debug.Print "Download ID Number:  " & sDownloadID
Debug.Print "-----------------------"

sDIC = sDownloadID & ".dic"
sLM = sDownloadID & ".lm"
sOnlineDIC = sResponseURL & Chr(47) & sDIC
sOnlineLM = sResponseURL & Chr(47) & sLM
Debug.Print "DIC link:  " & sOnlineDIC
Debug.Print "LM link:  " & sOnlineLM
Debug.Print "-----------------------"

oWebBrowser.Quit
oWebBrowser01.Quit

Call pfDownloadFile(sOnlineDIC, sFilePathDIC)
Call pfDownloadFile(sOnlineLM, sFilePathLM)

Debug.Print "DIC saved to:  " & sFilePathDIC
Debug.Print "LM saved to:  " & sFilePathLM
Debug.Print "-----------------------"

End Sub


Public Function pfGetFile(sFileName As String) As String
'============================================================================
' Name        : pfGetFile
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfGetFile(sFileName)
' Description : read binary file As a string value
'============================================================================

Dim bFileContents() As Byte, iFileNumber As Integer

ReDim bFileContents(FileLen(sFileName) - 1)
iFileNumber = FreeFile

Open sFileName For Binary As iFileNumber
Get iFileNumber, , bFileContents

Close iFileNumber

pfGetFile = StrConv(bFileContents, vbUnicode)

End Function
'******************* upload - end

Public Sub pfDoFolder(Folder As Variant)
'============================================================================
' Name        : pfDoFolder
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfDoFolder(Folder)
' Description : cycles through subfolders of /Prepared/####/ and makes find/replaces listed below
'============================================================================

Dim oFolderObject As Object, oRootFolder As Object, oCurrentFile As Object, oSubfolder As Object
Dim oWordApp As New Word.Application, oTranscriptionWD As New Word.Document
Dim sCorpusPath As String, sDICPath As String, sLMPath As String
Dim sDestinationURL As String, sFieldName As String, sSubfolder As String
Dim oCurrentFileString As String, sFileExtension As String, sFolderName As String
 

sDestinationURL = "http://www.speech.cs.cmu.edu/cgi-bin/tools/lmtool/run"
sFieldName = "corpus"

For Each oSubfolder In Folder.SubFolders

     sCourtDatesID = oSubfolder.ParentFolder.Name
     pfDoFolder oSubfolder
     'sCourtDatesID = oSubfolder.Name
     
     Debug.Print "Now processing Job No. " & sCourtDatesID; "..."
     
     sSubfolder = oSubfolder.Name
     
     Debug.Print sSubfolder
     
     If InStr(1, sSubfolder, "WorkingFiles") Then
        sCorpusPath = "S:\UnprocessedAudio\Prepared\" & sCourtDatesID & "\WorkingFiles\Base-corpus.txt"
        sDICPath = "S:\UnprocessedAudio\Prepared\" & sCourtDatesID & "\WorkingFiles\Base.dic"
        sLMPath = "S:\UnprocessedAudio\Prepared\" & sCourtDatesID & "\WorkingFiles\Base.lm"
    
        Set oWordApp = CreateObject("Word.Application")
        Set oTranscriptionWD = oWordApp.Documents.Open("S:\UnprocessedAudio\Prepared\" & sCourtDatesID & "\WorkingFiles\base.transcription")
        
        With oTranscriptionWD
            With .Application.Selection.Find
                .Text = "$"
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
            
            With .Application.Selection.Find
                .Text = Chr(34)
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
        
            .Application.Selection.WholeStory
            .Application.Selection.Range.Case = wdLowerCase
            
        End With
        
        oTranscriptionWD.Save
        oTranscriptionWD.Close
        oWordApp.Quit
        
        Call pfUploadFile(sDestinationURL, sCorpusPath, sFieldName)
        Debug.Print "Done processing Job No. " & sCourtDatesID; "..."
 
     End If
     
 Next
 
End Sub






