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


'@Ignore EmptyMethod
'@Ignore ProcedureNotUsed
Private Sub emptyFunction()
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID
    Dim oWordDoc As New Word.Document
    Dim oWordApp As New Word.Application
    Dim oWordDoc1 As New Word.Document
        
    'To Fix Bookmarks:
    'Final Transcript Outline:
        'Cover All
        'General Index
            'Date
                'General Events
        'Witness Index
            'normal
        'Exhibit Index
            'normal
        'Cover Date
            'Transcript Body
        'Cover Date
            'Transcript Body
        'Certificate
        'Table of Authorities
        
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
End Sub

