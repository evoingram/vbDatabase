Attribute VB_Name = "Stage1"
'@Folder("Database.Production.Modules")
Option Compare Database
Option Explicit
'============================================================================
'class module cmStage1:
'variables:
'   NONE

'functions:
'fAssignPS:                                 Description:  prompts to assign file in ProjectSend
'                                       Arguments:    NONE
'pfEnterNewJob:                             Description:  import job info to db from xlsm file
'                                       Arguments:    NONE
'fCheckTempCustomersCustomers:              Description:  retrieve info from TempCustomers/Customers
'                                       Arguments:    NONE
'fCheckTempCasesCases:                      Description:  retrieve info from TempCases/Cases
'                                       Arguments:    NONE
'fInsertCalculatedFieldintoTempCourtDates:  Description:  insert several calculated fields into tempcourtdates
'                                       Arguments:    NONE
'fAudioPlayPromptTyping:                    Description:  prompt to play audio in /Audio/folder
'                                       Arguments:    NONE
'fProcessAudioParent:                       Description:  process audio in express scribe
'                                       Arguments:    NONE
'fPlayAudioParent:                          Description:  play audio as appropriate
'                                       Arguments:    NONE
'fPlayAudioFolder:                          Description:  plays audio folder
'                                       Arguments:    HostFolder
'fProcessAudioFolder:                       Description:  process audio in /Audio/ folder
'                                       Arguments:    HostFolder
'pfPriceQuoteEmail:                         Description:  generates price quote and sends via e-mail
'                                       Arguments:    NONE
'pfStage1Ppwk:                              Description:  completes all stage 1 tasks
'                                       Arguments:    NONE
'fWunderlistAddNewJob:                      Description:  adds new job task list to wunderlist w/ due dates
'                                       Arguments:    NONE
'autointake:                                Description:  process new job email when received
'                                       Arguments:    NONE
'NewOLEntry:                                Description:  checks outlook folder for new job email
'                                       Arguments:    NONE
'ResetDisplay:                              Description:  part of scrolling marquee notification
'                                       Arguments:    NONE
'ScrollingMarquee:                          Description:  scrolling marquee notification for new job
'                                       Arguments:    NONE
'MinimizeNavigationPane:                    Description:  part of scrolling marquee notification
'                                       Arguments:    NONE
        
'============================================================================

Public Sub fAssignPS()
    '============================================================================
    ' Name        : fAssignPS
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fAssignPS
    ' Description : prompts to assign file in ProjectSend
    '============================================================================
    Dim sQuestion As String
    Dim sAnswer As String
    Dim sBrowserPath As String

    sBrowserPath = """C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe"""
    sQuestion = "Do you want to assign this file in ProjectSend?"

    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    If sAnswer = vbNo Then                       'Code for No
        MsgBox "Files in ProjectSend will not be assigned to the client."
    Else                                         'Code for yes, opens PS in chrome
        Shell (sBrowserPath & " -url https://www.aquoco.co/ProjectSend/index.php")
    End If

End Sub

Public Sub pfEnterNewJob()
    '============================================================================
    ' Name        : pfEnterNewJob
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfEnterNewJob
    ' Description : import job info to db from xlsm file
    '============================================================================

    Dim sTurnaroundTimesCD As String
    Dim sInvoiceNumber As String
    Dim sNewCourtDatesRowSQL As String
    Dim sOrderingID As String
    Dim sCurrentJobSQL As String
    Dim sTempJobSQL As String
    Dim sStatusesEntrySQL As String
    Dim sCasesID As String
    Dim sCurrentTempApp As String
    Dim sTurnaround As String
    Dim sAnswer As String
    Dim sQuestion As String
    Dim sAppNumber As String
    Dim sTempCustomersSQL As String
    Dim sFactoring As String
    Dim sFiled As String
    Dim sBrandingTheme As String
    Dim sUnitPrice As String
    Dim sIRC As String
    Dim sAccountCode As String
    
    Dim oExcelApp As Object
    
    Dim x As Long
    
    Dim dInvoiceDate As Date
    
    Dim rstTempJob As DAO.Recordset
    Dim rstCurrentJob As DAO.Recordset
    Dim rstCurrentCasesID As DAO.Recordset
    Dim rstTempCourtDates As DAO.Recordset
    Dim rstTempCases As DAO.Recordset
    Dim rstTempCustomers As DAO.Recordset
    Dim rstCurrentStatusesEntry As DAO.Recordset
    Dim rstMaxCasesID As DAO.Recordset
    Dim rstCourtDatesID As DAO.Recordset
    
    Dim oExcelWB As Excel.Workbook
    Dim oExcelMacroWB As Excel.Workbook

    Dim cJob As New Job

    Set oExcelApp = CreateObject("Excel.Application")

    Set oExcelMacroWB = oExcelApp.Application.Workbooks.Open(cJob.DocPath.OrderFormCustomersXLSM)
    oExcelMacroWB.Application.DisplayAlerts = False
    oExcelMacroWB.Application.Visible = False
    oExcelMacroWB.SaveAs Replace(cJob.DocPath.OrderFormCustomersXLSM, ".xlsm", ".xlsx"), FileFormat:=xlWorkbookDefault
    oExcelMacroWB.Close True
    Set oExcelMacroWB = Nothing

    Set oExcelWB = oExcelApp.Application.Workbooks.Open(FileName:=cJob.DocPath.OrderFormCustomersXLSX, Local:=True)
    oExcelWB.Application.DisplayAlerts = False
    oExcelWB.Application.Visible = False
    oExcelWB.SaveAs Replace(cJob.DocPath.OrderFormCustomersXLSX, ".xlsx", ".csv"), FileFormat:=6

    oExcelWB.Close True
    Set oExcelWB = Nothing
    
    On Error Resume Next:   On Error GoTo 0
    CurrentDb.TableDefs.Refresh

    Set oExcelMacroWB = oExcelApp.Application.Workbooks.Open(FileName:=cJob.DocPath.OrderFormXLSM, Local:=True)
    oExcelMacroWB.Application.DisplayAlerts = False
    oExcelMacroWB.Application.Visible = False
    oExcelMacroWB.SaveAs Replace(cJob.DocPath.OrderFormXLSM, ".xlsm", ".xlsx"), FileFormat:=xlWorkbookDefault
    oExcelMacroWB.Close True
    Set oExcelMacroWB = Nothing

    Set oExcelWB = oExcelApp.Application.Workbooks.Open(FileName:=cJob.DocPath.OrderFormXLSX, Local:=True)
    oExcelWB.Application.DisplayAlerts = False
    oExcelWB.Application.Visible = False
    oExcelWB.SaveAs Replace(cJob.DocPath.OrderFormXLSX, ".xlsx", ".csv"), FileFormat:=6
    oExcelWB.Close True
    Set oExcelWB = Nothing

 
    On Error Resume Next:   On Error GoTo 0
    CurrentDb.TableDefs.Refresh

    DoCmd.TransferText TransferType:=acImportDelim, TableName:="TempCourtDates", _
                       FileName:=cJob.DocPath.OrderFormCSV, HasFieldNames:=True
    CurrentDb.TableDefs.Refresh

    On Error Resume Next:   On Error GoTo 0
    CurrentDb.TableDefs.Refresh

    DoCmd.TransferText TransferType:=acImportDelim, TableName:="TempCustomers", _
                       FileName:=cJob.DocPath.OrderFormCustomersCSV, HasFieldNames:=True

    On Error Resume Next:   On Error GoTo 0
    CurrentDb.TableDefs.Refresh

    Set rstTempCourtDates = CurrentDb.OpenRecordset("TempCourtDates")
    rstTempCourtDates.MoveFirst
    sJurisdiction = rstTempCourtDates.Fields("JurisDiction").Value
    sAudioLength = rstTempCourtDates.Fields("AudioLength").Value
    sTurnaround = rstTempCourtDates.Fields("TurnaroundTimesCD").Value
    rstTempCourtDates.Close

    Select Case sTurnaround
    Case "45"
        sUnitPrice = 64
        sIRC = 96
    Case "30"
        sUnitPrice = 58
        sIRC = 78
    Case "14"
        sUnitPrice = 59
        sIRC = 7
    Case "7"
        sUnitPrice = 60
        sIRC = 8
    Case "3"
        sUnitPrice = 42
        sIRC = 90
    Case "1"
        sUnitPrice = 61
        sIRC = 14
    End Select


    Select Case sJurisdiction
    Case "eScribers"
        sUnitPrice = 33
        sIRC = 95
    Case "FDA", "Food and Drug Administration"
        sUnitPrice = 37
        sIRC = 41
    Case "Weber"
        sUnitPrice = 36
        sIRC = 65
    Case "J&J"
        sUnitPrice = 36
        sIRC = 43
    Case "Non-Court", "NonCourt", "noncourt", "non-court"
        sUnitPrice = 49
        sIRC = 86
    Case "KCI", "King County Indigent", "KCI King County Superior Court", _
         "KCI Snohomish County Superior Court"
        sUnitPrice = 40
        sIRC = 56
    End Select

    'TODO: insert filed/factoring boxes in xlsm/csv
    sFiled = InputBox("Are we filing this, yes or no?")
    sFactoring = InputBox("Are we factoring this, yes or no?")

    Select Case sFiled
    
    Case "yes", "Yes", "YES", "Yeah", "yeah", "yea", "YEAH", "YEA", "Y", "y", _
         "yep", "Yep", "YEP", "YA", "Ya", "ya"   'filed
                
        Select Case sFactoring
        Case "yes", "Yes", "YES", "Yeah", "yeah", "yea", "YEAH", "YEA", _
             "Y", "y", "yep", "Yep", "YEP", "YA", "Ya", "ya" 'no deposit
            sFactoring = True
            sBrandingTheme = 6
        Case "no", "NO", "No", "nah", "Nah", "NAH", "nope", "Nope", _
             "NOPE", "N", "n"                    'with deposit
            sFactoring = False
            sBrandingTheme = 8
        End Select
        
    Case "no", "NO", "No", "nah", "Nah", "NAH", "nope", "NOPE", "Nope", "N", "n" 'not filed
        
        Select Case sFactoring
        Case "yes", "Yes", "YES", "Yeah", "yeah", "yea", "YEAH", "YEA", _
             "Y", "y", "yep", "Yep", "YEP", "YA", "Ya", "ya" 'no deposit
            sFactoring = True
            Select Case sJurisdiction
            Case "J&J", "J&J Court Transcribers", "J&J Court"
                sBrandingTheme = 10
            Case "eScribers", "AVT", "AVTranz", "eScribers NH", _
                 "eScribers Bankruptcy"
                sBrandingTheme = 11
            Case "FDA", "Food and Drug Administration", "Weber"
                sBrandingTheme = 12
            Case "NonCourt", "Non-Court", "Noncourt", "NONCOURT"
                sBrandingTheme = 1
            Case Else
                sBrandingTheme = 7
            End Select
        Case "no", "NO", "No", "nah", "Nah", "NAH", "nope", "Nope", _
             "NOPE", "n", "N"                    'with deposit
            sFactoring = False
            Select Case sJurisdiction
            Case "NonCourt", "Non-Court", "Noncourt", "NONCOURT"
                sBrandingTheme = 2
            Case Else
                sBrandingTheme = 9
            End Select
        End Select
    End Select

    'place info into tempcourtdates and tempcases
    Set rstTempCourtDates = CurrentDb.OpenRecordset("TempCourtDates")
    rstTempCourtDates.MoveFirst
    sTurnaround = rstTempCourtDates.Fields("TurnaroundTimesCD").Value
    rstTempCourtDates.Close
    dInvoiceDate = (Date + sTurnaround) - 2
    dDueDate = (Date + sTurnaround) - 2
    sAccountCode = 400
    CurrentDb.Execute "UPDATE TempCourtDates SET [InvoiceDate] = " & dInvoiceDate & ", [DueDate] = " & dDueDate & ", [AccountCode] = " & sAccountCode & _
               ", [UnitPrice] = " & sUnitPrice & ", [InventoryRateCode] = " & sIRC & ", [BrandingTheme] = " & sBrandingTheme & _
               " WHERE [ID] = " & sCourtDatesID & ";"

    'rstTempCourtDates.Edit
    'rstTempCourtDates.Fields("InvoiceDate").Value = dInvoiceDate
    'rstTempCourtDates.Fields("DueDate").Value = dDueDate
    'rstTempCourtDates.Fields("AccountCode").Value = sAccountCode
    'rstTempCourtDates.Fields("UnitPrice").Value = sUnitPrice
    'rstTempCourtDates.Fields("InventoryRateCode").Value = sIRC
    'rstTempCourtDates.Fields("BrandingTheme").Value = sBrandingTheme
    'rstTempCourtDates.Update
    'rstTempCourtDates.Close
    'Set rstTempCourtDates = CurrentDb.OpenRecordset("TempCourtDates")
    'Set rstTempCases = CurrentDb.OpenRecordset("TempCases")

    'rstTempCases.AddNew
    'rstTempCases.Fields("HearingTitle").Value = rstTempCourtDates.Fields("HearingTitle").Value
    'rstTempCases.Fields("Party1").Value = rstTempCourtDates.Fields("Party1").Value
    'rstTempCases.Fields("Party1Name").Value = rstTempCourtDates.Fields("Party1Name").Value
    'rstTempCases.Fields("Party2").Value = rstTempCourtDates.Fields("Party2").Value
    'rstTempCases.Fields("Party2Name").Value = rstTempCourtDates.Fields("Party2Name").Value
    'rstTempCases.Fields("CaseNumber1").Value = rstTempCourtDates.Fields("CaseNumber1").Value
    'rstTempCases.Fields("CaseNumber2").Value = rstTempCourtDates.Fields("CaseNumber2").Value
    'rstTempCases.Fields("Jurisdiction").Value = rstTempCourtDates.Fields("Jurisdiction").Value
    'rstTempCases.Fields("Judge").Value = rstTempCourtDates.Fields("Judge").Value
    'rstTempCases.Fields("JudgeTitle").Value = rstTempCourtDates.Fields("JudgeTitle").Value
    'rstTempCases.Fields("Notes").Value = rstTempCourtDates.Fields("Notes").Value
    'rstTempCases.Update
    'rstTempCases.Close
    'rstTempCourtDates.Close
    sNewCourtDatesRowSQL = "INSERT INTO TempCases (HearingTitle, Party1, Party1Name, Party2, Party2Name, CaseNumber1, CaseNumber2, " & _
                           "Jurisdiction, Judge, JudgeTitle, Notes) SELECT HearingTitle, Party1, Party1Name, Party2, Party2Name, CaseNumber1, CaseNumber2, " & _
                           "Jurisdiction, Judge, JudgeTitle, Notes FROM [TempCourtDates];"
    CurrentDb.Execute (sNewCourtDatesRowSQL)

    'delete blank lines
    CurrentDb.Execute "DELETE FROM TempCustomers WHERE [Company] = " & Chr(34) & Chr(34) & ";"
    CurrentDb.Execute "DELETE FROM TempCourtDates WHERE [AudioLength] IS NULL;"
    CurrentDb.Execute "DELETE FROM TempCases WHERE [Party1] = " & Chr(34) & Chr(34) & ";"


    'Perform the import
    sNewCourtDatesRowSQL = "INSERT INTO CourtDates (HearingDate, HearingStartTime, HearingEndTime, AudioLength, Location, TurnaroundTimesCD, " & _
                           "InvoiceNo, DueDate, UnitPrice, InvoiceDate, InventoryRateCode, AccountCode, BrandingTheme) SELECT HearingDate, HearingStartTime, " & _
                           "HearingEndTime, AudioLength, Location, TurnaroundTimesCD, InvoiceNo, DueDate, UnitPrice, InvoiceDate, InventoryRateCode, AccountCode, " & _
                           "BrandingTheme FROM [TempCourtDates];"
    CurrentDb.Execute (sNewCourtDatesRowSQL)

    ' store courtdatesID
    Set rstCourtDatesID = CurrentDb.OpenRecordset("SELECT MAX(ID) as IDNo FROM CourtDates")
    sCourtDatesID = rstCourtDatesID.Fields("IDNo").Value
    rstCourtDatesID.Close
    sCourtDatesID = str(CurrentDb.OpenRecordset("SELECT MAX(ID) FROM CourtDates"))
    [Forms]![NewMainMenu]![ProcessJobSubformNMM].[Form]![JobNumberField].Value = sCourtDatesID

    'TODO: What is going on here?
    Call fCheckTempCustomersCustomers
    Call fCheckTempCasesCases

    sTempJobSQL = "SELECT * FROM TempCustomers;"
    Set rstTempJob = CurrentDb.OpenRecordset(sTempJobSQL)
    
    sCurrentJobSQL = "SELECT * FROM CourtDates WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
    Set rstCurrentJob = CurrentDb.OpenRecordset(sCurrentJobSQL)

    rstTempJob.MoveFirst
    sOrderingID = rstTempJob.Fields("AppID").Value

    If IsNull(rstCurrentJob!OrderingID) Then
        CurrentDb.Execute "UPDATE CourtDates SET OrderingID = " & sOrderingID & " WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
    End If
    rstTempJob.Close
    rstCurrentJob.Close
    Set rstTempJob = Nothing
    Set rstCurrentJob = Nothing

    'TODO: not sure what
    Call fGenerateInvoiceNumber
    Call fInsertCalculatedFieldintoTempCourtDates

    'import casesID & CourtdatesID into tempcourtdates
    sCurrentJobSQL = "SELECT * FROM CourtDates WHERE ID = " & sCourtDatesID & ";"
    sTempJobSQL = "SELECT * FROM TempCourtDates;"
    sStatusesEntrySQL = "SELECT * FROM Statuses WHERE [CourtDatesID] = " & sCourtDatesID & ";"
    CurrentDb.Execute "INSERT INTO Statuses (CourtDatesID) SELECT CourtDatesID FROM TempCourtDates;"

    'Set rstStatuses = currentdb.OpenRecordset("Statuses")
    'rstStatuses.AddNew
    'rstStatuses.Fields("CourtDatesID").Value = sCourtDatesID
    'rstStatuses.Update
    'rstStatuses.Close
    'Set rstStatuses = Nothing
    Set rstTempJob = CurrentDb.OpenRecordset(sTempJobSQL)
    Set rstCurrentJob = CurrentDb.OpenRecordset(sCurrentJobSQL)
    Set rstCurrentStatusesEntry = CurrentDb.OpenRecordset(sStatusesEntrySQL)
    rstCurrentJob.MoveFirst
    Do Until rstCurrentJob.EOF
        Set rstTempJob = CurrentDb.OpenRecordset(sTempJobSQL)
        sTurnaroundTimesCD = rstTempJob.Fields("TurnaroundTimesCD")
        sInvoiceNumber = rstTempJob.Fields("InvoiceNo")
        Set rstMaxCasesID = CurrentDb.OpenRecordset("SELECT MAX(ID) FROM Cases;")
        sCasesID = rstMaxCasesID.Fields(0).Value
        rstMaxCasesID.Close
    
        CurrentDb.Execute "UPDATE TempCourtDates SET [CasesID] = " & sCasesID & " WHERE [CourtDatesID] = " & sCourtDatesID & ";"
        'rstTempJob.Edit
        'rstTempJob.Fields("CasesID") = sCasesID
        'rstTempJob.Update
        'sCasesID = rstTempJob.Fields("CasesID")
    
        CurrentDb.Execute "UPDATE TempCourtDates SET [CourtDatesID] = " & sCourtDatesID & " WHERE [InvoiceNo] = " & sInvoiceNumber & ";"
        '"SELECT * FROM TempCourtDates WHERE [InvoiceNo]=" & sInvoiceNumber & ";"
        'Set rstTempCDs = CurrentDb.OpenRecordset("TempCourtDates")
        'rstTempCDs.Edit
        'rstTempCDs.Fields("CourtDatesID").Value = sCourtDatesID
        'rstTempCDs.Update
        'rstTempCDs.Close
        'Set rstTempCDs = Nothing
        'CurrentDb.Execute "UPDATE TempCustomers SET [CourtDatesID] = " & sCourtDatesID & " WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
    
        CurrentDb.Execute "UPDATE TempCustomers SET [CourtDatesID] = " & sCourtDatesID & ";"
        'Set rstTempCDs = CurrentDb.OpenRecordset("TempCustomers")
        'rstTempCDs.Edit
        'rstTempCDs.Fields("CourtDatesID").Value = sCourtDatesID
        'rstTempCDs.Update
        'rstTempCDs.Close
        'Set rstTempCDs = Nothing
    
        CurrentDb.Execute "UPDATE CourtDates SET [CasesID] = " & sCasesID & " WHERE [ID] = " & sCourtDatesID & ";"
        'Set rstTempCDs = CurrentDb.OpenRecordset("SELECT * FROM CourtDates WHERE [ID] = " & sCourtDatesID & ";")
        'rstTempCDs.Edit
        'If sCasesID <> "" Then rstTempCDs.Fields("CasesID").Value = sCasesID
        'rstTempCDs.Update
        'rstTempCDs.Close
        'Set rstTempCDs = Nothing
    
        CurrentDb.Execute "UPDATE CourtDates SET [TurnaroundTimesCD] = " & sTurnaroundTimesCD & " WHERE [ID] = " & sCourtDatesID & ";"
        'Set rstTempCDs = CurrentDb.OpenRecordset("SELECT * FROM CourtDates WHERE [ID] = " & sCourtDatesID & ";")
        'rstTempCDs.Edit
        'rstTempCDs.Fields("TurnaroundTimesCD").Value = sTurnaroundTimesCD
        'rstTempCDs.Update
        'rstTempCDs.Close
        'Set rstTempCDs = Nothing
    
        CurrentDb.Execute "UPDATE CourtDates SET [InvoiceNo] = " & sInvoiceNumber & " WHERE [ID] = " & sCourtDatesID & ";"
        'Set rstTempCDs = CurrentDb.OpenRecordset("SELECT * FROM CourtDates WHERE [ID] = " & sCourtDatesID & ";")
        'rstTempCDs.Edit
        'rstTempCDs.Fields("InvoiceNo").Value = sInvoiceNumber
        'rstTempCDs.Update
        'rstTempCDs.Close
        'Set rstTempCDs = Nothing
    
        If IsNull(rstCurrentJob!StatusesID) Then
    
            rstCurrentStatusesEntry.Edit
            sStatusesID = rstCurrentStatusesEntry.Fields("ID").Value
            rstCurrentStatusesEntry.Update
            CurrentDb.Execute "UPDATE CourtDates SET StatusesID = " & sStatusesID & " WHERE [ID] = " & sCourtDatesID & ";"
            CurrentDb.Execute "UPDATE Statuses SET ContactsEntered = True, JobEntered = True WHERE [CourtDatesID] = " & sCourtDatesID & ";"
        End If
    
        rstCurrentJob.MoveNext
    
    Loop

    rstCurrentStatusesEntry.Close

    Call pfCheckFolderExistence 'checks for job folders/rough draft

    'import appearancesId from tempcustomers into courtdates
    sTempCustomersSQL = "SELECT * FROM TempCustomers;"

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    sCurrentJobSQL = "SELECT * FROM CourtDates WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
    Set rstTempJob = CurrentDb.OpenRecordset(sTempCustomersSQL)
    Set rstCurrentJob = CurrentDb.OpenRecordset(sCurrentJobSQL)

    x = 1

    rstTempJob.MoveFirst

    Do Until rstTempJob.EOF

        sCurrentTempApp = rstTempJob.Fields("AppID").Value
        sAppNumber = "App" & x
    
        If Not rstTempJob.EOF Or sCurrentTempApp <> "" Or Not IsNull(sCurrentTempApp) Then
            Select Case sAppNumber
            Case "App1", "App2", "App3", "App4", "App5", "App6"
                CurrentDb.Execute "UPDATE CourtDates SET " & sAppNumber & " = " & sCurrentTempApp & " WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
            Case Else
                Exit Do
            End Select
        
            'Set rstTempCDs = CurrentDb.OpenRecordset("SELECT * FROM CourtDates WHERE [ID] = " & sCourtDatesID & ";") '
            'rstTempCDs.Edit
            'If sAppNumber = "App7" Then
            '    rstTempCDs.Update
            '    rstTempCDs.Close
            '    Set rstTempCDs = Nothing
            '    GoTo ExitLoop
            'Else
            '    rstTempCDs.Fields(sAppNumber).Value = sCurrentTempApp
            '    rstTempCDs.Update
            '    rstTempCDs.Close
            '    Set rstTempCDs = Nothing
            'End If
            'rstTempJob.MoveNext
        
        Else:
            Exit Do
        End If
        x = x + 1
    Loop
    'rstCurrentJob.Close
    'rstTempJob.Close

    CurrentDb.Execute "INSERT INTO AGShortcuts (CourtDatesID, CasesID) SELECT CourtDatesID, CasesID FROM TempCourtDates;"

    Call fIsFactoringApproved                    'create new invioce
    Call pfGenerateJobTasks                      'generates job tasks
    Call pfPriorityPointsAlgorithm               'gives tasks priority points
    Call fProcessAudioParent                     'process audio in audio folder

    CurrentDb.Execute "DELETE FROM TempCourtDates", dbFailOnError
    CurrentDb.Execute "DELETE FROM TempCustomers", dbFailOnError
    CurrentDb.Execute "DELETE FROM TempCases", dbFailOnError

    'update statuses dependent on jurisdiction:
    'AddTrackingNumber, GenerateShippingEM, ShippingXMLs, BurnCD, FileTranscript,NoticeofService,SpellingsEmail

    Set rstCurrentCasesID = CurrentDb.OpenRecordset("SELECT * FROM Cases WHERE ID=" & sCasesID & ";")
    sJurisdiction = rstCurrentCasesID.Fields("Jurisdiction").Value
    rstCurrentCasesID.Close
    CurrentDb.Execute "UPDATE Statuses SET AddTrackingNumber = True, GenerateShippingEM = True, ShippingXMLs = True, " & _
               "BurnCD = True, FileTranscript = True, NoticeofService = True WHERE [CourtDatesID] = " & sCourtDatesID & ";"
    'CurrentDb.Execute "UPDATE Statuses SET WHERE [CourtDatesID] = " & sCourtDatesID & ";"
    'CurrentDb.Execute "UPDATE Statuses SET  WHERE [CourtDatesID] = " & sCourtDatesID & ";"
    'CurrentDb.Execute "UPDATE Statuses SET  WHERE [CourtDatesID] = " & sCourtDatesID & ";"
    'CurrentDb.Execute "UPDATE Statuses SET  WHERE [CourtDatesID] = " & sCourtDatesID & ";"
    'CurrentDb.Execute "UPDATE Statuses SET  WHERE [CourtDatesID] = " & sCourtDatesID & ";"

    Select Case sJurisdiction
    Case "Weber Nevada", "Weber Bankruptcy", "Weber Oregon", "Food and Drug Administration", "FDA", "AVT", _
         "eScribers", "AVTranz", "eScribers NH", "eScribers Bankruptcy", "J&J", "J&J Court Transcribers", "J&J Court"
        CurrentDb.Execute "UPDATE Statuses SET SpellingsEmail = True WHERE [CourtDatesID] = " & sCourtDatesID & ";"
    End Select

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

    Call pfGenericExportandMailMerge("Case", "Stage1s\OrderConfirmation")
    Call pfSendWordDocAsEmail("OrderConfirmation", "Transcript Order Confirmation") 'Order Confrmation Email

    sQuestion = "Would you like to complete stage 1 for job number " & sCourtDatesID & "?"
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    If sAnswer = vbNo Then                       'Code for No
        MsgBox "No paperwork will be processed."
    Else                                         'Code for yes
        Call pfStage1Ppwk
    End If

    Call fPlayAudioFolder(cJob.DocPath.JobDirectoryA) 'code for processing audio
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    MsgBox "Thanks, job entered!  Job number is " & sCourtDatesID & " if you want to process it!"

    Call pfClearGlobals
End Sub

Public Sub fCheckTempCustomersCustomers()
    '============================================================================
    ' Name        : fCheckTempCustomersCustomers
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fCheckTempCustomersCustomers
    ' Description : retrieve info from TempCustomers/Customers
    '============================================================================

    Dim sCheckTCuAgainstCuSQL As String
    Dim tcFirstName As String
    Dim tcLastName As String
    Dim tcCompany As String
    Dim tcMrMs As String
    Dim tcJobTitle As String
    Dim tcBusinessPhone As String
    Dim tcAddress As String
    Dim tcCity As String
    Dim tcZIP As String
    Dim tcState As String
    Dim tcNotes As String
    Dim tcFactoringApproved As String
    Dim tcCID As String

    Dim rstTempCustomers As DAO.Recordset
    Dim rstCheckTCuvCu As DAO.Recordset
    Dim rstCustomers As DAO.Recordset
    
    Set rstTempCustomers = CurrentDb.OpenRecordset("TempCustomers")

    If Not (rstTempCustomers.EOF And rstTempCustomers.BOF) Then

        rstTempCustomers.MoveFirst
    
    
        Do Until rstTempCustomers.EOF = True
        
        
            If rstTempCustomers.Fields("LastName").Value <> "" Then
                tcLastName = rstTempCustomers.Fields("LastName").Value
            End If
            If rstTempCustomers.Fields("FirstName").Value <> "" Then
                tcFirstName = rstTempCustomers.Fields("FirstName").Value
            End If
            If rstTempCustomers.Fields("Company").Value <> "" Then
                tcCompany = rstTempCustomers.Fields("Company").Value
            End If
            If rstTempCustomers.Fields("AppID").Value <> "" Then
                tcCID = rstTempCustomers.Fields("AppID").Value
            End If
            If rstTempCustomers.Fields("Company").Value <> "" Then
                tcCompany = rstTempCustomers("Company").Value
            End If
            If rstTempCustomers.Fields("MrMs").Value <> "" Then
                tcMrMs = rstTempCustomers.Fields("MrMs").Value
            End If
            If rstTempCustomers.Fields("JobTitle").Value <> "" Then
                tcJobTitle = rstTempCustomers.Fields("JobTitle").Value
            End If
            If rstTempCustomers.Fields("BusinessPhone").Value <> "" Then
                tcBusinessPhone = rstTempCustomers.Fields("BusinessPhone").Value
            End If
            If rstTempCustomers.Fields("Address").Value <> "" Then
                tcAddress = rstTempCustomers.Fields("Address").Value
            End If
            If rstTempCustomers.Fields("City").Value <> "" Then
                tcCity = rstTempCustomers.Fields("City").Value
            End If
            If rstTempCustomers.Fields("State").Value <> "" Then
                tcState = rstTempCustomers.Fields("State").Value
            End If
            If rstTempCustomers.Fields("ZIP").Value <> "" Then
                tcZIP = rstTempCustomers.Fields("ZIP").Value
            End If
            If rstTempCustomers.Fields("Notes").Value <> "" Then
                tcNotes = rstTempCustomers.Fields("Notes").Value
            End If
            If rstTempCustomers.Fields("FactoringApproved").Value <> "" Then
                tcFactoringApproved = rstTempCustomers.Fields("FactoringApproved").Value
            End If
    
        
            'query to check TempCustomers against Customers
            sCheckTCuAgainstCuSQL = "SELECT Customers.ID As AppID, Customers.LastName, Customers.FirstName, Customers.Company, Customers.Address, Customers.City, Customers.State, Customers.ZIP, Customers.MrMs, Customers.EmailAddress, Customers.JobTitle, Customers.BusinessPhone, Customers.MobilePhone, Customers.FaxNumber, Customers.Notes, Customers.FactoringApproved FROM Customers WHERE (((Customers.LastName) like " & Chr(34) & "*" & tcLastName & "*" & Chr(34) & ") AND ((Customers.FirstName) like " & Chr(34) & "*" & tcFirstName & "*" & Chr(34) & ") AND ((Customers.Company) like " & Chr(34) & "*" & tcCompany & "*" & Chr(34) & "));"
            Set rstCheckTCuvCu = CurrentDb.OpenRecordset(sCheckTCuAgainstCuSQL)
         
            If rstCheckTCuvCu.EOF Then           'if they are new customers do the following
        
                'Set rstCustomers = CurrentDb.OpenRecordset("SELECT * From Customers;")
                CurrentDb.Execute "INSERT INTO Customers (LastName, FirstName, Company, MrMs, JobTitle, BusinessPhone, Address, City, State, ZIP, FactoringApproved, Notes) VALUES (" & _
                                  tcLastName & ", " & tcFirstName & ", " & tcCompany & ", " & tcMrMs & ", " & tcJobTitle & ", " & tcBusinessPhone & ", " & tcAddress & ", " & tcCity & ", " & tcState & _
                                  ", " & tcZIP & ", " & tcFactoringApproved & ", " & "notes" & ");"

                'rstCustomers.AddNew
                'rstCustomers.Fields("LastName").Value = tcLastName
                'rstCustomers.Fields("FirstName").Value = tcFirstName
                'rstCustomers.Fields("Company").Value = tcCompany
                '"Ancel, Glink, Diamond, Bush, DiCianni & Krafthefer"
                'rstCustomers.Fields("MrMs").Value = tcMrMs
                'rstCustomers.Fields("JobTitle").Value = tcJobTitle
                'rstCustomers.Fields("BusinessPhone").Value = tcBusinessPhone
                'rstCustomers.Fields("Address").Value = tcAddress
                'rstCustomers.Fields("City").Value = tcCity
                'rstCustomers.Fields("State").Value = tcState
                'rstCustomers.Fields("ZIP").Value = tcZIP
            
                'rstCustomers.Fields("FactoringApproved").Value = tcFactoringApproved
                'rstCustomers.Fields("Notes").Value = "notes"
                tcCID = rstCheckTCuvCu.Fields("AppID").Value
                'rstCustomers.Update
                'rstCustomers.Close
        
            Else                                 'if they are previous customers, do the following
        
                tcCID = rstCheckTCuvCu.Fields("AppID").Value
                tcCompany = rstCheckTCuvCu.Fields("Company").Value
            
                tcMrMs = rstCheckTCuvCu.Fields("MrMs").Value
                tcLastName = rstCheckTCuvCu.Fields("LastName").Value
                tcFirstName = rstCheckTCuvCu.Fields("FirstName").Value
                tcJobTitle = rstCheckTCuvCu.Fields("JobTitle").Value
            
        
            End If
            'do for everyone
            CurrentDb.Execute "UPDATE TempCustomers SET AppID = " & tcCID & ", " & "Company = " & tcCompany & ", " & "MrMs = " & tcMrMs _
                            & ", " & "LastName= " & tcLastName & ", " & "FirstName = " & tcFirstName & ", " & "JobTitle = " & tcJobTitle & ", " & "BusinessPhone = " & _
                              tcBusinessPhone & ", " & "Address = " & tcAddress & ", " & "City = " & tcCity & ", " & "State = " & tcState & ", " & "ZIP = " & tcZIP _
                            & ", " & "Notes= " & tcNotes & ", " & "FactoringApproved = " & tcFactoringApproved & " WHERE [CourtDates].[ID] = " & sCourtDatesID & " AND AppID = '';"
            'rstTempCustomers.Edit
            'rstTempCustomers.Fields("AppID").Value = tcCID
            'rstTempCustomers.Fields("Company").Value = tcCompany
            'rstTempCustomers.Fields("MrMs").Value = tcMrMs
            'rstTempCustomers.Fields("LastName").Value = tcLastName
            'rstTempCustomers.Fields("FirstName").Value = tcFirstName
            'rstTempCustomers.Fields("JobTitle").Value = tcJobTitle
            'rstTempCustomers.Fields("BusinessPhone").Value = tcBusinessPhone
            'rstTempCustomers.Fields("Address").Value = tcAddress
            'rstTempCustomers.Fields("City").Value = tcCity
            'rstTempCustomers.Fields("State").Value = tcState
            'rstTempCustomers.Fields("ZIP").Value = tcZIP
            'rstTempCustomers.Fields("Notes").Value = tcNotes
            'rstTempCustomers.Fields("FactoringApproved").Value = tcFactoringApproved
            'Company, MrMs, LastName, FirstName,JobTitle,BusinessPhone,Address,City,State,ZIP,Notes,FactoringApproved
        
            'rstTempCustomers.Update
            'rstTempCustomers.MoveNext
        
        Loop
    
    Else
    End If

    rstCheckTCuvCu.Close
    Set rstCheckTCuvCu = Nothing

    rstTempCustomers.Close
    Set rstTempCustomers = Nothing

End Sub

Public Sub fCheckTempCasesCases()
    '============================================================================
    ' Name        : fCheckTempCasesCases
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fCheckTempCasesCases
    ' Description : retrieve info from TempCases/Cases
    '============================================================================

    Dim sCheckTCaAgainstCaSQL As String
    Dim sNewCasesIDSQL As String
    Dim tcsCourtDatesID As String
    Dim sCasesID As String
    Dim tcHearingTitle As String
    Dim tcParty1 As String
    Dim tcParty1Name As String
    Dim tcParty2 As String
    Dim tcParty2Name As String
    Dim tcCaseNumber1 As String
    Dim tcCaseNumber2 As String
    Dim tcJurisdiction As String
    Dim tcJudge As String
    Dim tcJudgeTitle As String
    
    Dim rstTempCases As DAO.Recordset
    Dim rstCheckTCavCa As DAO.Recordset
    Dim rstMaxCasesID As DAO.Recordset
    Dim rstCurrentJob As DAO.Recordset

    Set rstTempCases = CurrentDb.OpenRecordset("TempCases")
    rstTempCases.MoveFirst

    sCasesID = rstTempCases.Fields("CasesID").Value
    tcHearingTitle = rstTempCases.Fields("HearingTitle").Value
    tcParty1 = rstTempCases.Fields("Party1").Value
    tcParty1Name = rstTempCases.Fields("Party1Name").Value
    tcParty2 = rstTempCases.Fields("Party2").Value
    tcParty2Name = rstTempCases.Fields("Party2Name").Value
    tcCaseNumber1 = rstTempCases.Fields("CaseNumber1").Value
    tcCaseNumber2 = rstTempCases.Fields("CaseNumber2").Value
    tcJurisdiction = rstTempCases.Fields("Jurisdiction").Value
    tcJudge = rstTempCases.Fields("Judge").Value
    tcJudgeTitle = rstTempCases.Fields("JudgeTitle").Value

    'query to check TempCases against Cases
    sCheckTCaAgainstCaSQL = "SELECT Cases.ID As CasesID, Cases.CaseNumber1 as CaseNumber1, Cases.Party1 as Party1, Cases.Jurisdiction as Jurisdiction, Cases.Party2 as Party2, Cases.CaseNumber2 as CaseNumber2, Cases.Party1Name as Party1Name, Cases.Party2Name as Party2Name, Cases.HearingTitle as HearingTitle, Cases.Judge as Judge, Cases.JudgeTitle as JudgeTitle FROM Cases " & _
                            "WHERE ((Cases.CaseNumber1) like '*" & tcCaseNumber1 & "*') AND ((Cases.Party1) like '*" & tcParty1 & "*') AND ((Cases.Jurisdiction) like '*" & tcJurisdiction & "*');"

    Set rstCheckTCavCa = CurrentDb.OpenRecordset(sCheckTCaAgainstCaSQL)

    If rstCheckTCavCa.RecordCount < 1 Then       'if no match

        sNewCasesIDSQL = "INSERT INTO Cases (HearingTitle, Party1, Party1Name, Party2, Party2Name, CaseNumber1, CaseNumber2, Jurisdiction, Judge, JudgeTitle) SELECT HearingTitle, " & _
                         "Party1, Party1Name, Party2, Party2Name, CaseNumber1, CaseNumber2, Jurisdiction, Judge, JudgeTitle FROM [TempCases];"
        
        CurrentDb.Execute (sNewCasesIDSQL)
    
        Set rstMaxCasesID = CurrentDb.OpenRecordset("SELECT MAX(ID) as CasesID From Cases;")
    
        rstMaxCasesID.MoveFirst
        vCasesID = rstMaxCasesID.Fields("CasesID").Value
        sCasesID = rstMaxCasesID.Fields("CasesID").Value
        rstMaxCasesID.Close
    
        Set rstMaxCasesID = Nothing
        rstCheckTCavCa.Close
        rstTempCases.Close
    
    Else                                         'if there is a match

        Set rstCheckTCavCa = CurrentDb.OpenRecordset(sCheckTCaAgainstCaSQL)
        rstCheckTCavCa.MoveFirst
    
        sCasesID = rstCheckTCavCa.Fields("CasesID").Value
        tcHearingTitle = rstCheckTCavCa.Fields("HearingTitle").Value
        tcParty1 = rstCheckTCavCa.Fields("Party1").Value
        tcParty1Name = rstCheckTCavCa.Fields("Party1Name").Value
        tcParty2 = rstCheckTCavCa.Fields("Party2").Value
        tcParty2Name = rstCheckTCavCa.Fields("Party2Name").Value
        tcCaseNumber1 = rstCheckTCavCa.Fields("CaseNumber1").Value
        tcCaseNumber2 = rstCheckTCavCa.Fields("CaseNumber2").Value
        tcJurisdiction = rstCheckTCavCa.Fields("Jurisdiction").Value
        tcJudge = rstCheckTCavCa.Fields("Judge").Value
        tcJudgeTitle = rstCheckTCavCa.Fields("JudgeTitle").Value
    
        rstCheckTCavCa.Close
    
        Set rstTempCases = CurrentDb.OpenRecordset("TempCases")
        rstTempCases.Edit
    
        rstTempCases.Fields("CasesID").Value = sCasesID
        rstTempCases.Fields("HearingTitle").Value = tcHearingTitle
        rstTempCases.Fields("Party1").Value = tcParty1
        rstTempCases.Fields("Party1Name").Value = tcParty1Name
        rstTempCases.Fields("Party2").Value = tcParty2
        rstTempCases.Fields("Party2Name").Value = tcParty2Name
        rstTempCases.Fields("CaseNumber1").Value = tcCaseNumber1
        rstTempCases.Fields("CaseNumber2").Value = tcCaseNumber2
        rstTempCases.Fields("Jurisdiction").Value = tcJurisdiction
        rstTempCases.Fields("Judge").Value = tcJudge
        rstTempCases.Fields("JudgeTitle").Value = tcJudgeTitle
        rstTempCases.Update                      'update record
        rstTempCases.Close
    
    End If


    vCasesID = sCasesID
    Set rstCurrentJob = Nothing
    Set rstCheckTCavCa = Nothing
    Set rstTempCases = Nothing

    MsgBox "Checked for previous case info."
    
End Sub

Public Sub fInsertCalculatedFieldintoTempCourtDates()
    '============================================================================
    ' Name        : fInsertCalculatedFieldintoTempCourtDates
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fInsertCalculatedFieldintoTempCourtDates
    ' Description : insert several calculated fields into tempcourtdates
    '============================================================================
    Dim iTurnaroundTimesCD As Long
    Dim iAudioLength As Long
    Dim iEstimatedPageCount As Long
    Dim iUnitPriceID As Long
    Dim dInvoiceDate As Date
    Dim dExpectedBalanceDate As Date
    Dim dExpectedAdvanceDate As Date
    Dim dExpectedRebateDate As Date
    Dim cUnitPrice As Currency
    Dim sJurisdiction As String
    Dim sUnitPriceRateSrchSQL As String
    Dim InsertCustomersTempCourtDatesSQLstring As String
    
    Dim rs2 As DAO.Recordset
    Dim rstTempCourtDates As DAO.Recordset


    'calculate fields
    Set rstTempCourtDates = CurrentDb.OpenRecordset("TempCourtDates")
    iTurnaroundTimesCD = rstTempCourtDates.Fields("TurnaroundTimesCD").Value
    iAudioLength = rstTempCourtDates.Fields("AudioLength").Value
    sJurisdiction = rstTempCourtDates.Fields("Jurisdiction").Value

    'avail turnarounds 7 10 14 30 1 3
    'if jurisdiction contains and turnaround contains, for each different rate
    'avt rate 33 $1.35 or 35 $1.60, janet rate 37 $2.20, non-court rate 38 $2.00 per minute
    'regular 45 1 $6.05, 44 3 $5.45, 43 7 $4.85, 42 14 $4.25, 41 30 $3.65
    'volume 1 46 $6.65, 44 7 $5.45, 43 14 $4.85, 42 30 $4.25
    'copies for same 1.2, 1.05, 0.9, 0.9, 0.9
    'king county rate 40 3.10
        
    'Non -Court
    '    10 calendar-day turnaround, $2.00 per audio minute 49
    '    same day/overnight, $5.25 per page 42


    'Court Transcription
    '    electronic copy only (court receives hard copy where applicable)
    '    minimum 15 transcribed audio hours in one order
    '    45 calendar-day turnaround, $2.50/page
    '    30 calendar-day turnaround, $2.65/page 58
    '    14 calendar-day turnaround, $3.25/page 59
    '    07 calendar-day turnaround, $3.75/page 60
    '    03 calendar-day turnaround, $4.25/page 42

    dInvoiceDate = Date
    dExpectedBalanceDate = (Date + iTurnaroundTimesCD) - 2
    dExpectedAdvanceDate = (Date + iTurnaroundTimesCD) - 2
    dExpectedRebateDate = (Date + iTurnaroundTimesCD) + 28
    iEstimatedPageCount = ((iAudioLength / 60) * 45)

    Select Case True
    Case ((sJurisdiction) Like ("*" & "USBC" & "*")), ((sJurisdiction) Like ("*" & "superior court" & "*")), _
        ((sJurisdiction) Like ("*" & "Massachusetts" & "*")), ((sJurisdiction) Like ("*" & "Licensing" & "*")), _
        ((sJurisdiction) Like ("*" & "district court" & "*"))
        Select Case iTurnaroundTimesCD
        Case "45"
            iUnitPriceID = 64
        Case "30"
            iUnitPriceID = 58
        Case "14"
            iUnitPriceID = 59
        Case "7"
            iUnitPriceID = 60
        Case "3"
            iUnitPriceID = 42
        Case Else
            iUnitPriceID = 61
        End Select
    Case (sJurisdiction) Like ("*" & "non-court" & "*"), (sJurisdiction) Like ("*" & "noncourt" & "*")
        Select Case iTurnaroundTimesCD
        Case "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14"
            iUnitPriceID = 38
            iEstimatedPageCount = iAudioLength
        Case "2"
            iUnitPriceID = 49
        Case Else
            iUnitPriceID = 61
        End Select
    Case (sJurisdiction) Like ("*" & "Food and Drug Administration" & "*"), ((sJurisdiction) Like ("*" & "fda" & "*"))
        iUnitPriceID = 37
    Case ((sJurisdiction) Like ("*" & "KCI" & "*"))
        iUnitPriceID = 40
    Case ((sJurisdiction) Like ("*" & "Weber Oregon" & "*")), ((sJurisdiction) Like ("*" & "Weber Nevada" & "*")), _
        ((sJurisdiction) Like ("*" & "Weber Bankruptcy" & "*"))
        iUnitPriceID = 36
    Case ((sJurisdiction) Like ("*" & "eScribers" & "*")), ((sJurisdiction) Like ("*" & "Concord" & "*"))
        iUnitPriceID = 33
        
    End Select
    '
    'If iAudioLength > 865 Then
    '
    '    If ((sJurisdiction) Like ("*" & "USBC" & "*") And (iTurnaroundTimesCD) Like ("30")) Then
    '        iUnitPriceID = 58
    '    ElseIf ((sJurisdiction) Like ("*" & "USBC" & "*") And (iTurnaroundTimesCD) Like ("45")) Then iUnitPriceID = 64
    '    ElseIf ((sJurisdiction) Like ("*" & "USBC" & "*") And (iTurnaroundTimesCD) Like ("14")) Then iUnitPriceID = 59
    '    ElseIf ((sJurisdiction) Like ("*" & "USBC" & "*") And (iTurnaroundTimesCD) Like ("7")) Then iUnitPriceID = 60
    '    ElseIf ((sJurisdiction) Like ("*" & "USBC" & "*") And (iTurnaroundTimesCD) Like ("3")) Then iUnitPriceID = 42
    '    ElseIf ((sJurisdiction) Like ("*" & "USBC" & "*") And (iTurnaroundTimesCD) Like ("1")) Then iUnitPriceID = 61
    '    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (iTurnaroundTimesCD) Like ("45")) Then iUnitPriceID = 64
    '    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (iTurnaroundTimesCD) Like ("30")) Then iUnitPriceID = 58
    '    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (iTurnaroundTimesCD) Like ("14")) Then iUnitPriceID = 59
    '    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (iTurnaroundTimesCD) Like ("7")) Then iUnitPriceID = 60
    '    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (iTurnaroundTimesCD) Like ("3")) Then iUnitPriceID = 42
    '    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (iTurnaroundTimesCD) Like ("1")) Then iUnitPriceID = 61
    '    ElseIf ((sJurisdiction) Like ("*" & "non-court" & "*") And (iTurnaroundTimesCD) > ("2")) Then iUnitPriceID = 49
    '    ElseIf ((sJurisdiction) Like ("*" & "non-court" & "*") And (iTurnaroundTimesCD) Like ("1")) Then iUnitPriceID = 61
    '    ElseIf ((sJurisdiction) Like ("*" & "Food and Drug Administration" & "*")) Then iUnitPriceID = 37
    '    ElseIf ((sJurisdiction) Like ("*" & "fda" & "*")) Then iUnitPriceID = 37
    '    ElseIf ((sJurisdiction) Like ("*" & "KCI" & "*")) Then iUnitPriceID = 40
    '    ElseIf ((sJurisdiction) Like ("*" & "Weber Oregon" & "*")) Or ((sJurisdiction) Like ("*" & "Weber Nevada" & "*")) _
    '        Or ((sJurisdiction) Like ("*" & "Weber Bankruptcy" & "*")) Then iUnitPriceID = 36
    '    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (iTurnaroundTimesCD) Like ("30")) Then iUnitPriceID = 58
    '    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (iTurnaroundTimesCD) Like ("14")) Then iUnitPriceID = 59
    '    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (iTurnaroundTimesCD) Like ("7")) Then iUnitPriceID = 60
    '    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (iTurnaroundTimesCD) Like ("3")) Then iUnitPriceID = 42
    '    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (iTurnaroundTimesCD) Like ("1")) Then iUnitPriceID = 61
    '    Else
    '    End If
    '
    '
    '    If ((sJurisdiction) Like ("*" & "Licensing" & "*") And (iTurnaroundTimesCD) Like ("30")) Then
    '        iUnitPriceID = 58
    '    ElseIf ((sJurisdiction) Like ("*" & "Licensing" & "*") And (iTurnaroundTimesCD) Like ("14")) Then iUnitPriceID = 64
    '    ElseIf ((sJurisdiction) Like ("*" & "Licensing" & "*") And (iTurnaroundTimesCD) Like ("14")) Then iUnitPriceID = 59
    '    ElseIf ((sJurisdiction) Like ("*" & "Licensing" & "*") And (iTurnaroundTimesCD) Like ("7")) Then iUnitPriceID = 60
    '    ElseIf ((sJurisdiction) Like ("*" & "Licensing" & "*") And (iTurnaroundTimesCD) Like ("3")) Then iUnitPriceID = 42
    '    ElseIf ((sJurisdiction) Like ("*" & "Licensing" & "*") And (iTurnaroundTimesCD) Like ("1")) Then iUnitPriceID = 61
    '    Else
    '    End If
    '
    '    'if non-court, use audio length as page count for rate calculation
    '    If iUnitPriceID = "49" Then
    '        iEstimatedPageCount = iAudioLength
    '    Else
    '    End If
    '
    '    If ((sJurisdiction) Like ("*" & "eScribers" & "*")) Then iUnitPriceID = 33
    '
    '    If ((sJurisdiction) Like ("*" & "district court" & "*") And (iTurnaroundTimesCD) Like ("30")) Then
    '        iUnitPriceID = 58
    '    ElseIf ((sJurisdiction) Like ("*" & "district court" & "*") And (iTurnaroundTimesCD) Like ("14")) Then iUnitPriceID = 64
    '    ElseIf ((sJurisdiction) Like ("*" & "district court" & "*") And (iTurnaroundTimesCD) Like ("14")) Then iUnitPriceID = 59
    '    ElseIf ((sJurisdiction) Like ("*" & "district court" & "*") And (iTurnaroundTimesCD) Like ("7")) Then iUnitPriceID = 60
    '    ElseIf ((sJurisdiction) Like ("*" & "district court" & "*") And (iTurnaroundTimesCD) Like ("3")) Then iUnitPriceID = 42
    '    ElseIf ((sJurisdiction) Like ("*" & "district court" & "*") And (iTurnaroundTimesCD) Like ("1")) Then iUnitPriceID = 61
    '    Else
    '    End If
    '
    '    If (sJurisdiction) Like ("*" & "Concord" & "*") Then
    '        iUnitPriceID = 33
    '    Else
    '    End If
    '
    'Else
    '
    '    If ((sJurisdiction) Like ("*" & "USBC" & "*") And (iTurnaroundTimesCD) Like ("30")) Then
    '        iUnitPriceID = 39
    '    ElseIf ((sJurisdiction) Like ("*" & "USBC" & "*") And (iTurnaroundTimesCD) Like ("14")) Then iUnitPriceID = 41
    '    ElseIf ((sJurisdiction) Like ("*" & "USBC" & "*") And (iTurnaroundTimesCD) Like ("7")) Then iUnitPriceID = 62
    '    ElseIf ((sJurisdiction) Like ("*" & "USBC" & "*") And (iTurnaroundTimesCD) Like ("3")) Then iUnitPriceID = 57
    '    ElseIf ((sJurisdiction) Like ("*" & "USBC" & "*") And (iTurnaroundTimesCD) Like ("1")) Then iUnitPriceID = 61
    '    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (iTurnaroundTimesCD) Like ("30")) Then iUnitPriceID = 39
    '    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (iTurnaroundTimesCD) Like ("14")) Then iUnitPriceID = 41
    '    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (iTurnaroundTimesCD) Like ("7")) Then iUnitPriceID = 62
    '    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (iTurnaroundTimesCD) Like ("3")) Then iUnitPriceID = 57
    '    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (iTurnaroundTimesCD) Like ("1")) Then iUnitPriceID = 61
    '    ElseIf ((sJurisdiction) Like ("*" & "non-court" & "*") And (iTurnaroundTimesCD) > (2)) Then iUnitPriceID = 49
    '    ElseIf ((sJurisdiction) Like ("*" & "NonCourt" & "*") And (iTurnaroundTimesCD) > (2)) Then iUnitPriceID = 49
    '    ElseIf ((sJurisdiction) Like ("*" & "non-court" & "*") And (iTurnaroundTimesCD) Like ("1")) Then iUnitPriceID = 61
    '    ElseIf ((sJurisdiction) Like ("*" & "Food and Drug Administration" & "*")) Then iUnitPriceID = 37
    '    ElseIf ((sJurisdiction) Like ("*" & "fda" & "*")) Then iUnitPriceID = 37
    '    ElseIf ((sJurisdiction) Like ("*" & "KCI" & "*")) Then iUnitPriceID = 40
    '    ElseIf ((sJurisdiction) Like ("*" & "Weber Bankruptcy" & "*")) Then iUnitPriceID = 36
    '    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (iTurnaroundTimesCD) Like ("30")) Then iUnitPriceID = 39
    '    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (iTurnaroundTimesCD) Like ("14")) Then iUnitPriceID = 41
    '    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (iTurnaroundTimesCD) Like ("7")) Then iUnitPriceID = 62
    '    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (iTurnaroundTimesCD) Like ("3")) Then iUnitPriceID = 57
    '    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (iTurnaroundTimesCD) Like ("1")) Then iUnitPriceID = 61
    '
    '
    '    ElseIf ((sJurisdiction) Like ("*NonCourt*") And (iTurnaroundTimesCD) Like ("1")) Then iUnitPriceID = 61
    '    ElseIf ((sJurisdiction) Like ("*NonCourt*") And (iTurnaroundTimesCD) > ("2")) Then iUnitPriceID = 49
    '
    '    Else
    '
    '    End If
    '
    '
    '     If sJurisdiction Like "*NonCourt*" Then iUnitPriceID = 63
    '     If ((sJurisdiction) Like ("*" & "Licensing" & "*") And (iTurnaroundTimesCD) Like ("30")) Then
    '        iUnitPriceID = 39
    '    ElseIf ((sJurisdiction) Like ("*" & "Licensing" & "*") And (iTurnaroundTimesCD) Like ("14")) Then iUnitPriceID = 41
    '    ElseIf ((sJurisdiction) Like ("*" & "Licensing" & "*") And (iTurnaroundTimesCD) Like ("7")) Then iUnitPriceID = 62
    '    ElseIf ((sJurisdiction) Like ("*" & "Licensing" & "*") And (iTurnaroundTimesCD) Like ("3")) Then iUnitPriceID = 57
    '    ElseIf ((sJurisdiction) Like ("*" & "Licensing" & "*") And (iTurnaroundTimesCD) Like ("1")) Then iUnitPriceID = 61
    '    Else
    '    End If
    '
    '    'if non-court, use audio length as page count for rate calculation
    '    If iUnitPriceID = "38" Or iUnitPriceID = "46" Then
    '        iEstimatedPageCount = iAudioLength
    '    Else
    '    End If
    '
    '    If ((sJurisdiction) Like ("*" & "eScribers" & "*")) Then iUnitPriceID = 33
    '
    '    If ((sJurisdiction) Like ("*" & "district court" & "*") And (iTurnaroundTimesCD) Like ("30")) Then
    '        iUnitPriceID = 39
    '    ElseIf ((sJurisdiction) Like ("*" & "district court" & "*") And (iTurnaroundTimesCD) Like ("14")) Then iUnitPriceID = 41
    '    ElseIf ((sJurisdiction) Like ("*" & "district court" & "*") And (iTurnaroundTimesCD) Like ("7")) Then iUnitPriceID = 62
    '    ElseIf ((sJurisdiction) Like ("*" & "district court" & "*") And (iTurnaroundTimesCD) Like ("3")) Then iUnitPriceID = 57
    '    ElseIf ((sJurisdiction) Like ("*" & "district court" & "*") And (iTurnaroundTimesCD) Like ("1")) Then iUnitPriceID = 61
    '    Else
    '    End If
    '
    '    If (sJurisdiction) Like ("*" & "Concord" & "*") Then
    '        iUnitPriceID = 33
    '    Else
    '    End If
    '
    'End If

    'get proper rate now that we have UnitPrice ID
    sUnitPriceRateSrchSQL = "SELECT Rate from UnitPrice where ID = " & iUnitPriceID & ";"
    Set rs2 = CurrentDb.OpenRecordset(sUnitPriceRateSrchSQL)
    cUnitPrice = rs2.Fields("Rate").Value
    rs2.Close

    'calculate total price estimate
    sSubtotal = iEstimatedPageCount * cUnitPrice

    'set minimum charge
    If sSubtotal < 50 Then
        iUnitPriceID = 48
        sSubtotal = 50
    Else
    End If

    'insert calculated fields into tempcourtdates

    CurrentDb.Execute "UPDATE TempCourtDates SET InvoiceDate = " & dInvoiceDate & ", " & "UnitPrice = " & iUnitPriceID & ", " & "ExpectedRebateDate = " & dExpectedRebateDate _
                    & ", " & "ExpectedAdvanceDate = " & dExpectedAdvanceDate & ", " & "EstimatedPageCount = " & iEstimatedPageCount & ", " & "Subtotal = " & _
                      sSubtotal & " WHERE [CourtDatesID] = " & sCourtDatesID & ";"
    'rstTempCustomers.Edit
    '
    'Set rstTempCourtDates = CurrentDb.OpenRecordset("qSelect1TempCourtDates")
    'rstTempCourtDates.Edit
    '    rstTempCourtDates("InvoiceDate") = dInvoiceDate
    '    rstTempCourtDates("UnitPrice") = iUnitPriceID
    '    rstTempCourtDates("ExpectedRebateDate") = dExpectedRebateDate
    '    rstTempCourtDates("ExpectedAdvanceDate") = dExpectedAdvanceDate
    '    rstTempCourtDates("EstimatedPageCount") = iEstimatedPageCount
    '    rstTempCourtDates("Subtotal") = sSubtotal
    'rstTempCourtDates.Update
    'rstTempCourtDates.Close
    '

    MsgBox "Transcript Income Info:  " & Chr(13) & "Turnaround:  " & iTurnaroundTimesCD & " calendar days" _
                                               & Chr(13) & "Audio Length:  " & iAudioLength & " minutes" _
                                               & Chr(13) & "Estimated Page Count:  " & iEstimatedPageCount & " pages" _
                                               & Chr(13) & "Unit Price:  $" & cUnitPrice _
                                               & Chr(13) & "Expected Balance Payment Date:  " & dExpectedBalanceDate _
                                               & Chr(13) & "Expected Rebate Advance Date:  " & dExpectedAdvanceDate _
                                               & Chr(13) & "Expected Rebate Payment Date:  " & dExpectedRebateDate _
                                               & Chr(13) & "Expected Price Estimate:  $" & sSubtotal

End Sub

Public Sub fAudioPlayPromptTyping()
    '============================================================================
    ' Name        : fAudioPlayPromptTyping
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fAudioPlayPromptTyping
    ' Description : prompt to play audio in /Audio/folder
    '============================================================================

    Dim sQuestion As String
    Dim sAnswer As String

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    sQuestion = "Would you like to play the audio for job number " & sCourtDatesID & "?"
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    If sAnswer = vbNo Then                       'Code for No
        MsgBox "Audio will not be played."
    Else                                         'Code for yes
        Call fPlayAudioParent
    End If

End Sub

Public Sub fProcessAudioParent()
    '============================================================================
    ' Name        : fProcessAudioParent
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fProcessAudioParent
    ' Description : process audio in express scribe
    '============================================================================
    
    Dim cJob As New Job
    
    Call fProcessAudioFolder(cJob.DocPath.JobDirectoryA)

End Sub

Public Sub fPlayAudioParent()
    '============================================================================
    ' Name        : pfPlayAudioParent
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fPlayAudioParent
    ' Description : play audio as appropriate
    '============================================================================

    Dim cJob As New Job
    
    Call fPlayAudioFolder(cJob.DocPath.JobDirectoryA)


End Sub

Public Sub fPlayAudioFolder(ByVal sHostFolder As String)
    '============================================================================
    ' Name        : pfPlayAudioFolder
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fPlayAudioFolder(cJob.DocPath.JobDirectoryA)
    ' Description : plays audio folder
    '============================================================================

    Dim sExtension As String
    Dim sQuestion As String
    Dim sAnswer As String
    Dim sFileTypes() As String
    
    Dim blNotFirstIteration As Boolean
    
    Dim fiCurrentFile As File
    Dim foHFolder As Folder
    Dim FSO As Scripting.FileSystemObject
    
    Dim item As Variant
    
    Dim cJob As New Job

    Call pfCurrentCaseInfo                       'refresh transcript info

    Call pfAskforNotes

    Call pfAskforAudio

    sQuestion = "Would you like to play the audio for job number " & sCourtDatesID & "?"
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    If sAnswer = vbNo Then                       'Code for No

        MsgBox "Audio will not be played at this time."
    
    Else                                         'Code for yes
        
        If FSO Is Nothing Then Set FSO = New Scripting.FileSystemObject
    
        Set foHFolder = FSO.GetFolder(sHostFolder)
    
        'iterate through all files in the root of the main folder
        For Each fiCurrentFile In foHFolder.Files
          
            sExtension = FSO.GetExtensionName(fiCurrentFile.Path)
            GoTo Line2
            sFileTypes = Array("trs", "trm")
                
            For Each item In sFileTypes
                If fiCurrentFile Like "*trs*" Then GoTo Line2
                If fiCurrentFile Like "*trm*" Then GoTo Line2
            Next
                
            sFileTypes = Array("csx", "inf")
                
            For Each item In sFileTypes
                If fiCurrentFile Like "*csx*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-CourtSmartPlay.bat")
                If fiCurrentFile Like "*inf*" Then Exit For
            Next
                
            sFileTypes = Array("mp3", "mp4", "wav", "mpeg", "wma", "wmv", "divx", "m4v", "mov", "wmv")
                
            For Each item In sFileTypes
                If fiCurrentFile Like "*mp3*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribePlay.bat")
                If fiCurrentFile Like "*mp4*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribePlay.bat")
                If fiCurrentFile Like "*wav*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribePlay.bat")
                If fiCurrentFile Like "*mpeg*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribePlay.bat")
                If fiCurrentFile Like "*wma*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribePlay.bat")
                If fiCurrentFile Like "*wmv*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribePlay.bat")
                If fiCurrentFile Like "*divx*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribePlay.bat")
                If fiCurrentFile Like "*m4v*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribePlay.bat")
                If fiCurrentFile Like "*mov*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribePlay.bat")
                If fiCurrentFile Like "*wmv*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribePlay.bat")
            Next
                
        Next fiCurrentFile
          
    
Line2:
    End If
    Call pfClearGlobals
End Sub

Public Sub fProcessAudioFolder(ByVal HostFolder As String)
    '============================================================================
    ' Name        : pfProcessAudioFolder
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fProcessAudioFolder("\Production\2InProgress\" & sCourtDatesID & "\Audio\")
    ' Description : process audio in /Audio/ folder
    '============================================================================

    Dim sExtension As String
    Dim sQuestion As String
    Dim sAnswer As String
    Dim sFileTypes() As String
    
    Dim blNotFirstIteration As Boolean
    
    Dim fiCurrentFile As File
    Dim foHFolder As Folder
    Dim FSO As Scripting.FileSystemObject
    
    Dim item As Variant
    
    Dim cJob As New Job

    sQuestion = "Would you like to process the audio for job number " & sCourtDatesID & "?  Make sure the audio is in the \Audio\folder before proceeding."
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    If sAnswer = vbNo Then                       'Code for No
        MsgBox "Audio will not be processed."
    Else                                         'Code for yes
    
        If FSO Is Nothing Then Set FSO = New Scripting.FileSystemObject
    
        Set foHFolder = FSO.GetFolder(HostFolder)
    
        'iterate through all files in the root of the main folder
        
        For Each fiCurrentFile In foHFolder.Files
              
            sExtension = FSO.GetExtensionName(fiCurrentFile.Path)
            sFileTypes = Array("trs", "trm")
                    
            For Each item In sFileTypes
                If fiCurrentFile Like "*trs*" Then GoTo Line2
                If fiCurrentFile Like "*trm*" Then GoTo Line2
            Next
                    
            sFileTypes = Array("csx")
                    
            For Each item In sFileTypes
                If fiCurrentFile Like "*csx*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-CourtSmartPlay.bat")
            Next
                    
            sFileTypes = Array("mp3", "mp4", "wav", "mpeg", "wma", "wmv", "divx", "m4v", "mov", "wmv")
                    
            For Each item In sFileTypes
                If fiCurrentFile Like "*mp3*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribeAdd.bat")
                If fiCurrentFile Like "*mp4*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribeAdd.bat")
                If fiCurrentFile Like "*wav*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribeAdd.bat")
                If fiCurrentFile Like "*mpeg*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribeAdd.bat")
                If fiCurrentFile Like "*wma*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribeAdd.bat")
                If fiCurrentFile Like "*wmv*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribeAdd.bat")
                If fiCurrentFile Like "*divx*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribeAdd.bat")
                If fiCurrentFile Like "*m4v*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribeAdd.bat")
                If fiCurrentFile Like "*mov*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribeAdd.bat")
                If fiCurrentFile Like "*wmv*" Then Call Shell(cJob.DocPath.DBScripts & "Cortana\Audio-ExpressScribeAdd.bat")
            Next
                    
        Next fiCurrentFile
Line2:
    End If
End Sub

Public Sub pfPriceQuoteEmail()
    '============================================================================
    ' Name        : pfPriceQuoteEmail
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfPriceQuoteEmail
    ' Description : generates price quote and sends via e-mail
    '============================================================================

    Dim sQueryName As String
    Dim sSubtotal1 As String
    Dim sSubtotal2 As String
    Dim sSubtotal3 As String
    Dim sSubtotal4 As String
    Dim sPageRate4 As String
    Dim sPageRate3 As String
    Dim sPageRate2 As String
    Dim sPageRate1 As String
    Dim sPageRate8 As String
    Dim sPageRate7 As String
    Dim sPageRate6 As String
    Dim sPageRate5 As String
    Dim sPageRate As String
    Dim sPageRate9 As String
    Dim sSubtotal5 As String
    Dim sSubtotal6 As String
    Dim sPageRate10 As String

    Dim oWordAppDoc As New Word.Application
    Dim oOutlookApp As New Outlook.Application
    Dim oWordDoc As New Word.Document
    Dim oWordEditor As Word.editor
    Dim oWordApp As New Word.Application
    Dim Rng As Range
    
    Dim oOutlookMail As Object
    
    Dim qdfNew As QueryDef
    Dim qdObj As DAO.QueryDef

    Dim iDateDifference As Long
    Dim iPageCount As Long
    Dim iAudioLength As Long
    
    Dim dDeadline As Date
    
    Dim cJob As New Job

    dDeadline = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![txtDeadline]
    iAudioLength = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![txtAudioLength]

    sPageRate10 = "2.50"
    sPageRate6 = "2.65"
    sPageRate7 = "3.25"
    sPageRate8 = "3.75"
    sPageRate9 = "4.25"
    sPageRate5 = "5.25"

    sPageRate1 = "3.00"                          'get pagerate
    sPageRate2 = "3.50"
    sPageRate3 = "4.00"
    sPageRate4 = "4.75"

    iDateDifference = Int(DateDiff("d", Date, dDeadline))

    If iDateDifference < 4 And iDateDifference > 0 Then
        sPageRate = sPageRate5
    ElseIf iDateDifference < 8 And iDateDifference > 2 Then
        sPageRate = sPageRate9
    ElseIf iDateDifference < 15 And iDateDifference > 6 Then
        sPageRate = sPageRate8
    ElseIf iDateDifference < 31 And iDateDifference > 13 Then
        sPageRate = sPageRate7
    ElseIf iDateDifference > 30 And iDateDifference < 45 Then
        sPageRate = sPageRate6
    ElseIf iDateDifference > 44 Then
        sPageRate = sPageRate10
    End If
    '
    'Else
    '
    '        If iDateDifference < 4 And iDateDifference > 0 Then
    '            sPageRate = sPageRate5
    '        ElseIf iDateDifference < 8 And iDateDifference > 3 Then
    '            sPageRate = sPageRate4
    '        ElseIf iDateDifference < 15 And iDateDifference > 7 Then
    '            sPageRate = sPageRate3
    '        ElseIf iDateDifference < 31 And iDateDifference > 14 Then
    '            sPageRate = sPageRate2
    '        ElseIf iDateDifference < 30 Then
    '            sPageRate = sPageRate1
    '        Else
    '        End If
    '
    'End If


    iPageCount = Int((iAudioLength / 60) * 45)   'calculate PageCount

    'If iAudioLength > 885 Then

    sSubtotal6 = sPageRate5 * iPageCount
    sSubtotal4 = sPageRate9 * iPageCount
    sSubtotal3 = sPageRate8 * iPageCount
    sSubtotal2 = sPageRate7 * iPageCount
    sSubtotal1 = sPageRate6 * iPageCount
    sSubtotal5 = sPageRate10 * iPageCount
    '
    'Else
    '
    '    'calculate Subtotal1, Subtotal2, Subtotal3, Subtotal4
    '    sSubtotal1 = sPageRate1 * iPageCount
    '    sSubtotal2 = sPageRate2 * iPageCount
    '    sSubtotal3 = sPageRate3 * iPageCount
    '    sSubtotal4 = sPageRate4 * iPageCount
    '
    'End If

    sQueryName = "SELECT #" & dDeadline & "# AS Deadline, " & iAudioLength & " AS AudioLength, " & iPageCount & " AS PageCount, " _
               & sSubtotal1 & " AS Subtotal1, " & sSubtotal2 & " AS Subtotal2, " & _
                 sSubtotal3 & " AS Subtotal3, " & sSubtotal4 & " AS Subtotal4, " & sSubtotal5 & " AS Subtotal5;"
 
    On Error Resume Next
    With CurrentDb
        .QueryDefs.delete "tmpDataQry"
        Set qdfNew = .CreateQueryDef("tmpDataQry", sQueryName)
        .Close
    End With
    On Error GoTo 0

    DoCmd.OutputTo acOutputQuery, "tmpDataQry", acFormatXLSX, cJob.DocPath.PQXLS, False

    Set qdObj = Nothing
        
    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PQTemplate)

    'performs mail merge
    oWordDoc.Application.Visible = False
    oWordDoc.MailMerge.OpenDataSource Name:=cJob.DocPath.PQXLS, ReadOnly:=True
    oWordDoc.MailMerge.Execute
    oWordDoc.Application.ActiveDocument.SaveAs2 FileName:=cJob.DocPath.PQEmail
    oWordDoc.Application.ActiveDocument.Close

    'saves file in job number folder in in progress
    oWordDoc.Close SaveChanges:=wdSaveChanges

    'Set oOutlookApp = CreateObject("Outlook.Application")
    
    On Error Resume Next
    Set oWordApp = GetObject(, "Word.Application")

    If oWordApp Is Nothing Then
        Set oWordApp = CreateObject("Word.Application")
    End If

    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PQEmail)

    oWordDoc.Content.Copy

    Set oOutlookMail = oOutlookApp.CreateItem(0)
    With oOutlookMail
        .To = ""
        .CC = ""
        .BCC = ""
        .Subject = "Transcript Price Quote"
        .BodyFormat = olFormatRichText
        'Set oWordEditor = .GetInspector.WordEditor
        .GetInspector.WordEditor.Content.Paste
        .Display
    End With
    oWordDoc.Close
    oWordApp.Quit
    On Error GoTo 0
    Set oWordApp = Nothing
End Sub

Public Sub pfStage1Ppwk()
    'On Error GoTo eHandler
    '============================================================================
    ' Name        : pfStage1Ppwk
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfStage1Ppwk
    ' Description : completes all stage 1 tasks
    '============================================================================

    Dim sCourtRulesPath1 As String
    Dim sCourtRulesPath2 As String
    Dim sCourtRulesPath3 As String
    Dim sCourtRulesPath4 As String
    Dim sCourtRulesPath5 As String
    Dim sCourtRulesPath6 As String
    Dim sCourtRulesPath7 As String
    Dim sCourtRulesPath8 As String
    Dim sCourtRulesPath9 As String

    Dim sCourtRulesPath1a As String
    Dim sCourtRulesPath2a As String
    Dim sCourtRulesPath3a As String
    Dim sCourtRulesPath4a As String
    Dim sCourtRulesPath5a As String
    Dim sCourtRulesPath6a As String
    Dim sCourtRulesPath7a As String
    Dim sCourtRulesPath8a As String
    Dim sCourtRulesPath9a As String
    
    Dim sURL As String
    Dim sQuestion As String
    Dim sCourtRulesPath10a As String
    Dim sCourtRulesPath10 As String
    Dim sAnswer As String

    Dim cJob As New Job
    
    Call pfCurrentCaseInfo                       'refresh transcript info
    Call pfCheckFolderExistence                  'checks for job folder and creates it if not exists

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

    sCourtRulesPath1 = cJob.DocPath.TemplateFolder1 & "CourtRules-Bankruptcy-Rates.pdf"
    sCourtRulesPath2 = cJob.DocPath.TemplateFolder1 & "CourtRules-Bankruptcy-SafeguardingElectronicTranscripts.pdf"
    sCourtRulesPath3 = cJob.DocPath.TemplateFolder1 & "CourtRules-Bankruptcy-SampleTranscript.pdf"
    sCourtRulesPath4 = cJob.DocPath.TemplateFolder1 & "CourtRules-Bankruptcy-TranscriptFormatGuide-1.pdf"
    sCourtRulesPath5 = cJob.DocPath.TemplateFolder1 & "CourtRules-Bankruptcy-TranscriptFormatGuide-2.pdf"
    sCourtRulesPath6 = cJob.DocPath.TemplateFolder1 & "CourtRules-Bankruptcy-TranscriptRedactionQA.pdf"
    sCourtRulesPath7 = cJob.DocPath.TemplateFolder1 & "CourtRules-HowFileApprovedJurisdictions.pdf"
    sCourtRulesPath8 = cJob.DocPath.TemplateFolder1 & "CourtRules-WACounties.pdf"
    sCourtRulesPath9 = cJob.DocPath.TemplateFolder1 & "CourtRules-WACounties-2.pdf"
    sCourtRulesPath10 = cJob.DocPath.JurisdictionRefs & "Massachusetts\uniformtranscriptformat.pdf"

    sCourtRulesPath1a = cJob.DocPath.JobDirectoryN & "CourtRules-Bankruptcy-Rates.pdf"
    sCourtRulesPath2a = cJob.DocPath.JobDirectoryN & "CourtRules-Bankruptcy-SafeguardingElectronicTranscripts.pdf"
    sCourtRulesPath3a = cJob.DocPath.JobDirectoryN & "CourtRules-Bankruptcy-SampleTranscript.pdf"
    sCourtRulesPath4a = cJob.DocPath.JobDirectoryN & "CourtRules-Bankruptcy-TranscriptFormatGuide-1.pdf"
    sCourtRulesPath5a = cJob.DocPath.JobDirectoryN & "CourtRules-Bankruptcy-TranscriptFormatGuide-2.pdf"
    sCourtRulesPath6a = cJob.DocPath.JobDirectoryN & "CourtRules-Bankruptcy-TranscriptRedactionQA.pdf"
    sCourtRulesPath7a = cJob.DocPath.JobDirectoryN & "CourtRules-HowFileApprovedJurisdictions.pdf"
    sCourtRulesPath8a = cJob.DocPath.JobDirectoryN & "CourtRules-WACounties.pdf"
    sCourtRulesPath9a = cJob.DocPath.JobDirectoryN & "CourtRules-WACounties-2.pdf"
    sCourtRulesPath10a = cJob.DocPath.JobDirectoryN & "uniformtranscriptformat.pdf"

    Call pfSelectCoverTemplate                   'cover page prompt

    Call pfUpdateCheckboxStatus("CoverPage")
    Call pfGenericExportandMailMerge("Case", "Stage4s\TranscriptsReady")
    Call pfUpdateCheckboxStatus("TranscriptsReady")

    FileCopy sCourtRulesPath7, sCourtRulesPath7a

    Select Case True
    Case sJurisdiction Like "*AVT*", sJurisdiction Like "*AVTranz*", _
         sJurisdiction Like "*eScribers*", sJurisdiction Like "*FDA*", _
         sJurisdiction Like "Food and Drug Administration", _
         sJurisdiction Like "Weber Oregon", sJurisdiction Like "Weber Bankruptcy", _
         sJurisdiction Like "Weber Nevada"
        GoTo Line2
    Case sJurisdiction Like "*USBC*", sJurisdiction Like "*Bankruptcy*"
        FileCopy sCourtRulesPath1, sCourtRulesPath1a
        FileCopy sCourtRulesPath2, sCourtRulesPath2a
        FileCopy sCourtRulesPath3, sCourtRulesPath3a
        FileCopy sCourtRulesPath4, sCourtRulesPath4a
        FileCopy sCourtRulesPath5, sCourtRulesPath5a
        FileCopy sCourtRulesPath6, sCourtRulesPath6a
    Case sJurisdiction Like "*Superior Court*", sJurisdiction Like "*District Court*", sJurisdiction Like "*Supreme Court*"
        FileCopy sCourtRulesPath8, sCourtRulesPath8a
        FileCopy sCourtRulesPath9, sCourtRulesPath9a
    Case sJurisdiction Like "Massachusetts"
        FileCopy sCourtRulesPath10, sCourtRulesPath10a
    End Select

    '
    '
    'If sJurisdiction Like "*AVT*" Or sJurisdiction Like "*AVTranz*" Or sJurisdiction Like "*eScribers*" Then
    '    'FileCopy sCourtRulesPath9, sCourtRulesPath9a
    '    GoTo Line2
    'End If
    '
    'If sJurisdiction Like "*FDA*" Or sJurisdiction Like "Food and Drug Administration" Then
    '    'FileCopy sCourtRulesPath9, sCourtRulesPath9a
    '    GoTo Line2
    'End If
    '
    'If sJurisdiction Like "*USBC*" Or sJurisdiction Like "*Bankruptcy*" Then
    '    FileCopy sCourtRulesPath1, sCourtRulesPath1a
    '    FileCopy sCourtRulesPath2, sCourtRulesPath2a
    '    FileCopy sCourtRulesPath3, sCourtRulesPath3a
    '    FileCopy sCourtRulesPath4, sCourtRulesPath4a
    '    FileCopy sCourtRulesPath5, sCourtRulesPath5a
    '    FileCopy sCourtRulesPath6, sCourtRulesPath6a
    'End If
    '
    'If sJurisdiction Like "*Superior Court*" Or sJurisdiction Like "*District Court*" Or sJurisdiction Like "*Supreme Court*" Then
    '    FileCopy sCourtRulesPath8, sCourtRulesPath8a
    '    FileCopy sCourtRulesPath9, sCourtRulesPath9a
    'End If
    ''
    'If sJurisdiction Like "Weber Oregon" Or sJurisdiction Like "Weber Bankruptcy" Or sJurisdiction Like "Weber Nevada" Then
    '    GoTo Line2
    '    'FileCopy sCourtRulesPath9, sCourtRulesPath9a
    'End If
    '
    'If sJurisdiction Like "Massachusetts" Then FileCopy sCourtRulesPath10, sCourtRulesPath10a

    'Call pfCreateCDLabel 'cd label
    Call pfUpdateCheckboxStatus("CDLabel")

    'Call fCreatePELLetter 'package enclosed letter
    Call pfUpdateCheckboxStatus("PackageEnclosedLetter")


Line2:                                           'every jurisdiction converges here
    DoCmd.OpenQuery qXeroCSVQ, acViewNormal, acAdd 'export xero csv

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]


    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

    DoCmd.TransferText acExportDelim, , qSelectXero, cJob.DocPath.XeroCSV, True

    'TODO: xero api
    sURL = "https://go.xero.com/Import/Import.aspx?type=IMPORTTYPE/ARINVOICES"
    Application.FollowHyperlink (sURL)           'open xero website
    Call pfUpdateCheckboxStatus("InvoiceCompleted")


    Call pfInvoicesCSV                           'invoice creation prompt

    sURL = "https://go.xero.com/AccountsReceivable/Search.aspx?invoiceStatus=INVOICESTATUS%2fDRAFT&graphSearch=False"
    Application.FollowHyperlink (sURL)

    Call pfUpdateCheckboxStatus("InvoiceCompleted")

    Call fWunderlistAddNewJob

    sQuestion = "Want to send Adam an initial income report?"
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
    If sAnswer = vbNo Then                       'Code for No
        MsgBox "No initial income report will be sent.  You're done!"
    
    Else                                         'Code for yes
        Call pfGenericExportandMailMerge("Invoice", "Stage1s\CIDIncomeReport")
        Call pfCommunicationHistoryAdd("CIDIncomeReport")
        Call pfSendWordDocAsEmail("CIDIncomeReport", "Initial Income Notification") 'initial income report 'emails adam cid report

    End If

    sQuestion = "Want to send an order confirmation to the client?"
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    If sAnswer = vbNo Then                       'Code for No
        MsgBox "No confirmation will be sent.  You're done!"
    
    Else                                         'Code for yes

        Call pfGenericExportandMailMerge("Case", "Stage1s\OrderConfirmation")
        Call pfSendWordDocAsEmail("OrderConfirmation", "Transcript Order Confirmation") 'Order Confrmation Email
    
    End If


    MsgBox "Stage 1 complete."
    Call pfTypeRoughDraftF                       'type rough draft prompt
    Call pfClearGlobals
End Sub

Public Sub fWunderlistAddNewJob()
    '============================================================================
    ' Name        : fWunderlistAddNewJob
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fWunderlistAddNewJob()
    ' Description : add 1 task to a wunderlist list for general job due dates
    '               have it auto-set the next due date by stage
    '               4 tasks for each job, stage 1, 2, 3, 4
    '============================================================================
    'global variables lAssigneeID As Long, sDueDate As String, bStarred As Boolean
    '   bCompleted As Boolean, sTitle As String, sWLListID As String

    Dim sTitle As String
    Dim sDueDate As String
    Dim vErrorDetails As String
    Dim sURL As String
    Dim sEmail As String
    Dim sJSON As String
    Dim vErrorIssue As String
    Dim vErrorName As String
    Dim vErrorMessage As String
    Dim vErrorILink As String
    Dim sFile1 As String
    Dim sFile2 As String
    Dim sFile3 As String
    Dim sLine1 As String
    Dim sLine2 As String
    Dim sLine3 As String
    Dim sLists As String
    Dim sResponseText As String
    Dim apiWaxLRS As String

    Dim lFolderID As Long
    Dim iListID As Long
    
    Dim parsed As Dictionary
    
    Call pfCurrentCaseInfo

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

    sWLListID = 370524335                        'Ingram Household = 370524335
    'inbox = 370231796
    '1ToBeEntered = 388499976
    '2InProgress = 388499848
    '3Complete = 388499951
                        
    lAssigneeID = 88345676                       'erica / 86846933 adam
    bCompleted = "false"
    bStarred = "false"
    lFolderID = 13249242                         'id for "Production" folder

    sEmail = sCompanyEmail
    sTitle = sCourtDatesID

    'create a list JSON
    sJSON = "{" & Chr(34) & "title" & Chr(34) & ": " & Chr(34) & sTitle & Chr(34) & "}"
    
    'Debug.Print "sJSON-------------------------create a list JSON"
    'Debug.Print sJSON
    'Debug.Print "RESPONSETEXT--------------------------------------------"
    
    sURL = "https://a.wunderlist.com/api/v1/lists"
    
    With CreateObject("WinHttp.WinHttpRequest.5.1")
                
        .Open "POST", sURL, False
        .setRequestHeader "X-Access-Token", Environ("apiWunderlistT")
        .setRequestHeader "X-Client-ID", Environ("apiWunderlistUN")
        .setRequestHeader "Content-Type", "application/json"
        .send sJSON                              'send JSON to create empty list
        apiWaxLRS = .responseText
        .abort
    End With
    Set parsed = JsonConverter.ParseJson(apiWaxLRS)
    iListID = parsed("id")                       'get new list_id
    sTitle = parsed("title")
    
    'get folder ID
    
    'GET a.wunderlist.com/api/v1/folders to get list of all folders
    
    
    sURL = "https://a.wunderlist.com/api/v1/folders"
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "GET", sURL, False
        .setRequestHeader "X-Access-Token", Environ("apiWunderlistT")
        .setRequestHeader "X-Client-ID", Environ("apiWunderlistUN")
        .setRequestHeader "Content-Type", "application/json"
        .send
        '@Ignore AssignmentNotUsed
        apiWaxLRS = .responseText
        .abort
    End With
    
    
    '@Ignore AssignmentNotUsed
    apiWaxLRS = Left(apiWaxLRS, Len(apiWaxLRS) - 1)
    '@Ignore AssignmentNotUsed
    apiWaxLRS = Right(apiWaxLRS, Len(apiWaxLRS) - 1)
    
    Set parsed = JsonConverter.ParseJson(apiWaxLRS)
    vErrorName = parsed("id")                    '("value") 'second level array
    vErrorMessage = parsed("title")
    Dim rep As Object
    Set rep = parsed("list_ids")
    
    vErrorILink = ""
    Dim x As Long
    x = 1
    Dim ID As Variant
    For Each ID In rep                           ' third level objects
        If x = 1 Then
            vErrorILink = rep(x)
        Else
            vErrorILink = vErrorILink & "," & rep(x)
        End If
        x = x + 1
    Next
    vErrorIssue = parsed("revision")

    'put list in folder ID

    'PATCH a.wunderlist.com/api/v1/folders/:id to update folder by overwriting properties
    'params list_ids (list of list_ids), title, revision (required)
    
    vErrorILink = vErrorILink & "," & iListID
    vErrorILink = "[" & vErrorILink
    vErrorILink = vErrorILink & "]"

    sJSON = "{" & Chr(34) & _
                          "revision" & Chr(34) & ": " & vErrorIssue & ", " & Chr(34) & _
                          "title" & Chr(34) & ": " & Chr(34) & "Production" & Chr(34) & ", " & Chr(34) & _
                          "list_ids" & Chr(34) & ": " & vErrorILink _
                        & "}"
    
    sURL = "https://a.wunderlist.com/api/v1/folders/" & vErrorName
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "PATCH", sURL, False
        .setRequestHeader "X-Access-Token", Environ("apiWunderlistT")
        .setRequestHeader "X-Client-ID", Environ("apiWunderlistUN")
        .setRequestHeader "Content-Type", "application/json"
        .send sJSON
        '@Ignore AssignmentNotUsed
        apiWaxLRS = .responseText
        '@Ignore AssignmentNotUsed
        apiWaxLRS = Left(apiWaxLRS, Len(apiWaxLRS) - 1)
        '@Ignore AssignmentNotUsed
        apiWaxLRS = Right(apiWaxLRS, Len(apiWaxLRS) - 1)
        .abort
    End With
    
    'add 4 tasks to list:  Stage 1, Stage 2, Stage 3, Stage 4

    'POST a.wunderlist.com/api/v1/tasks
    'data:
    'list_id (required integer), title (required string), assignee_id (integer)
    'completed (boolean), due_date (string YYYY-MM-DD), starred (boolean)
    
    'auto-set task due dates
    'S1 = Today+2
    'S2 = DueDate-4
    'S3 = DueDate-2
    'S4 = DueDate
        
    'create a task add JSON
    sTitle = "Stage 1"
    bCompleted = "false"
    bStarred = "false"
    sDueDate = (Format((Date + 2), "yyyy-mm-dd"))
    
    sJSON = "{" & Chr(34) & _
                          "list_id" & Chr(34) & ": " & iListID & "," & Chr(34) & _
                          "title" & Chr(34) & ": " & Chr(34) & sTitle & Chr(34) & "," & Chr(34) & _
                          "assignee_id" & Chr(34) & ": " & lAssigneeID & "," & Chr(34) & _
                          "completed" & Chr(34) & ": " & bCompleted & "," & Chr(34) & _
                          "due_date" & Chr(34) & ": " & Chr(34) & sDueDate & Chr(34) & "," & Chr(34) & _
                          "starred" & Chr(34) & ": " & bStarred & _
                          "}"
    'Debug.Print "sJSON-----------------------------------Add Stage 1-4 Tasks"
    'Debug.Print sJSON
    'Debug.Print "RESPONSETEXT--------------------------------------------"
    
    sURL = "https://a.wunderlist.com/api/v1/tasks"
    
    With CreateObject("WinHttp.WinHttpRequest.5.1")
                        
        .Open "POST", sURL, False
        .setRequestHeader "X-Access-Token", Environ("apiWunderlistT")
        .setRequestHeader "X-Client-ID", Environ("apiWunderlistUN")
        .setRequestHeader "Content-Type", "application/json"
        .send sJSON                              'send JSON to create empty list
        apiWaxLRS = .responseText
        Debug.Print apiWaxLRS
        Debug.Print "Status:  " & .Status & "   |   " & "StatusText:  " & .StatusText
        Debug.Print "--------------------------------------------"
        .abort
    End With
    Set parsed = JsonConverter.ParseJson(apiWaxLRS)
    
    iListID = parsed("list_id")                  'get new list_id
    sTitle = parsed("title")


    'create a task add JSON
        
    sTitle = "Stage 2"
    bCompleted = "false"
    bStarred = "false"
    sDueDate = (Format((dDueDate - 4), "yyyy-mm-dd"))
    'dDueDate
    
    sJSON = "{" & Chr(34) & _
                          "list_id" & Chr(34) & ": " & iListID & "," & Chr(34) & _
                          "title" & Chr(34) & ": " & Chr(34) & sTitle & Chr(34) & "," & Chr(34) & _
                          "assignee_id" & Chr(34) & ": " & lAssigneeID & "," & Chr(34) & _
                          "completed" & Chr(34) & ": " & bCompleted & "," & Chr(34) & _
                          "due_date" & Chr(34) & ": " & Chr(34) & sDueDate & Chr(34) & "," & Chr(34) & _
                          "starred" & Chr(34) & ": " & bStarred & _
                          "}"
    
    sURL = "https://a.wunderlist.com/api/v1/tasks"
    
    With CreateObject("WinHttp.WinHttpRequest.5.1")
                
                
        .Open "POST", sURL, False
        .setRequestHeader "X-Access-Token", Environ("apiWunderlistT")
        .setRequestHeader "X-Client-ID", Environ("apiWunderlistUN")
        .setRequestHeader "Content-Type", "application/json"
        .send sJSON
        apiWaxLRS = .responseText
        
        .abort
    End With
    Set parsed = JsonConverter.ParseJson(apiWaxLRS)
    
    iListID = parsed("list_id")                  'get new list_id
    sTitle = parsed("title")

    'create a task add JSON
        
    sTitle = "Stage 3"
    bCompleted = "false"
    bStarred = "false"
    sDueDate = (Format((dDueDate - 3), "yyyy-mm-dd"))
    'dDueDate
    
    sJSON = "{" & Chr(34) & _
                          "list_id" & Chr(34) & ": " & iListID & "," & Chr(34) & _
                          "title" & Chr(34) & ": " & Chr(34) & sTitle & Chr(34) & "," & Chr(34) & _
                          "assignee_id" & Chr(34) & ": " & lAssigneeID & "," & Chr(34) & _
                          "completed" & Chr(34) & ": " & bCompleted & "," & Chr(34) & _
                          "due_date" & Chr(34) & ": " & Chr(34) & sDueDate & Chr(34) & "," & Chr(34) & _
                          "starred" & Chr(34) & ": " & bStarred & _
                          "}"
    
    sURL = "https://a.wunderlist.com/api/v1/tasks"
    
    With CreateObject("WinHttp.WinHttpRequest.5.1")
    
        .Open "POST", sURL, False
        .setRequestHeader "X-Access-Token", Environ("apiWunderlistT")
        .setRequestHeader "X-Client-ID", Environ("apiWunderlistUN")
        .setRequestHeader "Content-Type", "application/json"
        .send sJSON
        '@Ignore AssignmentNotUsed
        apiWaxLRS = .responseText
        .abort
    End With
    Set parsed = JsonConverter.ParseJson(apiWaxLRS)
    
    iListID = parsed("list_id")                  'get new list_id
    sTitle = parsed("title")

    'create a task add JSON
    
    sTitle = "Stage 4"
    bCompleted = "false"
    bStarred = "false"
    sDueDate = (Format((dDueDate - 1), "yyyy-mm-dd"))
    
    sJSON = "{" & Chr(34) & _
                          "list_id" & Chr(34) & ": " & iListID & "," & Chr(34) & _
                          "title" & Chr(34) & ": " & Chr(34) & sTitle & Chr(34) & "," & Chr(34) & _
                          "assignee_id" & Chr(34) & ": " & lAssigneeID & "," & Chr(34) & _
                          "completed" & Chr(34) & ": " & bCompleted & "," & Chr(34) & _
                          "due_date" & Chr(34) & ": " & Chr(34) & sDueDate & Chr(34) & "," & Chr(34) & _
                          "starred" & Chr(34) & ": " & bStarred & _
                          "}"
    
    sURL = "https://a.wunderlist.com/api/v1/tasks"
    
    With CreateObject("WinHttp.WinHttpRequest.5.1")
    
        .Open "POST", sURL, False
        .setRequestHeader "X-Access-Token", Environ("apiWunderlistT")
        .setRequestHeader "X-Client-ID", Environ("apiWunderlistUN")
        .setRequestHeader "Content-Type", "application/json"
        .send sJSON                              'send JSON to create empty list
        
        '@Ignore AssignmentNotUsed
        apiWaxLRS = .responseText
        .abort
    End With
    Set parsed = JsonConverter.ParseJson(apiWaxLRS)
    
    '@Ignore AssignmentNotUsed
    iListID = parsed("list_id")                  'get new list_id
    '@Ignore AssignmentNotUsed
    sTitle = parsed("title")

End Sub

Public Sub autointake()
    'TODO: autointake what's going on here
    'autoread email form into access db
    Dim sSubmissionDate As String
    Dim sEmailText As String
    Dim sSplitInfo() As String
    Dim sCSVInfo() As String
    Dim sInfoFields() As String
    Dim sAddress3A() As String
    Dim sYourNameA() As String
    Dim sAttorneyName() As String
    Dim sHearingDate As String
    Dim sCurrentAppString() As String
    Dim sAttorneyNameA() As String
    Dim sYourName As String
    Dim sFirstName As String
    Dim sLastName As String
    Dim sAFirst As String
    Dim sALast As String
    Dim tcCID As String
    Dim sAppNumber As String
    Dim vCasesID As String
    Dim sCurrentInput As String
    Dim sJobTitle As String
    Dim sBusinessPhone As String
    Dim sUnitPrice As String
    Dim sIRC As String
    Dim sFiled As String
    Dim sFactoring As String
    Dim sCompany As String
    Dim sEmail As String
    Dim sHardCopy As String
    Dim sTurnaround As String
    Dim sAudioLength As String
    Dim sAddress1 As String
    Dim sAddress2 As String
    Dim sParty1 As String
    Dim sParty2 As String
    Dim sCaseNumber1 As String
    Dim sCaseNumber2 As String
    Dim sJudge As String
    Dim sJurisdiction As String
    Dim sParty1Name As String
    Dim sParty2Name As String
    Dim sJudgeTitle As String
    Dim sHearingTitle As String
    Dim sHEnd As String
    Dim sHStart As String
    Dim sLocation As String
    Dim dInvoiceDate As String
    Dim dDueDate As String
    Dim dExpectedBalanceDate As String
    Dim dExpectedAdvanceDate As String
    Dim dExpectedRebateDate As String
    Dim iEstimatedPageCount As String
    Dim sAccountCode As String
    Dim sExtensionXLSM As String
    Dim sExtensionXLS As String
    Dim sFullPathXLS As String
    Dim sFullPathXLSM As String
    Dim sPartialPath As String
    Dim sTurnaroundTimesCD As String
    Dim sInvoiceNumber As String
    Dim sNewCourtDatesRowSQL As String
    Dim sOrderingID As String
    Dim sCurrentJobSQL As String
    Dim sTempJobSQL As String
    Dim sStatusesEntrySQL As String
    Dim sCasesID As String
    Dim sCurrentTempApp As String
    Dim sAddress3 As String
    Dim sLastA As String
    Dim sAnswer As String
    Dim sQuestion As String
    Dim sFirstA As String
    Dim sTempCustomersSQL As String

    Dim rstTempJob As DAO.Recordset
    Dim rstCurrentJob As DAO.Recordset
    Dim rstCurrentCasesID As DAO.Recordset
    Dim rstCurrentStatusesEntry As DAO.Recordset
    Dim rstStatuses As DAO.Recordset
    Dim rstMaxCasesID As DAO.Recordset
    Dim rstTempCDs As DAO.Recordset
    Dim rstOLP As DAO.Recordset
    Dim rstTempCourtDates As DAO.Recordset
    Dim rstTempCases As DAO.Recordset
    Dim rstTempCustomers As DAO.Recordset
    
    Dim oExcelWB As Excel.Workbook
    Dim oExcelMacroWB As Excel.Workbook
    
    Dim x As Long
    Dim y As Long
    
    Dim cJob As New Job


    Set rstOLP = CurrentDb.OpenRecordset("OLPaypalPayments")
    rstOLP.MoveFirst
    Do While rstOLP.EOF = False

        sEmailText = rstOLP.Fields("Contents").Value
        'split Contents at "|"
        sCSVInfo = Split(sEmailText, "|")
        'then split split contents
        sSplitInfo = str(sCSVInfo(1))
        sInfoFields = Split(sEmailText, ";")
        sSubmissionDate = Date
        sYourName = str(sInfoFields(0))
        sYourNameA = Split(sYourName, " ")
        sFirstName = sYourNameA(0)
        sLastName = sYourNameA(1)
        'split
        sAttorneyName = str(sInfoFields(1))
        sAttorneyNameA = Split(sYourName, " ")
        sFirstA = sAttorneyNameA(0)
        sLastA = sAttorneyNameA(1)
        'split
        sCompany = sInfoFields(2)
        sEmail = sInfoFields(3)
        sHardCopy = sInfoFields(4)
        sTurnaround = sInfoFields(5)
        sAudioLength = sInfoFields(6)
        sAddress1 = sInfoFields(7)
        sAddress2 = sInfoFields(8)
        sAddress3 = str(sInfoFields(9))
        sAddress3A = Split(sYourName, " ")
        sCity = sAddress3A(0)
        sState = sAddress3A(1)
        sZIP = sAddress3A(2)
        'split
        sParty1 = sInfoFields(10)
        sParty2 = sInfoFields(11)
        sCaseNumber1 = sInfoFields(12)
        sCaseNumber2 = sInfoFields(13)
        sJudge = sInfoFields(14)
        sJurisdiction = sInfoFields(15)
        sHearingDate = sInfoFields(16)
        'format
        sSubmissionDate = Date
    
        'ask for missing information to place in tempcourtdates
    
        sParty1Name = InputBox("Enter the title of Party 1 (Petitioner, Plaintiff, etc):")
        sParty2Name = InputBox("Enter the title of Party 2 (Defendant, Respondent, etc):")
        sJudgeTitle = InputBox("Enter the title of the judge:")
        sHearingTitle = InputBox("Enter the title of the hearing:")
        sHEnd = InputBox("Enter the hearing start time (##:## AM):")
        sHStart = InputBox("Enter the hearing end time (##:## AM):")
        sLocation = InputBox("Enter the city and state where this took place (Seattle, Washington):")
        dInvoiceDate = (Date + sTurnaround) - 2
        dDueDate = (Date + sTurnaround) - 2
        dExpectedBalanceDate = (Date + sTurnaround) - 2
        dExpectedAdvanceDate = (Date + sTurnaround) - 2
        dExpectedRebateDate = (Date + sTurnaround) + 28
        iEstimatedPageCount = ((sAudioLength / 60) * 45)
        sAccountCode = 400
        
        Select Case sTurnaround
        Case "45"
            sUnitPrice = 64
            sIRC = 96
                
        Case "30"
            sUnitPrice = 39
            sIRC = 17
                
        Case "14"
            sUnitPrice = 41
            sIRC = 19
                
        Case "7"
            sUnitPrice = 62
            sIRC = 20
                
        Case "3"
            sUnitPrice = 50
            sIRC = 84
                
        Case Else
            sUnitPrice = 61
            sIRC = 14
        End Select
        
        Select Case True
        Case sJurisdiction Like "*eScribers*"
            sUnitPrice = 33
            sIRC = 95
                
        Case sJurisdiction = "FDA", sJurisdiction = "Food and Drug Administration"
            sUnitPrice = 37
            sIRC = 41
                
        Case sJurisdiction Like "*Weber*", sJurisdiction Like "*J&J*"
            sUnitPrice = 36
            sIRC = 65
                
        Case sJurisdiction = "Non-Court", sJurisdiction = "NonCourt"
            sUnitPrice = 49
            sIRC = 86
                
        Case sJurisdiction Like "*KCI*"
            sUnitPrice = 40
            sIRC = 56
        
        End Select
        'calculate unitprice, inventoryratecode
        '        If sAudioLength >= 885 Then
        '            If sTurnaround = 30 Then sUnitPrice = 39
        '            If sTurnaround = 30 Then sIRC = 17
        '            If sTurnaround = 14 Then sUnitPrice = 41
        '            If sTurnaround = 14 Then sIRC = 19
        '            If sTurnaround = 7 Then sUnitPrice = 62
        '            If sTurnaround = 7 Then sIRC = 20
        '            If sTurnaround = 3 Then sUnitPrice = 50
        '            If sTurnaround = 3 Then sIRC = 84
        '            If sTurnaround = 1 Then sUnitPrice = 61
        '            If sTurnaround = 1 Then sIRC = 14
        '
        '        Else
        '            If sTurnaround = 30 Then sUnitPrice = 58
        '            If sTurnaround = 30 Then sIRC = 78
        '            If sTurnaround = 14 Then sUnitPrice = 59
        '            If sTurnaround = 14 Then sIRC = 7
        '            If sTurnaround = 7 Then sUnitPrice = 60
        '            If sTurnaround = 7 Then sIRC = 8
        '            If sTurnaround = 3 Then sUnitPrice = 42
        '            If sTurnaround = 3 Then sIRC = 90
        '            If sTurnaround = 1 Then sUnitPrice = 61
        '            If sTurnaround = 1 Then sIRC = 14
        '
        '            If sJurisdiction Like "*eScribers*" Then
        '                sUnitPrice = 33
        '                sIRC = 95
        '            End If
        '            If sJurisdiction = "FDA" Then
        '                sUnitPrice = 37
        '                sIRC = 41
        '            End If
        '            If sJurisdiction = "Food and Drug Administration" Then
        '                sUnitPrice = 37
        '                sIRC = 41
        '            End If
        '            If sJurisdiction Like "*Weber*" Then
        '                sUnitPrice = 36
        '                sIRC = 65
        '            End If
        '            If sJurisdiction Like "*J&J*" Then
        '                sUnitPrice = 36
        '                sIRC = 43
        '            End If
        '            If sJurisdiction = "Non-Court" Then
        '                sUnitPrice = 49
        '                sIRC = 86
        '            End If
        '            If sJurisdiction = "NonCourt" Then
        '                sUnitPrice = 49
        '                sIRC = 86
        '            End If
        '            If sJurisdiction Like "*KCI*" Then
        '                sUnitPrice = 40
        '                sIRC = 56
        '            End If
        '
        '        End If
        
        'TODO: What is going on here?
        sFiled = InputBox("Are we filing this, yes or no?")
        sFactoring = InputBox("Are we factoring this, yes or no?")
        
        Select Case sFiled
    
        Case "yes", "Yes", "YES", "Yeah", "yeah", "yea", "YEAH", "YEA", "Y", "y", _
             "yep", "Yep", "YEP", "YA", "Ya", "ya" 'filed
                
            Select Case sFactoring
            Case "yes", "Yes", "YES", "Yeah", "yeah", "yea", "YEAH", "YEA", _
                 "Y", "y", "yep", "Yep", "YEP", "YA", "Ya", "ya" 'no deposit
                sFactoring = True
                sBrandingTheme = 6
            Case "no", "NO", "No", "nah", "Nah", "NAH", "nope", "Nope", _
                 "NOPE", "N", "n"                'with deposit
                sFactoring = False
                sBrandingTheme = 8
            End Select
        
        Case "no", "NO", "No", "nah", "Nah", "NAH", "nope", "NOPE", "Nope", "N", "n" 'not filed
        
            Select Case sFactoring
            Case "yes", "Yes", "YES", "Yeah", "yeah", "yea", "YEAH", "YEA", _
                 "Y", "y", "yep", "Yep", "YEP", "YA", "Ya", "ya" 'no deposit
                sFactoring = True
                Select Case sJurisdiction
                Case "J&J", "J&J Court Transcribers", "J&J Court"
                    sBrandingTheme = 10
                Case "eScribers", "AVT", "AVTranz", "eScribers NH", _
                     "eScribers Bankruptcy"
                    sBrandingTheme = 11
                Case "FDA", "Food and Drug Administration", "Weber"
                    sBrandingTheme = 12
                Case "NonCourt", "Non-Court", "Noncourt", "NONCOURT"
                    sBrandingTheme = 1
                Case Else
                    sBrandingTheme = 7
                End Select
            Case "no", "NO", "No", "nah", "Nah", "NAH", "nope", "Nope", _
                 "NOPE", "n", "N"                'with deposit
                sFactoring = False
                Select Case sJurisdiction
                Case "NonCourt", "Non-Court", "Noncourt", "NONCOURT"
                    sBrandingTheme = 2
                Case Else
                    sBrandingTheme = 9
                End Select
            End Select
        End Select
        '
        '        If sFiled = "yes" Or sFiled = "Yes" Or sFiled = "Y" Or sFiled = "y" Then
        '
        '            If sFactoring = "yes" Or sFactoring = "Yes" Or sFactoring = "Y" Or sFactoring = "y" Then
        '                sFactoring = True
        '                sBrandingTheme = 6
        '            Else 'with deposit
        '                sFactoring = False
        '                sBrandingTheme = 8
        '            End If
        '
        '        Else 'not filed
        '
        '            If sFactoring = "yes" Or sFactoring = "Yes" Or sFactoring = "Y" Or sFactoring = "y" Then
        '                sFactoring = True
        '                If sJurisdiction Like "*J&J*" Then
        '                    sBrandingTheme = 10
        '                ElseIf sJurisdiction Like "*eScribers*" Then sBrandingTheme = 11
        '                ElseIf sJurisdiction Like "*FDA*" Or sJurisdiction Like "*Food and Drug Administration*" Then sBrandingTheme = 12
        '                ElseIf sJurisdiction Like "*Weber*" Then sBrandingTheme = 12
        '                ElseIf sJurisdiction Like "*NonCourt*" Or sJurisdiction Like "*Non-Court*" Then sBrandingTheme = 1
        '                Else: sBrandingTheme = 7
        '                End If
        '            Else 'with deposit
        '                sFactoring = False
        '                If sJurisdiction Like "*NonCourt*" Or sJurisdiction Like "*Non-Court*" Then
        '                    sBrandingTheme = 2
        '                Else: sBrandingTheme = 9
        '                End If
        '            End If
        '
        '
        '        End If
                
        'place info into tempcourtdates and tempcases
        Set rstTempCourtDates = CurrentDb.OpenRecordset("TempCourtDates")
        rstTempCourtDates.MoveFirst
        sTurnaround = rstTempCourtDates.Fields("TurnaroundTimesCD").Value
        rstTempCourtDates.Close
        dInvoiceDate = (Date + sTurnaround) - 2
        dDueDate = (Date + sTurnaround) - 2
        sAccountCode = 400

        CurrentDb.Execute "INSERT INTO TempCourtDates (SubmissionDate, FirstName, LastName, MrMs, AFirstName, ALastName, Company, Notes, EmailAddress, " & _
                   "HardCopy, Address1, Address2, City, State, ZIP, TurnaroundTimesCD, AudioLength, Party1, Party2, CaseNumber1, CaseNumber2, Judge, Jurisdiction, " & _
                   "HearingDate, Party1Name, Party2Name, JudgeTitle, HearingTitle, HearingEndTime, HearingStartTime, Location, InvoiceDate, DueDate, " & _
                   "AccountCode, UnitPrice, InventoryRateCode, BrandingTheme) VALUES (" & _
                   sSubmissionDate & ", " & sFirstName & ", " & sLastName & ", " & "Mrs" & ", " & sFirstA & ", " & sLastA & ", " & sCompany & ", " & sEmail & ", " & sCompanyEmail & _
                   ", " & sHardCopy & ", " & sAddress1 & ", " & sAddress2 & ", " & sCity & ", " & sState & ", " & sZIP & ", " & sTurnaround & ", " & sAudioLength & ", " & sParty1 & ", " & _
                   sParty2 & ", " & sCaseNumber1 & ", " & sCaseNumber2 & ", " & sJudge & ", " & sJurisdiction & ", " & sHearingDate & ", " & sParty1Name & ", " & sParty2Name & ", " & _
                   sJudgeTitle & ", " & sHearingTitle & ", " & sHEnd & ", " & sHStart & ", " & sLocation & ", " & dInvoiceDate & ", " & dDueDate & ", " & sAccountCode & ", " & sUnitPrice & ", " & _
                   sIRC & ", " & sBrandingTheme & ");"
        '
        '        'place info into tempcourtdates and tempcases
        '            Set rstTempCourtDates = CurrentDb.OpenRecordset("TempCourtDates")
        '                rstTempCourtDates.AddNew
        '                rstTempCourtDates.Fields("SubmissionDate").Value = sSubmissionDate
        '                rstTempCourtDates.Fields("FirstName").Value = sFirstName
        '                rstTempCourtDates.Fields("LastName").Value = sLastName
        '                rstTempCourtDates.Fields("MrMs").Value = "Mrs"
        '                rstTempCourtDates.Fields("AFirstName").Value = sFirstA
        '                rstTempCourtDates.Fields("ALastName").Value = sLastA
        '                rstTempCourtDates.Fields("Company").Value = sCompany
        '                rstTempCourtDates.Fields("Notes").Value = sEmail
        '                rstTempCourtDates.Fields("EmailAddress").Value = sCompanyEmail
        '                rstTempCourtDates.Fields("HardCopy").Value = sHardCopy
        '                rstTempCourtDates.Fields("Address1").Value = sAddress1
        '                rstTempCourtDates.Fields("Address2").Value = sAddress2
        '                rstTempCourtDates.Fields("City").Value = sCity
        '                rstTempCourtDates.Fields("State").Value = sState
        '                rstTempCourtDates.Fields("ZIP").Value = sZIP
        '                rstTempCourtDates.Fields("TurnaroundTimesCD").Value = sTurnaround
        '                rstTempCourtDates.Fields("AudioLength").Value = sAudioLength
        '                rstTempCourtDates.Fields("Address1").Value = sAddress1
        '                rstTempCourtDates.Fields("Address2").Value = sAddress2
        '                rstTempCourtDates.Fields("Party1").Value = sParty1
        '                rstTempCourtDates.Fields("Party2").Value = sParty2
        '                rstTempCourtDates.Fields("CaseNumber1").Value = sCaseNumber1
        '                rstTempCourtDates.Fields("CaseNumber2").Value = sCaseNumber2
        '                rstTempCourtDates.Fields("Judge").Value = sJudge
        '                rstTempCourtDates.Fields("Jurisdiction").Value = sJurisdiction
        '                rstTempCourtDates.Fields("HearingDate").Value = sHearingDate
        '                rstTempCourtDates.Fields("Party1Name").Value = sParty1Name
        '                rstTempCourtDates.Fields("Party2Name").Value = sParty2Name
        '                rstTempCourtDates.Fields("JudgeTitle").Value = sJudgeTitle
        '                rstTempCourtDates.Fields("HearingTitle").Value = sHearingTitle
        '                rstTempCourtDates.Fields("HearingEndTime").Value = sHEnd
        '                rstTempCourtDates.Fields("HearingStartTime").Value = sHStart
        '                rstTempCourtDates.Fields("Location").Value = sLocation
        '                rstTempCourtDates.Fields("InvoiceDate").Value = dInvoiceDate
        '                rstTempCourtDates.Fields("DueDate").Value = dDueDate
        '                rstTempCourtDates.Fields("AccountCode").Value = sAccountCode
        '                rstTempCourtDates.Fields("UnitPrice").Value = sUnitPrice
        '                rstTempCourtDates.Fields("InventoryRateCode").Value = sIRC
        '                rstTempCourtDates.Fields("BrandingTheme").Value = sBrandingTheme
        '                rstTempCourtDates.Update
        '    '           'SELECT FROM COURTDATES HearingDate, HearingStartTime, HearingEndTime, AudioLength, Location, TurnaroundTimesCD, InvoiceNo, DueDate, UnitPrice, InvoiceDate, InventoryRateCode, AccountCode, BrandingTheme FROM [TempCourtDates];"
        '
        'add to tempcases
        '                Set rstTempCases = CurrentDb.OpenRecordset("TempCases")
        '
        '                rstTempCases.AddNew
        '                rstTempCases.Fields("HearingTitle").Value = sHearingTitle
        '                rstTempCases.Fields("Party1").Value = sParty1
        '                rstTempCases.Fields("Party1Name").Value = sParty1Name
        '                rstTempCases.Fields("Party2").Value = sParty2
        '                rstTempCases.Fields("Party2Name").Value = sParty2Name
        '                rstTempCases.Fields("CaseNumber1").Value = sCaseNumber1
        '                rstTempCases.Fields("CaseNumber2").Value = sCaseNumber2
        '                rstTempCases.Fields("Jurisdiction").Value = sJurisdiction
        '                rstTempCases.Fields("Judge").Value = sJudge
        '                rstTempCases.Fields("JudgeTitle").Value = sJudgeTitle
        '                rstTempCases.Fields("Notes").Value = sEmail
        '                rstTempCases.Update
        '                rstTempCases.Close
        '                rstTempCourtDates.Close
                
        sNewCourtDatesRowSQL = "INSERT INTO TempCases (HearingTitle, Party1, Party1Name, Party2, Party2Name, CaseNumber1, CaseNumber2, " & _
                               "Jurisdiction, Judge, JudgeTitle, Notes) VALUES HearingTitle, Party1, Party1Name, Party2, Party2Name, CaseNumber1, CaseNumber2, " & _
                               "Jurisdiction, Judge, JudgeTitle, Notes FROM [TempCourtDates];"
        CurrentDb.Execute (sNewCourtDatesRowSQL)
        
        'enter apps into tempcustomers
        'ask how many appearances
        x = InputBox("How many appearances are there, 1 through 6?")
        'y = 1
            
        'loop questions for each number
        For y = 1 To x
            
            'add each appearance to tempcustomers
            sCurrentInput = InputBox("Please enter the appearance in the following fashion with semicolons separating each entry:" & Chr(13) & _
                                     "LastName;FirstName;Company;MrMs;JobTitle;BusinessPhone;Address;City;State;ZIP;Notes;FactoringApproved")
                
            'split what you input
            sCurrentAppString = Split(sCurrentInput, ";")
                
            'then separate split contents
            sLastName = sCurrentAppString(0)
            sFirstName = sCurrentAppString(1)
            sCompany = sCurrentAppString(2)
            sEmail = sCurrentAppString(3)
            sHardCopy = sCurrentAppString(4)
            sTurnaround = sCurrentAppString(5)
            sAudioLength = sCurrentAppString(6)
            sAddress1 = sCurrentAppString(7)
            sAddress2 = sCurrentAppString(8)
            sAddress3 = str(sCurrentAppString(9))
            'split
            sParty1 = sCurrentAppString(10)
            sParty2 = sCurrentAppString(11)
            sCaseNumber1 = sCurrentAppString(12)
            sCaseNumber2 = sCurrentAppString(13)
            sJudge = sCurrentAppString(14)
            sJurisdiction = sCurrentAppString(15)
            sHearingDate = sCurrentAppString(16)
                                                
                
            'enter into tempcustomers and tempcourtdates the appid after all questions answered
        
                
            'enter into tempcustomers
            '                Set rstTempCustomers = CurrentDb.OpenRecordset("TempCustomers")
            '                rstTempCustomers.AddNew
            '                rstTempCustomers.Fields("LastName").Value = sLastA
            '                rstTempCustomers.Fields("FirstName").Value = sFirstA
            '                rstTempCustomers.Fields("Company").Value = sCompany
            '                rstTempCustomers.Fields("MrMs").Value = sMrMs
            '                rstTempCustomers.Fields("JobTitle").Value = sJobTitle
            '                rstTempCustomers.Fields("BusinessPhone").Value = sBusinessPhone
            '                rstTempCustomers.Fields("Address").Value = sAddress1 & " " & sAddress2
            '                rstTempCustomers.Fields("City").Value = sCity
            '                rstTempCustomers.Fields("State").Value = sState
            '                rstTempCustomers.Fields("ZIP").Value = sZIP
            '                rstTempCustomers.Fields("Notes").Value = sEmail
            '                rstTempCustomers.Fields("FactoringApproved").Value = sFactoring
            '                tcCID = rstTempCustomers.Fields("ID").Value
            '                rstTempCustomers.Update
            '
                
            CurrentDb.Execute "INSERT INTO TempCourtDates (LastName, FirstName, Company, MrMs, JobTitle, BusinessPhone, Address, City, State, " & _
                       "ZIP, Notes, FactoringApproved) VALUES (" & _
                       sLastName & ", " & sFirstName & ", " & sLastName & ", " & sCompany & ", " & sMrMs & ", " & "" & ", " & "" & ", " & sAddress1 & " " & sAddress2 & ", " & _
                       sCity & ", " & sState & ", " & sZIP & ", " & sEmail & ", " & sFactoring & ");"
            'move to next appearance
        Next
    
        'run everything else like normal
        'delete blank lines
        CurrentDb.Execute "DELETE FROM TempCustomers WHERE [Company] = " & Chr(34) & Chr(34) & ";"
        'CurrentDb.Execute "DELETE FROM TempCourtDates WHERE [AudioLength] = " & Chr(34) & Chr(34) & ";"
        CurrentDb.Execute "DELETE FROM TempCases WHERE [Party1] = " & Chr(34) & Chr(34) & ";"
            
        'Perform the import
        sNewCourtDatesRowSQL = "INSERT INTO CourtDates (HearingDate, HearingStartTime, HearingEndTime, AudioLength, Location, TurnaroundTimesCD, InvoiceNo, DueDate, UnitPrice, InvoiceDate, InventoryRateCode, AccountCode, BrandingTheme) SELECT HearingDate, HearingStartTime, HearingEndTime, AudioLength, Location, TurnaroundTimesCD, InvoiceNo, DueDate, UnitPrice, InvoiceDate, InventoryRateCode, AccountCode, BrandingTheme FROM [TempCourtDates];"
        CurrentDb.Execute (sNewCourtDatesRowSQL)
            
            
        ' store courtdatesID
        sCourtDatesID = CurrentDb.OpenRecordset("SELECT @@IDENTITY")(0)
            
        [Forms]![NewMainMenu]![ProcessJobSubformNMM].[Form]![JobNumberField].Value = sCourtDatesID
            
        Call fCheckTempCustomersCustomers
        Call fCheckTempCasesCases
            
        sTempJobSQL = "SELECT * FROM TempCustomers;"
        Set rstTempJob = CurrentDb.OpenRecordset(sTempJobSQL)
                
        sCurrentJobSQL = "SELECT * FROM CourtDates WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
        Set rstCurrentJob = CurrentDb.OpenRecordset(sCurrentJobSQL)
            
        rstTempJob.MoveFirst
        sOrderingID = rstTempJob.Fields("AppID").Value
            
        If IsNull(rstCurrentJob!OrderingID) Then
            CurrentDb.Execute "UPDATE CourtDates SET OrderingID = " & sOrderingID & " WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
            rstTempJob.Close
            rstCurrentJob.Close
            Set rstTempJob = Nothing
            Set rstCurrentJob = Nothing
        End If
            
        Call fGenerateInvoiceNumber
        Call fInsertCalculatedFieldintoTempCourtDates
            
        'import casesID & CourtdatesID into tempcourtdates
        sCurrentJobSQL = "SELECT * FROM CourtDates WHERE ID = " & sCourtDatesID & ";"
        sTempJobSQL = "SELECT * FROM TempCourtDates;"
        sStatusesEntrySQL = "SELECT * FROM Statuses WHERE [CourtDatesID] = " & sCourtDatesID & ";"
        'CurrentDb.Execute "INSERT INTO Statuses (" & sCourtDatesID & ");"
        Set rstStatuses = CurrentDb.OpenRecordset("Statuses")
        rstStatuses.AddNew
        rstStatuses.Fields("CourtDatesID").Value = sCourtDatesID
        rstStatuses.Update
        rstStatuses.Close
        Set rstStatuses = Nothing
        Set rstTempJob = CurrentDb.OpenRecordset(sTempJobSQL)
        Set rstCurrentJob = CurrentDb.OpenRecordset(sCurrentJobSQL)
        Set rstCurrentStatusesEntry = CurrentDb.OpenRecordset(sStatusesEntrySQL)
        rstCurrentJob.MoveFirst
            
        Do Until rstCurrentJob.EOF
            
            sTurnaroundTimesCD = rstTempJob.Fields("TurnaroundTimesCD")
            sInvoiceNumber = rstTempJob.Fields("InvoiceNo")
            sCasesID = rstTempJob.Fields("CasesID")
                
                            
            CurrentDb.Execute "UPDATE TempCourtDates SET [CasesID] = " & sCasesID & " WHERE [CourtDatesID] = " & sCourtDatesID & ";"
            'rstTempJob.Edit
            'rstTempJob.Fields("CasesID") = sCasesID
            'rstTempJob.Update
            'sCasesID = rstTempJob.Fields("CasesID")
                
            CurrentDb.Execute "UPDATE TempCourtDates SET [CourtDatesID] = " & sCourtDatesID & " WHERE [InvoiceNo] = " & sInvoiceNumber & ";"
            '"SELECT * FROM TempCourtDates WHERE [InvoiceNo]=" & sInvoiceNumber & ";"
            'Set rstTempCDs = CurrentDb.OpenRecordset("TempCourtDates")
            'rstTempCDs.Edit
            'rstTempCDs.Fields("CourtDatesID").Value = sCourtDatesID
            'rstTempCDs.Update
            'rstTempCDs.Close
            'Set rstTempCDs = Nothing
            'CurrentDb.Execute "UPDATE TempCustomers SET [CourtDatesID] = " & sCourtDatesID & " WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
                
            CurrentDb.Execute "UPDATE TempCustomers SET [CourtDatesID] = " & sCourtDatesID & ";"
            'Set rstTempCDs = CurrentDb.OpenRecordset("TempCustomers")
            'rstTempCDs.Edit
            'rstTempCDs.Fields("CourtDatesID").Value = sCourtDatesID
            'rstTempCDs.Update
            'rstTempCDs.Close
            'Set rstTempCDs = Nothing
                
            CurrentDb.Execute "UPDATE CourtDates SET [CasesID] = " & sCasesID & " WHERE [ID] = " & sCourtDatesID & ";"
            'Set rstTempCDs = CurrentDb.OpenRecordset("SELECT * FROM CourtDates WHERE [ID] = " & sCourtDatesID & ";")
            'rstTempCDs.Edit
            'If sCasesID <> "" Then rstTempCDs.Fields("CasesID").Value = sCasesID
            'rstTempCDs.Update
            'rstTempCDs.Close
            'Set rstTempCDs = Nothing
                
            CurrentDb.Execute "UPDATE CourtDates SET [TurnaroundTimesCD] = " & sTurnaroundTimesCD & " WHERE [ID] = " & sCourtDatesID & ";"
            'Set rstTempCDs = CurrentDb.OpenRecordset("SELECT * FROM CourtDates WHERE [ID] = " & sCourtDatesID & ";")
            'rstTempCDs.Edit
            'rstTempCDs.Fields("TurnaroundTimesCD").Value = sTurnaroundTimesCD
            'rstTempCDs.Update
            'rstTempCDs.Close
            'Set rstTempCDs = Nothing
                
            CurrentDb.Execute "UPDATE CourtDates SET [InvoiceNo] = " & sInvoiceNumber & " WHERE [ID] = " & sCourtDatesID & ";"
            'Set rstTempCDs = CurrentDb.OpenRecordset("SELECT * FROM CourtDates WHERE [ID] = " & sCourtDatesID & ";")
            'rstTempCDs.Edit
            'rstTempCDs.Fields("InvoiceNo").Value = sInvoiceNumber
            'rstTempCDs.Update
            'rstTempCDs.Close
            'Set rstTempCDs = Nothing
                
            'CurrentDb.Execute "UPDATE TempCourtDates SET [CourtDatesID] = " & sCourtDatesID & " WHERE [TempCourtDates].[InvoiceNo] = " & sInvoiceNumber & ";"
            '"SELECT * FROM TempCourtDates WHERE [InvoiceNo]=" & sInvoiceNumber & ";"
            '                Set rstTempCDs = CurrentDb.OpenRecordset("TempCourtDates")
            '                rstTempCDs.Edit
            '                rstTempCDs.Fields("CourtDatesID").Value = sCourtDatesID
            '                rstTempCDs.Update
            '                rstTempCDs.Close
            '                Set rstTempCDs = Nothing
            'CurrentDb.Execute "UPDATE TempCustomers SET [CourtDatesID] = " & sCourtDatesID & " WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
            '
            '                Set rstTempCDs = CurrentDb.OpenRecordset("TempCustomers")
            '                rstTempCDs.Edit
            '                rstTempCDs.Fields("CourtDatesID").Value = sCourtDatesID
            '                rstTempCDs.Update
            '                rstTempCDs.Close
            '                Set rstTempCDs = Nothing
            '
            '
            '                'CurrentDb.Execute "UPDATE CourtDates SET [CasesID] = " & sCasesID & " WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
            '
            '                Set rstTempCDs = CurrentDb.OpenRecordset("SELECT * FROM CourtDates WHERE [ID] = " & sCourtDatesID & ";")
            '                rstTempCDs.Edit
            '                rstTempCDs.Fields("CasesID").Value = sCasesID
            '                rstTempCDs.Update
            '                rstTempCDs.Close
            '                Set rstTempCDs = Nothing
            '
                
            'CurrentDb.Execute "UPDATE CourtDates SET [TurnaroundTimesCD] = " & sTurnaroundTimesCD & " WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
            '
            '
            '                Set rstTempCDs = CurrentDb.OpenRecordset("SELECT * FROM CourtDates WHERE [ID] = " & sCourtDatesID & ";")
            '                rstTempCDs.Edit
            '                rstTempCDs.Fields("TurnaroundTimesCD").Value = sTurnaroundTimesCD
            '                rstTempCDs.Update
            '                rstTempCDs.Close
            '                Set rstTempCDs = Nothing
            '
                
            'CurrentDb.Execute "UPDATE CourtDates SET [InvoiceNo] = " & sInvoiceNumber & " WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
            '
            '                Set rstTempCDs = CurrentDb.OpenRecordset("SELECT * FROM CourtDates WHERE [ID] = " & sCourtDatesID & ";")
            '                rstTempCDs.Edit
            '                rstTempCDs.Fields("InvoiceNo").Value = sInvoiceNumber
            '                rstTempCDs.Update
            '                rstTempCDs.Close
            '                Set rstTempCDs = Nothing
            '
            '
                
                
            If IsNull(rstCurrentJob!StatusesID) Then
                
                rstCurrentStatusesEntry.Edit
                sStatusesID = rstCurrentStatusesEntry.Fields("ID")
                rstCurrentStatusesEntry.Update
                CurrentDb.Execute "UPDATE CourtDates SET StatusesID = " & sStatusesID & " WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
                CurrentDb.Execute "UPDATE Statuses SET ContactsEntered = True, JobEntered = True WHERE [CourtDatesID] = " & sCourtDatesID & ";"
                    
            End If
                
            rstCurrentJob.MoveNext
                
        Loop
            
        Call pfCheckFolderExistence              'checks for job folders/rough draft
            
        'import appearancesId from tempcustomers into courtdates
            
        sTempCustomersSQL = "SELECT * FROM TempCustomers;"
        sCurrentJobSQL = "SELECT * FROM CourtDates WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
            
        Set rstTempJob = CurrentDb.OpenRecordset(sTempCustomersSQL)
        Set rstCurrentJob = CurrentDb.OpenRecordset(sCurrentJobSQL)
            
        x = 1
            
        rstTempJob.MoveFirst
        '
            
        Do Until rstTempJob.EOF
            
            sCurrentTempApp = rstTempJob.Fields("AppID").Value
            sAppNumber = "App" & x
                
            If Not rstTempJob.EOF Or sCurrentTempApp <> "" Or Not IsNull(sCurrentTempApp) Then
                Select Case sAppNumber
                Case "App1", "App2", "App3", "App4", "App5", "App6"
                    CurrentDb.Execute "UPDATE CourtDates SET " & sAppNumber & " = " & sCurrentTempApp & " WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
                Case Else
                    Exit Do
                End Select
                    
                'Set rstTempCDs = CurrentDb.OpenRecordset("SELECT * FROM CourtDates WHERE [ID] = " & sCourtDatesID & ";") '
                'rstTempCDs.Edit
                'If sAppNumber = "App7" Then
                '    rstTempCDs.Update
                '    rstTempCDs.Close
                '    Set rstTempCDs = Nothing
                '    GoTo ExitLoop
                'Else
                '    rstTempCDs.Fields(sAppNumber).Value = sCurrentTempApp
                '    rstTempCDs.Update
                '    rstTempCDs.Close
                '    Set rstTempCDs = Nothing
                'End If
                'rstTempJob.MoveNext
                    
            Else:
                Exit Do
            End If
            x = x + 1
        Loop

        '            Do Until rstTempJob.EOF
        '
        '                sCurrentTempApp = rstTempJob.Fields("AppID").Value
        '                sAppNumber = "App" & x
        '
        '                If Not rstTempJob.EOF Or sCurrentTempApp <> "" Or Not IsNull(sCurrentTempApp) Then
        '
        '
        '                    'CurrentDb.Execute "UPDATE CourtDates SET " & sAppNumber & " = " & sCurrentTempApp & " WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
        '
        '                    Set rstTempCDs = CurrentDb.OpenRecordset("SELECT * FROM CourtDates WHERE [ID] = " & sCourtDatesID & ";")
        '                    rstTempCDs.Edit
        '                    rstTempCDs.Fields(sAppNumber).Value = sCurrentTempApp
        '                    rstTempCDs.Update
        '                    rstTempCDs.Close
        '                    Set rstTempCDs = Nothing
        '
        '
        '                    rstTempJob.MoveNext
        '                Else:
        '                    Exit Do
        '                End If
        '
        '
        '                x = x + 1
        '                rstTempJob.MoveNext
        '
        '
        '            Loop
        '
        'rstCurrentJob.Close
        'rstTempJob.Close
        'create new agshortcuts entry
        CurrentDb.Execute "INSERT INTO AGShortcuts (CourtDatesID, CasesID) SELECT CourtDatesID, CasesID FROM TempCourtDates;"
            
        Call fIsFactoringApproved                'create new invioce
        Call pfGenerateJobTasks                  'generates job tasks
        Call pfPriorityPointsAlgorithm           'gives tasks priority points
        Call fProcessAudioParent                 'process audio in audio folder
        
        CurrentDb.Execute "DELETE FROM TempCourtDates", dbFailOnError
        CurrentDb.Execute "DELETE FROM TempCustomers", dbFailOnError
        CurrentDb.Execute "DELETE FROM TempCases", dbFailOnError
            
        'update statuses dependent on jurisdiction:
        'AddTrackingNumber, GenerateShippingEM, ShippingXMLs, BurnCD, FileTranscript,NoticeofService,SpellingsEmail
            
        Set rstMaxCasesID = CurrentDb.OpenRecordset("SELECT MAX(ID) FROM Cases;")
            
        vCasesID = rstMaxCasesID.Fields(0).Value
            
        rstMaxCasesID.Close
            
        Set rstCurrentCasesID = CurrentDb.OpenRecordset("SELECT * FROM Cases WHERE ID=" & vCasesID & ";")
            
        sJurisdiction = rstCurrentCasesID.Fields("Jurisdiction").Value
            
        If sJurisdiction Like "Weber Nevada" Or sJurisdiction Like "Weber Bankruptcy" Or sJurisdiction Like "Weber Oregon" _
           Or sJurisdiction Like "Food and Drug Administration" Or sJurisdiction Like "*FDA*" Or sJurisdiction Like "*AVT*" _
           Or sJurisdiction Like "*eScribers*" Or sJurisdiction Like "*AVTranz*" Then
                
            CurrentDb.Execute "UPDATE Statuses SET AddTrackingNumber = True WHERE [CourtDatesID] = " & sCourtDatesID & ";"
            CurrentDb.Execute "UPDATE Statuses SET GenerateShippingEM = True WHERE [CourtDatesID] = " & sCourtDatesID & ";"
            CurrentDb.Execute "UPDATE Statuses SET ShippingXMLs = True WHERE [CourtDatesID] = " & sCourtDatesID & ";"
            CurrentDb.Execute "UPDATE Statuses SET BurnCD = True WHERE [CourtDatesID] = " & sCourtDatesID & ";"
            CurrentDb.Execute "UPDATE Statuses SET FileTranscript = True WHERE [CourtDatesID] = " & sCourtDatesID & ";"
            CurrentDb.Execute "UPDATE Statuses SET NoticeofService = True WHERE [CourtDatesID] = " & sCourtDatesID & ";"
            CurrentDb.Execute "UPDATE Statuses SET SpellingsEmail = True WHERE [CourtDatesID] = " & sCourtDatesID & ";"
            
        Else
        End If
            
        rstCurrentCasesID.Close
        sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
            
        Call pfGenericExportandMailMerge("Case", "Stage1s\OrderConfirmation")
        Call pfSendWordDocAsEmail("OrderConfirmation", "Transcript Order Confirmation") 'Order Confrmation Email
            
        sQuestion = "Would you like to complete stage 1 for job number " & sCourtDatesID & "?"
        sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
            
        If sAnswer = vbNo Then                   'Code for No
            MsgBox "No paperwork will be processed."
        Else                                     'Code for yes
            Call pfStage1Ppwk
        End If
        
        Call fPlayAudioFolder(cJob.DocPath.JobDirectoryA) 'code for processing audio
            
            
        MsgBox "Thanks, job entered!  Job number is " & sCourtDatesID & " if you want to process it!"
        Call pfClearGlobals
                
        rstOLP.MoveNext
    
    Loop
    
    'delete all from OLPayPalPayments
    sQuestion = "Jobs from email entered.  Ready to delete from table?"
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    If sAnswer = vbNo Then                       'Code for No
        MsgBox "No entries will be deleted."
    Else                                         'Code for yes
        CurrentDb.Execute "DELETE FROM OLPayPalPayments", dbFailOnError
    End If

    sQuestion = "Want to send an order confirmation to the client?"
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    If sAnswer = vbNo Then                       'Code for No
        MsgBox "No confirmation will be sent.  You're done!"
    
    Else                                         'Code for yes

        Call pfGenericExportandMailMerge("Case", "Stage1s\OrderConfirmation")
        Call pfSendWordDocAsEmail("OrderConfirmation", "Transcript Order Confirmation") 'Order Confrmation Email
    
    End If
End Sub

Public Sub NewOLEntry()
    'when new entry in OLPayPalPayments, run autointake function
    Dim sCount As DAO.Recordset

    Set sCount = CurrentDb.OpenRecordset("Select * from OLPaypalPayments;")
    If sCount.RecordCount > 0 Then
        Call autointake
        Call ScrollingMarquee
    
    Else
    End If

End Sub

Private Sub ResetDisplay()

    MinimizeNavigationPane
    'Me.lblFlash.Visible = False
    'Me.txtMarquee.Visible = False
    'Me.TimerInterval = 0
    '
    'Me.cmd10.Caption = "Scrolling Marquee Text"
    'Me.cmd10.ForeColor = RGB(63, 63, 63)
    'Me.cmd10.FontWeight = 400
    'Me.cmd10.FontSize = 12
    
End Sub

Private Sub ScrollingMarquee()
    Dim strText As String
    Dim n As Long
    ResetDisplay

    MinimizeNavigationPane
    
    'Sets the timer in motion for case 10 - scrolling text

    n = 10
    '@Ignore AssignmentNotUsed
    sCourtDatesID = DMax("[ID]", "CourtDates")
    
    'If Me.TimerInterval = 0 Then
    'Me.cmd10.Caption = "STOP Scrolling Marquee Text"
    'Me.cmd10.ForeColor = RGB(0, 32, 68)
    'Me.cmd10.FontWeight = 800
    'Me.cmd10.FontSize = 16
    'Me.TimerInterval = 100
    'Me.txtMarquee.Visible = True
    strText = "      IMPORTANT MESSAGE : You have a new job.  Please enter " & sCourtDatesID & " to process it or send an invoice . . . . "

    'Else
    'Me.TimerInterval = 0
    'Me.txtMarquee.Visible = False
    'Me.cmd10.Caption = "Scrolling Marquee Text"
    'Me.cmd10.ForeColor = RGB(0, 32, 68)
    'Me.cmd10.FontWeight = 400
    'Me.cmd10.FontSize = 12
    '@Ignore AssignmentNotUsed
    strText = ""
    'End If
    
End Sub

Public Sub MinimizeNavigationPane()

    On Error GoTo ErrHandler

    DoCmd.NavigateTo "acNavigationCategoryObjectType"
    DoCmd.Minimize
        
Exit_ErrHandler:
    Exit Sub
    
ErrHandler:
    MsgBox "Error " & Err.Number & " in HideNavigationPane routine : " & Err.Description, vbOKOnly + vbCritical
    Resume Exit_ErrHandler

End Sub


