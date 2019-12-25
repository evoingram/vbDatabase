VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("Database.Admin.Classes")
Option Compare Database
Option Explicit


Private sTableName As String

Public Property Get TableName() As String
    TableName = sTableName
End Property

Public Property Let TableName(sTableName1 As String)
    sTableName = sTableName1
End Property

'classes
'open/close worksheet
'open/close/edit/delete recordset
'open/close/create document
'manipulate tasks/status checklists
'connect/disconnect/upload/download ftp
'paypal
'folders

'Select Case sFiled
'Case "yes", "Yes", "YES", "Yeah", "yeah", "yea", "YEAH", "YEA", "Y", "y", "yep", "Yep", "YEP", "YA", "Ya", "ya"
'Case "no", "NO", "No", "nah", "Nah", "NAH"
'End Select
        
Public Function addRSet(sTableName As String) As String
    'add
    Dim sCasesID As String
    Dim db As Database
    Dim sPartialPath As String
    Dim sExtensionXLS As String
    Dim sExtensionXLSM As String
    Dim sFullPathXLS As String
    Dim sFullPathXLSM As String
    Dim oExcelApp As Object
    Dim oExcelMacroWB As Object
    Dim oExcelWB As Object
    Dim sFiled As String
    Dim sFactoring As String
    Dim sTurnaround As String
    Dim sAccountCode As String
    Dim sNewCourtDatesRowSQL As String
    Dim sTempJobSQL As String
    Dim sCurrentJobSQL As String
    Dim sTempCustomersSQL As String
    Dim sCurrentTempApp As String
    Dim sAppNumber As String

    Dim rstCourtDatesID As DAO.Recordset
    Dim rstTempJob As DAO.Recordset
    Dim rstCurrentJob As DAO.Recordset
    Dim rstTableName As DAO.Recordset
    Dim rstTempCourtDates As DAO.Recordset

    Dim x As Long

    Dim dInvoiceDate As Date
    
    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID

    Select Case sTableName

        Case "TempCourtDates"
            'add new case/transcript order to tempcourtdates
            Set oExcelApp = CreateObject("Excel.Application")
            
            Set oExcelMacroWB = oExcelApp.Application.Workbooks.Open(FileName:=cJob.DocPath.OrderFormXLSM, Local:=True)
            oExcelMacroWB.Application.DisplayAlerts = False
            oExcelMacroWB.Application.Visible = False
            oExcelMacroWB.SaveAs Replace(cJob.DocPath.OrderFormXLSM, ".xlsm", ".xls"), FileFormat:=xlWorkbookDefault
            oExcelMacroWB.Close True
            Set oExcelMacroWB = Nothing
            
            Set oExcelWB = oExcelApp.Application.Workbooks.Open(FileName:=cJob.DocPath.OrderFormXLSX, Local:=True)
            oExcelWB.Application.DisplayAlerts = False
            oExcelWB.Application.Visible = False
            oExcelWB.SaveAs Replace(cJob.DocPath.OrderFormXLSX, ".xls", ".csv"), FileFormat:=6
            oExcelWB.Close True
            Set oExcelWB = Nothing
             
            
            DoCmd.TransferText TransferType:=acImportDelim, TableName:="TempCourtDates", _
                               FileName:=cJob.DocPath.OrderFormCSV, HasFieldNames:=True
            CurrentDb.TableDefs.Refresh
            
            'generate necessary info
            Set rstTableName = CurrentDb.OpenRecordset("TempCourtDates")
            sFiled = rstTableName.Fields("Filing").Value
            sFactoring = rstTableName.Fields("Factoring").Value
            sTurnaround = rstTableName.Fields("TurnaroundTimesCD").Value
            dInvoiceDate = (Date + sTurnaround) - 2
            dDueDate = (Date + sTurnaround) - 2
            sAccountCode = 400
            
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
        Case "Non-Court", "NonCourt"
            sUnitPrice = 49
            sIRC = 86
        Case "KCI"
            sUnitPrice = 40
            sIRC = 56
        End Select
        
        Select Case sFiled
            
        Case "yes", "Yes", "YES", "Yeah", "yeah", "yea", "YEAH", "YEA", "Y", "y", _
             "yep", "Yep", "YEP", "YA", "Ya", "ya" 'filed
                        
            Select Case sFactoring
            Case "yes", "Yes", "YES", "Yeah", "yeah", "yea", "YEAH", "YEA", _
                 "Y", "y", "yep", "Yep", "YEP", "YA", "Ya", "ya" 'no deposit
                sFactoring = True
                sBrandingTheme = 6
            Case "no", "NO", "No", "nah", "Nah", "NAH", "nope", "Nope", _
                 "NOPE"                          'with deposit
                sFactoring = False
                sBrandingTheme = 8
            End Select
                
        Case "no", "NO", "No", "nah", "Nah", "NAH", "nope", "NOPE", "Nope" 'not filed
                
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
                 "NOPE"                          'with deposit
                sFactoring = False
                Select Case sJurisdiction
                Case "NonCourt", "Non-Court", "Noncourt", "NONCOURT"
                    sBrandingTheme = 2
                Case Else
                    sBrandingTheme = 9
                End Select
            End Select
        End Select
        
        'place just-calculated info into tempcourtdates
        rstTableName.MoveFirst
        rstTableName.Edit
        rstTableName.Fields("InvoiceDate").Value = dInvoiceDate
        rstTableName.Fields("DueDate").Value = dDueDate
        rstTableName.Fields("AccountCode").Value = sAccountCode
        rstTableName.Fields("UnitPrice").Value = sUnitPrice
        rstTableName.Fields("InventoryRateCode").Value = sIRC
        rstTableName.Fields("BrandingTheme").Value = sBrandingTheme
        rstTableName.Update
        
        'delete empty entries from table
        CurrentDb.Execute "DELETE FROM TempCourtDates WHERE [AudioLength] IS NULL;"
        
        Exit Function

    Case "Cases"
        Set rstTempCourtDates = CurrentDb.OpenRecordset(sTableName)
        'add new entry to cases from TempCourtDates
        rstTempCourtDates.MoveFirst
        rstTableName.AddNew
        sCasesID = rstTableName.Fields("ID").Value
        rstTableName.Fields("HearingTitle").Value = rstTempCourtDates.Fields("HearingTitle").Value
        rstTableName.Fields("Party1").Value = rstTempCourtDates.Fields("Party1").Value
        rstTableName.Fields("Party1Name").Value = rstTempCourtDates.Fields("Party1Name").Value
        rstTableName.Fields("Party2").Value = rstTempCourtDates.Fields("Party2").Value
        rstTableName.Fields("Party2Name").Value = rstTempCourtDates.Fields("Party2Name").Value
        rstTableName.Fields("CaseNumber1").Value = rstTempCourtDates.Fields("CaseNumber1").Value
        rstTableName.Fields("CaseNumber2").Value = rstTempCourtDates.Fields("CaseNumber2").Value
        rstTableName.Fields("Jurisdiction").Value = rstTempCourtDates.Fields("Jurisdiction").Value
        rstTableName.Fields("Judge").Value = rstTempCourtDates.Fields("Judge").Value
        rstTableName.Fields("JudgeTitle").Value = rstTempCourtDates.Fields("JudgeTitle").Value
        rstTableName.Fields("Notes").Value = rstTempCourtDates.Fields("Notes").Value
        rstTableName.Update
        
        'return cases table ID number
        addRSet = sCasesID
        Exit Function
        
    Case "CourtDates"
        'add new entry to courtdates from TempCourtDates
        sNewCourtDatesRowSQL = "INSERT INTO " & sTableName & " (HearingDate, HearingStartTime, " & _
                               "HearingEndTime, AudioLength, Location, TurnaroundTimesCD, InvoiceNo, DueDate, " & _
                               "UnitPrice, InvoiceDate, InventoryRateCode, AccountCode, BrandingTheme) SELECT " & _
                               "HearingDate, HearingStartTime, HearingEndTime, AudioLength, Location, " & _
                               "TurnaroundTimesCD, InvoiceNo, DueDate, UnitPrice, InvoiceDate, InventoryRateCode, " & _
                               "AccountCode, BrandingTheme FROM [TempCourtDates];"
        CurrentDb.Execute (sNewCourtDatesRowSQL)
        
        ' store courtdatesID
        Set rstCourtDatesID = CurrentDb.OpenRecordset("SELECT MAX(ID) as IDNo FROM " & sTableName)
        sCourtDatesID = rstCourtDatesID.Fields("IDNo").Value
        rstCourtDatesID.Close
        
        'return courtdates table ID number
        [Forms]![NewMainMenu]![ProcessJobSubformNMM].[Form]![JobNumberField].Value = sCourtDatesID
        addRSet = sCourtDatesID
        Exit Function
    
    Case "Customers"
        'add new appearances to tempcustomers
        'add new entry to customers from TempCustomers?
        Set oExcelApp = CreateObject("Excel.Application")
        
        Set oExcelMacroWB = oExcelApp.Application.Workbooks.Open(cJob.DocPath.OrderFormCustomersXLSM)
        oExcelMacroWB.Application.DisplayAlerts = False
        oExcelMacroWB.Application.Visible = False
        oExcelMacroWB.SaveAs Replace(cJob.DocPath.OrderFormCustomersXLSM, ".xlsm", ".xls"), FileFormat:=xlWorkbookDefault
        oExcelMacroWB.Close True
        Set oExcelMacroWB = Nothing
        
        Set oExcelWB = oExcelApp.Application.Workbooks.Open(FileName:=sFullPathXLS, Local:=True)
        oExcelWB.Application.DisplayAlerts = False
        oExcelWB.Application.Visible = False
        oExcelWB.SaveAs Replace(cJob.DocPath.OrderFormCustomersXLSX, ".xls", ".csv"), FileFormat:=6
        
        oExcelWB.Close True
        Set oExcelWB = Nothing
                
                
        'get appearances into database
        DoCmd.TransferText TransferType:=acImportDelim, TableName:="TempCustomers", _
                           FileName:=cJob.DocPath.OrderFormCustomersCSV, HasFieldNames:=True
                        
        Call fCheckTempCustomersCustomers
        Call fCheckTempCasesCases
        
        sTempJobSQL = "SELECT * FROM TempCustomers;"
        Set rstTempJob = CurrentDb.OpenRecordset(sTempJobSQL)
            
        sCurrentJobSQL = "SELECT * FROM CourtDates WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
        Set rstCurrentJob = CurrentDb.OpenRecordset(sCurrentJobSQL)
        
        'place customer ID # in correct job number in courtdates
        rstTempJob.MoveFirst
        sOrderingID = rstTempJob.Fields("AppID").Value
        
        If IsNull(rstCurrentJob!OrderingID) Then
            CurrentDb.Execute "UPDATE CourtDates SET OrderingID = " & sOrderingID & " WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
            rstCurrentJob.Close
        End If
        
        rstTempJob.MoveNext
        
        'import appearancesId from tempcustomers into courtdates
        sTempCustomersSQL = "SELECT * FROM TempCustomers;"
        sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
        sCurrentJobSQL = "SELECT * FROM CourtDates WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
        
        Set rstTempJob = CurrentDb.OpenRecordset(sTempCustomersSQL)
        'Set rstCurrentJob = db.OpenRecordset(sCurrentJobSQL)
        
        x = 1
        
        rstTempJob.MoveFirst
        
        Do Until rstTempJob.EOF
        
            sCurrentTempApp = rstTempJob.Fields("AppID").Value
            sAppNumber = "App" & x
            
            If Not rstTempJob.EOF Or sCurrentTempApp <> "" Or Not IsNull(sCurrentTempApp) Then
                'currentdb.Execute "UPDATE CourtDates SET " & sAppNumber & " = " & sCurrentTempApp & " WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
                
                'Set rstTempCDs = currentdb.OpenRecordset("SELECT * FROM CourtDates WHERE [ID] = " & sCourtDatesID & ";") '
                'rstCurrentJob.Edit
                
                Select Case sAppNumber
                Case "App1", "App2", "App3", "App4", "App5", "App6"
                    CurrentDb.Execute "UPDATE " & sTableName & " SET " & sAppNumber & " = " & sCurrentTempApp & " WHERE [CourtDates].[ID] = " & sCourtDatesID & ";"
                Case "App7"
                    Exit Do
                        
                End Select
                rstTempJob.MoveNext
            Else:
                Exit Do
            End If
            x = x + 1
        Loop

        'rstCurrentJob.Close
        rstTempJob.Close
        Set rstTempJob = Nothing
        'Set rstCurrentJob = Nothing

        addRSet = sCourtDatesID
        

    End Select

    rstTableName.Close
    rstTempCourtDates.Close
    Set rstTableName = Nothing
    Set rstTempCourtDates = Nothing

End Function

Public Sub editRSet(sTableName As String, sFieldName As String, sValue As String)
    Dim rstTableName As DAO.Recordset
    Dim db As Database
    Set db = CurrentDb
    Set rstTableName = db.OpenRecordset(sTableName)
    rstTableName.Edit
    rstTableName.Fields(sFieldName).Value = sValue
    rstTableName.Update
    rstTableName.Close
    Set rstTableName = Nothing
End Sub

Public Sub openRSet(sTableName As String)
    'open
    Dim db As Database
    Dim rstTableName As DAO.Recordset
    Set db = CurrentDb
    '@Ignore AssignmentNotUsed
    Set rstTableName = db.OpenRecordset(sTableName)
    
End Sub

Public Sub closeRSet(rstTableName As DAO.Recordset)

    'close
    rstTableName.Update
    rstTableName.Close
    Set rstTableName = Nothing

End Sub

Public Sub deleteRSet()
    'delete

End Sub


