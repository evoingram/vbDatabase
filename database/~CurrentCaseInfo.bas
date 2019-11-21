Attribute VB_Name = "~CurrentCaseInfo"
Option Compare Database

'============================================================================
' Name        : GetOrderingAttorneyInfo
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfGetOrderingAttorneyInfo
' Description : refreshes ordering attorney info for transcript
'============================================================================

'============================================================================
' Name        : pfClearGlobals
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfClearGlobals
' Description : clears all global variables
'============================================================================

    
'============================================================================
' Name        : pfCurrentCaseInfo
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfCurrentCaseInfo
' Description : refreshes global variables for current transcript
'============================================================================

'============================================================================
' Name        : fPPGenerateJSONInfo
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fPPGenerateJSONInfo
' Description : get info for invoice
'============================================================================

Public sParty1 As String, sCompany As String, sParty2 As String, sCourtDatesID As String, sInvoiceNumber As String
Public sParty1Name As String, sParty2Name As String, sInvoiceNo As String, sEmail As String, sDescription As String
Public sSubtotal As String, sInvoiceDate As String, sInvoiceTime As String, sPaymentTerms As String, sNote As String
Public sTerms As String, sMinimumAmount As String, vmMemo As String, vlURL As String, sTemplateID As String
Public sLine1 As String, sCity As String, sState As String, sZIP As String, sQuantity As String, sValue As String
Public sInventoryRateCode As String, sIRC As String, sCaseNumber2 As String
Public sActualQuantity As String, sJurisdiction As String, sTurnaroundTime As String, sCaseNumber1 As String
Public sCustomerID As String, sAudioLength As String, sEstimatedPageCount As String, sStatusesID As String
Public dDueDate As Date, dExpectedAdvanceDate As Date, dExpectedRebateDate As Date, sPaymentSum As String
Public sFactoringApproved As String, sBrandingTheme As String, sFinalPrice As String
Public sClientTranscriptName As String, sCurrentTranscriptName As String
Public sBalanceDue As String, sFactoringCost As String, svURL As String, sLinkToCSV As String
Public sFirstName As String, sLastName As String, dHearingDate As Date, sMrMs As String
Public sName As String, sAddress1 As String, sAddress2 As String, sNotes As String
Public sTime As String, sTime1 As String, sCasesID As String, vCasesID As String
Public vStatus As String, sUnitPrice As String
Public sHearingLocation As String, sStartTime As String, sEndTime As String

Public lngNumOfHrs As Long, lngNumOfMins As Long, lngNumOfSecsRem As Long, lngNumOfSecs As Long
Public lngNumOfHrs1 As Long, lngNumOfMins1 As Long, lngNumOfSecsRem1 As Long, lngNumOfSecs1 As Long
Public i As Long

'Public SharedRecognizer As SpSharedRecognizer
'Public theRecognizers As ISpeechObjectTokens

Public oWordApp As Word.Document, oWordDoc As Word.Application

Public Const qnTRCourtQ As String = "TR-Court-Q"
Public Const qnShippingOptionsQ As String = "ShippingOptionsQ"
Public Const qnViewJobFormAppearancesQ As String = "ViewJobFormAppearancesQ"
Public Const qnTRCourtUnionAppAddrQ As String = "TR-Court-Union-AppAddr"
Public Const qnOrderingAttorneyInfo As String = "OrderingAttorneyInfo"

Public Const sCompanyEmail As String = "inquiries@aquoco.co"
Public Const sCompanyFirstName As String = "Erica"
Public Const sCompanyLastName As String = "Ingram"
Public Const sCompanyName As String = "A Quo Co."
Public Const sPCountryCode As String = "001"
Public Const sCompanyNationalNumber As String = "2064785028"
Public Const sCompanyAddress As String = "320 West Republican Street Suite 207"
Public Const sCompanyCity As String = "Seattle"
Public Const sCompanyState As String = "WA"
Public Const sCompanyZIP As String = "98119"
Public Const sZCountryCode As String = "US"
Public Const slURL As String = "\\hubcloud\evoingram\Administration\Marketing\LOGO-5inch-by-1.22inches.jpg"
Public Const sPPTemplateName As String = "Amount only"
Public Const sTermDays As String = "30"
Public lAssigneeID As Long, sDueDate As String, bStarred As String, bCompleted As String, sTitle As String, sWLListID As String


Public Function pfCurrentCaseInfo()
    
'============================================================================
' Name        : pfCurrentCaseInfo
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfCurrentCaseInfo
' Description : refreshes global variables for current transcript
'============================================================================

Dim rstTRCourtUnionAA As DAO.Recordset
Dim db As Database
Dim qdf As QueryDef

Const qnTRCourtQ As String = "TR-Court-Q"
Const qnShippingOptionsQ As String = "ShippingOptionsQ"
Const qnViewJobFormAppearancesQ As String = "ViewJobFormAppearancesQ"
Const qnTRCourtUnionAppAddrQ As String = "TR-Court-Union-AppAddr"
Const qnOrderingAttorneyInfo As String = "OrderingAttorneyInfo"

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
Set db = CurrentDb

            
            Set qdf = db.QueryDefs("TR-Court-Union-AppAddr")
            qdf.Parameters(0) = sCourtDatesID
            Set rstTRCourtUnionAA = qdf.OpenRecordset
            
            If Not rstTRCourtUnionAA.EOF Then
            If Not rstTRCourtUnionAA.Fields("TR-AppAddrQ.ID").Value Like "" Then
            sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
            
            sCustomerID = rstTRCourtUnionAA.Fields("TR-AppAddrQ.ID").Value
            sTurnaroundTime = rstTRCourtUnionAA.Fields("TurnaroundTimesCD").Value
            sEstimatedPageCount = rstTRCourtUnionAA.Fields("EstimatedPageCount").Value
            
            sHearingLocation = rstTRCourtUnionAA.Fields("Location").Value
            sStartTime = rstTRCourtUnionAA.Fields("HearingStartTime").Value
            sEndTime = rstTRCourtUnionAA.Fields("HearingEndTime").Value
            sAudioLength = rstTRCourtUnionAA.Fields("AudioLength").Value
            dHearingDate = rstTRCourtUnionAA.Fields("HearingDate").Value
            sInvoiceNumber = rstTRCourtUnionAA.Fields("InvoiceNo").Value
            sCaseNumber1 = rstTRCourtUnionAA.Fields("CaseNumber1").Value
            sCaseNumber2 = rstTRCourtUnionAA.Fields("CaseNumber2").Value
            sParty1 = rstTRCourtUnionAA.Fields("Party1").Value
            sParty1Name = rstTRCourtUnionAA.Fields("Party1Name").Value
            sCompany = rstTRCourtUnionAA.Fields("Company").Value
            sJurisdiction = rstTRCourtUnionAA.Fields("Jurisdiction").Value
            sParty2 = rstTRCourtUnionAA.Fields("Party2").Value
            sParty2Name = rstTRCourtUnionAA.Fields("Party2Name").Value
            sActualQuantity = rstTRCourtUnionAA.Fields("ActualQuantity").Value
            sFactoringApproved = rstTRCourtUnionAA.Fields("FactoringApproved").Value
            
            If Not rstTRCourtUnionAA.Fields("ExpectedAdvanceDate").Value = "" Then
                dExpectedAdvanceDate = Format(rstTRCourtUnionAA.Fields("ExpectedAdvanceDate").Value, "Short Date") 'Format(rstTRCourtUnionAA.fields("ExpectedAdvanceDate").Value, "Short Date")
            End If
            
            If Not rstTRCourtUnionAA.Fields("ExpectedRebateDate").Value = "" Then
                dExpectedRebateDate = Format(rstTRCourtUnionAA.Fields("ExpectedRebateDate").Value, "Short Date")
            End If
            
            dDueDate = Format(rstTRCourtUnionAA.Fields("DueDate").Value, "Short Date")
            
            If Not rstTRCourtUnionAA.Fields("Subtotal").Value = "" Then
                sSubtotal = rstTRCourtUnionAA.Fields("Subtotal").Value
            Else
                sSubtotal = "0"
            End If
            
            sPaymentSum = Nz(rstTRCourtUnionAA.Fields("PaymentSum").Value, 0)
            sFinalPrice = rstTRCourtUnionAA.Fields("FinalPrice").Value
            sFactoringCost = rstTRCourtUnionAA.Fields("FactoringCost").Value
            sBalanceDue = sFinalPrice - sPaymentSum
            
        
    End If
    
    rstTRCourtUnionAA.Close
    Set qdf = Nothing
    db.Close
    
End If

End Function
    
    
Public Function pfGetOrderingAttorneyInfo()
    
'============================================================================
' Name        : GetOrderingAttorneyInfo
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfGetOrderingAttorneyInfo
' Description : refreshes ordering attorney info for transcript
'============================================================================

Dim rstOrderingAttyInfo As DAO.Recordset
Dim db As Database
Dim qdf As QueryDef

Const qnTRCourtQ As String = "TR-Court-Q"
Const qnShippingOptionsQ As String = "ShippingOptionsQ"
Const qnViewJobFormAppearancesQ As String = "ViewJobFormAppearances"
Const qnTRCourtUnionAppAddrQ As String = "TR-Court-Union-AppAddr"
Const qnOrderingAttorneyInfo As String = "OrderingAttorneyInfo"
'Const sCourtDatesID As String = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
'GetProperty(cCurrentCI, sCourtDatesID)

        
sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

Set db = CurrentDb
Set qdf = db.QueryDefs("OrderingAttorneyInfo")
qdf.Parameters(0) = sCourtDatesID

Set rstOrderingAttyInfo = qdf.OpenRecordset

If Not rstOrderingAttyInfo.EOF Then

    If Not rstOrderingAttyInfo.Fields("CourtDatesID").Value Like "" Then
                
        
            sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
            sFactoringApproved = rstOrderingAttyInfo.Fields("FactoringApproved").Value
            
            sFirstName = rstOrderingAttyInfo("FirstName").Value
            sLastName = rstOrderingAttyInfo("LastName").Value
            sName = rstOrderingAttyInfo("FirstName").Value & " " & rstOrderingAttyInfo("LastName").Value
            sAddress1 = rstOrderingAttyInfo("Company").Value
            sAddress2 = rstOrderingAttyInfo("Address").Value
            sLine1 = rstOrderingAttyInfo("Address").Value
            sCity = rstOrderingAttyInfo("City").Value
            sState = rstOrderingAttyInfo("State").Value
            sZIP = rstOrderingAttyInfo("ZIP").Value
            sQuantity = rstOrderingAttyInfo("OAIQuantity").Value
            sSubtotal = rstOrderingAttyInfo("OAISubtotal").Value
            sUnitPrice = rstOrderingAttyInfo("OAIUnitPrice").Value
            
            
            If Not rstOrderingAttyInfo.Fields("Notes").Value Like "" Then
                sEmail = rstOrderingAttyInfo("Notes").Value
                sNotes = rstOrderingAttyInfo("Notes").Value
            End If
            
            If Not rstOrderingAttyInfo.Fields("OAIInvoiceNo").Value Like "" Then
                sInvoiceNumber = rstOrderingAttyInfo.Fields("OAIInvoiceNo").Value
            End If
            
            If Not rstOrderingAttyInfo.Fields("CustomersID").Value Like "" Then
                vOrderingID = rstOrderingAttyInfo.Fields("CustomersID").Value
            End If
            
            sCompany = rstOrderingAttyInfo.Fields("Company").Value
            sMrMs = rstOrderingAttyInfo.Fields("MrMs").Value
            sLastName = rstOrderingAttyInfo.Fields("LastName").Value
            sFirstName = rstOrderingAttyInfo.Fields("FirstName").Value
            
    End If


rstOrderingAttyInfo.Close
Set qdf = Nothing

Set rstOrderingAttyInfo = CurrentDb.OpenRecordset("SELECT Rate FROM UnitPrice WHERE ID=" & sUnitPrice & ";")
sUnitPrice = rstOrderingAttyInfo("Rate").Value

rstOrderingAttyInfo.Close

db.Close

End If
    
End Function


  
Public Function pfSecondCurrentCaseInfo()

Dim rstTRCourtUnionAA As DAO.Recordset
Dim db As Database
Dim qdf As QueryDef
Dim sCourtDatesID As String

Const qnTRCourtQ As String = "TR-Court-Q"
Const qnShippingOptionsQ As String = "ShippingOptionsQ"
Const qnViewJobFormAppearancesQ As String = "ViewJobFormAppearancesQ"
Const qnTRCourtUnionAppAddrQ As String = "TR-Court-Union-AppAddr"
Const qnOrderingAttorneyInfo As String = "OrderingAttorneyInfo"

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

Set db = CurrentDb
Set qdf = db.QueryDefs("TR-Court-Union-AppAddr")
qdf.Parameters(0) = sCourtDatesID
Set rstTRCourtUnionAA = qdf.OpenRecordset

If Not rstTRCourtUnionAA.EOF Then

    If Not rstTRCourtUnionAA.Fields("TR-AppAddrQ.ID").Value Like "" Then
    
            sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
            sCustomerID = rstTRCourtUnionAA.Fields("TR-AppAddrQ.ID").Value
            sTurnaroundTime = rstTRCourtUnionAA.Fields("TurnaroundTimesCD").Value
            sEstimatedPageCount = rstTRCourtUnionAA.Fields("EstimatedPageCount").Value
            sAudioLength = rstTRCourtUnionAA.Fields("AudioLength").Value
            dHearingDate = rstTRCourtUnionAA.Fields("HearingDate").Value
            sInvoiceNumber = rstTRCourtUnionAA.Fields("InvoiceNo").Value
            sCaseNumber1 = rstTRCourtUnionAA.Fields("CaseNumber1").Value
            sCaseNumber2 = rstTRCourtUnionAA.Fields("CaseNumber2").Value
            sParty1 = rstTRCourtUnionAA.Fields("Party1").Value
            sParty1Name = rstTRCourtUnionAA.Fields("Party1Name").Value
            sCompany = rstTRCourtUnionAA.Fields("Company").Value
            sJurisdiction = rstTRCourtUnionAA.Fields("Jurisdiction").Value
            sParty2 = rstTRCourtUnionAA.Fields("Party2").Value
            sParty2Name = rstTRCourtUnionAA.Fields("Party2Name").Value
            sActualQuantity = rstTRCourtUnionAA.Fields("ActualQuantity").Value
            
            If Not rstTRCourtUnionAA.Fields("ExpectedAdvanceDate").Value = "" Then
                dExpectedAdvanceDate = Format(rstTRCourtUnionAA.Fields("ExpectedAdvanceDate").Value, "Short Date") 'Format(rstTRCourtUnionAA.fields("ExpectedAdvanceDate").Value, "Short Date")
            End If
            
            If Not rstTRCourtUnionAA.Fields("ExpectedRebateDate").Value = "" Then
                dExpectedRebateDate = Format(rstTRCourtUnionAA.Fields("ExpectedRebateDate").Value, "Short Date")
            End If
            
            dDueDate = Format(rstTRCourtUnionAA.Fields("DueDate").Value, "Short Date")
            sSubtotal = rstTRCourtUnionAA.Fields("Subtotal").Value
            sPaymentSum = Nz(rstTRCourtUnionAA.Fields("PaymentSum").Value, 0)
            sFinalPrice = rstTRCourtUnionAA.Fields("FinalPrice").Value
            sFactoringCost = rstTRCourtUnionAA.Fields("FactoringCost").Value
            sBalanceDue = sFinalPrice - sPaymentSum
            
        
    End If
    
    rstTRCourtUnionAA.Close
    Set qdf = Nothing
    db.Close
    
End If

End Function


Public Function pfGetCaseInfoQDFRecordset()


Dim rs1 As DAO.Recordset
Dim db As Database
Dim qdf As QueryDef

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

Set db = CurrentDb
Set qdf = db.QueryDefs("TR-Court-Union-AppAddr")
qdf.Parameters(0) = sCourtDatesID
Set rs1 = qdf.OpenRecordset

If Not rs1.EOF Then

        
        If Not rs1.Fields("TR-AppAddrQ.ID").Value Like "" Then
            sCustomerID = rs1.Fields("TR-AppAddrQ.ID").Value
        End If
        sTurnaroundTime = rs1.Fields("TurnaroundTimesCD").Value
        sEstimatedPageCount = rs1.Fields("EstimatedPageCount").Value
        sAudioLength = rs1.Fields("AudioLength").Value
        dHearingDate = rs1.Fields("HearingDate").Value
        
        If Not rs1.Fields("InvoiceNo").Value Like "" Then
            sInvoiceNumber = rs1.Fields("InvoiceNo").Value
        End If
        
        sCaseNumber1 = rs1.Fields("CaseNumber1").Value
        sCaseNumber2 = rs1.Fields("CaseNumber2").Value
        sParty1 = rs1.Fields("Party1").Value
        sParty1Name = rs1.Fields("Party1Name").Value
        sCompany = rs1.Fields("Company").Value
        sJurisdiction = rs1.Fields("Jurisdiction").Value
        sParty2 = rs1.Fields("Party2").Value
        sParty2Name = rs1.Fields("Party2Name").Value
        sActualQuantity = rs1.Fields("ActualQuantity").Value
        
        If Not rs1.Fields("ExpectedAdvanceDate").Value = "" Then
            dExpectedAdvanceDate = Format(rs1.Fields("ExpectedAdvanceDate").Value, "Short Date") 'Format(rs1.fields("ExpectedAdvanceDate").Value, "Short Date")
        End If
        
        If Not rs1.Fields("ExpectedRebateDate").Value = "" Then
            dExpectedRebateDate = Format(rs1.Fields("ExpectedRebateDate").Value, "Short Date")
        End If
        
        dDueDate = Format(rs1.Fields("DueDate").Value, "Short Date")
        sSubtotal = rs1.Fields("Subtotal").Value
        sPaymentSum = Nz(rs1.Fields("PaymentSum").Value, 0)
        sFinalPrice = rs1.Fields("FinalPrice").Value
        sFactoringCost = rs1.Fields("FactoringCost").Value
        sBalanceDue = sFinalPrice - sPaymentSum
    
End If

rs1.Close
Set qdf = Nothing
db.Close

End Function





Function fPPGenerateJSONInfo()
'============================================================================
' Name        : fPPGenerateJSONInfo
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fPPGenerateJSONInfo
' Description : get info for invoice
'============================================================================
        
Dim rstTRQPlusCases As DAO.Recordset
Dim db As Database
Dim qdf As QueryDef

Const sCompanyEmail As String = "inquiries@aquoco.co"
Const sCompanyFirstName As String = "Erica"
Const sCompanyLastName As String = "Ingram"
Const sCompanyName As String = "A Quo Co."
Const sPCountryCode As String = "001"
Const sCompanyNationalNumber As String = "2064785028"
Const sCompanyAddress As String = "320 West Republican Street Suite 207"
Const sCompanyCity As String = "Seattle"
Const sCompanyState As String = "WA"
Const sCompanyZIP As String = "98119"
Const sZCountryCode As String = "US"
Const slURL As String = "\\hubcloud\evoingram\Administration\Marketing\LOGO-5inch-by-1.22inches.jpg"
Const sPPTemplateName As String = "Amount only"
Const sTermDays As String = "30"

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
Call pfGetCaseInfoQDFRecordset

Set db = CurrentDb
Set qdf = db.QueryDefs("TRInvoiQPlusCases")
qdf.Parameters(0) = sCourtDatesID
Set rstTRQPlusCases = qdf.OpenRecordset

If Not rstTRQPlusCases.EOF Then
    
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

    
    sInvoiceNumber = rstTRQPlusCases.Fields("TRInv.GetInvoiceNoFromCDID.InvoiceNo").Value
    vInvoiceID = rstTRQPlusCases.Fields("TRInv.PPID").Value '"INV2-C8EE-ZVC5-5U36-MF27" 'INV2-K8L5-ML2R-2GLL-7KW6 '
    'vInvoiceID = .vInvoiceID
    sSubtotal = rstTRQPlusCases.Fields("TRInv.UnitPrice").Value
    sInvoiceDate = (Format((Date + 28), "yyyy-mm-dd")) & " PST"
    sInvoiceTime = (Format(Now(), "hh:mm:ss"))
    sMinimumAmount = "1" 'rstTRQPlusCases.Fields("").value
    sQuantity = rstTRQPlusCases.Fields("TRInv.ActualQuantity").Value
    sValue = rstTRQPlusCases.Fields("TRInv.UnitPrice").Value
    sBrandingTheme = rstTRQPlusCases.Fields("BrandingTheme").Value

    dDueDate = Format(rstTRQPlusCases.Fields("TRInvoiceCasesQ.DueDate").Value, "Short Date")
    sPaymentSum = Nz(rstTRQPlusCases.Fields("TRInv.PaymentSum").Value, 0)
    sFinalPrice = rstTRQPlusCases.Fields("TRInv.FinalPrice").Value
    sFactoringCost = rstTRQPlusCases.Fields("TRInv.FactoringCost").Value
    sBalanceDue = sFinalPrice - sPaymentSum
    sIRC = rstTRQPlusCases.Fields("TRInv.InventoryRateCode").Value
    
    sDescription = "Job No.:  " & sCourtDatesID & "  |  " & _
    "Invoice No.:  " & sInvoiceNumber & "\n" & _
    sParty1 & " v " & sParty2 & "\n" & _
    "Case Nos.:  " & sCaseNumber1 & " " & sCaseNumber2 & "\n" & _
    "Hearing Date:  " & dHearingDate & "\n" & _
    "Approx. " & sAudioLength & " minutes" & "  |  " & _
    "Turnaround Time:  " & sTurnaroundTime & " calendar days"
    
    If sBrandingTheme = "1" Then 'WRTS NC Factoring
        sPaymentTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
            "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
            "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
            "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/ServiceA.html." 'rstTRQPlusCases.Fields("").value
        
        sNote = "Your transcript is attached to this invoice.  We will upload this transcript to our repository for your 24/7 access and " & _
        "mail out and/or file as appropriate.  Thank you for your business." 'rstTRQPlusCases.Fields("").value
        
        sTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
            "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
            "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
            "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/ServiceA.html." 'rstTRQPlusCases.Fields("").value
            
        vmMemo = sCourtDatesID & " " & sInvoiceNo
    
    ElseIf sBrandingTheme = "2" Then  'WRTS NC 100 Deposit
        sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
            "The turnaround as described above will begin once this invoice is paid.  Full terms of service listed at https://www.aquoco.co/ServiceA.html."
        'rstTRQPlusCases.Fields("")value
        
        sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
        "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
        "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
        'rstTRQPlusCases.Fields("")value
        
        sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
        
        vmMemo = sCourtDatesID & " " & sInvoiceNo
    
    ElseIf sBrandingTheme = "3" Then  'WRTS C 50 Deposit Filed Non-BK
        sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
            "The turnaround as described above will begin once this invoice is paid."
        'rstTRQPlusCases.Fields("")value
        
        sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
        "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
        "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
        'rstTRQPlusCases.Fields("")value
        
        sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
        
        vmMemo = sCourtDatesID & " " & sInvoiceNo
    
    ElseIf sBrandingTheme = "4" Then  'WRTS C 50 Deposit Filed BK
    
        sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
            "The turnaround as described above will begin once this invoice is paid."
        'rstTRQPlusCases.Fields("")value
        
        sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
        "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
        "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
        'rstTRQPlusCases.Fields("")value
        
        sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
        
        vmMemo = sCourtDatesID & " " & sInvoiceNo
    
    ElseIf sBrandingTheme = "5" Then  'WRTS C 50 Deposit Not Filed
        sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
            "The turnaround as described above will begin once this invoice is paid."
        'rstTRQPlusCases.Fields("")value
        
        sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
        "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
        "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
        'rstTRQPlusCases.Fields("")value
        
        sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
        
        vmMemo = sCourtDatesID & " " & sInvoiceNo
    
    ElseIf sBrandingTheme = "6" Then  'WRTS C Factoring Filed
        sPaymentTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
            "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
            "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
            "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/ServiceA.html." 'rstTRQPlusCases.Fields("")value
        
        sNote = "Your transcript is attached to this invoice.  We will upload this transcript to our repository for your 24/7 access and " & _
        "mail out and/or file as appropriate.  Thank you for your business." 'rstTRQPlusCases.Fields("")value
        
        sTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
            "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
            "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
            "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/ServiceA.html." 'rstTRQPlusCases.Fields("")value
            
        vmMemo = sCourtDatesID & " " & sInvoiceNo
    
    ElseIf sBrandingTheme = "7" Then  'WRTS C Factoring Not Filed
        sPaymentTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
            "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
            "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
            "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/ServiceA.html." 'rstTRQPlusCases.Fields("")value
        
        sNote = "Your transcript is attached to this invoice.  We will upload this transcript to our repository for your 24/7 access and " & _
        "mail out and/or file as appropriate.  Thank you for your business." 'rstTRQPlusCases.Fields("")value
        
        sTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
            "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
            "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
            "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/ServiceA.html." 'rstTRQPlusCases.Fields("")value
            
        vmMemo = sCourtDatesID & " " & sInvoiceNo
    
    ElseIf sBrandingTheme = "8" Then  'WRTS C 100 Deposit Filed
        sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
            "The turnaround as described above will begin once this invoice is paid."
        'rstTRQPlusCases.Fields("")value
        
        sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
        "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
        "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
        'rstTRQPlusCases.Fields("")value
        
        sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
        
        vmMemo = sCourtDatesID & " " & sInvoiceNo
    
    ElseIf sBrandingTheme = "9" Then  'WRTS C 100 Deposit Not Filed
        sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
            "The turnaround as described above will begin once this invoice is paid."
        'rstTRQPlusCases.Fields("")value
        
        sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
        "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
        "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
        'rstTRQPlusCases.Fields("")value
        
        sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
        
        vmMemo = sCourtDatesID & " " & sInvoiceNo
    
    ElseIf sBrandingTheme = "10" Then  'WRTS JJ Factoring
        sPaymentTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
            "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
            "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
            "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/ServiceA.html." 'rstTRQPlusCases.Fields("")value
        
        sNote = "Your transcript is attached to this invoice.  We will upload this transcript to our repository for your 24/7 access and " & _
        "mail out and/or file as appropriate.  Thank you for your business." 'rstTRQPlusCases.Fields("")value
        
        sTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
            "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
            "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
            "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/ServiceA.html." 'rstTRQPlusCases.Fields("")value
            
        vmMemo = sCourtDatesID & " " & sInvoiceNo
    
    ElseIf sBrandingTheme = "11" Then  'WRTS Tabula Not Factored/Filed
        sPaymentTerms = sCourtDatesID & " " & sInvoiceNo 'rstTRQPlusCases.Fields("")value
        sNote = "Thank you for your business." 'rstTRQPlusCases.Fields("")value
        sTerms = "Thank you for your business." 'rstTRQPlusCases.Fields("")value
        
        vmMemo = sCourtDatesID & " " & sInvoiceNo
    
    ElseIf sBrandingTheme = "12" Then  'WRTS AMOR Factoring
        sPaymentTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
            "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
            "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
            "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/ServiceA.html." 'rstTRQPlusCases.Fields("")value
        
        sNote = "Your transcript is attached to this invoice.  We will upload this transcript to our repository for your 24/7 access and " & _
        "mail out and/or file as appropriate.  Thank you for your business." 'rstTRQPlusCases.Fields("")value
        
        sTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
            "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
            "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
            "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/ServiceA.html." 'rstTRQPlusCases.Fields("")value
            
        vmMemo = sCourtDatesID & " " & sInvoiceNo
    
    End If
            
                
    End If
    
    rstTRQPlusCases.Close
    Set qdf = Nothing
    db.Close
  
End Function

Function pfClearGlobals()

'============================================================================
' Name        : pfClearGlobals
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfClearGlobals
' Description : clears all global variables
'============================================================================
sParty1 = ""
sCompany = ""
sParty2 = ""
sCourtDatesID = ""
sInvoiceNumber = ""
sParty1Name = ""
sParty2Name = ""
sInvoiceNo = ""
sEmail = ""
sDescription = ""
sSubtotal = ""
sInvoiceDate = ""
sInvoiceTime = ""
sPaymentTerms = ""
sNote = ""
sTerms = ""
sMinimumAmount = ""
vmMemo = ""
vlURL = ""
sTemplateID = ""
sLine1 = ""
sCity = ""
sState = ""
sZIP = ""
sQuantity = ""
sValue = ""
sInventoryRateCode = ""
sIRC = ""
sActualQuantity = ""
sJurisdiction = ""
sTurnaroundTime = ""
sCaseNumber1 = ""
sCaseNumber2 = ""
Set oWordApp = Nothing
Set oWordDoc = Nothing
sCustomerID = ""
sAudioLength = ""
sEstimatedPageCount = ""
sStatusesID = ""
sFinalPrice = ""
sPaymentSum = ""
sBalanceDue = ""
sFactoringCost = ""
svURL = ""
sLinkToCSV = ""
sFactoringApproved = ""
sBrandingTheme = ""
sFirstName = ""
sLastName = ""
sMrMs = ""
sName = ""
sAddress1 = ""
sAddress2 = ""
sNotes = ""
HyperlinkString = ""
rtfstringbody = ""
sTime = ""
sTime1 = ""
sClientTranscriptName = ""
sCurrentTranscriptName = ""
sLocation = ""
sStartTime = ""
sEndTime = ""

End Function
