Attribute VB_Name = "~CurrentCaseInfo"
'@Folder("Database.General.Modules")
Option Compare Database
Option Explicit
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

Public sParty1 As String
Public sCompany As String
Public sParty2 As String
Public sCourtDatesID As String
Public sInvoiceNumber As String
Public sParty1Name As String
Public sParty2Name As String
Public sInvoiceNo As String
Public sEmail As String
Public sDescription As String
Public sSubtotal As String
Public sInvoiceDate As String
Public dInvoiceDate As Date
Public sInvoiceTime As String
Public sPaymentTerms As String
Public sNote As String
Public sTerms As String
Public sMinimumAmount As String
Public vmMemo As String
Public vlURL As String
Public sTemplateID As String
Public sLine1 As String
Public sCity As String
Public sState As String
Public sZIP As String
Public sQuantity As String
Public sValue As String
Public sInventoryRateCode As String
Public sIRC As String
Public sCaseNumber2 As String
Public sActualQuantity As String
Public sJurisdiction As String
Public sTurnaroundTime As String
Public sCaseNumber1 As String
Public sCustomerID As String
Public sAudioLength As String
Public sEstimatedPageCount As String
Public sStatusesID As String
Public dDueDate As Date
Public dExpectedAdvanceDate As Date
Public dExpectedRebateDate As Date
Public dExpectedBalanceDate As Date
Public sPaymentSum As String
Public sFactoringApproved As String
Public sBrandingTheme As String
Public sFinalPrice As String
Public sClientTranscriptName As String
Public sCurrentTranscriptName As String
Public sBalanceDue As String
Public sFactoringCost As String
Public svURL As String
Public sLinkToCSV As String
Public sFirstName As String
Public sLastName As String
Public dHearingDate As Date
Public sMrMs As String
Public sName As String
Public sAddress1 As String
Public sAddress2 As String
Public sNotes As String
Public sTime As String
Public sTime1 As String
Public sCasesID As String
Public vCasesID As String
Public vStatus As String
Public sUnitPrice As String
Public sOrderingID As String
Public sInvoiceID As String
Public sHearingLocation As String
Public sStartTime As String
Public sEndTime As String
Public HyperlinkString As String
Public rtfStringBody As String
Public sLocation As String
Public sPPID As String
Public lngNumOfHrs As Long
Public lngNumOfMins As Long
Public lngNumOfSecsRem As Long
Public lngNumOfSecs As Long
Public lngNumOfHrs1 As Long
Public lngNumOfMins1 As Long
Public lngNumOfSecsRem1 As Long
Public lngNumOfSecs1 As Long
Public i As Long

'Public SharedRecognizer As SpSharedRecognizer
'Public theRecognizers As ISpeechObjectTokens

Public oWordApp As Word.Document
Public oWordDoc As Word.Application

Public Const qnTRCourtQ As String = "TR-Court-Q"
Public Const qnShippingOptionsQ As String = "ShippingOptionsQ"
Public Const qnViewJobFormAppearancesQ As String = "ViewJobFormAppearancesQ"
Public Const qnTRCourtUnionAppAddrQ As String = "TR-Court-Union-AppAddr"
Public Const qnOrderingAttorneyInfo As String = "OrderingAttorneyInfo"
Public Const qnQInfobyInvNo As String = "QInfobyInvoiceNumber"
Public Const qTempShippingOptions As String = "TempShippingOptionsQ"

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

'TODO: CurrentCaseInfo module, why isn't this used?  Come back
'@Ignore ConstantNotUsed
Public Const slURL As String = "\\hubcloud\evoingram\Administration\Marketing\LOGO-5inch-by-1.22inches.jpg"
'@Ignore ConstantNotUsed
Public Const sPPTemplateName As String = "Amount only"
Public Const sTermDays As String = "30"
Public lAssigneeID As Long
Public sDueDate As String
Public bStarred As String
Public bCompleted As String
Public sTitle As String
Public sWLListID As String

Public Sub pfCurrentCaseInfo()
    
    '============================================================================
    ' Name        : pfCurrentCaseInfo
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfCurrentCaseInfo
    ' Description : refreshes global variables for current transcript
    '============================================================================

    Dim cJob As Job
    Set cJob = New Job

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID
    sTurnaroundTime = cJob.TurnaroundTime
    sCasesID = cJob.CaseID
    sEstimatedPageCount = cJob.EstimatedPageCount
    sHearingLocation = cJob.Location
    sStartTime = Format(cJob.HearingStartTime, "h:mm AM/PM")
    sEndTime = Format(cJob.HearingEndTime, "h:mm AM/PM")
    sAudioLength = cJob.AudioLength
    dHearingDate = Format(cJob.HearingDate, "mm-dd-yyyy")
    sInvoiceNumber = cJob.InvoiceNo
    sCaseNumber1 = cJob.CaseInfo.CaseNumber1
    sCaseNumber2 = cJob.CaseInfo.CaseNumber2
    sParty1 = cJob.CaseInfo.Party1
    sParty1Name = cJob.CaseInfo.Party1Name
    sCompany = cJob.App0.Company
    sJurisdiction = cJob.CaseInfo.Jurisdiction
    sParty2 = cJob.CaseInfo.Party2
    sParty2Name = cJob.CaseInfo.Party2Name
    sActualQuantity = cJob.ActualQuantity
    sFactoringApproved = cJob.App0.FactoringApproved
    dExpectedAdvanceDate = Format(cJob.ExpectedAdvanceDate, "mm-dd-yyyy")
    dExpectedRebateDate = Format(cJob.ExpectedRebateDate, "mm-dd-yyyy")
    dDueDate = Format(cJob.DueDate, "mm-dd-yyyy")
    sSubtotal = cJob.Subtotal
    sPaymentSum = Nz(cJob.PaymentSum, 0)
    sFinalPrice = cJob.FinalPrice
    sFactoringCost = cJob.FactoringCost
    sBrandingTheme = cJob.BrandingTheme
    sPPID = cJob.PPID
    sBalanceDue = sFinalPrice - sPaymentSum
    sInventoryRateCode = cJob.InventoryRateCode
    sIRC = cJob.InventoryRateCode

End Sub

Public Sub pfGetOrderingAttorneyInfo()
    
    '============================================================================
    ' Name        : GetOrderingAttorneyInfo
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfGetOrderingAttorneyInfo
    ' Description : refreshes ordering attorney info for transcript
    '============================================================================

    Dim rstOrderingAttyInfo As DAO.Recordset
    Dim cJob As Job
    Set cJob = New Job
        
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID
    sFactoringApproved = cJob.App0.FactoringApproved
    sFirstName = cJob.App0.FirstName
    sLastName = cJob.App0.LastName
    sName = sFirstName & " " & sLastName
    sAddress1 = cJob.App0.Company
    sAddress2 = cJob.App0.Address
    sLine1 = cJob.App0.Address
    sCity = cJob.App0.City
    sState = cJob.App0.State
    sZIP = cJob.App0.ZIP
    sQuantity = cJob.Quantity
    sSubtotal = cJob.Subtotal
    sUnitPrice = cJob.UnitPrice
    sEmail = cJob.App0.Notes
    sNotes = cJob.App0.Notes
    sInvoiceNumber = cJob.InvoiceNo
    sOrderingID = cJob.sApp0
    sCompany = cJob.App0.Company
    sMrMs = cJob.App0.MrMs
    Set rstOrderingAttyInfo = CurrentDb.OpenRecordset("SELECT Rate FROM UnitPrice WHERE ID=" & sUnitPrice & ";")
    sUnitPrice = rstOrderingAttyInfo("Rate").Value
    rstOrderingAttyInfo.Close
    Debug.Print sMrMs
    Debug.Print sCompany
    
End Sub

Public Sub fPPGenerateJSONInfo()
    '============================================================================
    ' Name        : fPPGenerateJSONInfo
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fPPGenerateJSONInfo
    ' Description : get info for invoice
    '============================================================================

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    Call pfCurrentCaseInfo
    sInvoiceID = sPPID                           '"INV2-C8EE-ZVC5-5U36-MF27" 'INV2-K8L5-ML2R-2GLL-7KW6 '
    sSubtotal = sUnitPrice
    sInvoiceDate = (Format((Date + 28), "yyyy-mm-dd")) & " PST"
    sInvoiceTime = (Format(Now(), "hh:mm:ss"))
    sMinimumAmount = "1"                         'rstTRQPlusCases.Fields("").value
    sValue = sUnitPrice
    sDescription = "Job No.:  " & sCourtDatesID & "  |  " & _
                   "Invoice No.:  " & sInvoiceNumber & "\n" & _
                   sParty1 & " v " & sParty2 & "\n" & _
                   "Case Nos.:  " & sCaseNumber1 & " " & sCaseNumber2 & "\n" & _
                   "Hearing Date:  " & dHearingDate & "\n" & _
                   "Approx. " & sAudioLength & " minutes" & "  |  " & _
                   "Turnaround Time:  " & sTurnaroundTime & " calendar days"

    If sBrandingTheme = "1" Then                 'WRTS NC Factoring
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

    ElseIf sBrandingTheme = "2" Then             'WRTS NC 100 Deposit
        sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
                        "The turnaround as described above will begin once this invoice is paid.  Full terms of service listed at https://www.aquoco.co/ServiceA.html."
        'rstTRQPlusCases.Fields("")value
    
        sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
                "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
                "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
        'rstTRQPlusCases.Fields("")value
    
        sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
    

    ElseIf sBrandingTheme = "3" Then             'WRTS C 50 Deposit Filed Non-BK
        sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
                        "The turnaround as described above will begin once this invoice is paid."
        'rstTRQPlusCases.Fields("")value
    
        sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
                "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
                "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
        'rstTRQPlusCases.Fields("")value
    
        sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
    
        vmMemo = sCourtDatesID & " " & sInvoiceNo

    ElseIf sBrandingTheme = "4" Then             'WRTS C 50 Deposit Filed BK

        sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
                        "The turnaround as described above will begin once this invoice is paid."
        'rstTRQPlusCases.Fields("")value
    
        sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
                "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
                "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
        'rstTRQPlusCases.Fields("")value
    
        sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
    

    ElseIf sBrandingTheme = "5" Then             'WRTS C 50 Deposit Not Filed
        sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
                        "The turnaround as described above will begin once this invoice is paid."
        'rstTRQPlusCases.Fields("")value
    
        sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
                "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
                "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
        'rstTRQPlusCases.Fields("")value
    
        sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
    

    ElseIf sBrandingTheme = "6" Then             'WRTS C Factoring Filed
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
        

    ElseIf sBrandingTheme = "7" Then             'WRTS C Factoring Not Filed
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
        

    ElseIf sBrandingTheme = "8" Then             'WRTS C 100 Deposit Filed
        sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
                        "The turnaround as described above will begin once this invoice is paid."
        'rstTRQPlusCases.Fields("")value
    
        sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
                "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
                "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
        'rstTRQPlusCases.Fields("")value
    
        sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
    
    ElseIf sBrandingTheme = "9" Then             'WRTS C 100 Deposit Not Filed
        sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
                        "The turnaround as described above will begin once this invoice is paid."
        'rstTRQPlusCases.Fields("")value
    
        sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
                "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
                "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
        'rstTRQPlusCases.Fields("")value
    
        sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
    

    ElseIf sBrandingTheme = "10" Then            'WRTS JJ Factoring
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
        

    ElseIf sBrandingTheme = "11" Then            'WRTS Tabula Not Factored/Filed
        sPaymentTerms = sCourtDatesID & " " & sInvoiceNo 'rstTRQPlusCases.Fields("")value
        sNote = "Thank you for your business."   'rstTRQPlusCases.Fields("")value
        sTerms = "Thank you for your business."  'rstTRQPlusCases.Fields("")value

    ElseIf sBrandingTheme = "12" Then            'WRTS AMOR Factoring
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
        
    End If
    vmMemo = sCourtDatesID & " " & sInvoiceNo

End Sub

Public Sub pfClearGlobals()

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
    rtfStringBody = ""
    sTime = ""
    sTime1 = ""
    sClientTranscriptName = ""
    sCurrentTranscriptName = ""
    sLocation = ""
    sStartTime = ""
    sEndTime = ""

End Sub

Public Sub pfCurrentCaseInfo1()
    'everything from this sub down to bottom of module can be deleted after known safe
    '============================================================================
    ' Name        : pfCurrentCaseInfo
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfCurrentCaseInfo
    ' Description : refreshes global variables for current transcript
    '============================================================================

    Dim rstTRCourtUnionAA As DAO.Recordset
    Dim qdf As QueryDef


    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    Set qdf = CurrentDb.QueryDefs(qnTRCourtUnionAppAddrQ)
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
    
    End If

End Sub

Public Sub pfGetOrderingAttorneyInfo1()
    
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
    'Const sCourtDatesID As String = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    'GetProperty(cCurrentCI, sCourtDatesID)

        
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

    Set db = CurrentDb
    Set qdf = db.QueryDefs(qnOrderingAttorneyInfo)
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
                sOrderingID = rstOrderingAttyInfo.Fields("CustomersID").Value
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
    
End Sub

Public Sub pfGetCaseInfoQDFRecordset()


    Dim rs1 As DAO.Recordset
    Dim db As Database
    Dim qdf As QueryDef

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

    Set db = CurrentDb
    Set qdf = db.QueryDefs(qnTRCourtUnionAppAddrQ)
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

End Sub

Public Sub fPPGenerateJSONInfo1()
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


    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    Call pfCurrentCaseInfo

    Set db = CurrentDb
    Set qdf = db.QueryDefs("TRInvoiQPlusCases")
    qdf.Parameters(0) = sCourtDatesID
    Set rstTRQPlusCases = qdf.OpenRecordset

    If Not rstTRQPlusCases.EOF Then
        sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    
        sInvoiceNumber = rstTRQPlusCases.Fields("TRInv.GetInvoiceNoFromCDID.InvoiceNo").Value
        sInvoiceID = rstTRQPlusCases.Fields("TRInv.PPID").Value '"INV2-C8EE-ZVC5-5U36-MF27" 'INV2-K8L5-ML2R-2GLL-7KW6 '
        'sInvoiceID = .sInvoiceID
        sSubtotal = rstTRQPlusCases.Fields("TRInv.UnitPrice").Value
        sInvoiceDate = (Format((Date + 28), "yyyy-mm-dd")) & " PST"
        sInvoiceTime = (Format(Now(), "hh:mm:ss"))
        sMinimumAmount = "1"                     'rstTRQPlusCases.Fields("").value
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
    
        If sBrandingTheme = "1" Then             'WRTS NC Factoring
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
    
        ElseIf sBrandingTheme = "2" Then         'WRTS NC 100 Deposit
            sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
                            "The turnaround as described above will begin once this invoice is paid.  Full terms of service listed at https://www.aquoco.co/ServiceA.html."
            'rstTRQPlusCases.Fields("")value
        
            sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
                    "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
                    "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
            'rstTRQPlusCases.Fields("")value
        
            sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
        
            vmMemo = sCourtDatesID & " " & sInvoiceNo
    
        ElseIf sBrandingTheme = "3" Then         'WRTS C 50 Deposit Filed Non-BK
            sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
                            "The turnaround as described above will begin once this invoice is paid."
            'rstTRQPlusCases.Fields("")value
        
            sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
                    "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
                    "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
            'rstTRQPlusCases.Fields("")value
        
            sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
        
            vmMemo = sCourtDatesID & " " & sInvoiceNo
    
        ElseIf sBrandingTheme = "4" Then         'WRTS C 50 Deposit Filed BK
    
            sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
                            "The turnaround as described above will begin once this invoice is paid."
            'rstTRQPlusCases.Fields("")value
        
            sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
                    "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
                    "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
            'rstTRQPlusCases.Fields("")value
        
            sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
        
            vmMemo = sCourtDatesID & " " & sInvoiceNo
    
        ElseIf sBrandingTheme = "5" Then         'WRTS C 50 Deposit Not Filed
            sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
                            "The turnaround as described above will begin once this invoice is paid."
            'rstTRQPlusCases.Fields("")value
        
            sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
                    "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
                    "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
            'rstTRQPlusCases.Fields("")value
        
            sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
        
            vmMemo = sCourtDatesID & " " & sInvoiceNo
    
        ElseIf sBrandingTheme = "6" Then         'WRTS C Factoring Filed
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
    
        ElseIf sBrandingTheme = "7" Then         'WRTS C Factoring Not Filed
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
    
        ElseIf sBrandingTheme = "8" Then         'WRTS C 100 Deposit Filed
            sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
                            "The turnaround as described above will begin once this invoice is paid."
            'rstTRQPlusCases.Fields("")value
        
            sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
                    "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
                    "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
            'rstTRQPlusCases.Fields("")value
        
            sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
        
            vmMemo = sCourtDatesID & " " & sInvoiceNo
    
        ElseIf sBrandingTheme = "9" Then         'WRTS C 100 Deposit Not Filed
            sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
                            "The turnaround as described above will begin once this invoice is paid."
            'rstTRQPlusCases.Fields("")value
        
            sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
                    "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
                    "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
            'rstTRQPlusCases.Fields("")value
        
            sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
        
            vmMemo = sCourtDatesID & " " & sInvoiceNo
    
        ElseIf sBrandingTheme = "10" Then        'WRTS JJ Factoring
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
    
        ElseIf sBrandingTheme = "11" Then        'WRTS Tabula Not Factored/Filed
            sPaymentTerms = sCourtDatesID & " " & sInvoiceNo 'rstTRQPlusCases.Fields("")value
            sNote = "Thank you for your business." 'rstTRQPlusCases.Fields("")value
            sTerms = "Thank you for your business." 'rstTRQPlusCases.Fields("")value
        
            vmMemo = sCourtDatesID & " " & sInvoiceNo
    
        ElseIf sBrandingTheme = "12" Then        'WRTS AMOR Factoring
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
  
End Sub


