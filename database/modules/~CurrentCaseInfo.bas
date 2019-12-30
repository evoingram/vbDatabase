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
' Name        : fPPGenerateJSONInfo
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fPPGenerateJSONInfo
' Description : get info for invoice
'============================================================================

Public sCourtDatesID As String
Public sEmail As String
Public sDescription As String
Public sInvoiceDate As String
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
Public sValue As String
Public sCustomerID As String
Public sStatusesID As String
Public sClientTranscriptName As String
Public sCurrentTranscriptName As String
Public svURL As String
Public sLinkToCSV As String
Public sFirstName As String
Public sLastName As String
Public sMrMs As String
Public sName As String
Public sAddress1 As String
Public sAddress2 As String
Public sNotes As String
Public sTime As String
Public sTime1 As String
Public vCasesID As String
Public vStatus As String
Public sUnitPrice As String
Public sOrderingID As String
Public sInvoiceID As String
Public HyperlinkString As String
Public rtfStringBody As String
Public sLocation As String
Public sDueDate As String
Public bStarred As String
Public bCompleted As String
Public sTitle As String
Public sWLListID As String

Public lAssigneeID As Long
Public lngNumOfHrs As Long
Public lngNumOfMins As Long
Public lngNumOfSecsRem As Long
Public lngNumOfSecs As Long
Public lngNumOfHrs1 As Long
Public lngNumOfMins1 As Long
Public lngNumOfSecsRem1 As Long
Public lngNumOfSecs1 As Long
Public i As Long

Public dExpectedBalanceDate As Date

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
Public Const qTRIQPlusCases As String = "TRInvoiQPlusCases"
Public Const qFCSVQ As String = "FactoringCSVQuery"
Public Const qSelectXero As String = "SelectXero"
Public Const qXeroCSVQ As String = "XeroCSVQuery"
Public Const qUPQ As String = "UnitPriceQuery"

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
Public Const sDrive As String = "T:\"

'@Ignore ConstantNotUsed
Public Const slURL As String = "\\hubcloud\evoingram\Administration\Marketing\LOGO-5inch-by-1.22inches.jpg"
'@Ignore ConstantNotUsed
Public Const sPPTemplateName As String = "Amount only"
Public Const sTermDays As String = "30"


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
    
    sFirstName = cJob.App0.FirstName
    sLastName = cJob.App0.LastName
    sName = sFirstName & " " & sLastName
    sAddress1 = cJob.App0.Company
    sAddress2 = cJob.App0.Address
    sLine1 = cJob.App0.Address
    sCity = cJob.App0.City
    sState = cJob.App0.State
    sZIP = cJob.App0.ZIP
    sUnitPrice = cJob.UnitPrice
    sEmail = cJob.App0.Notes
    sNotes = cJob.App0.Notes
    sOrderingID = cJob.sApp0
    sMrMs = cJob.App0.MrMs
    sUnitPrice = cJob.PageRate
    
End Sub

Public Sub fPPGenerateJSONInfo()
    '============================================================================
    ' Name        : fPPGenerateJSONInfo
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fPPGenerateJSONInfo
    ' Description : get info for invoice
    '============================================================================

    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID
    
    
    
    'PPID = "INV2-C8EE-ZVC5-5U36-MF27" 'INV2-K8L5-ML2R-2GLL-7KW6 '
    'sInvoiceDate = (Format((Date + 28), "yyyy-mm-dd")) & " PST"
    
    sInvoiceTime = (Format(Now(), "hh:mm:ss"))
    sMinimumAmount = "1"
    sValue = sUnitPrice
    sDescription = "Job No.:  " & sCourtDatesID & "  |  " & _
                   "Invoice No.:  " & cJob.CaseInfo.CaseNumber1 & "\n" & _
                   cJob.CaseInfo.Party1 & " v " & cJob.CaseInfo.Party2 & "\n" & _
                   "Case Nos.:  " & cJob.CaseInfo.CaseNumber1 & " " & cJob.CaseInfo.CaseNumber2 & "\n" & _
                   "Hearing Date:  " & cJob.HearingDate & "\n" & _
                   "Approx. " & cJob.AudioLength & " minutes" & "  |  " & _
                   "Turnaround Time:  " & cJob.TurnaroundTime & " calendar days"

    If cJob.BrandingTheme = "1" Then                 'WRTS NC Factoring
        sPaymentTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
                        "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
                        "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
                        "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/.  Click on 'Rates', then 'Terms of Service'." 'rstTRQPlusCases.Fields("").value
    
        sNote = "Your transcript is attached to this invoice.  We will upload this transcript to our repository for your 24/7 access and " & _
                "mail out and/or file as appropriate.  Thank you for your business." 'rstTRQPlusCases.Fields("").value
    
        sTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
                 "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
                 "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
                 "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/.  Click on 'Rates', then 'Terms of Service'." 'rstTRQPlusCases.Fields("").value

    ElseIf cJob.BrandingTheme = "2" Then             'WRTS NC 100 Deposit
        sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
                        "The turnaround as described above will begin once this invoice is paid.  Full terms of service listed at https://www.aquoco.co/.  Click on 'Rates', then 'Terms of Service'."
        'rstTRQPlusCases.Fields("")value
    
        sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
                "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
                "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
        'rstTRQPlusCases.Fields("")value
    
        sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
    

    ElseIf cJob.BrandingTheme = "3" Then             'WRTS C 50 Deposit Filed Non-BK
        sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
                        "The turnaround as described above will begin once this invoice is paid."
        'rstTRQPlusCases.Fields("")value
    
        sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
                "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
                "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
        'rstTRQPlusCases.Fields("")value
    
        sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
    
        vmMemo = sCourtDatesID & " " & cJob.InvoiceNo

    ElseIf cJob.BrandingTheme = "4" Then             'WRTS C 50 Deposit Filed BK

        sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
                        "The turnaround as described above will begin once this invoice is paid."
        'rstTRQPlusCases.Fields("")value
    
        sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
                "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
                "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
        'rstTRQPlusCases.Fields("")value
    
        sTerms = "Full terms of service listed at https://www.aquoco.co/ServiceA.html."
    

    ElseIf cJob.BrandingTheme = "5" Then             'WRTS C 50 Deposit Not Filed
        sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
                        "The turnaround as described above will begin once this invoice is paid."
        'rstTRQPlusCases.Fields("")value
    
        sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
                "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
                "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
        'rstTRQPlusCases.Fields("")value
    
        sTerms = "Full terms of service listed at https://www.aquoco.co/.  Click on 'Rates', then 'Terms of Service'."
    

    ElseIf cJob.BrandingTheme = "6" Then             'WRTS C Factoring Filed
        sPaymentTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
                        "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
                        "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
                        "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/.  Click on 'Rates', then 'Terms of Service'." 'rstTRQPlusCases.Fields("")value
    
        sNote = "Your transcript is attached to this invoice.  We will upload this transcript to our repository for your 24/7 access and " & _
                "mail out and/or file as appropriate.  Thank you for your business." 'rstTRQPlusCases.Fields("")value
    
        sTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
                 "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
                 "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
                 "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/.  Click on 'Rates', then 'Terms of Service'." 'rstTRQPlusCases.Fields("")value
        

    ElseIf cJob.BrandingTheme = "7" Then             'WRTS C Factoring Not Filed
        sPaymentTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
                        "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
                        "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
                        "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/.  Click on 'Rates', then 'Terms of Service'." 'rstTRQPlusCases.Fields("")value
    
        sNote = "Your transcript is attached to this invoice.  We will upload this transcript to our repository for your 24/7 access and " & _
                "mail out and/or file as appropriate.  Thank you for your business." 'rstTRQPlusCases.Fields("")value
    
        sTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
                 "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
                 "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
                 "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/.  Click on 'Rates', then 'Terms of Service'." 'rstTRQPlusCases.Fields("")value
        

    ElseIf cJob.BrandingTheme = "8" Then             'WRTS C 100 Deposit Filed
        sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
                        "The turnaround as described above will begin once this invoice is paid."
        'rstTRQPlusCases.Fields("")value
    
        sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
                "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
                "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
        'rstTRQPlusCases.Fields("")value
    
        sTerms = "Full terms of service listed at https://www.aquoco.co/.  Click on 'Rates', then 'Terms of Service'."
    
    ElseIf cJob.BrandingTheme = "9" Then             'WRTS C 100 Deposit Not Filed
        sPaymentTerms = "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript." & "  " & _
                        "The turnaround as described above will begin once this invoice is paid."
        'rstTRQPlusCases.Fields("")value
    
        sNote = "After completion, the transcript will be e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access." & _
                "And if you have any questions or if we can be of any more assistance, please do not hesitate to contact us (inquiries@aquoco.co).  " & _
                "If I have any spellings questions or things like that (hopefully not), I will let you know.  Thank you for your business."
        'rstTRQPlusCases.Fields("")value
    
        sTerms = "Full terms of service listed at https://www.aquoco.co/.  Click on 'Rates', then 'Terms of Service'."
    

    ElseIf cJob.BrandingTheme = "10" Then            'WRTS JJ Factoring
        sPaymentTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
                        "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
                        "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
                        "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/.  Click on 'Rates', then 'Terms of Service'." 'rstTRQPlusCases.Fields("")value
    
        sNote = "Your transcript is attached to this invoice.  We will upload this transcript to our repository for your 24/7 access and " & _
                "mail out and/or file as appropriate.  Thank you for your business." 'rstTRQPlusCases.Fields("")value
    
        sTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
                 "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
                 "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
                 "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/.  Click on 'Rates', then 'Terms of Service'." 'rstTRQPlusCases.Fields("")value
        

    ElseIf cJob.BrandingTheme = "11" Then            'WRTS Tabula Not Factored/Filed
        sPaymentTerms = sCourtDatesID & " " & cJob.InvoiceNo 'rstTRQPlusCases.Fields("")value
        sNote = "Thank you for your business."   'rstTRQPlusCases.Fields("")value
        sTerms = "Thank you for your business."  'rstTRQPlusCases.Fields("")value

    ElseIf cJob.BrandingTheme = "12" Then            'WRTS AMOR Factoring
        sPaymentTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
                        "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
                        "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
                        "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/.  Click on 'Rates', then 'Terms of Service'." 'rstTRQPlusCases.Fields("")value
    
        sNote = "Your transcript is attached to this invoice.  We will upload this transcript to our repository for your 24/7 access and " & _
                "mail out and/or file as appropriate.  Thank you for your business." 'rstTRQPlusCases.Fields("")value
    
        sTerms = "Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  " & _
                 "Please pay within 28 days. 5% interest if payment received after 28 calendar days of" & _
                 "invoice date, additional 1% interest added every 7th calendar day after day 28 up " & _
                 "to a maximum of 12%.  Full terms of service listed at https://www.aquoco.co/.  Click on 'Rates', then 'Terms of Service'." 'rstTRQPlusCases.Fields("")value
        
    End If
    vmMemo = sCourtDatesID & " " & cJob.InvoiceNo

End Sub

Public Sub pfClearGlobals()

    '============================================================================
    ' Name        : pfClearGlobals
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfClearGlobals
    ' Description : clears all global variables
    '============================================================================
    sCourtDatesID = ""
    sEmail = ""
    sDescription = ""
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
    sValue = ""
    Set oWordApp = Nothing
    Set oWordDoc = Nothing
    sCustomerID = ""
    sStatusesID = ""
    svURL = ""
    sLinkToCSV = ""
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

End Sub
