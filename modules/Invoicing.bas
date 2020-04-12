Attribute VB_Name = "Invoicing"
'@Folder("Database.Admin.Modules")
Option Compare Database
Option Explicit

'============================================================================
'class module cmInvoice

'variables:
'   NONE

'functions:
'ApplyShipDateTrackingNumber:   Description:  functions like ApplyPayPalPayment for shipping
'                                         checks outlook email table for ShipDate & tracking number and adds to courtdates
'                           Arguments:    NONE
'ApplyPayPalPayment:            Description:  applies found PayPal payment to job ##
'                           Arguments:    NONE
'fTranscriptExpensesAfter:      Description:  logs post-completion transcript expenses
'                                                 ink x actualquantity (after job completed)   |
'                                                 paper x actualquantity (after job completed)   |
'                                                 Vendor, ExpensesDate, Amount, Memo
'                           Arguments:    NONE
'fTranscriptExpensesBeginning:  Description:  pre-completion transcript expenses
'                                         covers x 2 per volume & copy (beginning)
'                                         velobind (beginning)
'                                         1 CD or 2 CDs if superior court (beginning)
'                                         1 cd sleeve or 2 (beginning)
'                                         1 business card (beginning)
'                                         Vendor, ExpensesDate, Amount, Memo
'                           Arguments:    NONE
'fUpdateFactoringDates:         Description:  updates various factoring dates in CourtDates table
'                           Arguments:    NONE
'fPaymentAdd:                   Description:  adds payment to Payments table
'                           Arguments:    sInvoiceNumber
'fAutoCalculateFactorInterest:  Description:  add 1% after every 7 days payment not made
'                           Arguments:    NONE
'fDepositPaymentReceived:       Description:  does some things after a deposit is paid
'                           Arguments:    NONE
'fIsFactoringApproved:          Description:  checks if factoring is approved for customer
'                           Arguments:    NONE
'fGenerateInvoiceNumber:        Description:  generates invoice number
'                           Arguments:    NONE
'fShippingExpenseEntry:         Description:  auto generate shipping expense entry when shipping xml generated; needs sVendorName, dExpenseIncurred, iExpenseAmount, sExpenseMemo
'                                         generates shipping expenses and enters them into the expenses table
'                           Arguments:    sTrackingNumber
        
'============================================================================

Public Sub fApplyShipDateTrackingNumber()

    '============================================================================
    ' Name        : fApplyShipDateTrackingNumber
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fApplyShipDateTrackingNumber
    ' Description : functions like ApplyPayPalPayment for shipping
    '               checks outlook email table for ShipDate & tracking number and adds to courtdates
    '============================================================================

    Dim rstCourtDates As DAO.Recordset
    Dim rstOLPayPalPmt As DAO.Recordset
    Dim x As Long
    Dim vShipDate As String
    Dim sTrackingNumber As String
    Dim formDOM As DOMDocument60                 'Currently opened xml file
    Dim ixmlRoot As IXMLDOMElement
    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID


    Set rstCourtDates = CurrentDb.OpenRecordset("SELECT ID, ShipDate, TrackingNumber FROM CourtDates")
    Set rstOLPayPalPmt = CurrentDb.OpenRecordset("OLPayPalPayments")
    
    If Not (rstCourtDates.EOF And rstCourtDates.BOF) Then 'For each CourtDates.ID

        rstCourtDates.MoveFirst
    
        Do Until rstCourtDates.EOF = True
        
            'get tracking number and ship date
            rstCourtDates.Fields("ID").Value = sCourtDatesID
      
            If Not (rstOLPayPalPmt.EOF And rstOLPayPalPmt.BOF) Then 'For each row in OLPayPalPayments
        
                rstOLPayPalPmt.MoveFirst
            
                Do Until rstOLPayPalPmt.EOF = True
            
                    'will return 15 in x; If not found it will return 0
                    x = InStr(rstOLPayPalPmt!Contents, "Endicia")
                    
                    If x > 0 Then
                    
                        Call pfCommunicationHistoryAdd("TrackingNumber") 'comms history entry for paypal email
                        
                        Do While Len(Dir(cJob.DocPath.ShippingOutputFolder)) > 0
                        
                            Set formDOM = New DOMDocument60          'Open the xml file
                            formDOM.resolveExternals = False         'using schema yes/no true/false
                            formDOM.validateOnParse = False          'Parser validate document?  Still parses well-formed XML
                            formDOM.Load (cJob.DocPath.ShippingOutputFolder & Dir(cJob.DocPath.ShippingOutputFolder))
                            Set ixmlRoot = formDOM.DocumentElement   'Get document reference
                            sTrackingNumber = ixmlRoot.SelectSingleNode("//DAZzle/Package/PIC").Text
                            Set formDOM = Nothing
                            Set ixmlRoot = Nothing
                            rstCourtDates.Fields("ShipDate").Value = Format(Now, "mm-dd-yyyy")
                            rstCourtDates.Fields("TrackingNumber").Value = sTrackingNumber
                            
                        Loop
                        
                        
                    Else
                    End If
                  
                    rstOLPayPalPmt.MoveNext
                
                Loop
            
                vShipDate = vbNullString
                sCourtDatesID = vbNullString
                sTrackingNumber = vbNullString
                
                rstCourtDates.MoveNext
                
            Else
            End If
        
        Loop
    
    Else

        MsgBox "There are no more packages to process."
    
    End If

    MsgBox "Finished searching for tracking numbers."

    rstOLPayPalPmt.Close
    Set rstOLPayPalPmt = Nothing
    rstCourtDates.Close                          'Close the recordset
    Set rstCourtDates = Nothing                  'Clean up

    sCourtDatesID = vbNullString
End Sub

Public Sub fApplyPayPalPayment()

    '============================================================================
    ' Name        : fApplyPayPalPayment
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fApplyPayPalPayment
    ' Description : applies found PayPal payment to job ##
    '============================================================================

    Dim rstCourtDates As DAO.Recordset
    Dim rstOLPayPalPmt As DAO.Recordset
    Dim x As Long

    'TODO: narrow these dao recordsets down
    
    Dim vAmount As String
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID
    
    Set rstCourtDates = CurrentDb.OpenRecordset("SELECT ID, InvoiceNo FROM CourtDates")
    Set rstOLPayPalPmt = CurrentDb.OpenRecordset("OLPayPalPayments")
    
    If Not (rstCourtDates.EOF And rstCourtDates.BOF) Then 'For each CourtDates.ID

        rstCourtDates.MoveFirst
    
        Do Until rstCourtDates.EOF = True
    
            sCourtDatesID = rstCourtDates.Fields("ID").Value
        
            If Not (rstOLPayPalPmt.EOF And rstOLPayPalPmt.BOF) Then 'For each row in OLPayPalPayments
                rstOLPayPalPmt.MoveFirst
                Do Until rstOLPayPalPmt.EOF = True
            
                    'will return 15 in x; If not found it will return 0
                    x = InStr(rstOLPayPalPmt!Contents, sCourtDatesID)
                    
                    If x > 0 Then
                        vAmount = 0
                        Call pfCommunicationHistoryAdd("PayPalPayment") 'comms history entry for paypal email
                    
                        Call fPaymentAdd(cJob.InvoiceNo, vAmount) 'apply payment to sCourtDatesID
                        rstOLPayPalPmt.delete    'delete in OLPayPalPayments
                    Else
                    
                        Call pfCommunicationHistoryAdd("PayPalInvoiceSent") 'comms history entry for paypal email
                        rstOLPayPalPmt.delete    'delete in OLPayPalPayments
                    
                    End If
       
                    rstOLPayPalPmt.MoveNext
                
                Loop
            
            Else
            End If

            sCourtDatesID = vbNullString
        
            rstCourtDates.MoveNext
        Loop
            
    Else

        MsgBox "There are no PayPal payments to process."
    
    End If

    Debug.Print "Finished looping through PayPal Payments."

    rstOLPayPalPmt.Close
    Set rstOLPayPalPmt = Nothing
    rstCourtDates.Close                          'Close the recordset
    Set rstCourtDates = Nothing                  'Clean up

    sCourtDatesID = vbNullString
End Sub

Public Sub fGenerateInvoiceNumber()
    '============================================================================
    ' Name        : fGenerateInvoiceNumber
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fGenerateInvoiceNumber
    ' Description : generates invoice number
    '============================================================================

    Dim rstTempCourtDates As DAO.Recordset
    Dim rstMaxCourtDates As DAO.Recordset
    Dim rstCourtDates As DAO.Recordset
    Dim sQuestion As String
    Dim sAnswer As String
    Dim sInvoiceNumber As String

    Set rstMaxCourtDates = CurrentDb.OpenRecordset("SELECT MAX(InvoiceNo) FROM CourtDates;")
    sInvoiceNumber = rstMaxCourtDates.Fields(0).Value

    Set rstCourtDates = CurrentDb.OpenRecordset("SELECT MAX(ID) as CourtDatesID FROM CourtDates;")
    rstCourtDates.MoveFirst
    sCourtDatesID = rstCourtDates.Fields("CourtDatesID").Value

    sQuestion = "The most recent invoice number is " & sInvoiceNumber & _
                ".  Would you like to use " & sInvoiceNumber + 1 & " as your next invoice number?"
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    If sAnswer = vbNo Then                       'Code for No
        sInvoiceNumber = InputBox("Enter your next Invoice Number.  The most recent one was " & sInvoiceNumber & ".")
    Else                                         'Code for yes
        sInvoiceNumber = sInvoiceNumber + 1
    End If
    
    'insert calculated fields into tempFPtable
    Set rstTempCourtDates = CurrentDb.OpenRecordset("qSelect1TempCourtDates")
    rstTempCourtDates.Edit
    rstTempCourtDates.Fields("InvoiceNo") = sInvoiceNumber
    rstTempCourtDates.Update
    rstTempCourtDates.Close

    rstCourtDates.Close

    Set rstCourtDates = CurrentDb.OpenRecordset("SELECT * FROM CourtDates WHERE [ID]=" & sCourtDatesID & ";")
    rstCourtDates.Edit
    rstCourtDates.Fields("InvoiceNo") = sInvoiceNumber
    rstCourtDates.Update

    rstCourtDates.Close
    rstMaxCourtDates.Close

    sCourtDatesID = vbNullString
End Sub

Public Sub fIsFactoringApproved()

    '============================================================================
    ' Name        : fIsFactoringApproved
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fIsFactoringApproved
    ' Description : checks if factoring is approved for customer
    '============================================================================

    Dim cJob As Job
    Set cJob = New Job
    Dim svURL As String

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID

    svURL = "https://www.paypal.com"
    DoCmd.OpenQuery qUPQ, acViewNormal, acReadOnly
    DoCmd.OpenQuery "INVUpdateUnitPriceQuery", acViewNormal, acEdit
    DoCmd.Close acQuery, qUPQ
    DoCmd.Close acQuery, "InvUpdateUnitPriceQuery"

    'if approved for factoring
    
    If cJob.App0.FactoringApproved = "True" Then                         'if box checked, then do this

        'send price quote
        Call pfGenericExportandMailMerge("Invoice", "Stage1s\PriceQuote")
        Call pfCommunicationHistoryAdd("PriceQuote")
        
        'send adam CIDIncomeReport
        Call pfGenericExportandMailMerge("Invoice", "Stage1s\CIDIncomeReport")
        Call pfCommunicationHistoryAdd("CIDIncomeReport")
        Call pfSendWordDocAsEmail("CIDIncomeReport", "Initial Income Notification") 'initial income report 'emails adam cid report
        
    Else                                         'if factoring not approved
                 
        'send deposit invoice (HTML copy in email and PDF copy)
        Call pfGenericExportandMailMerge("Invoice", "Stage1s\DepositInvoice") 'ISSUE LOCAL GENERATED INVOICE
        Application.FollowHyperlink (svURL)      'ISSUE UPDATED INVOICE
        Call pfCommunicationHistoryAdd("DepositInvoice") 'LOG UPDATED INVOICE
        Call pfGenericExportandMailMerge("Invoice", "Stage1s\CIDIncomeReport") 'ISSUE INCOME REPORT
        Call pfCommunicationHistoryAdd("CIDIncomeReport") 'LOG INCOME REPORT
        Call pfSendWordDocAsEmail("CIDIncomeReport", "Initial Income Notification") 'initial income report 'emails adam cid report
        Call fPPDraft                            'issues paypal invoice
        Call fSendPPEmailDeposit                 'sends paypal invoice email for deposit.
        
        
        
        
    End If

    sCourtDatesID = vbNullString

End Sub

Public Sub fDepositPaymentReceived()

    '============================================================================
    ' Name        : fDepositPaymentReceived
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fDepositPaymentReceived
    ' Description : does some things after a deposit is paid
    '============================================================================
    
    Dim vAmount As String
    Dim sQuestion As String
    Dim sAnswer As String

    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID

    'calculate refund
    vAmount = InputBox("How much was the payment?  Their invoice totals up to $" & cJob.FinalPrice & ".")
    
    sQuestion = "Are you sure that's correct?  The payment was in the amount of $" & vAmount & "?"
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    If sAnswer = vbNo Then                       'Code for No

        sQuestion = "Do you want to enter a payment?"
        sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
    
        If sAnswer = vbNo Then                   'Code for No

            GoTo Exitif
        
        Else                                     'Code for Yes

            vAmount = InputBox("How much was the payment?  Their invoice totals up to $" & cJob.FinalPrice & ".")
            Call fPaymentAdd(cJob.InvoiceNo, vAmount)
            Call pfGenericExportandMailMerge("Invoice", "Stage1s\DepositPaid")
            Call pfCommunicationHistoryAdd("DepositPaid")
        
        End If
    
    Else                                         'Code for Yes

        vAmount = InputBox("How much was the payment?  Their invoice totals up to $" & cJob.FinalPrice & ".")
        Call fPaymentAdd(cJob.InvoiceNo, vAmount)
        Call pfGenericExportandMailMerge("Invoice", "Stage1s\DepositPaid")
        Call pfCommunicationHistoryAdd("DepositPaid")

    End If
Exitif:
    sCourtDatesID = vbNullString
End Sub

Public Sub pfAutoCalculateFactorInterest()

    '============================================================================
    ' Name        : pfAutoCalculateFactorInterest
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call pfAutoCalculateFactorInterest
    ' Description : add 1% after every 7 days payment not made
    '============================================================================

    Dim sFactoringApproved As String
    Dim sSubtotal As String
    Dim sPaymentSum As String
    Dim sFactoringCost As String
    
    Dim dInvoiceDate As Date
    
    Dim iOrderingID As Long
    Dim iDateDifference As Long
    
    Dim rstCustomers As DAO.Recordset
    Dim rstCourtDates As DAO.Recordset
    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID
    
    Set rstCourtDates = CurrentDb.OpenRecordset("CourtDates")

    If Not (rstCourtDates.EOF And rstCourtDates.BOF) Then 'For each CourtDates.ID

        rstCourtDates.MoveFirst
    
        Do Until rstCourtDates.EOF = True
    
            sCourtDatesID = rstCourtDates.Fields("ID").Value
            iOrderingID = rstCourtDates.Fields("OrderingID").Value
            dInvoiceDate = rstCourtDates.Fields("InvoiceDate").Value
            sSubtotal = rstCourtDates.Fields("FinalPrice").Value
            sPaymentSum = Nz(rstCourtDates.Fields("PaymentSum").Value, 0)
        
            If Not IsNull(dInvoiceDate) Then
        
                If IsNull(sPaymentSum) Then
            
                    Set rstCustomers = CurrentDb.OpenRecordset("SELECT * FROM Customers WHERE [ID] = " & iOrderingID & ";")
                    sFactoringApproved = rstCustomers.Fields("FactoringApproved").Value
                    rstCustomers.Close
                
                    Set rstCustomers = Nothing
                
                    If sFactoringApproved = "True" Then
                
                        iDateDifference = DateDiff("d", Now, dInvoiceDate)
                    
                        If iDateDifference < 0 And iDateDifference > -7 Then
                            sFactoringCost = 0
                        ElseIf iDateDifference < -7 And iDateDifference > -14 Then
                            sFactoringCost = sSubtotal * 0.01
                        ElseIf iDateDifference < -14 And iDateDifference > -21 Then
                            sFactoringCost = sSubtotal * 0.02
                        ElseIf iDateDifference < -21 And iDateDifference > -28 Then
                            sFactoringCost = sSubtotal * 0.03
                        ElseIf iDateDifference < -28 And iDateDifference > -35 Then
                            sFactoringCost = sSubtotal * 0.04
                        ElseIf iDateDifference < -35 Then
                            sFactoringCost = sSubtotal * 0.05
                        Else
                        End If
                    
                        rstCourtDates.Edit
                        rstCourtDates.Fields("FactoringCost").Value = sFactoringCost
                        rstCourtDates.Update
                    
                    Else
                    End If
                
                Else
                End If
            
            Else
            End If
        
            rstCourtDates.MoveNext
        Loop
    
        rstCourtDates.Close
        Set rstCourtDates = Nothing
        Set rstCustomers = Nothing

    Else
    End If
    sCourtDatesID = vbNullString

End Sub

Public Sub fPaymentAdd(sInvoiceNumber As String, vAmount As String)

    '============================================================================
    ' Name        : fPaymentAdd
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fPaymentAdd(sInvoiceNumber, vAmount)
    ' Description : adds payment to Payments table
    '============================================================================

    Dim sTableHyperlink As String
    
    Dim rstPaymentAdd As DAO.Recordset

    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID

    '@Ignore AssignmentNotUsed
    sTableHyperlink = sCourtDatesID & "-PaymentMade" & "#" & cJob.DocPath.PaymentMade & "#"

    Set rstPaymentAdd = CurrentDb.OpenRecordset("Payments")

    rstPaymentAdd.AddNew
        rstPaymentAdd("InvoiceNo").Value = sInvoiceNumber
        rstPaymentAdd("Amount").Value = vAmount
        rstPaymentAdd("RemitDate").Value = Date
    rstPaymentAdd.Update

    rstPaymentAdd.Close

    Call fManualPPPayment
    
    sCourtDatesID = vbNullString
End Sub

Public Sub fUpdateFactoringDates()

    '============================================================================
    ' Name        : fUpdateFactoringDates
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fUpdateFactoringDates
    ' Description : updates various factoring dates in CourtDates table
    '============================================================================

    Dim sExpectedAdvanceAmount As String
    Dim sExpectedRebateAmount As String
    Dim sCDCalcUpdateSQL As String
    
    Dim dExpectedBalanceDate As Date
    
    Dim iUnitPriceID As Long

    Dim rstUnitPriceRate As DAO.Recordset
    
    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID

    sExpectedRebateAmount = (cJob.FinalPrice * 0.18)
    sExpectedAdvanceAmount = (cJob.FinalPrice * 0.8)
    dExpectedBalanceDate = (Date + cJob.TurnaroundTime) - 2
    
    
    'TODO: update rates here
    'avail turnarounds 7 10 14 30 1 3
    'if jurisdiction contains and turnaround contains, for each different rate
    'avt rate 33 $1.35 or 35 $1.60, janet rate 37 $2.20, non-court rate 38 $2.00 per minute
    'regular 45 1 $6.05, 44 3 $5.45, 43 7 $4.85, 42 14 $4.25, 41 30 $3.65
    'volume 1 46 $6.65, 44 7 $5.45, 43 14 $4.85, 42 30 $4.25
    'copies for same 1.2, 1.05, 0.9, 0.9, 0.9
    'king county rate 40 3.10
        
    'Non -Court

    '    10 calendar-day turnaround, $2.00 per audio minute 49
    '    same day/overnight, $5.25 per page 61


    'Court Transcription

    '    30 calendar-day turnaround, $2.65/page 39
    '    14 calendar-day turnaround, $3.50/page 41
    '    07 calendar-day turnaround, $4.00/page 62
    '    03 calendar-day turnaround, $4.75/page 57
    '    same day/overnight turnaround, $5.25/page 61
    
    
    'insert calculated fields into courtdates
    sCDCalcUpdateSQL = "UPDATE CourtDates SET [ExpectedRebateDate] = " & Format(cJob.ExpectedRebateDate, "mm-dd-yyyy") & ", [ExpectedAdvanceDate] = " & Format(cJob.ExpectedAdvanceDate, "mm-dd-yyyy") & ", [Subtotal] = " & cJob.Subtotal & " WHERE ID = " & sCourtDatesID & ";"

    CurrentDb.Execute sCDCalcUpdateSQL
    CurrentDb.Close
    
    

    sCourtDatesID = vbNullString
End Sub

Public Sub fTranscriptExpensesBeginning()

    '============================================================================
    ' Name        : fTranscriptExpensesBeginning
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fTranscriptExpensesBeginning
    ' Description : pre-completion transcript expenses
    '               covers x 2 per volume & copy (beginning)
    '               velobind (beginning)
    '               1 CD or 2 CDs if superior court (beginning)
    '               1 cd sleeve or 2 (beginning)
    '               1 business card (beginning)
    '               Vendor, ExpensesDate, Amount, Memo
    '============================================================================
        
    Dim rstExpensesAdd As DAO.Recordset
    Dim rstCourtDatesSet As DAO.Recordset
    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID
    
    Set rstExpensesAdd = CurrentDb.OpenRecordset("Expenses")

    If Int(cJob.EstimatedPageCount) > 200 Then
 
        rstExpensesAdd.AddNew                    'back cover
            rstExpensesAdd("Vendor").Value = "Got Print"
            rstExpensesAdd("ExpensesDate").Value = Now
            rstExpensesAdd("Amount").Value = 0.6
            rstExpensesAdd("Memo").Value = "Back Cover"
            rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
            rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
        rstExpensesAdd.Update

        rstExpensesAdd.AddNew                    'back cover
            rstExpensesAdd("Vendor").Value = "Got Print"
            rstExpensesAdd("ExpensesDate").Value = Now
            rstExpensesAdd("Amount").Value = 0.6
            rstExpensesAdd("Memo").Value = "Back Cover"
            rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
            rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
        rstExpensesAdd.Update

        rstExpensesAdd.AddNew                    'front cover
            rstExpensesAdd("Vendor").Value = "Amazon"
            rstExpensesAdd("ExpensesDate").Value = Now
            rstExpensesAdd("Amount").Value = 0.22
            rstExpensesAdd("Memo").Value = "front cover"
            rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
            rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
        rstExpensesAdd.Update

        rstExpensesAdd.AddNew                    'front cover
            rstExpensesAdd("Vendor").Value = "Amazon"
            rstExpensesAdd("ExpensesDate").Value = Now
            rstExpensesAdd("Amount").Value = 0.22
            rstExpensesAdd("Memo").Value = "front cover"
            rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
            rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
        rstExpensesAdd.Update

        rstExpensesAdd.AddNew                    'CD
            rstExpensesAdd("Vendor").Value = "Amazon"
            rstExpensesAdd("ExpensesDate").Value = Now
            rstExpensesAdd("Amount").Value = 0.18
            rstExpensesAdd("Memo").Value = "CD"
            rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
            rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
        rstExpensesAdd.Update

        rstExpensesAdd.AddNew                    'CD sleeve
            rstExpensesAdd("Vendor").Value = "Amazon"
            rstExpensesAdd("ExpensesDate").Value = Now
            rstExpensesAdd("Amount").Value = 0.09
            rstExpensesAdd("Memo").Value = "CD Sleeve"
            rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
            rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
        rstExpensesAdd.Update

        rstExpensesAdd.AddNew                    'shipping envelope
            rstExpensesAdd("Vendor").Value = "Amazon"
            rstExpensesAdd("ExpensesDate").Value = Now
            rstExpensesAdd("Amount").Value = 0.32
            rstExpensesAdd("Memo").Value = "Shipping Envelope"
            rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
            rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
        rstExpensesAdd.Update

        rstExpensesAdd.AddNew                    'Velobind
            rstExpensesAdd("Vendor").Value = "Amazon"
            rstExpensesAdd("ExpensesDate").Value = Now
            rstExpensesAdd("Amount").Value = 0.72
            rstExpensesAdd("Memo").Value = "Velobind"
            rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
            rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
        rstExpensesAdd.Update

        rstExpensesAdd.AddNew                    'Velobind
            rstExpensesAdd("Vendor").Value = "Amazon"
            rstExpensesAdd("ExpensesDate").Value = Now
            rstExpensesAdd("Amount").Value = 0.72
            rstExpensesAdd("Memo").Value = "Velobind"
            rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
            rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
        rstExpensesAdd.Update

        rstExpensesAdd.AddNew                    'business card
            rstExpensesAdd("Vendor").Value = "Got Print"
            rstExpensesAdd("ExpensesDate").Value = Now
            rstExpensesAdd("Amount").Value = 0.05
            rstExpensesAdd("Memo").Value = "Business Card"
            rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
            rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
        rstExpensesAdd.Update
 
        rstExpensesAdd.AddNew                    'shipping label
            rstExpensesAdd("Vendor").Value = "Amazon"
            rstExpensesAdd("ExpensesDate").Value = Now
            rstExpensesAdd("Amount").Value = 0.15
            rstExpensesAdd("Memo").Value = "Shipping Label"
            rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
            rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
        rstExpensesAdd.Update
    Else
 
        rstExpensesAdd.AddNew                    'back cover
            rstExpensesAdd("Vendor").Value = "Got Print"
            rstExpensesAdd("ExpensesDate").Value = Now
            rstExpensesAdd("Amount").Value = 0.6
            rstExpensesAdd("Memo").Value = "Back Cover"
            rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
            rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
        rstExpensesAdd.Update

        rstExpensesAdd.AddNew                    'front cover
            rstExpensesAdd("Vendor").Value = "Amazon"
            rstExpensesAdd("ExpensesDate").Value = Now
            rstExpensesAdd("Amount").Value = 0.22
            rstExpensesAdd("Memo").Value = "front cover"
            rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
            rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
        rstExpensesAdd.Update

        rstExpensesAdd.AddNew                    'CD
            rstExpensesAdd("Vendor").Value = "Amazon"
            rstExpensesAdd("ExpensesDate").Value = Now
            rstExpensesAdd("Amount").Value = 0.18
            rstExpensesAdd("Memo").Value = "CD"
            rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
            rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
        rstExpensesAdd.Update

        rstExpensesAdd.AddNew                    'CD sleeve
            rstExpensesAdd("Vendor").Value = "Amazon"
            rstExpensesAdd("ExpensesDate").Value = Now
            rstExpensesAdd("Amount").Value = 0.09
            rstExpensesAdd("Memo").Value = "CD Sleeve"
            rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
            rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
        rstExpensesAdd.Update

        rstExpensesAdd.AddNew                    'shipping envelope
            rstExpensesAdd("Vendor").Value = "Amazon"
            rstExpensesAdd("ExpensesDate").Value = Now
            rstExpensesAdd("Amount").Value = 0.32
            rstExpensesAdd("Memo").Value = "Shipping Envelope"
            rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
            rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
        rstExpensesAdd.Update

        rstExpensesAdd.AddNew                    'Velobind
            rstExpensesAdd("Vendor").Value = "Amazon"
            rstExpensesAdd("ExpensesDate").Value = Now
            rstExpensesAdd("Amount").Value = 0.72
            rstExpensesAdd("Memo").Value = "Velobind"
            rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
            rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
        rstExpensesAdd.Update

        rstExpensesAdd.AddNew                    'business card
            rstExpensesAdd("Vendor").Value = "Got Print"
            rstExpensesAdd("ExpensesDate").Value = Now
            rstExpensesAdd("Amount").Value = 0.05
            rstExpensesAdd("Memo").Value = "Business Card"
            rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
            rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
        rstExpensesAdd.Update
    
    End If

    rstExpensesAdd.AddNew                        'shipping label
        rstExpensesAdd("Vendor").Value = "Amazon"
        rstExpensesAdd("ExpensesDate").Value = Now
        rstExpensesAdd("Amount").Value = 0.15
        rstExpensesAdd("Memo").Value = "Shipping Label"
        rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
        rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
    rstExpensesAdd.Update
    
    rstExpensesAdd.Close
        
    MsgBox "Static Expenses Added!"
    
    
    sCourtDatesID = vbNullString
End Sub

Public Sub fTranscriptExpensesAfter()

    '============================================================================
    ' Name        : fTranscriptExpensesAfter
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fTranscriptExpensesAfter
    ' Description : logs post-completion transcript expenses
    '               ink x actualquantity (after job completed)   |
    '               paper x actualquantity (after job completed)   |
    '               Vendor, ExpensesDate, Amount, Memo
    '============================================================================
    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID
    
    'TODO: Fix hard copy expenses calculation
    
    Dim rstExpensesAdd As DAO.Recordset

    Set rstExpensesAdd = CurrentDb.OpenRecordset("Expenses")

    rstExpensesAdd.AddNew                        'static
        rstExpensesAdd("Vendor").Value = "internet rent etc"
        rstExpensesAdd("ExpensesDate").Value = Now
        rstExpensesAdd("Amount").Value = 0.09 * cJob.EstimatedPageCount
        rstExpensesAdd("Memo").Value = "internet rent electricity website"
        rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
        rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
    rstExpensesAdd.Update

    rstExpensesAdd.AddNew                        'paper
        rstExpensesAdd("Vendor").Value = "OfficeSupply.com"
        rstExpensesAdd("ExpensesDate").Value = Now
        rstExpensesAdd("Amount").Value = 0.01 * cJob.EstimatedPageCount
        rstExpensesAdd("Memo").Value = "paper"
        rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
        rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
    rstExpensesAdd.Update

    rstExpensesAdd.AddNew                        'ink
        rstExpensesAdd("Vendor").Value = "OfficeSupply.com"
        rstExpensesAdd("ExpensesDate").Value = Now
        rstExpensesAdd("Amount").Value = 0.008
        rstExpensesAdd("Memo").Value = "ink"
        rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
        rstExpensesAdd("InvoiceNo").Value = cJob.InvoiceNo
    rstExpensesAdd.Update
    
    rstExpensesAdd.Close

    MsgBox "Dynamic Expenses Added!"
    
    sCourtDatesID = vbNullString
End Sub

Public Sub fShippingExpenseEntry(sTrackingNumber As String)
    '============================================================================
    ' Name        : fShippingExpenseEntry
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fShippingExpenseEntry(sTrackingNumber)
    ' Description : auto generate shipping expense entry when shipping xml generated;
    'needs sVendorName, dExpenseIncurred, iExpenseAmount, sExpenseMemo
    '               generates shipping expenses and enters them into the expenses table
    '============================================================================

    Dim sVendorName As String
    Dim sExpenseMemo As String
    
    Dim dExpenseIncurred As Date
    Dim iExpenseAmount As Long

    Dim qdf1 As QueryDef
    Dim rstExpenses As DAO.Recordset
    Dim rstShippingExpenseEntry As DAO.Recordset

    Dim formDOM As DOMDocument60                 'Currently opened xml file
    Dim ixmlRoot As IXMLDOMElement
    
    Dim cJob As Job
    Set cJob = New Job
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    cJob.FindFirst "ID=" & sCourtDatesID

    'Set qdf1 = CurrentDb.QueryDefs(qnTRCourtQ)
    'qdf1.Parameters(0) = sCourtDatesID
    'Set rstShippingExpenseEntry = qdf1.OpenRecordset


    'sVendorName = rstShippingExpenseEntry.Fields("OrderingID").Value

    dExpenseIncurred = Date
    
    Do While Len(Dir(cJob.DocPath.ShippingOutputFolder)) > 0
    
        Set formDOM = New DOMDocument60          'Open the xml file
        formDOM.resolveExternals = False         'using schema yes/no true/false
        formDOM.validateOnParse = False          'Parser validate document?  Still parses well-formed XML
        formDOM.Load (cJob.DocPath.ShippingOutputFolder & Dir(cJob.DocPath.ShippingOutputFolder))
        Set ixmlRoot = formDOM.DocumentElement   'Get document reference
        iExpenseAmount = ixmlRoot.SelectSingleNode("//DAZzle/Package/FinalPostage").Text
        sTrackingNumber = ixmlRoot.SelectSingleNode("//DAZzle/Package/PIC").Text
        Set formDOM = Nothing
        Set ixmlRoot = Nothing
        
    Loop
    
    sExpenseMemo = "Shipping Job No. " & sCourtDatesID & " Invoice " & cJob.InvoiceNo & " Tracking " & sTrackingNumber

    'generates new entry for shipping XML in expenses table
    Set rstExpenses = CurrentDb.OpenRecordset("Expenses")

    rstExpenses.AddNew
        rstExpenses("sVendorName").Value = cJob.App0.ID
        rstExpenses("dExpenseIncurred").Value = dExpenseIncurred
        rstExpenses("iExpenseAmount").Value = iExpenseAmount
        rstExpenses("sExpenseMemo").Value = sExpenseMemo
    rstExpenses.Update
    
    rstExpenses.Close
    
    
    sCourtDatesID = vbNullString
End Sub
