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


Sub fApplyShipDateTrackingNumber()

'============================================================================
' Name        : fApplyShipDateTrackingNumber
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fApplyShipDateTrackingNumber
' Description : functions like ApplyPayPalPayment for shipping
'               checks outlook email table for ShipDate & tracking number and adds to courtdates
'============================================================================

Dim rstCourtDates As DAO.Recordset, rstOLPayPalPmt As DAO.Recordset
Dim x As Integer
Dim vShipDate As String, vTrackingNumber As String


Set rstCourtDates = CurrentDb.OpenRecordset("SELECT ID, ShipDate, TrackingNumber FROM CourtDates")
Set rstOLPayPalPmt = CurrentDb.OpenRecordset("OLPayPalPayments")
    
If Not (rstCourtDates.EOF And rstCourtDates.BOF) Then 'For each CourtDates.ID

    rstCourtDates.MoveFirst
    
    Do Until rstCourtDates.EOF = True
    
        rstCourtDates.Fields("ShipDate").Value = vShipDate
        rstCourtDates.Fields("ID").Value = sCourtDatesID
      
        If Not (rstOLPayPalPmt.EOF And rstOLPayPalPmt.BOF) Then 'For each row in OLPayPalPayments
        
            rstOLPayPalPmt.MoveFirst
            
            Do Until rstOLPayPalPmt.EOF = True
            
                'will return 15 in x; If not found it will return 0
                x = InStr(rstOLPayPalPmt!Contents, "Endicia")
                    
                If x > 0 Then
                    
                        Call pfCommunicationHistoryAdd("TrackingNumber") 'comms history entry for paypal email
                    'get tracking number and ship date
                    
                    'update courtdates shipdate and tracking number come back
                    
                Else
                End If
                  
                rstOLPayPalPmt.MoveNext
                
            Loop
            
                vShipDate = ""
                sCourtDatesID = ""
                vTrackingNumber = ""
                
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
rstCourtDates.Close 'Close the recordset
Set rstCourtDates = Nothing 'Clean up

End Sub
Sub fApplyPayPalPayment()

'============================================================================
' Name        : fApplyPayPalPayment
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fApplyPayPalPayment
' Description : applies found PayPal payment to job ##
'============================================================================

Dim rstCourtDates As DAO.Recordset, rstOLPayPalPmt As DAO.Recordset
Dim x As Integer


Dim vAmount As String
Set rstCourtDates = CurrentDb.OpenRecordset("SELECT ID, InvoiceNo FROM CourtDates")
Set rstOLPayPalPmt = CurrentDb.OpenRecordset("OLPayPalPayments")
    
If Not (rstCourtDates.EOF And rstCourtDates.BOF) Then 'For each CourtDates.ID

    rstCourtDates.MoveFirst
    
    Do Until rstCourtDates.EOF = True
    
        sInvoiceNumber = rstCourtDates.Fields("InvoiceNo").Value
        sCourtDatesID = rstCourtDates.Fields("ID").Value
        
        If Not (rstOLPayPalPmt.EOF And rstOLPayPalPmt.BOF) Then 'For each row in OLPayPalPayments
            rstOLPayPalPmt.MoveFirst
            Do Until rstOLPayPalPmt.EOF = True
            
                'will return 15 in x; If not found it will return 0
                x = InStr(rstOLPayPalPmt!Contents, sCourtDatesID)
                    
                If x > 0 Then
                   vAmount = 0
                    Call pfCommunicationHistoryAdd("PayPalPayment") 'comms history entry for paypal email
                    
                    Call fPaymentAdd(sInvoiceNumber, vAmount) 'apply payment to sCourtDatesID
                    rstOLPayPalPmt.delete 'delete in OLPayPalPayments
                Else
                    
                    Call pfCommunicationHistoryAdd("PayPalInvoiceSent") 'comms history entry for paypal email
                    rstOLPayPalPmt.delete 'delete in OLPayPalPayments
                    
                End If
       
                rstOLPayPalPmt.MoveNext
                
            Loop
            
        Else
        End If
        
        sInvoiceNumber = ""
        sCourtDatesID = ""
        
        rstCourtDates.MoveNext
    Loop
            
Else

    MsgBox "There are no PayPal payments to process."
    
End If

MsgBox "Finished looping through PayPal Payments."

rstOLPayPalPmt.Close
Set rstOLPayPalPmt = Nothing
rstCourtDates.Close 'Close the recordset
Set rstCourtDates = Nothing 'Clean up

End Sub

Sub fGenerateInvoiceNumber()
'============================================================================
' Name        : fGenerateInvoiceNumber
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fGenerateInvoiceNumber
' Description : generates invoice number
'============================================================================

Dim rstTempCourtDates As DAO.Recordset, rstMaxCourtDates As DAO.Recordset, rstCourtDates As DAO.Recordset
Dim sQuestion As String, sAnswer As String

Set rstMaxCourtDates = CurrentDb.OpenRecordset("SELECT MAX(InvoiceNo) FROM CourtDates;")
sInvoiceNumber = rstMaxCourtDates.Fields(0).Value

Set rstCourtDates = CurrentDb.OpenRecordset("SELECT MAX(ID) as CourtDatesID FROM CourtDates;")
rstCourtDates.MoveFirst
sCourtDatesID = rstCourtDates.Fields("CourtDatesID").Value

sQuestion = "The most recent invoice number is " & sInvoiceNumber & _
".  Would you like to use " & sInvoiceNumber + 1 & " as your next invoice number?"
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

If sAnswer = vbNo Then 'Code for No
    sInvoiceNumber = InputBox("Enter your next Invoice Number.  The most recent one was " & sInvoiceNumber & ".")
Else 'Code for yes
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

End Sub



Sub fIsFactoringApproved()

'============================================================================
' Name        : fIsFactoringApproved
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fIsFactoringApproved
' Description : checks if factoring is approved for customer
'============================================================================





Dim vFA As String

Call pfGetOrderingAttorneyInfo

svURL = "https://www.paypal.com"

DoCmd.OpenQuery "UnitPriceQuery", acViewNormal, acReadOnly
DoCmd.OpenQuery "INVUpdateUnitPriceQuery", acViewNormal, acEdit


'check if approved for factoring
vFA = sFactoringApproved


If vFA = "True" Then 'if box checked, then do this

    'send price quote
        Call pfGenericExportandMailMerge("Invoice", "Stage1s\PriceQuote")
        Call pfCommunicationHistoryAdd("PriceQuote")
        
    'send adam CIDIncomeReport
        Call pfGenericExportandMailMerge("Invoice", "Stage1s\CIDIncomeReport")
        Call pfCommunicationHistoryAdd("CIDIncomeReport")
        Call pfSendWordDocAsEmail("CIDIncomeReport", "Initial Income Notification") 'initial income report 'emails adam cid report
        
Else 'if factoring not approved
                 
    'send deposit invoice (HTML copy in email and PDF copy)
        Call pfGenericExportandMailMerge("Invoice", "Stage1s\DepositInvoice") 'ISSUE LOCAL GENERATED INVOICE
        Application.FollowHyperlink (svURL) 'ISSUE UPDATED INVOICE
        Call pfCommunicationHistoryAdd("DepositInvoice") 'LOG UPDATED INVOICE
        Call pfGenericExportandMailMerge("Invoice", "Stage1s\CIDIncomeReport") 'ISSUE INCOME REPORT
        Call pfCommunicationHistoryAdd("CIDIncomeReport") 'LOG INCOME REPORT
        Call pfSendWordDocAsEmail("CIDIncomeReport", "Initial Income Notification") 'initial income report 'emails adam cid report
        Call fPPDraft 'issues paypal invoice
        Call fSendPPEmailDeposit 'sends paypal invoice email for deposit.
        
        
        
        
End If


End Sub


Sub fDepositPaymentReceived()

'============================================================================
' Name        : fDepositPaymentReceived
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fDepositPaymentReceived
' Description : does some things after a deposit is paid
'============================================================================

Dim rstQInfoInvNo As DAO.Recordset
Dim db As Database
Dim qdf As QueryDef



Dim vAmount As String, sQuestion As String, sAnswer As String

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

Set db = CurrentDb
Set qdf = db.QueryDefs("QInfobyInvoiceNumber")
qdf.Parameters(0) = sCourtDatesID
Set rstQInfoInvNo = qdf.OpenRecordset

sInvoiceNumber = rstQInfoInvNo.Fields("InvoiceNo").Value
sFinalPrice = rstQInfoInvNo.Fields("FinalPrice").Value

'calculate refund
vAmount = InputBox("How much was the payment?  Their invoice totals up to $" & sFinalPrice & ".")
    
sQuestion = "Are you sure that's correct?  The payment was in the amount of $" & vAmount & "?"
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

If sAnswer = vbNo Then 'Code for No

    sQuestion = "Do you want to enter a payment?"
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
    
    If sAnswer = vbNo Then 'Code for No

        GoTo Exitif
        
    Else 'Code for Yes

        vAmount = InputBox("How much was the payment?  Their invoice totals up to $" & sFinalPrice & ".")
        Call fPaymentAdd(sInvoiceNumber, vAmount)
        Call pfGenericExportandMailMerge("Invoice", "Stage1s\DepositPaid")
        Call pfCommunicationHistoryAdd("DepositPaid")
        
    End If
    
Else 'Code for Yes

    vAmount = InputBox("How much was the payment?  Their invoice totals up to $" & sFinalPrice & ".")
    Call fPaymentAdd(sInvoiceNumber, vAmount)
    Call pfGenericExportandMailMerge("Invoice", "Stage1s\DepositPaid")
    Call pfCommunicationHistoryAdd("DepositPaid")

End If
Exitif:
End Sub
Public Sub pfAutoCalculateFactorInterest()

'============================================================================
' Name        : pfAutoCalculateFactorInterest
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfAutoCalculateFactorInterest
' Description : add 1% after every 7 days payment not made
'============================================================================

Dim rstCustomers As DAO.Recordset, rstCourtDates As DAO.Recordset
Dim dInvoiceDate As Date
Dim iOrderingID As Integer, iDateDifference As Integer

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

End Sub
Sub fPaymentAdd(sInvoiceNumber As String, vAmount As String)

'============================================================================
' Name        : fPaymentAdd
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fPaymentAdd(sInvoiceNumber, vAmount)
' Description : adds payment to Payments table
'============================================================================

Dim db As DAO.Database
Dim rstPaymentAdd As DAO.Recordset
Dim sTableHyperlink As String, sPaymentMadePath As String


Call pfCurrentCaseInfo  'refresh transcript info

Set db = CurrentDb

sPaymentMadePath = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-PaymentMade.docx"
sTableHyperlink = sCourtDatesID & "-PaymentMade" & "#" & sPaymentMadePath & "#"

Set rstPaymentAdd = db.OpenRecordset("Payments")

rstPaymentAdd.AddNew
rstPaymentAdd("InvoiceNo").Value = sInvoiceNumber
rstPaymentAdd("Amount").Value = vAmount
rstPaymentAdd("RemitDate").Value = Date
rstPaymentAdd.Update

rstPaymentAdd.Close
db.Close

Call fManualPPPayment
Call pfClearGlobals
End Sub

Sub fUpdateFactoringDates()

'============================================================================
' Name        : fUpdateFactoringDates
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fUpdateFactoringDates
' Description : updates various factoring dates in CourtDates table
'============================================================================

Dim rstUnitPriceRate As DAO.Recordset
Dim dInvoiceDate As Date
Dim db As Database
Dim iUnitPriceID As Integer
Dim sExpectedAdvanceAmount As String, sExpectedRebateAmount As String, dExpectedBalanceDate As String
Dim sUnitPrice As String, sUnitPriceRateSQL As String, sCDCalcUpdateSQL As String

Call pfCurrentCaseInfo  'refresh transcript info

sExpectedRebateAmount = (sFinalPrice * 0.18)
sExpectedAdvanceAmount = (sFinalPrice * 0.8)
dInvoiceDate = Date
dExpectedBalanceDate = (Date + sTurnaroundTime) - 2
dExpectedRebateDate = DateAdd("d", 28, dExpectedAdvanceDate) 'dExpectedAdvanceDate + 28
dExpectedAdvanceDate = (Date + sTurnaroundTime) - 2
sEstimatedPageCount = ((sAudioLength / 60) * 45)


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

'    30 calendar-day turnaround, $3.00/page 39
'    14 calendar-day turnaround, $3.50/page 41
'    07 calendar-day turnaround, $4.00/page 62
'    03 calendar-day turnaround, $4.75/page 57
'    same day/overnight turnaround, $5.25/page 61


If sAudioLength > 865 Then
    
    If ((sJurisdiction) Like ("*" & "fda" & "*")) Then
        iUnitPriceID = 37
    ElseIf ((sJurisdiction) Like ("*" & "KCI" & "*")) Then iUnitPriceID = 40
    ElseIf ((sJurisdiction) Like ("*" & "AVT" & "*")) Then iUnitPriceID = 33
    ElseIf ((sJurisdiction) Like ("*" & "bankruptcy" & "*") And (sTurnaroundTime) Like ("30")) Then iUnitPriceID = 58
    ElseIf ((sJurisdiction) Like ("*" & "bankruptcy" & "*") And (sTurnaroundTime) Like ("14")) Then iUnitPriceID = 59
    ElseIf ((sJurisdiction) Like ("*" & "bankruptcy" & "*") And (sTurnaroundTime) Like ("7")) Then iUnitPriceID = 60
    ElseIf ((sJurisdiction) Like ("*" & "bankruptcy" & "*") And (sTurnaroundTime) Like ("3")) Then iUnitPriceID = 42
    ElseIf ((sJurisdiction) Like ("*" & "bankruptcy" & "*") And (sTurnaroundTime) Like ("1")) Then iUnitPriceID = 61
    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (sTurnaroundTime) Like ("30")) Then iUnitPriceID = 58
    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (sTurnaroundTime) Like ("14")) Then iUnitPriceID = 59
    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (sTurnaroundTime) Like ("7")) Then iUnitPriceID = 60
    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (sTurnaroundTime) Like ("3")) Then iUnitPriceID = 42
    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (sTurnaroundTime) Like ("1")) Then iUnitPriceID = 61
    ElseIf ((sJurisdiction) Like ("*" & "non-court" & "*") And (sTurnaroundTime) Like ("10")) Then iUnitPriceID = 49
    ElseIf ((sJurisdiction) Like ("*" & "non-court" & "*") And (sTurnaroundTime) Like ("1")) Then iUnitPriceID = 61
    ElseIf ((sJurisdiction) Like ("*" & "noncourt" & "*") And (sTurnaroundTime) Like ("10")) Then iUnitPriceID = 49
    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*")) Then iUnitPriceID = 37
    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (sTurnaroundTime) Like ("30")) Then iUnitPriceID = 58
    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (sTurnaroundTime) Like ("14")) Then iUnitPriceID = 59
    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (sTurnaroundTime) Like ("7")) Then iUnitPriceID = 60
    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (sTurnaroundTime) Like ("3")) Then iUnitPriceID = 42
    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (sTurnaroundTime) Like ("1")) Then iUnitPriceID = 61
    End If


Else
    
    If ((sJurisdiction) Like ("*" & "fda" & "*")) Then
        iUnitPriceID = 37
    ElseIf ((sJurisdiction) Like ("*" & "KCI" & "*")) Then iUnitPriceID = 40
    ElseIf ((sJurisdiction) Like ("*" & "AVT" & "*")) Then iUnitPriceID = 33
    ElseIf ((sJurisdiction) Like ("*" & "bankruptcy" & "*") And (sTurnaroundTime) Like ("30")) Then iUnitPriceID = 39
    ElseIf ((sJurisdiction) Like ("*" & "bankruptcy" & "*") And (sTurnaroundTime) Like ("14")) Then iUnitPriceID = 41
    ElseIf ((sJurisdiction) Like ("*" & "bankruptcy" & "*") And (sTurnaroundTime) Like ("7")) Then iUnitPriceID = 62
    ElseIf ((sJurisdiction) Like ("*" & "bankruptcy" & "*") And (sTurnaroundTime) Like ("3")) Then iUnitPriceID = 57
    ElseIf ((sJurisdiction) Like ("*" & "bankruptcy" & "*") And (sTurnaroundTime) Like ("1")) Then iUnitPriceID = 61
    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (sTurnaroundTime) Like ("30")) Then iUnitPriceID = 39
    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (sTurnaroundTime) Like ("14")) Then iUnitPriceID = 41
    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (sTurnaroundTime) Like ("7")) Then iUnitPriceID = 62
    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (sTurnaroundTime) Like ("3")) Then iUnitPriceID = 57
    ElseIf ((sJurisdiction) Like ("*" & "superior court" & "*") And (sTurnaroundTime) Like ("1")) Then iUnitPriceID = 61
    ElseIf ((sJurisdiction) Like ("*" & "non-court" & "*") And (sTurnaroundTime) Like ("10")) Then iUnitPriceID = 49
    ElseIf ((sJurisdiction) Like ("*" & "non-court" & "*") And (sTurnaroundTime) Like ("1")) Then iUnitPriceID = 61
    ElseIf ((sJurisdiction) Like ("*" & "noncourt" & "*") And (sTurnaroundTime) Like ("10")) Then iUnitPriceID = 49
    ElseIf ((sJurisdiction) Like ("*" & "Food and Drug Administration" & "*")) Then iUnitPriceID = 37
    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (sTurnaroundTime) Like ("30")) Then iUnitPriceID = 39
    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (sTurnaroundTime) Like ("14")) Then iUnitPriceID = 41
    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (sTurnaroundTime) Like ("7")) Then iUnitPriceID = 62
    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (sTurnaroundTime) Like ("3")) Then iUnitPriceID = 57
    ElseIf ((sJurisdiction) Like ("*" & "Massachusetts" & "*") And (sTurnaroundTime) Like ("1")) Then iUnitPriceID = 61
    End If

End If

'if non-court, use audio length as page count for rate calculation
If iUnitPriceID Like "49" And sTurnaroundTime <> "1" Then sEstimatedPageCount = sAudioLength

'get proper rate
sUnitPriceRateSQL = "SELECT Rate from UnitPrice where ID = " & iUnitPriceID & ";"

Set db = CurrentDb
Set rstUnitPriceRate = db.OpenRecordset(sUnitPriceRateSQL)
sUnitPrice = rstUnitPriceRate.Fields("Rate").Value
rstUnitPriceRate.Close

If iUnitPriceID = 49 Then sEstimatedPageCount = sAudioLength

sSubtotal = sEstimatedPageCount * sUnitPrice 'calculate total price estimate


If sSubtotal < 50 Then 'set minimum charge
    iUnitPriceID = 48
    sSubtotal = 50
End If

'insert calculated fields into courtdates
sCDCalcUpdateSQL = "UPDATE CourtDates SET [ExpectedRebateDate] = " & dExpectedRebateDate & ", [ExpectedAdvanceDate] = " & dExpectedAdvanceDate & ", [Subtotal] = " & sSubtotal & " WHERE ID = " & sCourtDatesID & ";"

db.Execute sCDCalcUpdateSQL
db.Close
Call pfClearGlobals

End Sub




Sub fTranscriptExpensesBeginning()

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
        
Dim db As DAO.Database
Dim rstExpensesAdd As DAO.Recordset, rstCourtDatesSet As DAO.Recordset
Dim vEPC As String

Call pfCurrentCaseInfo  'refresh transcript info

Set db = CurrentDb
Set rstExpensesAdd = db.OpenRecordset("Expenses")
    
vEPC = sEstimatedPageCount
vEPC = Int(vEPC)

If vEPC > 200 Then
 
     rstExpensesAdd.AddNew 'back cover
     rstExpensesAdd("Vendor").Value = "Got Print"
     rstExpensesAdd("ExpensesDate").Value = Now
     rstExpensesAdd("Amount").Value = 0.6
     rstExpensesAdd("Memo").Value = "Back Cover"
     rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
     rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
     rstExpensesAdd.Update

     rstExpensesAdd.AddNew 'back cover
     rstExpensesAdd("Vendor").Value = "Got Print"
     rstExpensesAdd("ExpensesDate").Value = Now
     rstExpensesAdd("Amount").Value = 0.6
     rstExpensesAdd("Memo").Value = "Back Cover"
     rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
     rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
     rstExpensesAdd.Update

     rstExpensesAdd.AddNew 'front cover
     rstExpensesAdd("Vendor").Value = "Amazon"
     rstExpensesAdd("ExpensesDate").Value = Now
     rstExpensesAdd("Amount").Value = 0.22
     rstExpensesAdd("Memo").Value = "front cover"
     rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
     rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
     rstExpensesAdd.Update

     rstExpensesAdd.AddNew 'front cover
     rstExpensesAdd("Vendor").Value = "Amazon"
     rstExpensesAdd("ExpensesDate").Value = Now
     rstExpensesAdd("Amount").Value = 0.22
     rstExpensesAdd("Memo").Value = "front cover"
     rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
     rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
     rstExpensesAdd.Update

     rstExpensesAdd.AddNew 'CD
     rstExpensesAdd("Vendor").Value = "Amazon"
     rstExpensesAdd("ExpensesDate").Value = Now
     rstExpensesAdd("Amount").Value = 0.18
     rstExpensesAdd("Memo").Value = "CD"
     rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
     rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
     rstExpensesAdd.Update

    rstExpensesAdd.AddNew 'CD sleeve
     rstExpensesAdd("Vendor").Value = "Amazon"
     rstExpensesAdd("ExpensesDate").Value = Now
     rstExpensesAdd("Amount").Value = 0.09
     rstExpensesAdd("Memo").Value = "CD Sleeve"
     rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
     rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
     rstExpensesAdd.Update

     rstExpensesAdd.AddNew 'shipping envelope
     rstExpensesAdd("Vendor").Value = "Amazon"
     rstExpensesAdd("ExpensesDate").Value = Now
     rstExpensesAdd("Amount").Value = 0.32
     rstExpensesAdd("Memo").Value = "Shipping Envelope"
     rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
     rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
     rstExpensesAdd.Update

     rstExpensesAdd.AddNew 'Velobind
     rstExpensesAdd("Vendor").Value = "Amazon"
     rstExpensesAdd("ExpensesDate").Value = Now
     rstExpensesAdd("Amount").Value = 0.72
     rstExpensesAdd("Memo").Value = "Velobind"
     rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
     rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
     rstExpensesAdd.Update

     rstExpensesAdd.AddNew 'Velobind
     rstExpensesAdd("Vendor").Value = "Amazon"
     rstExpensesAdd("ExpensesDate").Value = Now
     rstExpensesAdd("Amount").Value = 0.72
     rstExpensesAdd("Memo").Value = "Velobind"
     rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
     rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
     rstExpensesAdd.Update

     rstExpensesAdd.AddNew 'business card
     rstExpensesAdd("Vendor").Value = "Got Print"
     rstExpensesAdd("ExpensesDate").Value = Now
     rstExpensesAdd("Amount").Value = 0.05
     rstExpensesAdd("Memo").Value = "Business Card"
     rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
     rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
     rstExpensesAdd.Update
 
     rstExpensesAdd.AddNew 'shipping label
     rstExpensesAdd("Vendor").Value = "Amazon"
     rstExpensesAdd("ExpensesDate").Value = Now
     rstExpensesAdd("Amount").Value = 0.15
     rstExpensesAdd("Memo").Value = "Shipping Label"
     rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
     rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
     rstExpensesAdd.Update
 Else
 
    rstExpensesAdd.AddNew 'back cover
    rstExpensesAdd("Vendor").Value = "Got Print"
    rstExpensesAdd("ExpensesDate").Value = Now
    rstExpensesAdd("Amount").Value = 0.6
    rstExpensesAdd("Memo").Value = "Back Cover"
    rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
    rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
    rstExpensesAdd.Update

    rstExpensesAdd.AddNew 'front cover
    rstExpensesAdd("Vendor").Value = "Amazon"
    rstExpensesAdd("ExpensesDate").Value = Now
    rstExpensesAdd("Amount").Value = 0.22
    rstExpensesAdd("Memo").Value = "front cover"
    rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
    rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
    rstExpensesAdd.Update

    rstExpensesAdd.AddNew 'CD
    rstExpensesAdd("Vendor").Value = "Amazon"
    rstExpensesAdd("ExpensesDate").Value = Now
    rstExpensesAdd("Amount").Value = 0.18
    rstExpensesAdd("Memo").Value = "CD"
    rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
    rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
    rstExpensesAdd.Update

    rstExpensesAdd.AddNew 'CD sleeve
    rstExpensesAdd("Vendor").Value = "Amazon"
    rstExpensesAdd("ExpensesDate").Value = Now
    rstExpensesAdd("Amount").Value = 0.09
    rstExpensesAdd("Memo").Value = "CD Sleeve"
    rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
    rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
    rstExpensesAdd.Update

    rstExpensesAdd.AddNew 'shipping envelope
    rstExpensesAdd("Vendor").Value = "Amazon"
    rstExpensesAdd("ExpensesDate").Value = Now
    rstExpensesAdd("Amount").Value = 0.32
    rstExpensesAdd("Memo").Value = "Shipping Envelope"
    rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
    rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
    rstExpensesAdd.Update

    rstExpensesAdd.AddNew 'Velobind
    rstExpensesAdd("Vendor").Value = "Amazon"
    rstExpensesAdd("ExpensesDate").Value = Now
    rstExpensesAdd("Amount").Value = 0.72
    rstExpensesAdd("Memo").Value = "Velobind"
    rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
    rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
    rstExpensesAdd.Update

    rstExpensesAdd.AddNew 'business card
    rstExpensesAdd("Vendor").Value = "Got Print"
    rstExpensesAdd("ExpensesDate").Value = Now
    rstExpensesAdd("Amount").Value = 0.05
    rstExpensesAdd("Memo").Value = "Business Card"
    rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
    rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
    rstExpensesAdd.Update
    
End If

rstExpensesAdd.AddNew 'shipping label
rstExpensesAdd("Vendor").Value = "Amazon"
rstExpensesAdd("ExpensesDate").Value = Now
rstExpensesAdd("Amount").Value = 0.15
rstExpensesAdd("Memo").Value = "Shipping Label"
rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
rstExpensesAdd.Update
rstExpensesAdd.Close
        
MsgBox "Static Expenses Added!"
Call pfClearGlobals
End Sub
Sub fTranscriptExpensesAfter()

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

Dim db As DAO.Database
Dim rstExpensesAdd As DAO.Recordset

Call pfCurrentCaseInfo  'refresh transcript info

Set db = CurrentDb
Set rstExpensesAdd = db.OpenRecordset("Expenses")

rstExpensesAdd.AddNew 'static
rstExpensesAdd("Vendor").Value = "internet rent etc"
rstExpensesAdd("ExpensesDate").Value = Now
rstExpensesAdd("Amount").Value = 0.09 * sEstimatedPageCount
rstExpensesAdd("Memo").Value = "internet rent electricity website"
rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
rstExpensesAdd.Update

rstExpensesAdd.AddNew 'paper
rstExpensesAdd("Vendor").Value = "OfficeSupply.com"
rstExpensesAdd("ExpensesDate").Value = Now
rstExpensesAdd("Amount").Value = 0.01 * sEstimatedPageCount
rstExpensesAdd("Memo").Value = "paper"
rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
rstExpensesAdd.Update

rstExpensesAdd.AddNew 'ink
rstExpensesAdd("Vendor").Value = "OfficeSupply.com"
rstExpensesAdd("ExpensesDate").Value = Now
rstExpensesAdd("Amount").Value = 0.008
rstExpensesAdd("Memo").Value = "ink"
rstExpensesAdd("CourtDatesID").Value = sCourtDatesID
rstExpensesAdd("InvoiceNo").Value = sInvoiceNumber
rstExpensesAdd.Update
rstExpensesAdd.Close

MsgBox "Dynamic Expenses Added!"
Call pfClearGlobals
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
Dim db As DAO.Database
Dim rstExpenses As DAO.Recordset, rstShippingExpenseEntry As DAO.Recordset
Dim dExpenseIncurred As Date
Dim iExpenseAmount As Integer
Dim sVendorName As String, sExpenseMemo As String

Dim qdf1 As QueryDef

Set db = CurrentDb

Call pfCurrentCaseInfo  'refresh transcript info

'come back need query for shipping iExpenseAmount + info from CourtDates and Customers


Set qdf1 = db.QueryDefs("TR-Court-Q")
qdf1.Parameters(0) = sCourtDatesID
Set rstShippingExpenseEntry = qdf1.OpenRecordset


sVendorName = rstShippingExpenseEntry.Fields("OrderingID").Value
sInvoiceNumber = rstShippingExpenseEntry.Fields("InvoiceNo").Value


dExpenseIncurred = Date
iExpenseAmount = 0
sExpenseMemo = "Shipping Job No. " & sCourtDatesID & " Invoice " & sInvoiceNumber & " Tracking " & sTrackingNumber

'generates new entry for shipping XML in expenses table
Set rstExpenses = CurrentDb.OpenRecordset("Expenses")

rstExpenses.AddNew
rstExpenses("sVendorName").Value = sVendorName
rstExpenses("dExpenseIncurred").Value = dExpenseIncurred
rstExpenses("iExpenseAmount").Value = iExpenseAmount
rstExpenses("sExpenseMemo").Value = sExpenseMemo
rstExpenses.Update
Call pfClearGlobals
End Sub


