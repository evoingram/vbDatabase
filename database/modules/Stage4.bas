Attribute VB_Name = "Stage4"
'@Folder("Database.Production.Modules")
Option Compare Database
Option Explicit

'============================================================================
'class module cmStage4

'variables:
'   NONE

'functions:

    'pfStage4Ppwk:          Description:  completes all stage 4 tasks
    '                       Arguments:    NONE
    'pfNewZip:              Description:  creates empty ZIP file
    '                       Arguments:    sPath
    'fTranscriptDeliveryF:  Description:  parent function to deliver transcript electronically in various ways depending on jurisdiction
    '                       Arguments:    NONE
    'fAudioDone:            Description:  completes audio in express scribe
    '                       Arguments:    NONE
    'fRunXLSMacro:          Description:  parent function to ZIP various necessary files going to customer
    '                       Arguments:    sFile, sMacroName
    'pfSendTrackingEmail:   Description:  generates tracking number e-mail for customer
    '                       Arguments:    NONE
    'fZIPTranscripts:       Description:  zips transcripts folder in I:\####\
    '                       Arguments:    NONE
    'fZIPAudioTranscripts:  Description:  zips audio & transcripts folders in I:\####\
    '                       Arguments:    NONE
    'fZIPAudio:             Description:  zips audio folder in I:\####\
    '                       Arguments:    NONE
    'fUploadZIPsPrompt:     Description:  asks if you want to upload ZIPs to FTP
    '                       Arguments:    NONE
    'fUploadtoFTP:          Description:  uploads ZIPs to ftp
    '                       Arguments:    NONE
    'fGenerateZIPsF:        Description:  parent function to ZIP various necessary files going to customer
    '                        Arguments:   NONE
    'fEmailtoPrint:         Description:  sends an email to print@aquoco.co to be printed
    '                       Arguments:    sFiletoEmailPath
    'fDistiller:            Description:  distills for PDFs
    '                       Arguments:    sExportTopic
    'fPrint2upPDF:          Description:  prints 2-up transcript PDF
    '                       Arguments:    NONE
    'fPrint4upPDF:          Description:  prints 4-up transcript PDF
    '                       Arguments:    NONE
    'fAcrobatKCIInvoice:    Description:  inserts page count into KCI invoice
    '                       Arguments:    NONE
    'pfUpload:              Description:  sends to website ftp
    '                       Arguments:    mySession
    'fPrivatePrint:         Description:  prompts to send necessary transcript files to print@aquoco.co to be printed
    '                       Arguments:    NONE
    'fExportRecsToXML:      Description : exports ShippingOptionsQ entries to XML
    '                       Arguments:    NONE
    'fAppendXMLFiles:       Description : appends XML files
    '                       Arguments:    NONE
    'fCourtofAppealsIXML:   Description : creates Court of Appeals XML for shipping
    '                       Arguments:    NONE
    
'============================================================================



Public Sub pfStage4Ppwk()
'============================================================================
' Name        : pfStage4Ppwk
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfStage4Ppwk
' Description : completes all stage 4 tasks / button name is cmdStage4Paperwork
'============================================================================

Dim db As Database
Dim rs1 As DAO.Recordset
Dim qdf1 As QueryDef, qdf As QueryDef
Dim sAnswer As String, sQuestion As String
Dim sIPCompletedFolderPath As String, sCompletedFolderPath As String
Dim sFactoredChkBxSQL As String, sBillingURL As String
Dim sPaymentDueDate As Date
Call pfCurrentCaseInfo  'refresh transcript info

Call pfGetOrderingAttorneyInfo
Call pfCheckFolderExistence 'checks for job folder and creates it if not exists

If sJurisdiction Like "*AVT*" Then
    'paypal commands
    Call fPPDraft

    Call pfAcrobatGetNumPages(sCourtDatesID) 'GETS OFFICIAL PAGE COUNT AND UPDATES ACTUALQUANTITY
    
    Set db = CurrentDb
    Set qdf1 = CurrentDb.QueryDefs("PaymentQueryInvoiceInfo")
    Set qdf1.Parameters(0) = sCourtDatesID
    Set rs1 = qdf1.OpenRecordset
    
    sFinalPrice = rs1.Fields("FinalPrice").Value 'STORE FINAL PRICE IN VARIABLE
    sInvoiceNumber = rs1.Fields("InvoiceNo").Value 'STORE INVOICE NUMBER IN VARIABLE
    
    Set qdf1 = CurrentDb.QueryDefs("TR-AppAddr-Q")
    Set qdf1.Parameters(0) = sCourtDatesID
    Set rs1 = qdf1.OpenRecordset
    
    sFactoringApproved = rs1.Fields("FactoringApproved").Value 'STORE FACTORING APPROVED YES/NO IN VARIABLE
    
    Set rs1 = db.OpenRecordset("BalanceofPaymentsPerInvoiceQuery")
    sPaymentSum = Nz(rs1.Fields("PaymentSum").Value, 0) 'STORE SUM OF ALL PAYMENTS/REFUNDS IN VARIABLE
    
    Set qdf1 = CurrentDb.QueryDefs("UpdateInvoiceFPaymentDueDateQuery")
    Set qdf1.Parameters(0) = sCourtDatesID
    qdf1.Execute
    
    MsgBox "Time to deliver.  Next we will do factoring."
    
    Call pfAutoCalculateFactorInterest 'CALCULATES FACTORING COST TO US FOR DAYS FROM INVOICEDATE AND UPDATES DB
    Call fUpdateFactoringDates 'UPDATES CALCULATED DATES/AMOUNTS, ADVANCE/REBATE IN COURTDATES
    Call pfGenericExportandMailMerge("Invoice", "Stage4s\PP-FactoredInvoiceEmail") 'GENERATE FACTORED PP INVOICE EMAIL
    Call pfGenericExportandMailMerge("Invoice", "Stage4s\FactoredInvoiceLite") 'GENERATE FACTORED CLIENT INVOICE
    Call fSendPPEmailFactored 'paypal command
    Call pfCommunicationHistoryAdd("FactoredInvoiceLite") 'LOG FACTORED CLIENT INVOICE
    Call pfInvoicesCSV 'RUNS FACTORING AND XERO INVOICE CSVS
    Call fFactorInvoicEmailF 'GENERATES FACTOR INVOICE EMAIL, FACTOR INVOICE, LOGS IT, AND GOES TO FACTOR WEBSITE
    
    MsgBox "Go upload your Xero invoice and factoring CSVs." 'GO DO THIS AT THIS TIME
    
    Call pfGenericExportandMailMerge("Invoice", "Stage4s\CIDFinalIncomeReport") ' RUN INCOME REPORT
    Call pfCommunicationHistoryAdd("CIDFinalIncomeReport") 'LOG INCOME REPORT
    Call pfSendWordDocAsEmail("CIDFinalIncomeReport", "Final Income Notification") 'final income report 'emails adam cid report
    
    rs1.Close
    
ElseIf sJurisdiction Like "*eScribers*" Then
    
    'paypal commands
    Call fPPDraft

    Call pfGetOrderingAttorneyInfo
    Call pfCurrentCaseInfo 'refresh transcript info
    Call pfAcrobatGetNumPages(sCourtDatesID) 'GETS OFFICIAL PAGE COUNT AND UPDATES ACTUALQUANTITY
    
    Set db = CurrentDb
    Set qdf1 = CurrentDb.QueryDefs("PaymentQueryInvoiceInfo")
    Set qdf1.Parameters(0) = sCourtDatesID
    Set rs1 = qdf1.OpenRecordset
    
    sFinalPrice = rs1.Fields("FinalPrice").Value 'STORE FINAL PRICE IN VARIABLE
    sInvoiceNumber = rs1.Fields("InvoiceNo").Value 'STORE INVOICE NUMBER IN VARIABLE
    
    Set qdf1 = CurrentDb.QueryDefs("TR-AppAddr-Q")
    Set qdf1.Parameters(0) = sCourtDatesID
    Set rs1 = qdf1.OpenRecordset
    
    sFactoringApproved = rs1.Fields("FactoringApproved").Value 'STORE FACTORING APPROVED YES/NO IN VARIABLE
    
    Set rs1 = db.OpenRecordset("BalanceofPaymentsPerInvoiceQuery")
    sPaymentSum = Nz(rs1.Fields("PaymentSum").Value, 0) 'STORE SUM OF ALL PAYMENTS/REFUNDS IN VARIABLE
    
    Set qdf1 = CurrentDb.QueryDefs("UpdateInvoiceFPaymentDueDateQuery")
    Set qdf1.Parameters(0) = sCourtDatesID
    qdf1.Execute
    
    MsgBox "Time to deliver.  Next we will do factoring."
    
    Call pfAutoCalculateFactorInterest 'CALCULATES FACTORING COST TO US FOR DAYS FROM INVOICEDATE AND UPDATES DB
    Call fUpdateFactoringDates 'UPDATES CALCULATED DATES/AMOUNTS, ADVANCE/REBATE IN COURTDATES
    Call pfGenericExportandMailMerge("Invoice", "Stage4s\FactoredInvoiceLite") 'GENERATE FACTORED CLIENT INVOICE
    Call pfGenericExportandMailMerge("Invoice", "Stage4s\PP-FactoredInvoiceEmail") 'GENERATE FACTORED PP INVOICE EMAIL
    Call fSendPPEmailFactored 'paypal command
    Call pfCommunicationHistoryAdd("FactoredInvoiceLite") 'LOG FACTORED CLIENT INVOICE
    Call pfInvoicesCSV 'RUNS FACTORING AND XERO INVOICE CSVS
    Call fFactorInvoicEmailF 'GENERATES FACTOR INVOICE EMAIL, FACTOR INVOICE, LOGS IT, AND GOES TO FACTOR WEBSITE
    
    MsgBox "Go upload your Xero invoice and factoring CSVs." 'GO DO THIS AT THIS TIME
    
    Call pfGenericExportandMailMerge("Invoice", "Stage4s\CIDFinalIncomeReport") ' RUN INCOME REPORT
    Call pfCommunicationHistoryAdd("CIDFinalIncomeReport") 'LOG INCOME REPORT
    Call pfSendWordDocAsEmail("CIDFinalIncomeReport", "Final Income Notification") 'final income report 'emails adam cid report
    
    rs1.Close
    
ElseIf sJurisdiction Like "*FDA*" Then
    
    'paypal commands
    Call fPPDraft
    
    Call pfAcrobatGetNumPages(sCourtDatesID) 'GETS OFFICIAL PAGE COUNT AND UPDATES ACTUALQUANTITY
    Call pfGetOrderingAttorneyInfo 'STORE FACTORING APPROVED YES/NO IN VARIABLE
    
    Set db = CurrentDb
    Set qdf1 = CurrentDb.QueryDefs("BalanceofPaymentsPerInvoiceQuery")
    Set qdf1.Parameters(0) = sCourtDatesID
    Set rs1 = qdf1.OpenRecordset
    
    sPaymentSum = Nz(rs1.Fields("PaymentSum").Value, 0) 'STORE SUM OF ALL PAYMENTS/REFUNDS IN VARIABLE
    
    Set qdf1 = CurrentDb.QueryDefs("UpdateInvoiceFPaymentDueDateQuery")
    Set qdf1.Parameters(0) = sCourtDatesID
    qdf1.Execute
    
    MsgBox "Next we will do factoring and then deliver."
    
    Call pfAutoCalculateFactorInterest 'CALCULATES FACTORING COST TO US FOR DAYS FROM INVOICEDATE AND UPDATES DB
    Call fUpdateFactoringDates 'UPDATES CALCULATED DATES/AMOUNTS, ADVANCE/REBATE IN COURTDATES
    Call pfGenericExportandMailMerge("Invoice", "Stage4s\FactoredInvoiceLite") 'GENERATE FACTORED CLIENT INVOICE
    Call pfGenericExportandMailMerge("Invoice", "Stage4s\PP-FactoredInvoiceEmail") 'GENERATE FACTORED PP INVOICE EMAIL
    Call fSendPPEmailFactored 'paypal command
    Call pfCommunicationHistoryAdd("FactoredInvoiceLite") 'LOG FACTORED CLIENT INVOICE
    Call pfInvoicesCSV 'RUNS FACTORING AND XERO INVOICE CSVS
    Call fFactorInvoicEmailF 'GENERATES FACTOR INVOICE EMAIL, FACTOR INVOICE, LOGS IT, AND GOES TO FACTOR WEBSITE
    
    MsgBox "Go upload your Xero invoice and factoring CSVs." 'GO DO THIS AT THIS TIME
    
    Call pfGenericExportandMailMerge("Invoice", "Stage4s\CIDFinalIncomeReport") ' RUN INCOME REPORT
    Call pfCommunicationHistoryAdd("CIDFinalIncomeReport") 'LOG INCOME REPORT
    Call pfSendWordDocAsEmail("CIDFinalIncomeReport", "Final Income Notification") 'final income report 'emails adam cid report
    
ElseIf sJurisdiction Like "*Food and Drug Administration*" Then
    
    'paypal commands
    Call fPPDraft
    Call pfAcrobatGetNumPages(sCourtDatesID) 'GETS OFFICIAL PAGE COUNT AND UPDATES ACTUALQUANTITY
    Call pfGetOrderingAttorneyInfo 'STORE FACTORING APPROVED YES/NO IN VARIABLE
    
    Set db = CurrentDb
    Set qdf1 = CurrentDb.QueryDefs("BalanceofPaymentsPerInvoiceQuery")
    Set qdf1.Parameters(0) = sCourtDatesID
    Set rs1 = qdf1.OpenRecordset
    
    sPaymentSum = Nz(rs1.Fields("PaymentSum").Value, 0) 'STORE SUM OF ALL PAYMENTS/REFUNDS IN VARIABLE
    
    Set qdf1 = CurrentDb.QueryDefs("UpdateInvoiceFPaymentDueDateQuery")
    Set qdf1.Parameters(0) = sCourtDatesID
    qdf1.Execute
    
    MsgBox "Next we will do factoring and then deliver."
    
    Call pfAutoCalculateFactorInterest 'CALCULATES FACTORING COST TO US FOR DAYS FROM INVOICEDATE AND UPDATES DB
    Call fUpdateFactoringDates 'UPDATES CALCULATED DATES/AMOUNTS, ADVANCE/REBATE IN COURTDATES
    Call pfGenericExportandMailMerge("Invoice", "Stage4s\FactoredInvoiceLite") 'GENERATE FACTORED CLIENT INVOICE
    Call pfGenericExportandMailMerge("Invoice", "Stage4s\PP-FactoredInvoiceEmail") 'GENERATE FACTORED PP INVOICE EMAIL
    Call fSendPPEmailFactored 'paypal command
    Call pfCommunicationHistoryAdd("FactoredInvoiceLite") 'LOG FACTORED CLIENT INVOICE
    Call pfInvoicesCSV 'RUNS FACTORING AND XERO INVOICE CSVS
    Call fFactorInvoicEmailF 'GENERATES FACTOR INVOICE EMAIL, FACTOR INVOICE, LOGS IT, AND GOES TO FACTOR WEBSITE
    
    MsgBox "Go upload your Xero invoice and factoring CSVs." 'GO DO THIS AT THIS TIME
    
    Call pfGenericExportandMailMerge("Invoice", "Stage4s\CIDFinalIncomeReport") ' RUN INCOME REPORT
    Call pfCommunicationHistoryAdd("CIDFinalIncomeReport") 'LOG INCOME REPORT
    Call pfSendWordDocAsEmail("CIDFinalIncomeReport", "Final Income Notification") 'final income report 'emails adam cid report
    
ElseIf sJurisdiction Like "*Weber*" Then
    
    'paypal commands
    Call fPPDraft
        
    Call pfAcrobatGetNumPages(sCourtDatesID) 'GETS OFFICIAL PAGE COUNT AND UPDATES ACTUALQUANTITY
    Call pfGetOrderingAttorneyInfo 'STORE FACTORING APPROVED YES/NO IN VARIABLE
    
    Set db = CurrentDb
    Set qdf1 = CurrentDb.QueryDefs("BalanceofPaymentsPerInvoiceQuery")
    Set qdf1.Parameters(0) = sCourtDatesID
    Set rs1 = qdf1.OpenRecordset
    
    sPaymentSum = Nz(rs1.Fields("PaymentSum").Value, 0) 'STORE SUM OF ALL PAYMENTS/REFUNDS IN VARIABLE
    
    Set qdf1 = CurrentDb.QueryDefs("UpdateInvoiceFPaymentDueDateQuery")
    Set qdf1.Parameters(0) = sCourtDatesID
    qdf1.Execute
    
    MsgBox "Next we will do factoring and then deliver."
    
    Call pfAutoCalculateFactorInterest 'CALCULATES FACTORING COST TO US FOR DAYS FROM INVOICEDATE AND UPDATES DB
    Call fUpdateFactoringDates 'UPDATES CALCULATED DATES/AMOUNTS, ADVANCE/REBATE IN COURTDATES
    Call pfGenericExportandMailMerge("Invoice", "Stage4s\FactoredInvoiceLite") 'GENERATE FACTORED CLIENT INVOICE
    Call pfGenericExportandMailMerge("Invoice", "Stage4s\PP-FactoredInvoiceEmail") 'GENERATE FACTORED PP INVOICE EMAIL
    Call fSendPPEmailFactored 'paypal command
    Call pfCommunicationHistoryAdd("FactoredInvoiceLite") 'LOG FACTORED CLIENT INVOICE
    Call pfInvoicesCSV 'RUNS FACTORING AND XERO INVOICE CSVS
    Call fFactorInvoicEmailF 'GENERATES FACTOR INVOICE EMAIL, FACTOR INVOICE, LOGS IT, AND GOES TO FACTOR WEBSITE
    
    MsgBox "Go upload your Xero invoice and factoring CSVs." 'GO DO THIS AT THIS TIME
    
    Call pfGenericExportandMailMerge("Invoice", "Stage4s\CIDFinalIncomeReport") ' RUN INCOME REPORT
    Call pfCommunicationHistoryAdd("CIDFinalIncomeReport") 'LOG INCOME REPORT
    Call pfSendWordDocAsEmail("CIDFinalIncomeReport", "Final Income Notification") 'final income report 'emails adam cid report
    
Else

    'Call fPrivatePrint
    Call fTranscriptExpensesBeginning 'LOGS STATIC PER-TRANSCRIPT EXPENSES
    Call pfAcrobatGetNumPages(sCourtDatesID) 'GETS OFFICIAL PAGE COUNT AND UPDATES ACTUALQUANTITY
    
     Set qdf1 = CurrentDb.QueryDefs(qnTRCourtUnionAppAddrQ)
     
    Set qdf1.Parameters(0) = sCourtDatesID
    Set rs1 = qdf1.OpenRecordset
     sFinalPrice = rs1.Fields("FinalPrice").Value 'STORE FINAL PRICE IN VARIABLE
    sInvoiceNumber = rs1.Fields("InvoiceNo").Value 'STORE INVOICE NUMBER IN VARIABLE
     
    Set qdf1 = CurrentDb.QueryDefs("TR-AppAddr-Q")
    
     
    Set qdf1.Parameters(0) = sCourtDatesID
    Set rs1 = qdf1.OpenRecordset
    
    sFactoringApproved = rs1.Fields("FactoringApproved").Value
    rs1.Close
    Set qdf1 = CurrentDb.QueryDefs("BalanceofPaymentsPerInvoiceQuery")
    Set qdf1.Parameters(0) = sCourtDatesID
    Set rs1 = qdf1.OpenRecordset
    
    sPaymentSum = Nz(rs1.Fields("PaymentSum").Value, 0)
    sBalanceDue = sFinalPrice - sPaymentSum  'DETERMINES IF DELIVERY HELD OR NO AND REFUND OR BALANCE DUE
    
    If sFactoringApproved = True Then 'IF FACTORING APPROVED, DO THE FOLLOWING
        
        Set qdf = CurrentDb.QueryDefs("UpdateInvoiceFPaymentDueDateQuery")
        Set qdf.Parameters(0) = sCourtDatesID
        qdf.Execute
        
        MsgBox "Time to deliver.  Next we will do factoring."
        
       
        Call fPPDraft 'paypal command
        Call pfAutoCalculateFactorInterest 'CALCULATES FACTORING COST TO US FOR DAYS FROM INVOICEDATE AND UPDATES DB
        Call fUpdateFactoringDates 'UPDATES CALCULATED DATES/AMOUNTS, ADVANCE/REBATE IN COURTDATES
        Call pfGenericExportandMailMerge("Invoice", "Stage4s\FactoredInvoice") 'GENERATE FACTORED CLIENT INVOICE
        Call pfGenericExportandMailMerge("Invoice", "Stage4s\PP-FactoredInvoiceEmail") 'GENERATE PP INVOICE EMAIL
        Call fSendPPEmailFactored 'paypal command
        Call pfCommunicationHistoryAdd("FactoredInvoice") 'LOG FACTORED CLIENT INVOICE
        Call pfInvoicesCSV 'RUNS FACTORING AND XERO INVOICE CSVS
        Call fFactorInvoicEmailF 'GENERATES FACTOR INVOICE EMAIL, FACTOR INVOICE, LOGS IT, AND GOES TO FACTOR WEBSITE
        
        MsgBox "Go upload your Xero invoice and factoring CSVs." 'GO DO THIS AT THIS TIME
        
        Call pfGenericExportandMailMerge("Invoice", "Stage4s\CIDFinalIncomeReport") ' RUN INCOME REPORT
        Call pfCommunicationHistoryAdd("CIDFinalIncomeReport") 'LOG INCOME REPORT
        Call fTranscriptExpensesAfter 'LOGS DYNAMIC PER-TRANSCRIPT EXPENSES
        Call fTranscriptDeliveryF
        Call pfSendWordDocAsEmail("CIDFinalIncomeReport", "Final Income Notification") 'final income report 'emails adam cid report
        
        sQuestion = "Expenses logged.  Have you factored the transcript?"
        sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
        
        If sAnswer = vbNo Then 'Code for No
        
            MsgBox "Transcript will not be factored."
            
        Else
        
            Set db = CurrentDb
            
            sFactoredChkBxSQL = "update [CourtDates] set Factored =(Yes) WHERE ID=" & sCourtDatesID & ";"
            db.Execute sFactoredChkBxSQL
            MsgBox "Transcript has been marked factored."
            
        End If
        
    Else

        If sBalanceDue < 10 And sBalanceDue > 0 Then
        
            Call fTranscriptExpensesAfter 'LOGS DYNAMIC PER-TRANSCRIPT EXPENSES
            
            Set qdf = CurrentDb.QueryDefs("UpdateInvoicePPaymentDueDateQuery") 'UPDATE PAYMENTDUEDATE & INVOICEDATE
            Set qdf.Parameters(0) = sCourtDatesID
            qdf.Execute
            
            MsgBox "They owe less than $10.  Time to deliver."
            
            sBillingURL = "https://www.paypal.com"
            Application.FollowHyperlink (sBillingURL) 'ISSUE UPDATED INVOICE
            
            Call pfGenericExportandMailMerge("Invoice", "Stage4s\BalanceDue") 'RUN BALANCE DUE INVOICE
            Call pfGenericExportandMailMerge("Invoice", "Stage4s\PP-BalanceDueInvoiceEmail") 'GENERATE PP INVOICE EMAIL
            Call pfCommunicationHistoryAdd("BalanceDue")  'LOG BALANCE DUE REPORT
            Call pfInvoicesCSV 'RUNS FACTORING AND XERO INVOICE CSVS
            
            'balance due commands paypal
            Call fPPGetInvoiceInfo
            Call fPPUpdate
            Call fSendPPEmailBalanceDue
            
            MsgBox "Go upload your Xero invoice and factoring CSVs." 'GO DO THIS AT THIS TIME
            
            Call pfGenericExportandMailMerge("Invoice", "Stage4s\CIDFinalIncomeReport") ' RUN INCOME REPORT
            Call pfCommunicationHistoryAdd("CIDFinalIncomeReport") 'LOG INCOME REPORT
            Call pfSendWordDocAsEmail("CIDFinalIncomeReport", "Final Income Notification") 'final income report 'emails adam cid report
            Call fTranscriptDeliveryF
    
            
        ElseIf sBalanceDue <= 0 Then
        
            Call pfGenericExportandMailMerge("Invoice", "Stage4s\Refund") 'REPORT FOR ISSUING REFUND, paypal CSV
            Call pfGenericExportandMailMerge("Invoice", "Stage4s\PP-RefundMadeEmail") 'GENERATE PP INVOICE EMAIL
            Call pfCommunicationHistoryAdd("Refund") 'LOG ISSUING THE REFUND
            
            MsgBox "Issue refund in the amount of " & sBalanceDue & " for invoice number  " & sInvoiceNumber & " at PayPal.  Thank you."
            sBillingURL = "https://www.paypal.com"
            Application.FollowHyperlink (sBillingURL) 'ISSUE REFUND
            
            Call fPaymentAdd(sInvoiceNumber, "-" & sBalanceDue)  'FOR RECORDING REFUND
            Call fTranscriptDeliveryF
            
            'refund commands PAYPAL
            Call fPPGetInvoiceInfo
            Call fPPRefund
            Call pfSendWordDocAsEmail("PP-RefundMadeEmail", "Refund Issued")
            
        
        ElseIf sBalanceDue > 10 Then
            Set rs1 = CurrentDb.OpenRecordset("SELECT CourtDatesID, PaymentDueDate FROM InvoicePPaymentDueDateQuery WHERE CourtDatesID = " & sCourtDatesID & ";")
            sPaymentDueDate = rs1.Fields("PaymentDueDate").Value
            rs1.Close
        
             CurrentDb.Execute "UPDATE CourtDates SET PaymentDueDate = " & sPaymentDueDate & " WHERE ID = " & sCourtDatesID & ";"
            MsgBox "Hold delivery.  Send an invoice in the amount of $" & sBalanceDue & " at PayPal.  Thank you."
            sBillingURL = "https://www.paypal.com"
            Application.FollowHyperlink (sBillingURL) 'ISSUE UPDATED INVOICE
            
            Call pfGenericExportandMailMerge("Invoice", "Stage4s\BalanceDue") 'RUN BALANCE DUE INVOICE
            Call pfCommunicationHistoryAdd("BalanceDue")  'LOG BALANCE DUE REPORT
        
            'balance due commands paypal
            Call fPPGetInvoiceInfo
            Call fPPUpdate
            Call fSendPPEmailBalanceDue
            
        End If
        
    End If
    
    If (sJurisdiction) Like ("*" & "KCI" & "*") Then
        MsgBox "This transcript will be paid by the State, so we'll generate their invoice now."
        Call fAcrobatKCIInvoice
    End If
        
    Call pfUpdateCheckboxStatus("FileTranscript")
    
    Forms![NewMainMenu]![ProcessJobSubformNMM].SourceObject = "PJShippingInfo"
    Forms![NewMainMenu]![ProcessJobSubformNMM].Requery
    
    Call pfUpdateCheckboxStatus("ShippingXMLs")
    
    sBillingURL = "https://go.xero.com/AccountsReceivable/Search.aspx?invoiceStatus=INVOICESTATUS%2fDRAFT&graphSearch=False"
    Application.FollowHyperlink (sBillingURL) 'GO TO XERO WEBSITE
    
    Call pfUpdateCheckboxStatus("InvoiceCompleted") 'CHECK INVOICE BOX
    
End If

'when done, move folder to /completed/ and change document hyperlinks from /in progress/ in communicationhistory
sQuestion = "Do you want to move this job to /3Complete/?"
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

If sAnswer = vbNo Then 'Code for No
    MsgBox "Job " & sCourtDatesID & " will not be completed."
Else
    Set db = CurrentDb
    sIPCompletedFolderPath = "I:\" & sCourtDatesID & "\*.*"
    sCompletedFolderPath = "T:\Production\3Complete\" & sCourtDatesID & "\*.*"
    Shell "cmd /c move '" & sIPCompletedFolderPath & "' " & sCompletedFolderPath & _
        ", vbNormalFocus"
    db.Execute "Update CommunicationHistory Set [FileHyperlink] = Replace(FileHyperlink, " & "'2InProgress\" & sCourtDatesID & "', '3Complete\" & sCourtDatesID & "') WHERE fileHyperLink LIKE '*2InProgress\" & sCourtDatesID & "*';"

    MsgBox "Job " & sCourtDatesID & " has been moved to /3Complete/ and document history hyperlinks have been updated."
End If

MsgBox "Stage 4 complete."
Call pfClearGlobals
End Sub
Public Sub pfNewZip(sPath As String)
'============================================================================
' Name        : pfNewZip
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfNewZip(sPath)
' Description : creates empty ZIP file
'============================================================================

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

If Len(Dir(sPath)) > 0 Then Kill sPath

Open sPath For Output As #1

Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
Close #1

End Sub

Sub fTranscriptDeliveryF()
'============================================================================
' Name        : fTranscriptDeliveryF
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fTranscriptDeliveryF
' Description : parent function to deliver transcript electronically in various ways depending on jurisdiction
'============================================================================
'
Dim sQuestion As String, sAnswer As String, sFiledNotFiledSQL As String
Dim sPDFFinalTranscript As String, sWordFinalTranscript As String, sWorkingCopyPath As String
Dim sInvoiceWD As String, sInvoicePDF As String
Dim db As Database
Dim oWordApp As New Word.Application, oWordDoc As New Word.Document

Call pfCurrentCaseInfo  'refresh transcript info

sPDFFinalTranscript = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-FINAL.pdf"
sWordFinalTranscript = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-FINAL.docx"
sWorkingCopyPath = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-WorkingCopy.docx"
sInvoiceWD = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-" & sInvoiceNumber & ".docx"
sInvoicePDF = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-" & sInvoiceNumber & ".pdf"

'checks for Audio, Transcripts, FTP, WorkingFiles subfolders and creates if not exists
Call pfCheckFolderExistence

sQuestion = "Have you filed or are you filing the transcript?"
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
If sJurisdiction = "*AVT*" Then

    Application.FollowHyperlink ("http://tabula.escribers.net/")
    GoTo ContractorFile

ElseIf sJurisdiction = "Food and Drug Administration" Then
    
    sAnswer = vbNo
    GoTo ContractorFile

ElseIf sJurisdiction = "*FDA*" Then
     
     sAnswer = vbNo
     GoTo ContractorFile

ElseIf sJurisdiction = "Weber Oregon" Then
    
    sAnswer = vbNo
    Application.FollowHyperlink ("https://app.therecordxchange.net/myjobs/active")
    GoTo ContractorFile

ElseIf sJurisdiction = "Weber Nevada" Then
    
    sAnswer = vbNo
    Application.FollowHyperlink ("https://app.therecordxchange.net/myjobs/active")
    GoTo ContractorFile

ElseIf sJurisdiction = "Weber Bankruptcy" Then
    
    sAnswer = vbNo
    Application.FollowHyperlink ("https://app.therecordxchange.net/myjobs/active")
    GoTo ContractorFile

Else
    
    If sJurisdiction = "King County Superior Court" Then
    
        Application.FollowHyperlink ("https://ac.courts.wa.gov/index.cfm?fa=efiling.home")
        
    ElseIf sJurisdiction = "District of Hawaii" Then
    
        Application.FollowHyperlink ("https://ecf.hib.uscourts.gov/cgi-bin/login.pl")
        Call pfSendWordDocAsEmail("TranscriptsReady", "Transcripts Ready", sPDFFinalTranscript, sWordFinalTranscript, sWorkingCopyPath)
        
    ElseIf sJurisdiction = "Eastern District of Pennsylvania" Then
    
        Application.FollowHyperlink ("https://ecf.paeb.uscourts.gov/cgi-bin/login.pl")
        'Call FileTranscriptSendEmail(sCompanyEmail)
        Call pfSendWordDocAsEmail("TranscriptsReady", "Transcripts Ready", sPDFFinalTranscript, sWordFinalTranscript, sWorkingCopyPath)
        
    ElseIf sJurisdiction = "District of Connecticut" Then
    
        Application.FollowHyperlink ("https://ecf.ctb.uscourts.gov/cgi-bin/login.pl")
    
    ElseIf sJurisdiction = "Southern District of Alabama" Then
    
        Application.FollowHyperlink ("https://ecf.alsb.uscourts.gov/cgi-bin/login.pl")
    
    ElseIf sJurisdiction = "Eastern District of Arkansas" Then
    
        Application.FollowHyperlink ("https://ecf.areb.uscourts.gov/cgi-bin/login.pl")
    
    ElseIf sJurisdiction = "Southern District of California" Then
    
        Application.FollowHyperlink ("https://ecf.casb.uscourts.gov/cgi-bin/login.pl")
    
    ElseIf sJurisdiction = "Eastern District of California" Then
    
        Application.FollowHyperlink ("https://efiling.caeb.uscourts.gov/LoginPage.aspx")
        'Call FileTranscriptSendEmail(sCompanyEmail)
        Call pfSendWordDocAsEmail("TranscriptsReady", "Transcripts Ready", sPDFFinalTranscript, sWordFinalTranscript, sWorkingCopyPath)
        
    ElseIf sJurisdiction = "Southern District of California" Then
    
        Application.FollowHyperlink ("https://ecf.casb.uscourts.gov/cgi-bin/login.pl")
        
    ElseIf sJurisdiction = "District of Hawaii" Then
    
        Application.FollowHyperlink ("https://efiling.caeb.uscourts.gov/LoginPage.aspx")
        'Call FileTranscriptSendEmail(sCompanyEmail)
        Call pfSendWordDocAsEmail("TranscriptsReady", "Transcripts Ready", sPDFFinalTranscript, sWordFinalTranscript, sWorkingCopyPath)
        
    ElseIf sJurisdiction = "Central District of Illinois" Then
    
        Application.FollowHyperlink ("https://ecf.ilcb.uscourts.gov/cgi-bin/login.pl")
    
    ElseIf sJurisdiction = "Southern District of Illinois" Then
    
        Application.FollowHyperlink ("https://ecf.ilsb.uscourts.gov/cgi-bin/login.pl")
    
    ElseIf sJurisdiction = "Northern District of Iowa" Then
    
        Application.FollowHyperlink ("https://ecf.ianb.uscourts.gov/cgi-bin/login.pl")
    
    ElseIf sJurisdiction = "District of Kansas" Then
    
        Application.FollowHyperlink ("https://ecf.ksb.uscourts.gov/cgi-bin/login.pl")
    
    ElseIf sJurisdiction = "Eastern District of Kentucky" Then
    
        Application.FollowHyperlink ("https://ecf.kyeb.uscourts.gov/cgi-bin/login.pl")
    
    ElseIf sJurisdiction = "Middle District of Louisiana" Then
    
        Application.FollowHyperlink ("https://ecf.lamb.uscourts.gov/cgi-bin/login.pl")
    
    ElseIf sJurisdiction = "Western District of Louisiana" Then
    
        Application.FollowHyperlink ("https://ecf.lawb.uscourts.gov/cgi-bin/login.pl")
    
    ElseIf sJurisdiction = "District of Minnesota" Then
    
        Application.FollowHyperlink ("https://ecf.mnb.uscourts.gov/cgi-bin/login.pl")
    
    ElseIf sJurisdiction = "District of Nebraska" Then
    
        Application.FollowHyperlink ("https://ecf.neb.uscourts.gov/cgi-bin/login.pl")
    
    ElseIf sJurisdiction = "District of New Mexico" Then
    
        Application.FollowHyperlink ("https://ecf.nmb.uscourts.gov/cgi-bin/login.pl")
    
    ElseIf sJurisdiction = "District of New York" Then
    
        Application.FollowHyperlink ("https://ecf.nynb.uscourts.gov/cgi-bin/login.pl")
    
    ElseIf sJurisdiction = "Middle District of North Carolina" Then
    
        Application.FollowHyperlink ("https://ecf.ncmb.uscourts.gov/cgi-bin/login.pl")
    
    ElseIf sJurisdiction = "District of North Dakota" Then
    
        Application.FollowHyperlink ("https://ecf.ndb.uscourts.gov/cgi-bin/login.pl")
    
    ElseIf sJurisdiction = "District of Oregon" Then
    
        Application.FollowHyperlink ("https://ecf.orb.uscourts.gov/cgi-bin/login.pl")
        Call pfSendWordDocAsEmail("TranscriptsReady", "Transcripts Ready", sPDFFinalTranscript, sWordFinalTranscript, sWorkingCopyPath)
        'Call FileTranscriptSendEmail(sCompanyEmail)
    
    ElseIf sJurisdiction = "District of Rhode Island" Then
    
        Application.FollowHyperlink ("https://ecf.rib.uscourts.gov/cgi-bin/login.pl")
        
    ElseIf sJurisdiction = "Western District of Washington" Then
    
        Application.FollowHyperlink ("https://ecf.wawb.uscourts.gov/cgi-bin/login.pl")
        
    Else
    
        Application.FollowHyperlink ("https://ac.courts.wa.gov/index.cfm?fa=efiling.home")
        
    End If
    
    'creates TranscriptReadyEmail
    Call pfGenericExportandMailMerge("Case", "Stage4s\TranscriptsReady")
    Call pfCommunicationHistoryAdd("TranscriptsReady")
    
    sQuestion = "Print transcript?"
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
    
    If sAnswer = vbNo Then 'Code for No
        MsgBox "Transcript will not print."
    Else 'Code for Yes
    
        Call fEmailtoPrint(sPDFFinalTranscript)
        Call fEmailtoPrint(sPDFFinalTranscript)
        
        Set oWordApp = Nothing
        Set oWordApp = GetObject(, "Word.Application")
                If oWordApp Is Nothing Then
            Set oWordApp = CreateObject("Word.Application")
        End If
        oWordApp.Application.Visible = False
        Set oWordDoc = oWordApp.Documents.Open(sInvoiceWD)
        oWordDoc.SaveAs2 FileName:=sInvoicePDF
        
        
        oWordApp.Quit
        Set oWordApp = Nothing
        
        Call pfSendWordDocAsEmail("TranscriptsReady", "Transcripts Ready", sPDFFinalTranscript, sWordFinalTranscript, sWorkingCopyPath, sInvoicePDF)
        
    End If
End If
ContractorFile:
    If sAnswer = vbNo Then
'Code for No
        MsgBox "Transcript will not be filed."
    Else
        Set db = CurrentDb
        sFiledNotFiledSQL = "update [CourtDates] set FiledNotFiled =(Yes) WHERE ID=" & sCourtDatesID & ";"
        db.Execute sFiledNotFiledSQL
        MsgBox "Transcript has been marked filed."
    End If
Call pfClearGlobals
End Sub

Sub fGenerateZIPsF()
'============================================================================
' Name        : fGenerateZIPsF
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fGenerateZIPsF
' Description : parent function to ZIP various necessary files going to customer
'============================================================================

Dim TranscriptWC As String, FindReplaceTranscriptD As String, FinalTranscriptPathwordWC As String
Dim FindReplaceTranscriptF As String, naceTranscriptP As String, vInvoiceFilePathPPDF As String
Dim sourceFile As String, destinationfile As String, sourcefile1 As String, destinationfile1 As String
Dim filecopied As Object
Dim MyNote As String, answer As String

Call pfCurrentCaseInfo  'refresh transcript info

Call pfCheckFolderExistence 'checks for job folder and creates it if not exists

If sJurisdiction Like "*Weber Nevada*" Or sJurisdiction Like "*Weber Bankruptcy*" Or sJurisdiction Like "*Weber Oregon*" Or sJurisdiction Like "*Food and Drug Administration*" Or sJurisdiction Like "*FDA*" Or sJurisdiction Like "*AVT*" Or sJurisdiction Like "*eScribers*" Or sJurisdiction Like "*AVTranz*" Then
    GoTo Line2
Else
End If

Call fCreateWorkingCopy
Call pfCreateRegularPDF

Line2:

Call fAudioDone

MsgBox "Check and make sure your transcript files look fine before hitting 'okay'." 'GO DO THIS AT THIS TIME

Call fPrint2upPDF
Call fPrint4upPDF

FileCopy "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-WordIndex.PDF", "I:\" & sCourtDatesID & "\Backups\" & sCourtDatesID & "-WordIndex.PDF"

MsgBox "Thank you.  Next, we will ZIP your files."

Call fZIPAudio 'zip audio folder
Call fZIPTranscripts 'zip transcripts folder
Call fZIPAudioTranscripts 'zip audio and transcripts folders together

MsgBox "Files Zipped.  Next, we will upload the ZIPs via FTP."

Call fUploadZIPsPrompt 'upload zips to ftp
'Call pfBurnCD 'burn CD

MyNote = "Do you want to open the job folder?"
answer = MsgBox(MyNote, vbQuestion + vbYesNo, "???")

If answer = vbNo Then 'Code for No

    MsgBox "Go to I:\ to open the job folder."
    
Else 'Code for yes

    Call Shell("explorer.exe" & " " & "I:\" & sCourtDatesID, vbNormalFocus)
    
End If

Call fAssignPS
Call pfClearGlobals
End Sub

Sub fUploadtoFTP()
'============================================================================
' Name        : fUploadtoFTP
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fUploadtoFTP
' Description : uploads ZIPs to ftp
'============================================================================

 Dim mySession As New Session

' Enable custom error handling
On Error Resume Next

pfUpload mySession

' Query for errors
If Err.Number <> 0 Then
    MsgBox "Error: " & Err.Description

    ' Clear the error
    Err.Clear
End If
 
' Disconnect, clean up
mySession.Dispose
 
' Restore default error handling
On Error GoTo 0

End Sub
 
Sub fUploadZIPsPrompt()
'============================================================================
' Name        : fUploadZIPsPrompt
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fUploadZIPsPrompt
' Description : asks if you want to upload ZIPs to FTP prompt to ftp zip yes or no
'============================================================================

Dim sAnswer As String, sQuestion As String
 
sQuestion = "Do you want to upload to FTP?"
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

If sAnswer = vbNo Then 'Code for No

    MsgBox "No files will be uploaded to FTP."

Else 'Code for yes
    
    Call fUploadtoFTP

End If

End Sub

Sub fZIPAudio()
'============================================================================
' Name        : fZIPAudio
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fZIPAudio
' Description : zips audio folder in I:\####\
'============================================================================

Dim sourceFile As String, destinationfile As String, sourcefile1 As String, destinationfile1 As String
Dim filecopied As Object
Dim FileNameZip1 As String, foldername1 As String, filenamezipFTPTRS As String, foldernameFTP As String
Dim dbVideoCollection As Database, rstCourtDates As DAO.Recordset, defpath As String, strDate As Date
Dim oApp As Object
'TODO: Universal Change dbVideoCollection database/other db names to proper name
Dim dbVideoCollection As Database
Set dbVideoCollection = CurrentDb
Set rstCourtDates = dbVideoCollection.OpenRecordset("CourtDates")
sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
defpath = CurrentProject.Path

If Right(defpath, 1) <> "\" Then
    defpath = defpath & "\"
End If

foldername1 = "I:\" & sCourtDatesID & "\Audio\"
'@Ignore AssignmentNotUsed, AssignmentNotUsed
foldernameFTP = "I:\" & sCourtDatesID & "\FTP"

strDate = Format(Now, " dd-mmm-yy h-mm-ss")
FileNameZip1 = "I:\" & sCourtDatesID & "\Backups\" & sCourtDatesID & "-Audio" & ".zip"
filenamezipFTPTRS = "I:\" & sCourtDatesID & "\FTP\" & sCourtDatesID & "-Audio" & ".zip"

Call pfNewZip(FileNameZip1) 'create empty zip file
Call pfNewZip(filenamezipFTPTRS) 'create empty zip file

Set oApp = CreateObject("Shell.Application")

'Copy the files to the compressed folder
oApp.Namespace("I:\" & sCourtDatesID & "\Backups\" & sCourtDatesID & "-Audio" & ".zip").CopyHere oApp.Namespace("I:\" & sCourtDatesID & "\Audio\").Items

foldernameFTP = "I:\" & sCourtDatesID & "\FTP\"    '<< Change
filenamezipFTPTRS = foldernameFTP & sCourtDatesID & "-Audio" & ".zip"

oApp.Namespace("I:\" & sCourtDatesID & "\FTP\" & sCourtDatesID & "-Audio" & ".zip").CopyHere oApp.Namespace("I:\" & sCourtDatesID & "\Audio\").Items


While oApp.Namespace("I:\" & sCourtDatesID & "\Backups\" & sCourtDatesID & "-Audio" & ".zip").Items.Count <> oApp.Namespace("I:\" & sCourtDatesID & "\Audio\").Items.Count

DoEvents
 Wend

'While oApp.Namespace("I:\" & sCourtDatesID & "\FTP\" & sCourtDatesID & "-Audio" & ".zip").Items.Count <> oApp.Namespace("I:\" & sCourtDatesID & "\Audio\").Items.Count
    'DoEvents
    
    'Wend
    
    
    
    'come back
    
    
MsgBox "You find the ZIP file here: " & FileNameZip1
End Sub

Sub fZIPAudioTranscripts()
'============================================================================
' Name        : fZIPAudioTranscripts
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fZIPAudioTranscripts
' Description : zips audio & transcripts folders in I:\####\
'============================================================================

Dim sourceFile As String, destinationfile As String, sourcefile1 As String, destinationfile1 As String
Dim strDate As String, defpath As String, foldernameaudio As String, foldernameTranscripts As String
Dim FileNameZip2 As String, FolderName2 As String, filenamezipFTP As String, foldernameFTP As String
Dim filecopied As Object, oApp As Object
Dim FileNameZipATRS As String, FileNameZipFTPATRS As String
sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

defpath = CurrentProject.Path
If Right(defpath, 1) <> "\" Then
    defpath = defpath & "\"
End If
strDate = Format(Now, " dd-mmm-yy h-mm-ss")
    
foldernameaudio = "I:\" & sCourtDatesID & "\Audio\"
foldernameTranscripts = "I:\" & sCourtDatesID & "\Transcripts\"
foldernameFTP = "I:\" & sCourtDatesID & "\FTP\"

FileNameZipATRS = "I:\" & sCourtDatesID & "\Backups\" & sCourtDatesID & "-AudioTranscripts" & ".zip"
FileNameZipFTPATRS = foldernameFTP & sCourtDatesID & "-AudioTranscripts" & ".zip"


Call pfNewZip(FileNameZipATRS) 'create empty zip files
Call pfNewZip(FileNameZipFTPATRS)

Set oApp = CreateObject("Shell.Application")

FolderName2 = (oApp.Namespace(foldernameaudio).Items.Count) + (oApp.Namespace(foldernameTranscripts).Items.Count)

'Copy the files to the compressed folder
oApp.Namespace(FileNameZipATRS).CopyHere oApp.Namespace(foldernameTranscripts).Items

While oApp.Namespace(foldernameTranscripts).Items.Count <> oApp.Namespace(FileNameZipATRS).Items.Count
    DoEvents
Wend
On Error GoTo 0
oApp.Namespace(FileNameZipATRS).CopyHere oApp.Namespace(foldernameaudio).Items

While oApp.Namespace(FileNameZipFTPATRS).Items.Count <> oApp.Namespace(FileNameZipATRS).Items.Count
    DoEvents
Wend
On Error GoTo 0
oApp.Namespace(FileNameZipFTPATRS).CopyHere oApp.Namespace(foldernameaudio).Items

While oApp.Namespace(foldernameaudio).Items.Count <> oApp.Namespace(FileNameZipFTPATRS).Items.Count
DoEvents
Wend
On Error GoTo 0
oApp.Namespace(FileNameZipFTPATRS).CopyHere oApp.Namespace(foldernameTranscripts).Items

On Error Resume Next
While oApp.Namespace(FileNameZipATRS).Items.Count <> oApp.Namespace(FileNameZipFTPATRS).Items.Count
    DoEvents
Wend
On Error GoTo 0





    

MsgBox "You find the zipfile here: " & FileNameZipATRS

End Sub

Sub fZIPTranscripts()
'============================================================================
' Name        : fZIPTranscripts
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fZIPTranscripts
' Description : zips transcripts folder in I:\####\
'============================================================================

Dim sourceFile As String, destinationfile As String, sourcefile1 As String, destinationfile1 As String
Dim filecopied As Object, oApp As Object
Dim FileNameZip1 As String, foldername1 As String, filenamezipFTPTRS As String, foldernameFTP As String
Dim dbVideoCollection As Database
Dim rstCourtDates As DAO.Recordset
Dim defpath As String
Dim strDate As Date
Set dbVideoCollection = CurrentDb
Set rstCourtDates = dbVideoCollection.OpenRecordset("CourtDates")
sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

defpath = CurrentProject.Path
If Right(defpath, 1) <> "\" Then
    defpath = defpath & "\"
End If

'@Ignore AssignmentNotUsed
foldername1 = "I:\" & sCourtDatesID & "\Transcripts\"
foldernameFTP = "I:\" & sCourtDatesID & "\FTP"

strDate = Format(Now, " dd-mmm-yy h-mm-ss")
FileNameZip1 = "I:\" & sCourtDatesID & "\Backups\" & sCourtDatesID & "-Transcripts" & ".zip"
filenamezipFTPTRS = "I:\" & sCourtDatesID & "\FTP\" & sCourtDatesID & "-Transcripts" & ".zip"

'Create empty Zip File
Call pfNewZip(FileNameZip1)
Call pfNewZip(filenamezipFTPTRS)

Set oApp = CreateObject("Shell.Application")

'Copy the files to the compressed folder
oApp.Namespace("I:\" & sCourtDatesID & "\Backups\" & sCourtDatesID & "-Transcripts" & ".zip").CopyHere oApp.Namespace("I:\" & sCourtDatesID & "\Transcripts\").Items

foldernameFTP = "I:\" & sCourtDatesID & "\FTP\"
filenamezipFTPTRS = foldernameFTP & sCourtDatesID & "-Transcripts" & ".zip"
oApp.Namespace("I:\" & sCourtDatesID & "\FTP\" & sCourtDatesID & "-Transcripts" & ".zip").CopyHere oApp.Namespace("I:\" & sCourtDatesID & "\Transcripts\").Items

On Error Resume Next
While oApp.Namespace("I:\" & sCourtDatesID & "\Backups\" & sCourtDatesID & "-Transcripts" & ".zip").Items.Count <> oApp.Namespace("I:\" & sCourtDatesID & "\Transcripts\").Items.Count
    DoEvents
Wend
On Error GoTo 0

On Error Resume Next
While oApp.Namespace("I:\" & sCourtDatesID & "\FTP\" & sCourtDatesID & "-Transcripts" & ".zip").Items.Count <> oApp.Namespace("I:\" & sCourtDatesID & "\Transcripts\").Items.Count
    DoEvents
Wend
On Error GoTo 0

MsgBox "You find the ZIP file here: " & FileNameZip1

End Sub
Sub fRunXLSMacro(sFile As String, sMacroName As String)
On Error GoTo eHandler
'============================================================================
' Name        : fGenerateZIPsF
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fRunXLSMacro(sFile, sMacroName)
' Description : runs XLS macro from XLS file path provided
'============================================================================

Dim oExcelApp As New Excel.Application, oExcelWkbk As New Excel.Workbook
Dim sFileName   As String
 
'Set oExcelApp = CreateObject("Excel.Application")
Set oExcelWkbk = oExcelApp.Workbooks.Open(sFile, True)
oExcelApp.Visible = True

sFileName = Right(sFile, Len(sFile) - InStrRev(sFile, "\"))

oExcelApp.Run sFileName & "!" & sMacroName


eHandlerX:

On Error Resume Next
oExcelWkbk.Close (True)
oExcelApp.Quit
On Error GoTo 0
Set oExcelWkbk = Nothing
Set oExcelApp = Nothing
Exit Function
 
eHandler:

MsgBox "The following error has occured." & vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Error Source: RunXLSMacro" & vbCrLf & _
        "Error Description: " & Err.Description, _
        vbCritical, "An Error has Occured!"
        
Resume eHandlerX

End Sub
Public Sub pfSendTrackingEmail()
'============================================================================
' Name        : pfSendTrackingEmail
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfSendTrackingEmail
' Description : generates tracking number e-mail for customer
'============================================================================

Dim Rng As Range
Dim vTrackingNumber As String, deliverySQLstring As String
Dim rs As DAO.Recordset



Call pfCurrentCaseInfo  'refresh transcript info

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
deliverySQLstring = "SELECT * FROM CourtDates WHERE [ID] = " & sCourtDatesID & ";"
'TODO: pfSendTrackingEmail get current values and delete following come back
Set rs = CurrentDb.OpenRecordset(deliverySQLstring)
vTrackingNumber = rs.Fields("TrackingNumber").Value
sParty1 = rs.Fields("Party1").Value
sParty2 = rs.Fields("Party2").Value
sCaseNumber1 = rs.Fields("CaseNumber1").Value
dHearingDate = rs.Fields("HearingDate").Value
sAudioLength = rs.Fields("AudioLength").Value

Call pfSendWordDocAsEmail("Shipped", "Transcript Shipped")
Call fWunderlistAdd(sCourtDatesID & ":  Package to Ship", Format(Now + 1, "yyyy-mm-dd"))
Call pfClearGlobals
End Sub
Sub fEmailtoPrint(sFiletoEmailPath As String)
'============================================================================
' Name        : fEmailtoPrint
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fEmailtoPrint(sFiletoEmailPath)
' Description : sends an email to print@aquoco.co to be printed
'               send email and add attachment yourself to print@aquoco.co
'============================================================================

Dim oOutlookApp As Outlook.Application, oOutlookMail As Outlook.MailItem


Set oOutlookApp = CreateObject("Outlook.Application")
Set oOutlookMail = oOutlookApp.CreateItem(0)

On Error Resume Next

With oOutlookMail
    .To = "print@aquoco.co"
    .CC = ""
    .BCC = ""
    .Subject = ""
    .HTMLBody = ""
    .Attachments.Add sFiletoEmailPath
End With

SendKeys "^{ENTER}"

On Error GoTo 0
Set oOutlookMail = Nothing
Set oOutlookApp = Nothing

End Sub

Public Sub fAudioDone()
'============================================================================
' Name        : fAudioDone
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fAudioDone
' Description : completes audio in express scribe
'============================================================================
sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

'If FSO Is Nothing Then Set FSO = New Scripting.FileSystemObject
   ' Set hFolder = FSO.GetFolder(HostFolder)

'iterate through all files in the root of the main folder
    'If Not blNotFirstIteration Then
      'For Each Fil In hFolder.Files
       '   FileExt = FSO.GetExtensionName(Fil.Path)
   ' FileTypes = Array("trs", "trm")
    
          'check if current file matches one of the specified file types ftr
      '    If Not IsError(Application.Match(FileExt, FileTypes, 0)) Then
              
        '    GoTo Line2
        '  End If
  ' FileTypes = Array("csx", "inf")
          
              'check if current file matches one of the specified file types courtsmart
        '  If Not IsError(Application.Match(FileExt, FileTypes, 0)) Then
        '    GoTo Line2
         ' End If
            ' check if current file matches one of the specified file types digital court player
                'to be added
          'Else
              'else try to open express scribe
        '    Call Shell("T:\Database\Scripts\Cortana\Audio-ExpressScribeDone.bat")
    '  Next Fil
      
'Line2:
  'Exit Do
End Sub
Sub fDistiller(sExportTopic As String)
'============================================================================
' Name        : fDistiller
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fDistiller(sExportTopic)
' Description : distills for PDFs
'               s2UpPSPath = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-" & sExportTopic & ".ps"
'               sFinalPDFPath = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-" & sExportTopic & ".pdf"
'============================================================================

Dim aaAcroApp As Acrobat.AcroApp
Dim aaAcroAVDoc As Acrobat.AcroAVDoc
Dim aaAcroPDDoc As Acrobat.AcroPDDoc
Dim pdTranscriptFinalDistiller As PdfDistiller
Dim sDistillerSettings As String, s2UpPSPath As String, sFinalPDFPath As String

s2UpPSPath = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-" & sExportTopic & ".ps"
sFinalPDFPath = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-" & sExportTopic & ".pdf"
sDistillerSettings = "C:\Program Files (x86)\Adobe\Acrobat 9.0\Acrobat\Settings\Standard.joboptions"

Set pdTranscriptFinalDistiller = New PdfDistiller
pdTranscriptFinalDistiller.FileToPdf strInputPostscript:=s2UpPSPath, strOutputPDF:=sFinalPDFPath, strJobOptions:=sDistillerSettings

Set pdTranscriptFinalDistiller = Nothing
aaAcroPDDoc.Close
aaAcroApp.CloseAllDocs

Set aaAcroPDDoc = Nothing
Set aaAcroAVDoc = Nothing
Set aaAcroApp = Nothing

End Sub
Sub fPrint2upPDF()
On Error GoTo eHandler
'============================================================================
' Name        : fPrint2upPDF
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fPrint2upPDF
' Description : prints 2-up transcript PDF
'============================================================================


Dim sTranscriptsFolderFinalPDF As String, sTranscriptsFolder2upPDF As String
Dim sTranscript2upPSPath As String, sJavascriptPrint As String, jobsettings As String
Dim sLogFilePath As String

Dim aaAcroApp As Acrobat.AcroApp
Dim aaAcroAVDoc As Acrobat.AcroAVDoc
Dim aaAcroPDDoc As Acrobat.AcroPDDoc
Dim bret As Variant
Dim pp As Object

Dim pdTranscriptFinalDistiller As PdfDistiller
Dim aaAFormApp As AFORMAUTLib.AFormApp

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
sTranscriptsFolderFinalPDF = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-FINAL.pdf"
sTranscriptsFolder2upPDF = "/i/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-Transcript-FINAL-2up.pdf"
sTranscript2upPSPath = "/i/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-Transcript-FINAL-2up.ps"

Set aaAcroApp = New AcroApp
Set aaAcroAVDoc = CreateObject("AcroExch.AVDoc")

If aaAcroAVDoc.Open(sTranscriptsFolderFinalPDF, "") Then
    aaAcroAVDoc.Maximize (1)
    
    Set aaAcroPDDoc = aaAcroAVDoc.GetPDDoc()
    Set aaAFormApp = CreateObject("AFormAut.App")
    
      sJavascriptPrint = "var pp=this.getPrintParams();" _
        & "pp.interactive=pp.constants.interactionLevel.automatic;" _
        & "pp.pageHandling=pp.constants.handling.nUp;" _
        & "pp.nUpPageOrders=pp.constants.nUpPageOrders.horizontal;" _
        & "pp.nUpAutoRotate=true;" _
        & "pp.nUpPageBorder=false;" _
        & "pp.nUpNumPagesV=2;" _
        & "pp.nUpNumPagesH=1;" _
        & "pp.fileName=" & Chr(34) & sTranscript2upPSPath & Chr(34) & ";" _
        & "this.print(pp);"
        '& "oPDFPrintSettings.bui=false;" _

    aaAFormApp.Fields.ExecuteThisJavascript sJavascriptPrint
    
    aaAcroPDDoc.Save PDSaveFull, sTranscript2upPSPath
    aaAcroPDDoc.Close
    aaAcroApp.CloseAllDocs
    
End If


sTranscript2upPSPath = "I:\" & sCourtDatesID & "\WorkingFiles\" & sCourtDatesID & "-Transcript-FINAL-2up.ps"
sTranscriptsFolder2upPDF = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-FINAL-2up.pdf"
jobsettings = "C:\Users\inqui\Standard1.joboptions"


Set pdTranscriptFinalDistiller = New PdfDistiller

pdTranscriptFinalDistiller.FileToPdf sTranscript2upPSPath, sTranscriptsFolder2upPDF, jobsettings ', jobsettings
'Debug.Print bret
sTranscriptsFolder2upPDF = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-FINAL-2up.pdf"
sTranscript2upPSPath = "I:\" & sCourtDatesID & "\Backups\" & sCourtDatesID & "-Transcript-FINAL-2up.pdf"

FileCopy sTranscriptsFolder2upPDF, sTranscript2upPSPath

Set pdTranscriptFinalDistiller = Nothing

eHandlerX:
Set aaAcroPDDoc = Nothing
Set aaAcroAVDoc = Nothing
Set aaAcroApp = Nothing

sLogFilePath = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-FINAL-2up.log"
'Check that file exists
If Len(Dir$(sLogFilePath)) > 0 Then
    'First remove readonly attribute, if set
    SetAttr sLogFilePath, vbNormal
    'Then delete the file
     Kill sLogFilePath
End If


MsgBox "2-up condensed transcript saved at " & sTranscript2upPSPath
Exit Sub

eHandler:
MsgBox Err.Number & ": " & Err.Description, vbCritical, "Error Detail"
GoTo eHandlerX
Resume
End Sub


Sub fPrint4upPDF()
On Error GoTo eHandler
'============================================================================
' Name        : fPrint4upPDF
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fPrint4upPDF
' Description : prints 4-up transcript PDF
'============================================================================

Dim sTranscriptsFolderFinalPDF As String, sTranscriptsFolder4upPDF As String, sTranscript4upPSPath As String
Dim sTranscript4upPDFPath As String, sAcrobatJobSettings As String, sJavascriptPrint As String
Dim sLogFilePath As String

Dim aaAcroApp As Acrobat.AcroApp
Dim aaAcroAVDoc As Acrobat.AcroAVDoc
Dim aaAcroPDDoc As Acrobat.AcroPDDoc
Dim pdTranscriptFinalDistiller As PdfDistiller
Dim aaAFormApp As AFORMAUTLib.AFormApp
Dim oPDFPrintSettings As Object


sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
sTranscriptsFolderFinalPDF = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-FINAL.pdf"
sTranscriptsFolder4upPDF = "/i/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-Transcript-FINAL-4up.pdf"
sTranscript4upPSPath = "/i/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-Transcript-FINAL-4up.ps"

Set aaAcroApp = New AcroApp
Set aaAcroAVDoc = CreateObject("AcroExch.AVDoc")

If aaAcroAVDoc.Open(sTranscriptsFolderFinalPDF, "") Then

    aaAcroAVDoc.Maximize (1)
    
    Set aaAcroPDDoc = aaAcroAVDoc.GetPDDoc()
    Set aaAFormApp = CreateObject("AFormAut.App")
    
      sJavascriptPrint = "var pp=this.getPrintParams();" _
        & "pp.interactive=pp.constants.interactionLevel.automatic;" _
        & "pp.pageHandling=pp.constants.handling.nUp;" _
        & "pp.nUpPageOrders=pp.constants.nUpPageOrders.horizontal;" _
        & "pp.nUpAutoRotate=true;" _
        & "pp.nUpPageBorder=false;" _
        & "pp.nUpNumPagesV=2;" _
        & "pp.nUpNumPagesH=2;" _
        & "pp.fileName=" & Chr(34) & sTranscript4upPSPath & Chr(34) & ";" _
        & "this.print(pp);"
        '& "oPDFPrintSettings.bui=false;" _


        
    aaAFormApp.Fields.ExecuteThisJavascript sJavascriptPrint
    aaAcroPDDoc.Save PDSaveFull, sTranscript4upPSPath
    aaAcroPDDoc.Close
    aaAcroApp.CloseAllDocs
    
End If

sTranscript4upPSPath = "I:\" & sCourtDatesID & "\WorkingFiles\" & sCourtDatesID & "-Transcript-FINAL-4up.ps"
sTranscript4upPDFPath = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-FINAL-4up.pdf"
sAcrobatJobSettings = "C:\Users\inqui\Standard1.joboptions"

Set pdTranscriptFinalDistiller = New PdfDistiller
pdTranscriptFinalDistiller.FileToPdf strInputPostscript:=sTranscript4upPSPath, strOutputPDF:=sTranscript4upPDFPath, strJobOptions:=sAcrobatJobSettings

sTranscriptsFolder4upPDF = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-FINAL-4up.pdf"
sTranscript4upPSPath = "I:\" & sCourtDatesID & "\Backups\" & sCourtDatesID & "-Transcript-FINAL-4up.pdf"

FileCopy sTranscriptsFolder4upPDF, sTranscript4upPSPath

Set pdTranscriptFinalDistiller = Nothing
eHandlerX:
Set aaAcroPDDoc = Nothing
Set aaAcroAVDoc = Nothing
Set aaAcroApp = Nothing


sLogFilePath = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-FINAL-4up.log"
'Check that file exists
If Len(Dir$(sLogFilePath)) > 0 Then
    'First remove readonly attribute, if set
    SetAttr sLogFilePath, vbNormal
    'Then delete the file
     Kill sLogFilePath
End If


MsgBox "4-up condensed transcript saved at " & sTranscript4upPDFPath
Exit Sub

eHandler:
MsgBox Err.Number & ": " & Err.Description, vbCritical, "Error Detail"
GoTo eHandlerX
Resume
End Sub

Sub fPrintKCIEnvelope()

Dim sQuestion As String, sAnswer As String, sEnvelopePath As String
sEnvelopePath = "T:\Database\Templates\Stage4s\Envelope-KCI.pdf"
sQuestion = "Print KCI envelope? (MAKE SURE ENVELOPE IS PRINT SIDE UP, ADHESIVE ON THE RIGHT INSIDE PRINTER TRAY)"
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???") '

If sAnswer = vbNo Then 'Code for No

    MsgBox "Envelope will not print."
    
Else 'Code for yes

    Call fEmailtoPrint(sEnvelopePath)
    
End If

End Sub



Sub fAcrobatKCIInvoice()
'============================================================================
' Name        : fAcrobatKCIInvoice
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fAcrobatKCIInvoice
' Description : inserts page count into KCI invoice
'============================================================================
'
On Error GoTo eHandler
Dim aaAcroApp As Acrobat.AcroApp
Dim aaAcroAVDoc As Acrobat.AcroAVDoc
Dim aaAcroPDDoc As Acrobat.AcroPDDoc
Dim aaAFormApp As AFORMAUTLib.AFormApp
Dim aaFoFiGroup As AFORMAUTLib.Fields
Dim aaFormField As AFORMAUTLib.Field

Dim sKCICompletedPath As String, sCaseName As String, sContactName As String

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
sKCICompletedPath = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-KCICompleted.pdf"

FileCopy "T:\Database\Templates\Stage4s\KCICompleted.pdf", sKCICompletedPath

Call pfCurrentCaseInfo  'refresh transcript info

sContactName = sFirstName & " " & sLastName
sCaseName = sParty1 & " v. " & sParty2

Set aaAcroApp = New AcroApp
Set aaAcroAVDoc = CreateObject("AcroExch.AVDoc")

If aaAcroAVDoc.Open(sKCICompletedPath, "") Then
    aaAcroAVDoc.Maximize (1)
    
    Set aaAcroPDDoc = aaAcroAVDoc.GetPDDoc()
    Set aaAFormApp = CreateObject("AFormAut.App")
    Set aaFoFiGroup = aaAFormApp.Fields
    
    For Each aaFormField In aaFoFiGroup
            If aaFormField.Name = "Case Name" Then aaFormField.Value = sCaseName
            If aaFormField.Name = "Trial Court" Then aaFormField.Value = sCaseNumber1
            If aaFormField.Name = "County" Then aaFormField.Value = sJurisdiction
            If aaFormField.Name = "COA No" Then aaFormField.Value = sCaseNumber2
            If aaFormField.Name = "Service Requested By" Then aaFormField.Value = sContactName
            If aaFormField.Name = "Original Report and 1 copy" Then aaFormField.Value = sActualQuantity
            If aaFormField.Name = "Date" Then aaFormField.Value = Date
    Next aaFormField
    
    sKCICompletedPath = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-KCICompleted1.pdf"
    
    aaAcroPDDoc.Save PDSaveFull, sKCICompletedPath
    aaAcroPDDoc.Close
End If

eHandlerX:

aaAcroAVDoc.Close True
Set aaAcroPDDoc = Nothing
Set aaAcroAVDoc = Nothing
Set aaAcroApp = Nothing

MsgBox "Done processing"
Exit Sub






eHandler:
MsgBox Err.Number & ": " & Err.Description, vbCritical, "Error Details"
GoTo eHandlerX
Resume
Call pfClearGlobals
End Sub


Public Sub pfUpload(ByRef mySession As Session)
'============================================================================
' Name        : pfUpload
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call pfUpload(mySession)
' Description : sends to website ftp
'============================================================================

Dim sAudioZIPPath As String, sTranscriptsZIPPath As String, sAudioTranscriptsZIPPath As String
Dim sFTPAudioTranscriptsZIPPath As String, sFTPTranscriptsZIPPath As String, sFTPAudioZIPPath As String
Dim mySessionOptions As New SessionOptions
Dim sIPCurrentJobFolder As String


Call pfCurrentCaseInfo  'refresh transcript info
    
With mySessionOptions 'set up session options

    .Protocol = Protocol_Ftp
    .HostName = "ftp.aquoco.co"
    .Username = Environ("ftpUserName")
    .password = Environ("ftpPassword")
    
End With

mySession.Open mySessionOptions 'connect

Dim myTransferOptions As New TransferOptions 'upload files
myTransferOptions.TransferMode = TransferMode_Binary

sIPCurrentJobFolder = "\\HUBCLOUD\evoingram\Production\2InProgress\" & sCourtDatesID & "\"
sFTPAudioTranscriptsZIPPath = "\\HUBCLOUD\evoingram\Production\2InProgress\" & sCourtDatesID & "\FTP\" & sCourtDatesID & "-AudioTranscripts" & ".zip"
sFTPTranscriptsZIPPath = "\\HUBCLOUD\evoingram\Production\2InProgress\" & sCourtDatesID & "\FTP\" & sCourtDatesID & "-Transcripts" & ".zip"
sFTPAudioZIPPath = "\\HUBCLOUD\evoingram\Production\2InProgress\" & sCourtDatesID & "\FTP\" & sCourtDatesID & "-Audio" & ".zip"
sAudioZIPPath = sIPCurrentJobFolder & sCourtDatesID & "-Audio" & ".zip"
sTranscriptsZIPPath = sIPCurrentJobFolder & sCourtDatesID & "-Transcripts" & ".zip"
sAudioTranscriptsZIPPath = sIPCurrentJobFolder & sCourtDatesID & "-AudioTranscripts" & ".zip"

Dim transferResult As TransferOperationResult
Dim transferResult2 As TransferOperationResult
Dim transferResult3 As TransferOperationResult

Set transferResult = mySession.PutFiles(sFTPAudioTranscriptsZIPPath, "/public_html/ProjectSend/upload/files/", False, myTransferOptions)
Set transferResult2 = mySession.PutFiles(sFTPTranscriptsZIPPath, "/public_html/ProjectSend/upload/files/", False, myTransferOptions)
Set transferResult3 = mySession.PutFiles(sFTPAudioZIPPath, "/public_html/ProjectSend/upload/files/", False, myTransferOptions)

transferResult.Check 'throw on any error
transferResult2.Check
transferResult3.Check
 

Dim transfer As TransferEventArgs 'display results
For Each transfer In transferResult.Transfers
    MsgBox "Upload of " & transfer.FileName & " succeeded"
Next
Call pfClearGlobals
End Sub


Sub fPrivatePrint()
'============================================================================
' Name        : fPrivatePrint
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fPrivatePrint
' Description : prompts to send necessary transcript files to print@aquoco.co to be printed
'============================================================================
'
Dim sCDLabelPath As String, s2upPath As String, s4upPath As String, sTranscriptPDFPath As String
Dim sQuestion As String, sAnswer As String

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
sCDLabelPath = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-CDLabel.PDF"
s2upPath = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-FINAL-2up.PDF"
s4upPath = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-FINAL-4up.PDF"
sTranscriptPDFPath = "I:\" & sCourtDatesID & "\Transcripts\" & sCourtDatesID & "-Transcript-FINAL.PDF"


'print 2-up (no without sfc intns)
sQuestion = "Print 2-up transcript?  Do not do so unless specifically requested."
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

If sAnswer = vbNo Then 'Code for No
    MsgBox "2-up transcript will not print."
Else 'Code for yes
    Call fEmailtoPrint(s2upPath)
End If


'print 4-up (no without sfc intns)
sQuestion = "Print 4-up transcript?  Do not do so unless specifically requested."
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

If sAnswer = vbNo Then 'Code for No
    MsgBox "4-up transcript will not print."
Else 'Code for yes
    Call fEmailtoPrint(s4upPath)
End If


'print transcript
sQuestion = "Print Transcript?"
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

If sAnswer = vbNo Then 'Code for No
    MsgBox "Transcript will not print."
Else 'Code for yes
    Call fEmailtoPrint(sTranscriptPDFPath)
End If


'print cd label
sQuestion = "Print CD Label? (MAKE SURE PAPER IS CORRECT IN PRINTER)"
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

If sAnswer = vbNo Then 'Code for No
    MsgBox "CD label will not print."
Else 'Code for yes
    Call fEmailtoPrint(sCDLabelPath)
End If


End Sub

Public Sub fExportRecsToXML()
On Error Resume Next
'============================================================================
' Name        : fExportRecsToXML
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fExportRecsToXML
' Description : exports ShippingOptionsQ entries to XML
'============================================================================

Dim qdf As DAO.QueryDef
Dim sTrackingNumber As String, sSavedXMLFileName As String
Dim prm As DAO.Parameter, rs As DAO.Recordset

Dim rstCommHistory As DAO.Recordset, rstShippingOptions As DAO.Recordset
Dim rstPkgType As DAO.Recordset, rstMailC As DAO.Recordset, rs1 As DAO.Recordset
Dim sNewSQL As String, SQLString As String, sMailClassNo As String
Dim sPackageTypeNo As String, sPackageType As String, sMailClass As String
    SQLString = "SELECT * FROM [ShippingOptions] WHERE [ShippingOptions].[CourtDatesID] = " & sCourtDatesID & ";"
    Set rs1 = CurrentDb.OpenRecordset(SQLString)
    sMailClassNo = rs1.Fields("MailClass").Value
    sPackageTypeNo = rs1.Fields("PackageType").Value
    rs1.Close
    
    '(SELECT MailClass FROM MailClass WHERE [ID] = " & sMailClassNo & ") as MailClass
    Set rstMailC = CurrentDb.OpenRecordset("SELECT MailClass FROM MailClass WHERE [ID] = " & sMailClassNo)
    sMailClass = rstMailC.Fields("MailClass").Value
    rstMailC.Close
    
    '(SELECT PackageType FROM PackageType WHERE [ID] = " & sPackageTypeNo & ") as PackageType
    Set rstPkgType = CurrentDb.OpenRecordset("SELECT PackageType FROM PackageType WHERE [ID] = " & sPackageTypeNo)
    sPackageType = rstPkgType.Fields("PackageType").Value
    rstPkgType.Close
    
    sNewSQL = "SELECT " & Chr(34) & sMailClass & Chr(34) & " as MailClass, " & Chr(34) & sPackageType & Chr(34) & " as PackageType, Width, Length, Depth, PriorityMailExpress1030, HolidayDelivery, SundayDelivery, SaturdayDelivery, SignatureRequired, Stealth, ReplyPostage, InsuredMail, COD, RestrictedDelivery, AdultSignatureRestricted, AdultSignatureRequired, ReturnReceipt, CertifiedMail, SignatureConfirmation, USPSTracking, CourtDatesIDLK as ReferenceID, ToName, ToAddress1, ToAddress2, ToCity, ToState, ToPostalCode, Value, Description, WeightOz, ActualWeight, ActualWeightText, ToEmail, ToPhone FROM [ShippingOptions] WHERE [ShippingOptions].[CourtDatesID] = " & sCourtDatesID & ";"

    Debug.Print (sNewSQL)
    
Set qdf = CurrentDb.QueryDefs(sNewSQL)

For Each prm In qdf.Parameters
    prm = Eval(prm.Name)
Next prm

Set rstShippingOptions = CurrentDb.OpenRecordset("ShippingOptions")

Do While rstShippingOptions.EOF = False
        
    SQLString = "SELECT * FROM [ShippingOptions] WHERE [ShippingOptions].[CourtDatesID] = " & sCourtDatesID & ";"
    Set rs = CurrentDb.OpenRecordset(SQLString)
    sMailClassNo = rs.Fields("MailClass").Value
    sPackageTypeNo = rs.Fields("PackageType").Value
    
    '(SELECT MailClass FROM MailClass WHERE [ID] = " & sMailClassNo & ") as MailClass
    Set rstMailC = CurrentDb.OpenRecordset("SELECT MailClass FROM MailClass WHERE [ID] = " & sMailClassNo)
    sMailClass = rstMailC.Fields("MailClass").Value
    rstMailC.Close
    
    '(SELECT PackageType FROM PackageType WHERE [ID] = " & sPackageTypeNo & ") as PackageType
    Set rstPkgType = CurrentDb.OpenRecordset("SELECT PackageType FROM PackageType WHERE [ID] = " & sPackageTypeNo)
    sPackageType = rstPkgType.Fields("PackageType").Value
    rstPkgType.Close
    sNewSQL = "SELECT CourtDatesIDLK, " & Chr(34) & sMailClass & Chr(34) & " as MailClass, " & Chr(34) & sPackageType & Chr(34) & " as PackageType, Width, Length, Depth, PriorityMailExpress1030, HolidayDelivery, SundayDelivery, SaturdayDelivery, SignatureRequired, Stealth, ReplyPostage, InsuredMail, COD, RestrictedDelivery, AdultSignatureRestricted, AdultSignatureRequired, ReturnReceipt, CertifiedMail, SignatureConfirmation, USPSTracking, ToName, ToAddress1, ToAddress2, ToCity, ToState, ToPostalCode, ToCountry, Description, Value, ToEmail, ToPhone FROM [ShippingOptions] WHERE [ShippingOptions].[CourtDatesID] = " & sCourtDatesID & ";"
    
    Debug.Print (sNewSQL)
    qdf.Sql = sNewSQL
    
    sCourtDatesID = rstShippingOptions.Fields("CourtDatesID").Value
    sTrackingNumber = rstShippingOptions.Fields("TrackingNumber").Value
    sSavedXMLFileName = "T:\Production\4ShippingXMLs\" & sCourtDatesID & "-" & sTrackingNumber & "-shipping.xml"
    Application.ExportXML acExportQuery, qdf.Name, sSavedXMLFileName 'export to XML
    
    rstShippingOptions.MoveNext

    'add shipping xml entry to comm history table
    sSavedXMLFileName = sCourtDatesID & "-ShippingXML" & "#" & sSavedXMLFileName & "#"
    Set rstCommHistory = CurrentDb.OpenRecordset("CommunicationHistory")
    
    rstCommHistory.AddNew
    rstCommHistory("FileHyperlink").Value = sSavedXMLFileName
    rstCommHistory("DateCreated").Value = Now
    rstCommHistory("CourtDatesID").Value = sCourtDatesID
    rstCommHistory.Update

    Call fShippingExpenseEntry(sTrackingNumber)
    Call fAppendXMLFiles
    
Loop

rstShippingOptions.Close
On Error GoTo 0
Set rstShippingOptions = Nothing
 
End Sub


Sub fAppendXMLFiles()
'============================================================================
' Name        : fAppendXMLFiles
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fAppendXMLFiles
' Description : appends XML files
'============================================================================

Dim file1 As New MSXML2.DOMDocument60, file2 As New MSXML2.DOMDocument60, file3 As New MSXML2.DOMDocument60
Dim appendNode As MSXML2.IXMLDOMNode
Dim FSO As New Scripting.FileSystemObject
Dim sXMLAfter As String, sXMLBefore As String

sXMLAfter = "T:\Database\Scripts\InProgressExcels\AfterXML.xml"
sXMLBefore = "T:\Database\Scripts\InProgressExcels\BeforeXML.xml"

' Load your xml files in to a DOM document
file1.Load sXMLBefore
file2.Load sXMLAfter

' iterate the childnodes of the second file, appending to the first file

For Each appendNode In file2.DocumentElement.ChildNodes
    file1.DocumentElement.appendChild appendNode
Next

For Each appendNode In file3.DocumentElement.ChildNodes
    file1.DocumentElement.appendChild appendNode
Next

' write combination to a new file
' if the specified filepath already exists, this will overwrite it'
FSO.CreateTextFile(file1 & "-FINISHED.xml", True, False).Write file1.XML

Set file1 = Nothing
Set file2 = Nothing
Set FSO = Nothing

End Sub


Sub fCourtofAppealsIXML()
On Error Resume Next
'============================================================================
' Name        : fCourtofAppealsIXML
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fCourtofAppealsIXML
' Description : creates Court of Appeals XML for shipping
'============================================================================

Dim sTSOCourtDatesID As String, sTempShippingOQ As String, sCOAXML As String
Dim sCOAXMLJF As String, sOutputXMLStringSQLFile As String, sOutputXMLStringSQL As String
Dim sTempShippingOQPath As String, sTempShipOptions As String, sTempShippingOQ1 As String
Dim sTempShipOptionsXLSM As String, sMacroName As String
Dim rstTempShippingOQ1 As DAO.Recordset, rstCommHistory As DAO.Recordset
Dim rstMailC As DAO.Recordset, rstPkgType As DAO.Recordset
Dim qdf As DAO.QueryDef, qdf1 As QueryDef
Dim oExcelApp As New Excel.Application, oExcelWkbk As New Excel.Workbook
Dim oExcelSheet As New Excel.Worksheet, oExcelWkbk2 As New Excel.Workbook
Dim sQueryName As String, sTSQExcelFileName As String, SQLString As String
Dim sMailClassNo As String, sPackageTypeNo As String, sMailClass As String
Dim sPackageType As String, sMailClass As String, sNewSQL As String
Dim rs1 As DAO.Recordset, rstShippingOptions As DAO.Recordset
Dim sCHHyperlinkXML As String

Call pfCurrentCaseInfo  'refresh transcript info

DoCmd.OpenQuery qnShippingOptionsQ, acViewNormal, acNormal 'pull up shipping query

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]



Call pfCurrentCaseInfo  'refresh transcript info
        
'@Ignore AssignmentNotUsed
sQueryName = "TempShippingOptionsQ"
'@Ignore AssignmentNotUsed
sTSQExcelFileName = "T:\Database\Scripts\InProgressExcels\TempShippingOptionsQ1.xlsm"

SQLString = "SELECT * FROM [ShippingOptions] WHERE [ShippingOptions].[CourtDatesID] = " & sCourtDatesID & ";"
Set rs1 = CurrentDb.OpenRecordset(SQLString)
sMailClassNo = rs1.Fields("MailClass").Value
sPackageTypeNo = rs1.Fields("PackageType").Value
rs1.Close

'(SELECT MailClass FROM MailClass WHERE [ID] = " & sMailClassNo & ") as MailClass
Set rstMailC = CurrentDb.OpenRecordset("SELECT MailClass FROM MailClass WHERE [ID] = " & sMailClassNo)
sMailClass = rstMailC.Fields("MailClass").Value
rstMailC.Close

'(SELECT PackageType FROM PackageType WHERE [ID] = " & sPackageTypeNo & ") as PackageType
Set rstPkgType = CurrentDb.OpenRecordset("SELECT PackageType FROM PackageType WHERE [ID] = " & sPackageTypeNo)
sPackageType = rstPkgType.Fields("PackageType").Value
rstPkgType.Close

sNewSQL = "SELECT " & Chr(34) & sMailClass & Chr(34) & " as MailClass, " & Chr(34) & sPackageType & Chr(34) & _
    " as PackageType, Width, Length, Depth, PriorityMailExpress1030, HolidayDelivery, SundayDelivery, SaturdayDelivery, SignatureRequired, " & _
    "Stealth, ReplyPostage, InsuredMail, COD, RestrictedDelivery, AdultSignatureRestricted, AdultSignatureRequired, ReturnReceipt, CertifiedMail, " & _
    "SignatureConfirmation, USPSTracking, CourtDatesIDLK as ReferenceID, " & Chr(34) & "Court of Appeals Div I Clerks Office," & Chr(34) & " AS ToName, " & _
    Chr(34) & "600 University St" & Chr(34) & " AS ToAddress1, " & Chr(34) & "One Union Square" & Chr(34) & " AS ToAddress2, " & Chr(34) & sCompanyCity & Chr(34) _
    & " AS ToCity, " & Chr(34) & sCompanyState & Chr(34) & " AS ToState, " & _
    Chr(34) & "98101" & Chr(34) & " AS ToPostalCode, Value, Description, WeightOz, ActualWeight, ActualWeightText, ToEmail, ToPhone " & _
    "FROM [ShippingOptions] WHERE [ShippingOptions].[CourtDatesID] = " & sCourtDatesID & ";"

sOutputXMLStringSQLFile = "\\HUBCLOUD\evoingram\Production\4ShippingXMLs\Output\" & sCourtDatesID & "-CoA-Output.xml"

Set rstShippingOptions = CurrentDb.OpenRecordset("SELECT * FROM ShippingOptions WHERE [ShippingOptions].[CourtDatesID] = " & sCourtDatesID & ";")
    
rstShippingOptions.Edit
Set rstShippingOptions.Fields("Output") = sOutputXMLStringSQLFile
rstShippingOptions.Update

'@Ignore AssignmentNotUsed
sTempShippingOQ1 = "TempShippingOptionsQ1"
'@Ignore AssignmentNotUsed
sTempShippingOQPath = "T:\Database\Scripts\InProgressExcels\TempShippingOptionsQ1.xlsm"

Set rstTempShippingOQ1 = CurrentDb.OpenRecordset(sNewSQL)
'@Ignore AssignmentNotUsed
sTSOCourtDatesID = rstTempShippingOQ1("ReferenceID").Value

Set oExcelApp = CreateObject("Excel.Application")
Set oExcelWkbk = oExcelApp.Workbooks.Open(sTempShippingOQPath)

sTempShipOptions = "TempShippingOptionsQ"
Set oExcelSheet = oExcelWkbk.Sheets(sTempShipOptions)
oExcelSheet.Cells(2, 1).Value = sOutputXMLStringSQLFile
oExcelSheet.Cells(2, 24).Value = "Court of Appeals Div I Clerk's Office"
oExcelSheet.Cells(2, 25).Value = "600 University Street"
oExcelSheet.Cells(2, 26).Value = "One Union Square"
oExcelSheet.Cells(2, 27).Value = sCompanyCity
oExcelSheet.Cells(2, 28).Value = sCompanyState
oExcelSheet.Cells(2, 29).Value = "98101"
oExcelWkbk.Save

oExcelSheet.Range("S2").CopyFromRecordset rstTempShippingOQ1

oExcelWkbk.Save
oExcelWkbk.Close SaveChanges:=True
qdf1.Close

rstTempShippingOQ1.Close
Set rstTempShippingOQ1 = Nothing
Set qdf1 = Nothing

sTempShipOptionsXLSM = "T:\Database\Scripts\InProgressExcels\TempShippingOptionsQ1.xlsm"
sMacroName = "ExportXMLCOA"

Call fRunXLSMacro(sTempShipOptionsXLSM, sMacroName)

sCOAXML = "T:\Production\4ShippingXMLs\" & sCourtDatesID & "-CourtofAppealsDivI-Shipping.xml"
sCOAXMLJF = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-CourtofAppealsDivI-Shipping.xml"

FileCopy sCOAXML, sCOAXMLJF
On Error GoTo 0

'add shipping xml entry to comm history table
sCHHyperlinkXML = sCourtDatesID & "CoADiv1-ShippingXML" & "#" & sCOAXML & "#"
Set rstCommHistory = CurrentDb.OpenRecordset("CommunicationHistory")

rstCommHistory.AddNew
rstCommHistory("FileHyperlink").Value = sCHHyperlinkXML
rstCommHistory("DateCreated").Value = Now
rstCommHistory("CourtDatesID").Value = sCourtDatesID
rstCommHistory.Update

'add another set of expenses for court of appeals package
Call fTranscriptExpensesBeginning
Call fTranscriptExpensesAfter

MsgBox "Exported COA XML and added entry to CommHistory table."
Call pfClearGlobals
End Sub



