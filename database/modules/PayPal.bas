Attribute VB_Name = "PayPal"
'@Folder("Database.Admin.Modules")
Option Explicit


'============================================================================
'class module cmPayPal

'variables:
'   Sleep(Milliseconds)

'functions:
'fPPGetInvoiceInfo:
'Description:  gets status of invoice
'arguments:    NONE

'fPPUpdate:
'Description:  updates PayPal invoice on PayPal website
'arguments:    NONE

'PPDraft:
'Description:  creates PayPal draft invoice on PayPal website
'arguments:    NONE

'fSendPPEmailBalanceDue:
'Description:  sends PP email for balance due
'arguments:    NONE

'fSendPPEmailDeposit:
'Description:  generates PP email for deposit
'arguments:    NONE

'fSendPPEmailFactored:
'Description:  generates factored invoice email for pp
'arguments:    NONE
        
'fPPGenerateJSONInfo:
'Description:  get info for invoice
'arguments:    NONE

'fManualPPPayment:
'Description:  marks invoice as paid with manual payment, like with check/cash
'arguments:    NONE

'fPayPalUpdateCheck:
'Description:  Check PP for update on invoice
'arguments:    NONE
        
'fPPRefund
'Description:  refund with pp
'arguments:    NONE
        
'PP Templates:
'deposit invoice (PP-DraftInvoiceEmail) fSendPPEmailDeposit
'payment receipt (PP-PaymentMadeEmail) vCHTopic PP-PaymentMadeEmail, vSubject "Payment Received"
'refund with invoice details (PP-RefundMadeEmail) pfSendWordDocAsEmail:  vCHTopic "Stage4s\PP-RefundMadeEmail", vSubject "Refund Issued"
'factoring invoice (PP-FactoredInvoiceEmail) fSendPPEmailFactored
'balance due invoice (PP-BalanceDueInvoiceEmail) fSendPPEmailBalanceDue
'invoice payment reminder
'============================================================================


Private x As String
Private sTemp As String

'TODO: fix PP/invoicing functions
'TODO: invoice # Word doc doesn't save properly w/ PP button

Public Sub fSendPPEmailFactored()
    'generates factored invoice email for pp
    Dim sInvoiceNumber As String
    Dim sName As String
    Dim vPPInvoiceNo As String
    Dim sHTMLPPB As String
    Dim vPPLink As String
    Dim sQuestion As String
    Dim sAnswer As String
    Dim sToEmail As String
    Dim sBuf As String
    Dim sTemp As String

    Dim oOutlookApp As Outlook.Application
    Dim oOutlookMail As Outlook.MailItem
    Dim oWordEditor As Word.editor
    Dim oWordApp As New Word.Application
    Dim oWordDoc As New Word.Document

    Dim qdf As QueryDef
    Dim rstQuery As DAO.Recordset
    
    Dim iFileNum As Long

    Dim cJob As Job
    Set cJob = New Job
    
    Call fPPGenerateJSONInfo                     'refreshes some necessary info

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

    Call pfGetOrderingAttorneyInfo               'refreshes some necessary info
    Call pfCurrentCaseInfo                       'refreshes some necessary info

    Set qdf = CurrentDb.QueryDefs(qTRIQPlusCases)
    qdf.Parameters(0) = sCourtDatesID
    Set rstQuery = qdf.OpenRecordset

    sInvoiceNumber = rstQuery.Fields("TRInv.CourtDates.InvoiceNo").Value
    sParty1 = rstQuery.Fields("Party1").Value
    sParty2 = rstQuery.Fields("Party2").Value
    sName = sFirstName & " " & sLastName

    If IsNull(rstQuery.Fields("TRinv.PPID").Value) Then

        Call fPPDraft
    End If

    vPPInvoiceNo = rstQuery.Fields("TRInv.PPID").Value
    vPPInvoiceNo = Right(vPPInvoiceNo, 20)
    vPPInvoiceNo = Replace(Replace(vPPInvoiceNo, " ", ""), "-", "")

    'create pp invoice link
    vPPLink = "https://www.paypal.com/invoice/p/#" & vPPInvoiceNo

    'create PPButton html (construct string, save it into text file in job folder)
    sHTMLPPB = "<html><body><br><br><a href =" & Chr(34) & vPPLink & Chr(34) & "><img src=" & Chr(34) _
      & "https://www.paypalobjects.com/webstatic/en_US/i/buttons/checkout-logo-large.png" & _
        Chr(34) & "alt=" & Chr(34) & "Check out with PayPal" & Chr(34) & "/></a></body></html>"

    'open txt file, save as html file

    Open cJob.DocPath.PPButtonT For Output As #1
    Write #1, sHTMLPPB
    Close #1

    iFileNum = FreeFile
    Open cJob.DocPath.PPButtonT For Input As iFileNum

    Do Until EOF(iFileNum)
        Line Input #iFileNum, sBuf
        sTemp = sTemp & sBuf & vbCrLf
    Loop

    Close iFileNum

    sTemp = Replace(sTemp, Chr(34) & "<html>", "<html>") 'doing it this way makes weird things happen to the text file
    sTemp = Replace(sTemp, "</html>" & Chr(34), "</html>") 'so these 3 sets of changes are necessary to form correct html
    sTemp = Replace(sTemp, Chr(34) & Chr(34), Chr(34))


    iFileNum = FreeFile
    Open cJob.DocPath.PPButtonT For Output As iFileNum

    Print #iFileNum, sTemp

    Close iFileNum

    Name cJob.DocPath.PPButtonT As cJob.DocPath.PPButton              'Save txt file as (if possible)

    'paste PPButton html into mail merged invoice/PQ at both bookmarks "PPButton" bookmark or #PPB1# AND "PPButton2" or #PPB2#
    Set oWordApp = CreateObject("Word.Application")
    oWordApp.Visible = False

    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPButton)

    oWordDoc.Content.Copy

    Set oWordApp = CreateObject("Word.Application")
    oWordApp.Visible = True

    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPFactoredInvoiceEmail)

    With oWordDoc.Application

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
            '.Application.Selection.PasteAndFormat (wdPasteDefault)
            .Application.Selection.PasteAndFormat (wdFormatOriginalFormatting)
        
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
            '.Application.Selection.Range.PasteSpecial DataType:=wdPasteHTML
            .Application.Selection.PasteAndFormat (wdFormatOriginalFormatting)
        End With
    
        'save invoice/PQ
        .ActiveDocument.SaveAs2 FileName:=cJob.DocPath.PPFactoredInvoiceEmail
    
    End With

    oWordDoc.Close
    oWordApp.Quit
    
    Set oWordApp = Nothing
    Set oWordDoc = Nothing
    

    'prompt to ask when pdf is made
    sQuestion = "Click yes after you have created your final invoice PDF at " & cJob.DocPath.InvoiceP
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")


    If sAnswer = vbNo Then                       'IF NO THEN THIS HAPPENS
    
        MsgBox "No invoice will be sent."
    
    Else                                         'if yes then this happens
    
        DoCmd.OutputTo acOutputQuery, qTRIQPlusCases, acFormatXLS, cJob.DocPath.InvoiceInfo, False
    
        'Set oWordApp = GetObject(cJob.DocPath.PPFIET, "Word.Document")
        Set oWordDoc = oWordApp.Documents.Add(cJob.DocPath.PPFIET)
        oWordApp.Application.Visible = False
    
        oWordDoc.MailMerge.OpenDataSource Name:=cJob.DocPath.InvoiceInfo, ReadOnly:=True
        oWordDoc.MailMerge.Execute
        oWordDoc.MailMerge.MainDocumentType = wdNotAMergeDocument
        oWordDoc.Application.ActiveDocument.SaveAs2 FileName:=cJob.DocPath.PPFactoredInvoiceEmail

    
        oWordApp.Quit
        Set oWordApp = Nothing
        Set oWordDoc = Nothing
    
        Set oWordApp = CreateObject("Word.Application")
        oWordApp.Visible = False
    
        Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPButton) 'copy button html file
    
        oWordDoc.Content.Copy
        oWordDoc.Close
        oWordApp.Quit
    
        Set oWordApp = Nothing
        Set oWordDoc = Nothing

        Set oWordApp = CreateObject("Word.Application")
        Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPFactoredInvoiceEmail)


        With oWordDoc.Application

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
        
                '.Application.Selection.PasteAndFormat (wdPasteDefault)
                .Application.Selection.PasteAndFormat (wdFormatOriginalFormatting) 'paste button in word doc at #PPB1#
        
            End With
    
            'save invoice
            oWordDoc.SaveAs FileName:=cJob.DocPath.PPFactoredInvoiceEmail
    
    
        End With
        oWordDoc.Close
        oWordApp.Quit
    
        Set oWordApp = Nothing
        Set oWordDoc = Nothing
        
        rstQuery.Close
        Set rstQuery = Nothing
        qdf.Close
        Set qdf = Nothing
        Set oOutlookApp = CreateObject("Outlook.Application")
        Set oOutlookMail = oOutlookApp.CreateItem(0)
        On Error Resume Next
        Set oWordApp = CreateObject("Word.Application")
        Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPFactoredInvoiceEmail)
        oWordDoc.Content.Copy
    
        With oOutlookMail                        'email that will also include pp button
            '@Ignore UnassignedVariableUsage
            .To = sToEmail
            .CC = sCompanyEmail
            .Subject = "Transcript Delivery & Invoice for " & sName & ", " & sParty1 & " v. " & sParty2
            .BodyFormat = olFormatRichText
            Set oWordEditor = .GetInspector.WordEditor
            .GetInspector.WordEditor.Content.Paste
            .Display
            .Attachments.Add (cJob.DocPath.InvoiceP)
        End With
    
        oWordDoc.Close
        oWordApp.Quit
        Set oWordApp = Nothing
        Set oWordDoc = Nothing
    End If
    On Error GoTo 0
    Call pfCommunicationHistoryAdd("PP Invoice Sent")
    Call pfClearGlobals
End Sub

Public Sub fPPDraft()
    '============================================================================
    ' Name        : fPPDraft
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fPPDraft()
    ' Description : creates PayPal draft invoice on PayPal website
    '============================================================================

    Dim sURL As String
    Dim sUserName As String
    Dim sPassword As String
    Dim sAuth As String
    Dim stringJSON As String
    Dim sEmail As String
    Dim vInvoiceID As String
    Dim apiWaxLRS As String
    Dim vErrorIssue As String
    Dim sInvoiceTime As String
    Dim vStatus As String
    Dim vTotal As String
    Dim sToken As String
    Dim json1 As String
    Dim json2 As String
    Dim json3 As String
    Dim json4 As String
    Dim json5 As String
    Dim vErrorName As String
    Dim vErrorMessage As String
    Dim vErrorILink As String
    Dim vErrorDetails As String
    Dim sFile1 As String
    Dim sFile2 As String
    Dim sText As String
    Dim sLine1 As String
    Dim sLine2 As String
    Dim vInventoryRateCode As String

    Dim oRequest As Object
    Dim Json As Object

    Dim rstRates As DAO.Recordset

    Dim resp As Variant
    Dim response As Variant
    Dim rep As Variant
    
    Dim vDetails As Object
    
    Dim parsed As Dictionary
    
    sIRC = 8
Beginning:
    Call pfGetOrderingAttorneyInfo
    Call fPPGenerateJSONInfo
    Set rstRates = CurrentDb.OpenRecordset("SELECT * FROM Rates WHERE [ID] = " & sIRC & ";")
    vInventoryRateCode = rstRates.Fields("Code").Value
    rstRates.Close

    'Debug.Print sCourtDatesID & " " & sInvoiceNumber

    sURL = "https://api.paypal.com/v1/oauth2/token/"
    sEmail = sCompanyEmail
    '  https://api.paypal.com/v1/oauth2/token \
    'sAuth = TextBase64Encode(myCn.GetConnection, "us-ascii") 'mycn.GetConnection
    sAuth = TextBase64Encode(Environ("ppUserName") & ":" & Environ("ppPassword"), "us-ascii") 'mycn.GetConnection
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        '.Visible = True
        .Open "POST", sURL, False
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Accept-Language", "en_US"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        '.setRequestHeader "grant_type=client_credentials"
        '-u "client_id:secret" \
        .setRequestHeader "Authorization", "Basic " + sAuth
        '.setRequestHeader "Authorization", "Bearer " & sAuth
        .send ("grant_type=client_credentials")
        apiWaxLRS = .responseText
        Debug.Print apiWaxLRS
        Set parsed = JsonConverter.ParseJson(apiWaxLRS)
        sToken = parsed.item("access_token")          'third level array
        .abort
        Debug.Print "--------------------------------------------"
    End With
      
    '@Ignore UnassignedVariableUsage
    json1 = "{" & Chr(34) & "merchant_info" & Chr(34) & ": {" & Chr(34) & _
                                                                        "email" & Chr(34) & ": " & Chr(34) & sCompanyEmail & Chr(34) & "," & Chr(34) & _
                                                                        "first_name" & Chr(34) & ": " & Chr(34) & sCompanyFirstName & Chr(34) & "," & Chr(34) & _
                                                                        "last_name" & Chr(34) & ": " & Chr(34) & sCompanyLastName & Chr(34) & "," & Chr(34) & _
                                                                        "business_name" & Chr(34) & ": " & Chr(34) & sCompanyName & Chr(34) & "," & Chr(34) & _
                                                                        "phone" & Chr(34) & ": {" & Chr(34) & _
                                                                        "country_code" & Chr(34) & ": " & Chr(34) & sPCountryCode & Chr(34) & "," & Chr(34) & _
                                                                        "national_number" & Chr(34) & ": " & Chr(34) & sCompanyNationalNumber & Chr(34) & "}," & Chr(34) & _
                                                                        "address" & Chr(34) & ": {" & Chr(34) & _
                                                                        "line1" & Chr(34) & ": " & Chr(34) & sCompanyAddress & Chr(34) & "," & Chr(34) & _
                                                                        "city" & Chr(34) & ": " & Chr(34) & sCompanyCity & Chr(34) & "," & Chr(34) & _
                                                                        "state" & Chr(34) & ": " & Chr(34) & sCompanyState & Chr(34) & "," & Chr(34) & _
                                                                        "postal_code" & Chr(34) & ": " & Chr(34) & sCompanyZIP & Chr(34) & "," & Chr(34) & _
                                                                        "country_code" & Chr(34) & ": " & Chr(34) & sZCountryCode & Chr(34) & "}},"
    json2 = Chr(34) & "billing_info" & Chr(34) & ": [{" & Chr(34) & _
                                                                  "email" & Chr(34) & ": " & Chr(34) & sEmail & Chr(34) & "," & Chr(34) & _
                                                                  "first_name" & Chr(34) & ": " & Chr(34) & sFirstName & Chr(34) & "," & Chr(34) & _
                                                                  "last_name" & Chr(34) & ": " & Chr(34) & sLastName & Chr(34) & "}]," & Chr(34) & _
                                                                  "shipping_info" & Chr(34) & ": {" & Chr(34) & _
                                                                  "first_name" & Chr(34) & ": " & Chr(34) & sFirstName & Chr(34) & "," & Chr(34) & _
                                                                  "last_name" & Chr(34) & ": " & Chr(34) & sLastName & Chr(34) & "," & Chr(34) & _
                                                                  "address" & Chr(34) & ": {" & Chr(34) & "line1" & Chr(34) & ": " & Chr(34) & sAddress2 & Chr(34) & "," & Chr(34) & _
                                                                  "city" & Chr(34) & ": " & Chr(34) & sCity & Chr(34) & "," & Chr(34) & _
                                                                  "state" & Chr(34) & ": " & Chr(34) & sState & Chr(34) & "," & Chr(34) & _
                                                                  "postal_code" & Chr(34) & ": " & Chr(34) & sZIP & Chr(34) & "," & Chr(34) & _
                                                                  "country_code" & Chr(34) & ": " & Chr(34) & "US" & Chr(34) & "}},"
    json3 = Chr(34) & "items" & Chr(34) & ": [" & _
                                        "{" & Chr(34) & _
                                        "name" & Chr(34) & ": " & Chr(34) & vInventoryRateCode & Chr(34) & "," & Chr(34) & _
                                        "description" & Chr(34) & ": " & Chr(34) & sDescription & Chr(34) & "," & Chr(34) & _
                                        "quantity" & Chr(34) & ": " & Chr(34) & sQuantity & Chr(34) & "," & Chr(34) & _
                                        "unit_price" & Chr(34) & ": {" & Chr(34) & _
                                        "currency" & Chr(34) & ": " & Chr(34) & "USD" & Chr(34) & "," & Chr(34) & _
                                        "value" & Chr(34) & ": " & Chr(34) & sUnitPrice & Chr(34) & "}," & Chr(34) & _
                                        "tax" & Chr(34) & ": {" & Chr(34) & _
                                        "name" & Chr(34) & ": " & Chr(34) & "Tax" & Chr(34) & "," & Chr(34) & _
                                        "percent" & Chr(34) & ": 0.00}}]," & Chr(34) & _
                                        "payment_term" & Chr(34) & ": {" & Chr(34) & "term_type" & Chr(34) & ": " & Chr(34) & "DUE_ON_DATE_SPECIFIED" & Chr(34) & "," & Chr(34) & _
                                        "due_date" & Chr(34) & ": " & Chr(34) & sInvoiceDate & " " _
                                      & sInvoiceTime & Chr(34) & "}," & Chr(34) & _
                                        "reference" & Chr(34) & ": " & Chr(34) & sCourtDatesID & Chr(34) & ","
    '"invoice_date" & Chr(34) & ": {" & Chr(34) & sInvoiceDate & Chr(34) & "}," & Chr(34) & _

    json4 = Chr(34) & _
                    "shipping_cost" & Chr(34) & ": {" & Chr(34) & _
                    "amount" & Chr(34) & ": {" & Chr(34) & _
                    "currency" & Chr(34) & ": " & Chr(34) & "USD" & Chr(34) & "," & Chr(34) & _
                    "value" & Chr(34) & ": " & Chr(34) & "0.00" & Chr(34) & "}}," & Chr(34) & _
                    "note" & Chr(34) & ": " & Chr(34) & sNote & Chr(34) & "," & Chr(34) & _
                    "terms" & Chr(34) & ": " & Chr(34) & sTerms & Chr(34) & "}" & "{" & Chr(34) & _
                    "allow_partial_payment" & Chr(34) & ": " & Chr(34) & "true" & Chr(34) & "}" & _
                    "{" & Chr(34) & _
                    "minimum_amount_due" & Chr(34) & ": " & Chr(34) & sMinimumAmount & "}" & "{" & Chr(34) & _
                    "tax_inclusive" & Chr(34) & ": " & Chr(34) & "true}" & _
                    "{" & Chr(34) & _
                    "merchant_memo" & Chr(34) & ": " & Chr(34) & vmMemo & "}" & "{" & Chr(34) & _
                    "logo_url" & Chr(34) & ": " & Chr(34) & vlURL & "}" & "{" & Chr(34) & _
                    "template_id" & Chr(34) & ": " & Chr(34) & sTemplateID & "}," & "{" & Chr(34) & "number" & Chr(34) & ": " & Chr(34) & sInvoiceNo & Chr(34) & "}"
    Debug.Print "JSON1--------------------------------------------"
    Debug.Print json1
    Debug.Print "JSON2--------------------------------------------"
    Debug.Print json2
    Debug.Print "JSON3--------------------------------------------"
    Debug.Print json3
    Debug.Print "JSON4--------------------------------------------"
    Debug.Print json4
    Debug.Print "RESPONSETEXT--------------------------------------------"
    sURL = "https://api.paypal.com/v1/invoicing/invoices"
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        '.Visible = True
        .Open "POST", sURL, False
        '.setRequestHeader "Accept", "application/json" 'application/x-www-form-urlencoded
        '.setRequestHeader "content-type", "application/x-www-form-urlencoded"
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & sToken
        json5 = json1 & json2 & json3 & json4
        Debug.Print json5
        .send json5
        apiWaxLRS = .responseText
        sToken = ""
        .abort
        Debug.Print apiWaxLRS
        'Debug.Print "--------------------------------------------"
    End With
    Set parsed = JsonConverter.ParseJson(apiWaxLRS)
    sInvoiceNumber = parsed.item("number")            'third level array
    vInvoiceID = parsed.item("id")                    'third level array
    vStatus = parsed.item("status")                   'third level array
    vTotal = parsed.item("total_amount").item("value")     'second level array
    vErrorName = parsed.item("name")                  '("value") 'second level array
    vErrorMessage = parsed.item("message")            '("value") 'second level array
    vErrorILink = parsed.item("information_link")     '("value") 'second level array
    '
    'Set vDetails = Parsed("details") 'second level array
    'For Each rep In vDetails ' third level objects
    '    vErrorIssue = rep("field")
    '    vErrorDetails = rep("issue")
    'Next
    Debug.Print "--------------------------------------------"
    Debug.Print "Error Name:  " & vErrorName
    Debug.Print "Error Message:  " & vErrorMessage
    Debug.Print "Error Info Link:  " & vErrorILink
    'Debug.Print "Error Field:  " & vErrorIssue
    'Debug.Print "Error Details:  " & vErrorDetails
    Debug.Print "--------------------------------------------"
    'Next
    Debug.Print "Invoice No.:  " & sInvoiceNumber & "   |   Invoice ID:  " & vInvoiceID
    Debug.Print "Status:  " & vStatus & "   |   Total:  " & vTotal
    Debug.Print "--------------------------------------------"

    'update PPID & PPStatus
    Dim sUpdatePPStatus As String
    Dim sUpdatePPID As String
    sUpdatePPStatus = "UPDATE CourtDates SET PPStatus = " & Chr(34) & vStatus & Chr(34) & " WHERE [ID] = " & sCourtDatesID & ";"
    CurrentDb.Execute sUpdatePPStatus
    sUpdatePPID = "UPDATE CourtDates SET PPID = " & Chr(34) & vInvoiceID & Chr(34) & " WHERE [ID] = " & sCourtDatesID & ";"
    CurrentDb.Execute sUpdatePPID
    Call pfClearGlobals
End Sub

Public Sub fPayPalUpdateCheck()
    '============================================================================
    ' Name        : fPayPalUpdateCheck
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fPayPalUpdateCheck()
    ' Description : check PayPal for update on invoice
    '============================================================================

    Dim sQueryName As String
    Dim sURL As String
    Dim sUserName As String
    Dim sPassword As String
    Dim sAuth As String
    Dim stringJSON As String
    Dim apiWaxLRS As String
    Dim vErrorIssue As String
    Dim sInvoiceTime As String
    Dim sInvoiceNo As String
    Dim sEmail As String
    Dim sFirstName As String
    Dim sLastName As String
    Dim sDescription As String
    Dim sInvoiceDate As String
    Dim sPaymentTerms As String
    Dim sCourtDatesID As String
    Dim sNote As String
    Dim sTerms As String
    Dim sMinimumAmount As String
    Dim vmMemo As String
    Dim vlURL As String
    Dim sTemplateID As String
    Dim vTotal As String
    Dim sCity As String
    Dim sState As String
    Dim sZIP As String
    Dim sQuantity As String
    Dim sValue As String
    Dim vInvoiceID As String
    Dim sInvoiceNumber As String
    Dim vStatus As String
    Dim sToken As String
    Dim json1 As String
    Dim json2 As String
    Dim json3 As String
    Dim json4 As String
    Dim json5 As String
    Dim vErrorName As String
    Dim vErrorMessage As String
    Dim vErrorILink As String
    Dim vErrorDetails As String
    Dim vTermDays As String
    Dim vDetails As String
    Dim sFile1 As String
    Dim sFile2 As String
    Dim sText As String
    Dim sLine1 As String
    Dim sLine2 As String
    Dim vPPStatus As String

    Dim resp As Variant
    Dim response As Variant
    Dim rep As Variant
    
    Dim oRequest As Object
    Dim Json As Object
    
    Dim qdf As QueryDef
    Dim qdf1 As QueryDef
    Dim rstQuery As DAO.Recordset
    Dim rstQuery1 As DAO.Recordset
    
    Dim parsed As Dictionary
    
    Call fPPGetInvoiceInfo

    sQueryName = "QPPStatus"
    Set rstQuery1 = CurrentDb.OpenRecordset(sQueryName)

    If Not (rstQuery1.EOF And rstQuery1.BOF) Then

        rstQuery1.MoveFirst
        
        Do Until rstQuery1.EOF = True
    
            sCourtDatesID = rstQuery1.Fields("ID").Value
            sInvoiceNumber = rstQuery1.Fields("InvoiceNo").Value
        
            If Not rstQuery1.Fields("PPID").Value = "" Or rstQuery1.Fields("PPID").Value = Null Then
                vInvoiceID = rstQuery1.Fields("PPID").Value
            Else
                GoTo NextJob
            End If
        
            vPPStatus = rstQuery1.Fields("PPStatus").Value 'DRAFT, SENT, SCHEDULED, PAID, MARKED_AS_PAID.
        
            'set vStatus of current invoice
                
            If vInvoiceID = "" Then
                GoTo NextJob
            Else
                vInvoiceID = Right(vInvoiceID, 20)
                vInvoiceID = Replace(Replace(vInvoiceID, " ", ""), "-", "")
            
                sURL = "https://api.paypal.com/v1/oauth2/token/"
                sEmail = sCompanyEmail
                '  https://api.paypal.com/v1/oauth2/token \
                'sAuth = TextBase64Encode(myCn.GetConnection, "us-ascii") 'mycn.GetConnection
                sAuth = TextBase64Encode(Environ("ppUserName") & ":" & Environ("ppPassword"), "us-ascii") 'mycn.GetConnection
                With CreateObject("WinHttp.WinHttpRequest.5.1")
                    '.Visible = True
                    .Open "POST", sURL, False
                    .setRequestHeader "Accept", "application/json"
                    .setRequestHeader "Accept-Language", "en_US"
                    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                    '.setRequestHeader "grant_type=client_credentials"
                    '-u "client_id:secret" \
                    .setRequestHeader "Authorization", "Basic " + sAuth
                    '.setRequestHeader "Authorization", "Bearer " & sAuth
                    .send ("grant_type=client_credentials")
                    apiWaxLRS = .responseText
                    Debug.Print apiWaxLRS
                    Set parsed = JsonConverter.ParseJson(apiWaxLRS)
                    sToken = parsed("access_token") 'third level array
                    .abort
                    'Debug.Print "--------------------------------------------"
                End With
                
              
                'Debug.Print "RESPONSETEXT--------------------------------------------"
                sURL = "https://api.paypal.com/v1/invoicing/invoices/" & vInvoiceID
                With CreateObject("WinHttp.WinHttpRequest.5.1")
                    '.Visible = True
                    .Open "GET", sURL, False
                    '.setRequestHeader "Accept", "application/json" 'application/x-www-form-urlencoded
                    '.setRequestHeader "content-type", "application/x-www-form-urlencoded"
                    .setRequestHeader "content-type", "application/json"
                    .setRequestHeader "Authorization", "Bearer " & sToken
                    ' json5 = json1 & json2 & json3
                    .send                        ' json5
                    apiWaxLRS = .responseText
                    sToken = ""
                    .abort
                    Debug.Print apiWaxLRS
                    Debug.Print "--------------------------------------------"
                End With
                Set parsed = JsonConverter.ParseJson(apiWaxLRS)
                sInvoiceNumber = parsed("number") 'third level array
                vInvoiceID = parsed("id")        'third level array
                vStatus = parsed("status")       'third level array 'can be either DRAFT, UNPAID, SENT, SCHEDULED, PARTIALLY_PAID, PAYMENT_PENDING, PAID, MARKED_AS_PAID,
                'CANCELLED, REFUNDED, PARTIALLY_REFUNDED, MARKED AS REFUNDED
                                                                        
                Debug.Print "--------------------------------------------"
            
            
                If vStatus <> vPPStatus And vStatus = "PAID" Or vStatus = "MARKED_AS_PAID" Then
            
                    Call pfSendWordDocAsEmail("PP-PaymentMadeEmail", "Payment Received") 'send/queue payment received receipt automatically
                
                    sQueryName = "SELECT PPStatus FROM CourtDates WHERE ID = " & sCourtDatesID & ";"
                
                    Set rstQuery = CurrentDb.OpenRecordset(sQueryName)
                
                    rstQuery.Edit
                    rstQuery.Fields("PPStatus").Value = vStatus
                    rstQuery.Update
                    rstQuery.Close
                    Set rstQuery = Nothing
                
                Else
                End If
        
        
            End If
NextJob:
            rstQuery1.MoveNext
        Loop
    
    End If

    rstQuery1.Close
    Set rstQuery1 = Nothing

    Call pfClearGlobals
End Sub

Public Sub fSendPPEmailBalanceDue()
    '============================================================================
    ' Name        : fPayPalUpdateCheck
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fPayPalUpdateCheck()
    ' Description : sends PayPal email for balance due
    '============================================================================
    
    Dim sName As String
    Dim vPPInvoiceNo As String
    Dim sHTMLPPB As String
    Dim vPPLink As String
    Dim sQuestion As String
    Dim sAnswer As String
    Dim sToEmail As String
    Dim sBuf As String
    Dim sTemp As String

    Dim oOutlookApp As Outlook.Application
    Dim oOutlookMail As Outlook.MailItem
    Dim oWordEditor As Word.editor
    Dim oWordApp As New Word.Application
    Dim oWordDoc As New Word.Document

    Dim qdf As QueryDef
    Dim rstQuery As DAO.Recordset
    
    Dim iFileNum As Long

    Dim cJob As Job
    Set cJob = New Job
    
    Call fPPGenerateJSONInfo

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]


    Call pfGetOrderingAttorneyInfo
    Call pfCurrentCaseInfo

    Set qdf = CurrentDb.QueryDefs(qTRIQPlusCases)
    qdf.Parameters(0) = sCourtDatesID
    Set rstQuery = qdf.OpenRecordset

    sParty1 = rstQuery.Fields("Party1").Value
    sParty2 = rstQuery.Fields("Party2").Value
    sName = sFirstName & " " & sLastName
    vPPInvoiceNo = rstQuery.Fields("TRInv.PPID").Value
    vPPInvoiceNo = Right(vPPInvoiceNo, 20)
    vPPInvoiceNo = Replace(Replace(vPPInvoiceNo, " ", ""), "-", "")

    'create pp invoice link
    vPPLink = "https://www.paypal.com/invoice/p/#" & vPPInvoiceNo

    'create PPButton html (construct string, save it into text file in job folder)
    sHTMLPPB = "<html><body><br><br><a href =" & Chr(34) & vPPLink & Chr(34) & "><img src=" & Chr(34) & "https://www.paypalobjects.com/webstatic/en_US/i/buttons/checkout-logo-large.png" & _
                                                                                                      Chr(34) & "alt=" & Chr(34) & "Check out with PayPal" & Chr(34) & "/></a></body></html>"
    'open txt file, save as html file

    Open cJob.DocPath.PPButtonT For Output As #1
    Write #1, sHTMLPPB
    Close #1

    iFileNum = FreeFile
    Open cJob.DocPath.PPButtonT For Input As iFileNum

    Do Until EOF(iFileNum)
        Line Input #iFileNum, sBuf
        sTemp = sTemp & sBuf & vbCrLf
    Loop

    Close iFileNum

    sTemp = Replace(sTemp, Chr(34) & "<html>", "<html>")
    sTemp = Replace(sTemp, "</html>" & Chr(34), "</html>")
    sTemp = Replace(sTemp, Chr(34) & Chr(34), Chr(34))

    'Save txt file as (if possible)

    iFileNum = FreeFile
    Open cJob.DocPath.PPButtonT For Output As iFileNum

    Print #iFileNum, sTemp

    Close iFileNum

    Name cJob.DocPath.PPButtonT As cJob.DocPath.PPButton

    'paste PPButton html into mail merged invoice/PQ at both bookmarks "PPButton" bookmark or #PPB1# AND "PPButton2" or #PPB2#

    Set oWordApp = CreateObject("Word.Application")
    oWordApp.Visible = False

    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPButton)

    oWordDoc.Content.Copy

    Set oWordApp = CreateObject("Word.Application")
    oWordApp.Visible = True

    Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPBalanceDue)

    With oWordDoc.Application

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
            '.Application.Selection.PasteAndFormat (wdPasteDefault)
            .Application.Selection.PasteAndFormat (wdFormatOriginalFormatting)
        
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
            .Application.Selection.PasteAndFormat (wdFormatOriginalFormatting)
        End With
    
        'save invoice
        .ActiveDocument.SaveAs2 FileName:=cJob.DocPath.PPBalanceDue
    
    End With

    oWordDoc.Close
    oWordApp.Quit
    
    Set oWordApp = Nothing
    Set oWordDoc = Nothing
    

    'prompt to ask when pdf is made
    sQuestion = "Click yes after you have created your final invoice PDF at " & cJob.DocPath.InvoiceP
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    If sAnswer = vbNo Then                       'IF NO THEN THIS HAPPENS

    
        DoCmd.OutputTo acOutputQuery, qTRIQPlusCases, acFormatXLS, cJob.DocPath.InvoiceInfo, False
    
        Set oWordApp = GetObject(cJob.DocPath.PPBalanceDueT, "Word.Document")
        Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPBalanceDueT)
    
        oWordApp.Application.Visible = False
    
        oWordDoc.MailMerge.OpenDataSource Name:=cJob.DocPath.InvoiceInfo, ReadOnly:=True
        oWordDoc.MailMerge.Execute
        oWordDoc.MailMerge.MainDocumentType = wdNotAMergeDocument
        oWordDoc.Application.ActiveDocument.SaveAs FileName:=cJob.DocPath.PPBalanceDue, FileFormat:=wdFormatDocument
    
        oWordApp.Quit
        Set oWordApp = Nothing
    
        Set oWordApp = CreateObject("Word.Application")
        oWordApp.Visible = False
    
        Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPButton)
    
        oWordDoc.Content.Copy
        oWordDoc.Close
        oWordApp.Quit
    
        Set oWordApp = Nothing
        Set oWordDoc = Nothing

        Set oWordApp = CreateObject("Word.Application")
        Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPBalanceDue)


        With oWordDoc.Application

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
        
                '.Application.Selection.PasteAndFormat (wdPasteDefault)
                .Application.Selection.PasteAndFormat (wdFormatOriginalFormatting)
        
            End With
    
    
            'save invoice
            oWordDoc.SaveAs2 FileName:=cJob.DocPath.PPBalanceDue
    
        End With

        oWordDoc.Close
        oWordApp.Quit
    
        Set oWordApp = Nothing
        Set oWordDoc = Nothing
    
    
        'generate PP-DraftInvoiceEmail in html format
        'oWordApp.Application.ActiveDocument.SaveAs filename:=cJob.DocPath.PPBalanceDueT, FileFormat:=wdFormatHTML
    
        rstQuery.Close
        Set rstQuery = Nothing
        qdf.Close
        Set qdf = Nothing
        Set oOutlookApp = CreateObject("Outlook.Application")
        Set oOutlookMail = oOutlookApp.CreateItem(0)
        On Error Resume Next
    
        Set oWordApp = CreateObject("Word.Application")
        Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPBalanceDue)
        oWordDoc.Content.Copy
    
        With oOutlookMail
            '@Ignore UnassignedVariableUsage
            .To = sToEmail
            .CC = sCompanyEmail
            .Subject = "Balance Due for " & sName & ", " & sParty1 & " v. " & sParty2
            .BodyFormat = olFormatRichText
            Set oWordEditor = .GetInspector.WordEditor
            .GetInspector.WordEditor.Content.Paste
            .Display
            .Attachments.Add (cJob.DocPath.InvoiceP)
        End With
        oWordDoc.Close
        oWordApp.Quit
        Set oWordApp = Nothing
        Set oWordDoc = Nothing
    
    
    Else                                         'if yes then this happens
    
    
        DoCmd.OutputTo acOutputQuery, qTRIQPlusCases, acFormatXLS, cJob.DocPath.InvoiceInfo, False
    
        '
        Set oWordApp = CreateObject("Word.Application")
        Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPBalanceDueT)
        oWordApp.Application.Visible = False
    
        oWordDoc.MailMerge.OpenDataSource Name:=cJob.DocPath.InvoiceInfo, ReadOnly:=True
        oWordDoc.MailMerge.Execute
        oWordDoc.MailMerge.MainDocumentType = wdNotAMergeDocument
        oWordDoc.Application.ActiveDocument.SaveAs2 FileName:=cJob.DocPath.PPBalanceDue

    
        oWordDoc.Close
        Set oWordApp = Nothing
        Set oWordDoc = Nothing
    
        Set oWordApp = CreateObject("Word.Application")
        oWordApp.Visible = False
    
        Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPButton)
    
        oWordDoc.Content.Copy
        oWordDoc.Close
        oWordApp.Quit
    
        Set oWordApp = Nothing
        Set oWordDoc = Nothing

        Set oWordApp = CreateObject("Word.Application")
        Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPBalanceDue)


        With oWordDoc.Application

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
        
                '.Application.Selection.PasteAndFormat (wdPasteDefault)
                .Application.Selection.PasteAndFormat (wdFormatOriginalFormatting)
        
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
                '.Application.Selection.Range.PasteSpecial DataType:=wdPasteHTML
                .Application.Selection.PasteAndFormat (wdFormatOriginalFormatting)
            End With
    
            'save invoice
            .ActiveDocument.SaveAs2 FileName:=cJob.DocPath.PPBalanceDue
    
    
        End With
        oWordDoc.Close
        oWordApp.Quit
    
        Set oWordApp = Nothing
        Set oWordDoc = Nothing
    
    
        'generate PP-DraftInvoiceEmail in html format
        'oWordApp.Application.ActiveDocument.SaveAs filename:=cJob.DocPath.PPBalanceDueT, FileFormat:=wdFormatHTML
    
        rstQuery.Close
        Set rstQuery = Nothing
        qdf.Close
        Set qdf = Nothing
        Set oOutlookApp = CreateObject("Outlook.Application")
        Set oOutlookMail = oOutlookApp.CreateItem(0)
        On Error Resume Next
    
        Set oWordApp = CreateObject("Word.Application")
        Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPBalanceDue)
        oWordDoc.Content.Copy
    
        With oOutlookMail
            '@Ignore UnassignedVariableUsage
            .To = sToEmail
            .CC = sCompanyEmail
            .Subject = "Balance Due for " & sName & ", " & sParty1 & " v. " & sParty2
            .BodyFormat = olFormatRichText
            Set oWordEditor = .GetInspector.WordEditor
            .GetInspector.WordEditor.Content.Paste
            .Display
            .Attachments.Add (cJob.DocPath.InvoiceP)
        End With
        oWordDoc.Close
        oWordApp.Quit
        Set oWordApp = Nothing
        Set oWordDoc = Nothing
    
    End If
    On Error GoTo 0
    Call pfCommunicationHistoryAdd("PP Invoice Sent")

End Sub

Public Sub fSendPPEmailDeposit()
    '============================================================================
    ' Name        : fPayPalUpdateCheck
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fPayPalUpdateCheck()
    ' Description : generates PayPal email for deposit
    '============================================================================
    
    Dim sName As String
    Dim vPPInvoiceNo As String
    Dim sHTMLPPB As String
    Dim vPPLink As String
    Dim sQuestion As String
    Dim sAnswer As String
    Dim sToEmail As String
    Dim sFileNameOut As String

    Dim oOutlookApp As Outlook.Application
    Dim oOutlookMail As Outlook.MailItem
    Dim oWordEditor As Word.editor
    Dim oWordApp As New Word.Application
    Dim oWordDoc As New Word.Document
    Dim oWordDoc1 As New Word.Document

    Dim qdf As QueryDef
    Dim rstQuery As DAO.Recordset
    Dim iFileNumIn As Long
    Dim iFileNumOut As Long
    
    Dim cJob As Job
    Set cJob = New Job
    
    'your invoice docx template MUST contain the phrase "#PPB1#" AND "#PPB2#" without the quotes somewhere on it.
    'your e-mail docx template MUST contain the phrase "#PPB1#" without the quotes somewhere on it.

    Call fPPGenerateJSONInfo                     'refreshes some info, not relevant for purposes of this code being on GH

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]


    Call pfGetOrderingAttorneyInfo               'refreshes some info, not relevant for purposes of this code being on GH

    'paste PPButton html into mail merged invoice/PQ at both bookmarks "PPButton" bookmark or #PPB1# AND "PPButton2" or #PPB2#

    Set qdf = CurrentDb.QueryDefs(qTRIQPlusCases)

    qdf.Parameters(0) = sCourtDatesID
    Set rstQuery = qdf.OpenRecordset
    sParty1 = rstQuery.Fields("Party1").Value
    sParty2 = rstQuery.Fields("Party2").Value
    vPPInvoiceNo = rstQuery.Fields("TRInv.PPID").Value
    vPPInvoiceNo = Right(vPPInvoiceNo, 20)
    vPPInvoiceNo = Replace(Replace(vPPInvoiceNo, " ", ""), "-", "")
    sName = sFirstName & " " & sLastName
    sToEmail = sNotes

    'create pp invoice link
    vPPLink = "https://www.paypal.com/invoice/p/#" & vPPInvoiceNo
    vPPInvoiceNo = ""

    'create PPButton html (construct string, save it into text file in job folder)
    sHTMLPPB = "<html><body><br><br><a href =" & Chr(34) & vPPLink & Chr(34) & _
                                                                             "><img src=" & Chr(34) & "https://www.paypalobjects.com/webstatic/en_US/i/buttons/checkout-logo-large.png" & _
                                                                             Chr(34) & "alt=" & Chr(34) & "Check out with PayPal" & Chr(34) & "/></a></body></html>"

    'open txt file, save as html file
    Open cJob.DocPath.PPButtonT For Output As #1

    Write #1, sHTMLPPB
    Close #1

    sFileNameOut = Replace(sHTMLPPB, Chr(34) & "<html>", "<html>")
    iFileNumOut = FreeFile
    Open cJob.DocPath.PPButtonT For Output As #1
    Print #1, sFileNameOut
    Close #1

    sFileNameOut = Replace(sHTMLPPB, "</html>" & Chr(34), "</html>")
    iFileNumOut = FreeFile
    Open cJob.DocPath.PPButtonT For Output As #1
    Print #1, sFileNameOut
    Close #1

    sFileNameOut = Replace(sHTMLPPB, Chr(34) & Chr(34), Chr(34))
    iFileNumOut = FreeFile
    Open cJob.DocPath.PPButtonT For Output As #1
    Print #1, sFileNameOut
    Close #1

    Name cJob.DocPath.PPButtonT As cJob.DocPath.PPButton

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
        .ActiveDocument.SaveAs2 FileName:=cJob.DocPath.InvoiceD
    
    End With

    oWordDoc.Close
    oWordApp.Quit
    
    Set oWordApp = Nothing
    Set oWordDoc = Nothing
    

    'prompt to ask when pdf of invoice is made
    sQuestion = "Click yes after you have created your final invoice PDF at " & cJob.DocPath.InvoiceP
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    If sAnswer = vbNo Then                       'IF NO THEN THIS HAPPENS

    
        MsgBox "No invoice will be sent at this time."
    
    
    Else                                         'if yes then this happens
    
    
        DoCmd.OutputTo acOutputQuery, qTRIQPlusCases, acFormatXLS, cJob.DocPath.InvoiceInfo, False
       
        Set oWordApp = CreateObject("Word.Application")
        Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPDIET)

        oWordApp.Application.Visible = False
        oWordDoc.MailMerge.OpenDataSource Name:=cJob.DocPath.InvoiceInfo, ReadOnly:=True
        oWordDoc.MailMerge.Execute
        oWordDoc.MailMerge.MainDocumentType = wdNotAMergeDocument
        oWordDoc.Application.ActiveDocument.SaveAs2 FileName:=cJob.DocPath.PPDraftInvoiceEmail

        Set oWordApp = CreateObject("Word.Application")
        oWordApp.Visible = False
    
        Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPButton)
    
        oWordDoc.Content.Copy
    
        Set oWordApp = CreateObject("Word.Application")
        Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPDraftInvoiceEmail)

        With oWordDoc.Application

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
        
                .Application.Selection.PasteAndFormat (wdFormatOriginalFormatting) 'paste button html file
        
            End With
           
            'save invoice
            oWordDoc.Save
    
        End With
        oWordDoc.Close
        oWordApp.Quit
    
        Set oWordApp = Nothing
        Set oWordDoc = Nothing
        
        rstQuery.Close
        Set rstQuery = Nothing
        qdf.Close
        Set qdf = Nothing
        Set oOutlookApp = CreateObject("Outlook.Application")
        Set oOutlookMail = oOutlookApp.CreateItem(0)
        On Error Resume Next
    
        Set oWordApp = CreateObject("Word.Application")
        Set oWordDoc = oWordApp.Documents.Open(cJob.DocPath.PPDraftInvoiceEmail) 'invoice email file in docx format
        oWordDoc.Content.Copy
    
        With oOutlookMail                        'now, you should have an e-mail with a PP button as well as an invoice with two PP buttons on it.
            .To = sToEmail
            .CC = sCompanyEmail
            .Subject = "Deposit Invoice for " & sName & ", " & sParty1 & " v. " & sParty2
            .BodyFormat = olFormatRichText
            Set oWordEditor = .GetInspector.WordEditor
            .GetInspector.WordEditor.Content.Paste
            .Display
            .Attachments.Add (cJob.DocPath.InvoiceP)
        End With
    
    End If
    oWordDoc.Close
    oWordApp.Quit
    On Error GoTo 0

    Set oOutlookApp = Nothing
    Set oOutlookMail = Nothing
    Set oWordEditor = Nothing
    Set oWordApp = Nothing
    Set oWordDoc = Nothing

    Call pfCommunicationHistoryAdd("PP Invoice Sent") 'record entry in comm history table for logs
    Call pfClearGlobals
End Sub

Public Sub fPPGetInvoiceInfo()
    '============================================================================
    ' Name        : fPPGetInvoiceInfo
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fPPGetInvoiceInfo
    ' Description : gets status of invoice
    '============================================================================

    Dim sURL As String
    Dim sUserName As String
    Dim sPassword As String
    Dim sAuth As String
    Dim stringJSON As String
    Dim apiWaxLRS As String
    Dim vErrorIssue As String
    Dim sInvoiceTime As String
    Dim sInvoiceNo As String
    Dim sEmail As String
    Dim sFirstName As String
    Dim sLastName As String
    Dim sDescription As String
    Dim sInvoiceDate As String
    Dim sPaymentTerms As String
    Dim sCourtDatesID As String
    Dim sNote As String
    Dim sTerms As String
    Dim sMinimumAmount As String
    Dim vmMemo As String
    Dim vlURL As String
    Dim sTemplateID As String
    Dim vTotal As String
    Dim sCity As String
    Dim sState As String
    Dim sZIP As String
    Dim sQuantity As String
    Dim sValue As String
    Dim vInvoiceID As String
    Dim sInvoiceNumber As String
    Dim vStatus As String
    Dim resp As Variant
    Dim response As Variant
    Dim rep As Variant
    Dim sToken As String
    Dim json1 As String
    Dim json2 As String
    Dim json3 As String
    Dim json4 As String
    Dim json5 As String
    Dim vErrorName As String
    Dim vErrorMessage As String
    Dim vErrorILink As String
    Dim vErrorDetails As String
    Dim vTermDays As String
    Dim vDetails As String

    Dim oRequest As Object
    Dim Json As Object
    
    Dim parsed As Dictionary

    Call fPPGenerateJSONInfo

    sURL = "https://api.paypal.com/v1/oauth2/token/"
    sEmail = sCompanyEmail
    '  https://api.paypal.com/v1/oauth2/token \
    'sAuth = TextBase64Encode(myCn.GetConnection, "us-ascii") 'mycn.GetConnection
    sAuth = TextBase64Encode(Environ("ppUserName") & ":" & Environ("ppPassword"), "us-ascii") 'mycn.GetConnection
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        '.Visible = True
        .Open "POST", sURL, False
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Accept-Language", "en_US"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        '.setRequestHeader "grant_type=client_credentials"
        '-u "client_id:secret" \
        .setRequestHeader "Authorization", "Basic " + sAuth
        '.setRequestHeader "Authorization", "Bearer " & sAuth
        .send ("grant_type=client_credentials")
        apiWaxLRS = .responseText
        Debug.Print apiWaxLRS
        Set parsed = JsonConverter.ParseJson(apiWaxLRS)
        sToken = parsed("access_token")          'third level array
        .abort
        'Debug.Print "--------------------------------------------"
    End With
    vInvoiceID = sPPID 'rstTRQPlusCases.Fields("TRInv.PPID").Value ' "INV2-C8EE-ZVC5-5U36-MF27" 'INV2-K8L5-ML2R-2GLL-7KW6
  
    'Debug.Print "RESPONSETEXT--------------------------------------------"
    sURL = "https://api.paypal.com/v1/invoicing/invoices/" & vInvoiceID
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        '.Visible = True
        .Open "GET", sURL, False
        '.setRequestHeader "Accept", "application/json" 'application/x-www-form-urlencoded
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & sToken
        ' json5 = json1 & json2 & json3
        .send                                    ' json5
        apiWaxLRS = .responseText
        sToken = ""
        .abort
        Debug.Print apiWaxLRS
        Debug.Print "--------------------------------------------"
    End With
    Set parsed = JsonConverter.ParseJson(apiWaxLRS)
    sInvoiceNumber = parsed("number")            'third level array
    vInvoiceID = parsed("id")                    'third level array
    vStatus = parsed("status")                   'third level array 'can be either DRAFT, UNPAID, SENT, SCHEDULED, PARTIALLY_PAID, PAYMENT_PENDING, PAID, MARKED_AS_PAID,
    'CANCELLED, REFUNDED, PARTIALLY_REFUNDED, MARKED AS REFUNDED
    vTotal = parsed("total_amount")("value")     'second level array
    vErrorName = parsed("name")                  '("value") 'second level array
    vErrorMessage = parsed("message")            '("value") 'second level array
    vErrorILink = parsed("information_link")     '("value") 'second level array
    vDetails = parsed("details")                 'second level array
    'For Each rep In vDetails ' third level objects
    vErrorIssue = parsed("field")
    vErrorDetails = parsed("issue")
    'Next
    Debug.Print "--------------------------------------------"
    Debug.Print "Error Name:  " & vErrorName
    Debug.Print "Error Message:  " & vErrorMessage
    Debug.Print "Error Info Link:  " & vErrorILink
    Debug.Print "Error Field:  " & vErrorIssue
    Debug.Print "Error Details:  " & vErrorDetails
    Debug.Print "Details:  " & vDetails
    Debug.Print "--------------------------------------------"
    '"id":"INV2-FZTH-K3T4-TM6Y-WC4U"
    '"number":"0003"
    '"status":"DRAFT"
    '"total_amount":{"currency":"USD","value":"3.00"},
    'Next
    Debug.Print "Invoice No.:  " & sInvoiceNumber & "   |   Invoice ID:  " & vInvoiceID
    Debug.Print "Status:  " & vStatus & "   |   Total:  " & vTotal
    Debug.Print "--------------------------------------------"


End Sub

Public Sub fPPRefund()
    '============================================================================
    ' Name        : fPPRefund
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fPPRefund()
    ' Description : refund with PayPal
    '============================================================================
    '
    Dim vAmount As String
    Dim sURL As String
    Dim sUserName As String
    Dim sPassword As String
    Dim sAuth As String
    Dim stringJSON As String
    Dim sEmail As String
    Dim apiWaxLRS As String
    Dim vErrorIssue As String
    Dim sInvoiceTime As String
    Dim sInvoiceNo As String
    Dim sFirstName As String
    Dim sLastName As String
    Dim sDescription As String
    Dim sInvoiceDate As String
    Dim sPaymentTerms As String
    Dim sCourtDatesID As String
    Dim sNote As String
    Dim sTerms As String
    Dim sMinimumAmount As String
    Dim vmMemo As String
    Dim vlURL As String
    Dim sTemplateID As String
    Dim vTotal As String
    Dim sLine1 As String
    Dim sCity As String
    Dim sState As String
    Dim sZIP As String
    Dim sQuantity As String
    Dim sValue As String
    Dim vInvoiceID As String
    Dim sInvoiceNumber As String
    Dim vStatus As String
    Dim resp As Variant
    Dim response As Variant
    Dim rep As Variant
    Dim sToken As String
    Dim json1 As String
    Dim json2 As String
    Dim json3 As String
    Dim json4 As String
    Dim json5 As String
    Dim vErrorName As String
    Dim vErrorMessage As String
    Dim vErrorILink As String
    Dim vErrorDetails As String
    Dim vTermDays As String
    Dim vDetails As String
    Dim vPPLink As String
    Dim sQuestion As String
    Dim sAnswer As String
    Dim vPPInvoiceNo As String
    
    Dim oRequest As Object
    Dim Json As Object
    
    Dim qdf As QueryDef
    Dim rstQuery As DAO.Recordset
    
    Dim parsed As Dictionary
    
    Dim cJob As Job
    Set cJob = New Job
    
    Call fPPGenerateJSONInfo

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

    Call pfGetOrderingAttorneyInfo
    Call pfCurrentCaseInfo

    Set qdf = CurrentDb.QueryDefs(qTRIQPlusCases)
    qdf.Parameters(0) = sCourtDatesID
    Set rstQuery = qdf.OpenRecordset

    sInvoiceNumber = rstQuery.Fields("TRInv.CourtDates.InvoiceNo").Value
    'get pp invoice ID
    vPPInvoiceNo = rstQuery.Fields("TRInv.PPID").Value
    sPaymentSum = rstQuery.Fields("TRInv.PaymentSum").Value
    sFinalPrice = rstQuery.Fields("TRInv.FinalPrice").Value
    vPPInvoiceNo = Right(vPPInvoiceNo, 20)
    vPPInvoiceNo = Replace(Replace(vPPInvoiceNo, " ", ""), "-", "")

    'create pp invoice link
    vPPLink = "https://www.paypal.com/invoice/p/#" & vPPInvoiceNo

    'calculate refund
    vAmount = InputBox("How much do you want to refund?  Their payments total up to " & sPaymentSum & " their final bill came to " & sFinalPrice & ".  This leaves a difference of " & (sPaymentSum - sFinalPrice) & " left owing to them.")
    
    sQuestion = "Are you sure that's correct?  You want to refund $" & vAmount & "?"
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

    If sAnswer = vbNo Then                       'Code for No

        sQuestion = "Do you want to refund anything?"
        sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
    
        If sAnswer = vbNo Then                   'Code for No

            GoTo Exitif
        Else

            vAmount = InputBox("How much do you want to refund?  Their payments total up to " & sPaymentSum & " their final bill came to " & sFinalPrice & ".  This leaves a difference of " & (sPaymentSum - sFinalPrice) & " left owing to them.")
    
        End If
    
    Else                                         'Code for yes
    End If

    sInvoiceDate = (Format((Date + 28), "yyyy-mm-dd")) & " PST"
    sInvoiceTime = (Format(Now(), "hh:mm:ss"))
    sURL = "https://api.paypal.com/v1/oauth2/token/"
    sEmail = sCompanyEmail
    '  https://api.paypal.com/v1/oauth2/token \
    'sAuth = TextBase64Encode(myCn.GetConnection, "us-ascii") 'mycn.GetConnection

    sAuth = TextBase64Encode(Environ("ppUserName") & ":" & Environ("ppPassword"), "us-ascii") 'mycn.GetConnection
    sPassword = ""
    sUserName = ""
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        '.Visible = True
        .Open "POST", sURL, False
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Accept-Language", "en_US"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        '.setRequestHeader "grant_type=client_credentials"
        '-u "client_id:secret" \
        .setRequestHeader "Authorization", "Basic " + sAuth
        '.setRequestHeader "Authorization", "Bearer " & sAuth
        .send ("grant_type=client_credentials")
        apiWaxLRS = .responseText
        Debug.Print apiWaxLRS
        Set parsed = JsonConverter.ParseJson(apiWaxLRS)
        sToken = parsed("access_token")          'third level array
        sAuth = ""
        .abort
        Debug.Print "--------------------------------------------"
    End With
       
       
    json1 = "{" & _
            Chr(34) & "date" & Chr(34) & ": " & Chr(34) & sInvoiceDate & " " & sInvoiceTime & Chr(34) & ",{" & Chr(34) & _
            "note" & Chr(34) & ": " & Chr(34) & "Refund as described for Invoice" & sInvoiceNumber & Chr(34) & "," & Chr(34) & _
            "amount" & Chr(34) & ": {" & Chr(34) & _
            "currency" & Chr(34) & ": " & Chr(34) & "USD" & Chr(34) & "," & Chr(34) & _
            "value" & Chr(34) & ": " & Chr(34) & vAmount & Chr(34) & "}}"
    Debug.Print "JSON1--------------------------------------------"
    Debug.Print json1
    'Debug.Print json4
    Debug.Print "RESPONSETEXT--------------------------------------------"
    sURL = "https://api.paypal.com/v1/invoicing/invoices/" & vInvoiceID & "/record-refund"
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        '.Visible = True
        .Open "POST", sURL, False
        '.setRequestHeader "Accept", "application/json" 'application/x-www-form-urlencoded
        '.setRequestHeader "content-type", "application/x-www-form-urlencoded"
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & sToken
        .send json1
        apiWaxLRS = .responseText
        sToken = ""
        .abort
        Debug.Print apiWaxLRS
        Debug.Print "--------------------------------------------"
    End With
    Set parsed = JsonConverter.ParseJson(apiWaxLRS)
    sInvoiceNumber = parsed("number")            'third level array
    vInvoiceID = parsed("id")                    'third level array
    vStatus = parsed("status")                   'third level array
    vTotal = parsed("total_amount")("value")     'second level array
    vErrorName = parsed("name")                  '("value") 'second level array
    vErrorMessage = parsed("message")            '("value") 'second level array
    vErrorILink = parsed("information_link")     '("value") 'second level array
    vDetails = parsed("details")                 'second level array
    'For Each rep In vDetails ' third level objects
    '   vErrorIssue = rep("field")
    '  vErrorDetails = rep("issue")
    'Next
    Debug.Print "--------------------------------------------"
    Debug.Print "Error Name:  " & vErrorName
    Debug.Print "Error Message:  " & vErrorMessage
    Debug.Print "Error Info Link:  " & vErrorILink
    'Debug.Print "Error Field:  " & vErrorIssue
    'Debug.Print "Error Details:  " & vErrorDetails
    Debug.Print "--------------------------------------------"
    '"id":"INV2-FZTH-K3T4-TM6Y-WC4U"
    '"number":"0003"
    '"status":"DRAFT"
    '"total_amount":{"currency":"USD","value":"3.00"},
    'Next
    Debug.Print "Invoice No.:  " & sInvoiceNumber & "   |   Invoice ID:  " & vInvoiceID
    Debug.Print "Status:  " & vStatus & "   |   Total:  " & vTotal
    Debug.Print "--------------------------------------------"

    If vStatus = "200" Or vStatus = "200 OK" Or ((vStatus) Like ("*" & "200" & "*")) Then

        'send client refund notice
        Call pfSendWordDocAsEmail("PP-RefundMadeEmail", "Refund Issued")
    
        'comm history entry
        Call pfCommunicationHistoryAdd("PP-RefundMadeEmail")
    
        'update db to show refund, a negative entry in payments
        Call fPaymentAdd(sInvoiceNumber, vAmount)
    
    
    Else
    End If
        
Exitif:
End Sub

Public Function TextBase64Encode(sText As String, sCharset As Variant) As Variant
    '============================================================================
    ' Name        : TextBase64Encode
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call TextBase64Encode(sText, sCharset)
    ' Description : encodes to base64
    '============================================================================
    '
    Dim aBinary As Variant

    With CreateObject("ADODB.Stream")
        .Type = 2                                ' adTypeText
        .Open
        .Charset = sCharset
        .WriteText sText
        .Position = 0
        .Type = 1                                ' adTypeBinary
        aBinary = .Read
        .Close
    End With
    With CreateObject("Microsoft.XMLDOM").createElement("objNode")
        .DataType = "bin.base64"
        .nodeTypedValue = aBinary
        TextBase64Encode = Replace(Replace(.Text, vbCr, ""), vbLf, "")
    End With

End Function

Public Sub fPPUpdate()
    '============================================================================
    ' Name        : fPPUpdate
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fPPUpdate
    ' Description : updates PayPal invoice on PayPal website
    '============================================================================

    Dim sURL As String
    Dim sAuth As String
    Dim stringJSON As String
    Dim sEmail As String
    Dim apiWaxLRS As String
    Dim vErrorIssue As String
    Dim sInvoiceTime As String
    Dim sInvoiceNo As String
    Dim sFirstName As String
    Dim sLastName As String
    Dim sDescription As String
    Dim sInvoiceDate As String
    Dim sPaymentTerms As String
    Dim sCourtDatesID As String
    Dim sNote As String
    Dim sTerms As String
    Dim sMinimumAmount As String
    Dim vmMemo As String
    Dim vlURL As String
    Dim sTemplateID As String
    Dim vTotal As String
    Dim sLine1 As String
    Dim sCity As String
    Dim sState As String
    Dim sZIP As String
    Dim sQuantity As String
    Dim sValue As String
    Dim vInvoiceID As String
    Dim sInvoiceNumber As String
    Dim vStatus As String
    Dim resp As Variant
    Dim response As Variant
    Dim rep As Variant
    Dim sToken As String
    Dim json1 As String
    Dim json2 As String
    Dim json3 As String
    Dim json4 As String
    Dim json5 As String
    Dim vErrorName As String
    Dim vErrorMessage As String
    Dim vErrorILink As String
    Dim vErrorDetails As String
    Dim vDetails As String

    Dim oRequest As Object
    Dim Json As Object
    
    Dim parsed As Dictionary
    

    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

    Call fPPGenerateJSONInfo
    Call pfGetOrderingAttorneyInfo

    sURL = "https://api.paypal.com/v1/oauth2/token/"
    sEmail = sCompanyEmail
    '  https://api.paypal.com/v1/oauth2/token \
    'sAuth = TextBase64Encode(myCn.GetConnection, "us-ascii") 'mycn.GetConnection
    
    sAuth = TextBase64Encode(Environ("ppUserName") & ":" & Environ("ppPassword"), "us-ascii") 'mycn.GetConnection
    
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        '.Visible = True
        .Open "POST", sURL, False
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Accept-Language", "en_US"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        '.setRequestHeader "grant_type=client_credentials"
        '-u "client_id:secret" \
        .setRequestHeader "Authorization", "Basic " + sAuth
        '.setRequestHeader "Authorization", "Bearer " & sAuth
        .send ("grant_type=client_credentials")
        apiWaxLRS = .responseText
        Debug.Print apiWaxLRS
        Set parsed = JsonConverter.ParseJson(apiWaxLRS)
        sToken = parsed("access_token")          'third level array
        sAuth = ""
        .abort
        Debug.Print "--------------------------------------------"
    End With
     
    json1 = "{" & _
            Chr(34) & "merchant_info" & Chr(34) & ": {" & Chr(34) & _
            "email" & Chr(34) & ": " & Chr(34) & sCompanyEmail & Chr(34) & "," & Chr(34) & _
            "first_name" & Chr(34) & ": " & Chr(34) & sCompanyFirstName & Chr(34) & "," & Chr(34) & _
            "last_name" & Chr(34) & ": " & Chr(34) & sCompanyLastName & Chr(34) & "," & Chr(34) & _
            "business_name" & Chr(34) & ": " & Chr(34) & sCompanyName & Chr(34) & "," & Chr(34) & _
            "phone" & Chr(34) & ": {" & Chr(34) & _
            "country_code" & Chr(34) & ": " & Chr(34) & sPCountryCode & Chr(34) & "," & Chr(34) & _
            "national_number" & Chr(34) & ": " & Chr(34) & sCompanyNationalNumber & Chr(34) & "}," & Chr(34) & _
            "address" & Chr(34) & ": {" & Chr(34) & _
            "line1" & Chr(34) & ": " & Chr(34) & sCompanyAddress & Chr(34) & "," & Chr(34) & _
            "city" & Chr(34) & ": " & Chr(34) & sCompanyCity & Chr(34) & "," & Chr(34) & _
            "state" & Chr(34) & ": " & Chr(34) & sCompanyState & Chr(34) & "," & Chr(34) & _
            "postal_code" & Chr(34) & ": " & Chr(34) & sCompanyZIP & Chr(34) & "," & Chr(34) & _
            "country_code" & Chr(34) & ": " & Chr(34) & sZCountryCode & Chr(34) & "}},"
    
    json2 = Chr(34) & "billing_info" & Chr(34) & ": [{" & Chr(34) & _
                                                                  "email" & Chr(34) & ": " & Chr(34) & sEmail & Chr(34) & "}],"
    
    '@Ignore UnassignedVariableUsage
    json3 = Chr(34) & "items" & Chr(34) & ": [" & "{" & Chr(34) & "name" & Chr(34) & ": " & Chr(34) & sDescription & Chr(34) & "," & Chr(34) & _
                                                                                                                                             "quantity" & Chr(34) & ": " & Chr(34) & sQuantity & Chr(34) & "," & Chr(34) & _
                                                                                                                                             "unit_price" & Chr(34) & ": {" & Chr(34) & _
                                                                                                                                             "currency" & Chr(34) & ": " & Chr(34) & "USD" & Chr(34) & "," & Chr(34) & _
                                                                                                                                             "value" & Chr(34) & ": " & Chr(34) & sUnitPrice & Chr(34) & "}}]," & Chr(34) & _
                                                                                                                                             "note" & Chr(34) & ": " & Chr(34) & sNote & Chr(34) & "," & Chr(34) & _
                                                                                                                                             "payment_term" & Chr(34) & ": {" & Chr(34) & _
                                                                                                                                             "term_type" & Chr(34) & ": " & Chr(34) & "NET_" & sTermDays & Chr(34) & "}," & Chr(34) & _
                                                                                                                                             "shipping_info" & Chr(34) & ": {" & Chr(34) & _
                                                                                                                                             "first_name" & Chr(34) & ": " & Chr(34) & sFirstName & Chr(34) & "," & Chr(34) & _
                                                                                                                                             "last_name" & Chr(34) & ": " & Chr(34) & sLastName & Chr(34) & "," & Chr(34) & _
                                                                                                                                             "business_name" & Chr(34) & ": " & Chr(34) & "WRTS Sample" & Chr(34) & "," & Chr(34) & _
                                                                                                                                             "address" & Chr(34) & ": {" & Chr(34) & _
                                                                                                                                             "line1" & Chr(34) & ": " & Chr(34) & sCompanyAddress & Chr(34) & "," & Chr(34) & _
                                                                                                                                             "city" & Chr(34) & ": " & Chr(34) & sCompanyCity & Chr(34) & "," & Chr(34) & _
                                                                                                                                             "state" & Chr(34) & ": " & Chr(34) & sCompanyState & Chr(34) & "," & Chr(34) & _
                                                                                                                                             "postal_code" & Chr(34) & ": " & Chr(34) & sCompanyZIP & Chr(34) & "," & Chr(34) & _
                                                                                                                                             "country_code" & Chr(34) & ": " & Chr(34) & "US" & Chr(34) & "}}}"
    Debug.Print "JSON1--------------------------------------------"
    Debug.Print json1
    Debug.Print "JSON2--------------------------------------------"
    Debug.Print json2
    Debug.Print "JSON3--------------------------------------------"
    Debug.Print json3
    'Debug.Print "JSON4--------------------------------------------"
    'Debug.Print json4
    Debug.Print "RESPONSETEXT--------------------------------------------"
    sURL = "https://api.paypal.com/v1/invoicing/invoices/" & vInvoiceID
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        '.Visible = True
        .Open "PUT", sURL, False
        '.setRequestHeader "Accept", "application/json" 'application/x-www-form-urlencoded
        '.setRequestHeader "content-type", "application/x-www-form-urlencoded"
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & sToken
        json5 = json1 & json2 & json3
        .send json5
        apiWaxLRS = .responseText
        sToken = ""
        .abort
        Debug.Print apiWaxLRS
        Debug.Print "--------------------------------------------"
    End With
    Set parsed = JsonConverter.ParseJson(apiWaxLRS)
    sInvoiceNumber = parsed("number")            'third level array
    vInvoiceID = parsed("id")                    'third level array
    vStatus = parsed("status")                   'third level array
    vTotal = parsed("total_amount")("value")     'second level array
    vErrorName = parsed("name")                  '("value") 'second level array
    vErrorMessage = parsed("message")            '("value") 'second level array
    vErrorILink = parsed("information_link")     '("value") 'second level array
    'vDetails = Parsed("details") 'second level array
    'For Each rep In vDetails ' third level objects
    '    vErrorIssue = rep("field")
    '    vErrorDetails = rep("issue")
    'Next
    Debug.Print "--------------------------------------------"
    Debug.Print "Error Name:  " & vErrorName
    Debug.Print "Error Message:  " & vErrorMessage
    Debug.Print "Error Info Link:  " & vErrorILink
    'Debug.Print "Error Field:  " & vErrorIssue
    'Debug.Print "Error Details:  " & vErrorDetails
    Debug.Print "--------------------------------------------"
    Debug.Print "Invoice No.:  " & sInvoiceNumber & "   |   Invoice ID:  " & vInvoiceID
    Debug.Print "Status:  " & vStatus & "   |   Total:  " & vTotal
    Debug.Print "--------------------------------------------"

    'update PPID & PPStatus
    Dim sUpdatePPStatus As String
    Dim sUpdatePPID As String
    
    sUpdatePPStatus = "UPDATE CourtDates SET PPStatus = " & Chr(34) & vStatus & Chr(34) & " WHERE [ID] = " & sCourtDatesID & ";"
    CurrentDb.Execute sUpdatePPStatus
    sUpdatePPID = "UPDATE CourtDates SET PPID = " & Chr(34) & vInvoiceID & Chr(34) & " WHERE [ID] = " & sCourtDatesID & ";"
    CurrentDb.Execute sUpdatePPID

End Sub

Public Sub fManualPPPayment()
    '============================================================================
    ' Name        : fManualPPPayment
    ' Author      : Erica L Ingram
    ' Copyright   : 2019, A Quo Co.
    ' Call command: Call fManualPPPayment
    ' Description : marks invoice as paid with manual payment, like with check/cash
    '============================================================================

    'curl -v -X POST https://api.sandbox.paypal.com/v1/invoicing/invoices/PPINV#/record-payment \
    '-H "Content-Type: application/json" \
    '-H "Authorization: Bearer Access-Token" \
    '-d '{
    '"method": "CASH", 'BANK_TRANSFER, CASH, CHECK, CREDIT_CARD, DEBIT_CARD, PAYPAL, WIRE_TRANSFER, OTHER.
    '"date": "2013-11-06 03:30:00 PST",
    '"note": "I got the payment by cash!",
    '"amount": {
    '  "currency": "USD",
    '  "value": "20.00"
    '}
    '}''
    Dim sURL As String
    Dim sAuth As String
    Dim stringJSON As String
    Dim sEmail As String
    Dim apiWaxLRS As String
    Dim vErrorIssue As String
    Dim sInvoiceTime As String
    Dim sInvoiceNo As String
    Dim sFirstName As String
    Dim sLastName As String
    Dim sDescription As String
    Dim sInvoiceDate As String
    Dim sPaymentTerms As String
    Dim sCourtDatesID As String
    Dim sNote As String
    Dim sTerms As String
    Dim sMinimumAmount As String
    Dim vmMemo As String
    Dim vlURL As String
    Dim sTemplateID As String
    Dim vTotal As String
    Dim sLine1 As String
    Dim sCity As String
    Dim sState As String
    Dim sZIP As String
    Dim sQuantity As String
    Dim sValue As String
    Dim vInvoiceID As String
    Dim sInvoiceNumber As String
    Dim vStatus As String
    Dim resp As Variant
    Dim response As Variant
    Dim rep As Variant
    Dim sToken As String
    Dim json1 As String
    Dim json2 As String
    Dim json3 As String
    Dim json4 As String
    Dim json5 As String
    Dim vErrorName As String
    Dim vErrorMessage As String
    Dim vErrorILink As String
    Dim vErrorDetails As String
    Dim vTermDays As String
    Dim vDetails As String
    Dim vMethod As String
    Dim vAmount As String
    
    Dim oRequest As Object
    Dim Json As Object
    
    Dim qdf As QueryDef
    Dim rstQInfoInvNo As DAO.Recordset

    Dim parsed As Dictionary

    Call fPPGenerateJSONInfo
    Call pfGetOrderingAttorneyInfo

    Set qdf = CurrentDb.QueryDefs("QInfobyInvoiceNumber")
    qdf.Parameters(0) = sCourtDatesID
    Set rstQInfoInvNo = qdf.OpenRecordset

    sInvoiceNumber = rstQInfoInvNo.Fields("InvoiceNo").Value
    sFinalPrice = rstQInfoInvNo.Fields("FinalPrice").Value

    sURL = "https://api.paypal.com/v1/oauth2/token/"
    sEmail = sCompanyEmail
    '  https://api.paypal.com/v1/oauth2/token \
    'sAuth = TextBase64Encode(myCn.GetConnection, "us-ascii") 'mycn.GetConnection
    
    sAuth = TextBase64Encode(Environ("ppUserName") & ":" & Environ("ppPassword"), "us-ascii") 'mycn.GetConnection
    
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        '.Visible = True
        .Open "POST", sURL, False
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Accept-Language", "en_US"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        '.setRequestHeader "grant_type=client_credentials"
        '-u "client_id:secret" \
        .setRequestHeader "Authorization", "Basic " + sAuth
        '.setRequestHeader "Authorization", "Bearer " & sAuth
        .send ("grant_type=client_credentials")
        apiWaxLRS = .responseText
        Debug.Print apiWaxLRS
        Set parsed = JsonConverter.ParseJson(apiWaxLRS)
        sToken = parsed("access_token")          'third level array
        sAuth = ""
        .abort
        Debug.Print "--------------------------------------------"
    End With
    vMethod = InputBox("What method was used to pay?  Select/type in either BANK_TRANSFER, CASH, CHECK, CREDIT_CARD, DEBIT_CARD, PAYPAL, WIRE_TRANSFER, or OTHER.")
    sInvoiceDate = (Format((Date + 28), "yyyy-mm-dd")) & " PST"
    sInvoiceTime = (Format(Now(), "hh:mm:ss"))
    vAmount = InputBox("How much was the payment?  Their invoice totals up to $" & sFinalPrice & ".")

    '@Ignore UnassignedVariableUsage
    json1 = "{" & _
            Chr(34) & "method" & Chr(34) & ": " & Chr(34) & vMethod & Chr(34) & "," & Chr(34) & _
            "date" & Chr(34) & ": " & Chr(34) & sInvoiceDate & " " & sInvoiceTime & Chr(34) & "," & Chr(34) & _
            "note" & Chr(34) & ": " & Chr(34) & sCourtDatesID & " " & sInvoiceNumber & Chr(34) & "," & Chr(34) & _
            "amount" & Chr(34) & ": {" & Chr(34) & _
            "currency" & Chr(34) & ": " & Chr(34) & "USD" & Chr(34) & "," & Chr(34) & _
            "value" & Chr(34) & ": " & Chr(34) & vAmount & Chr(34) & "}}"
    
    Debug.Print "JSON1--------------------------------------------"
    Debug.Print json1
    Debug.Print "RESPONSETEXT--------------------------------------------"
    sURL = "https://api.paypal.com/v1/invoicing/invoices/" & vInvoiceID
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        '.Visible = True
        .Open "PUT", sURL, False
        '.setRequestHeader "Accept", "application/json" 'application/x-www-form-urlencoded
        '.setRequestHeader "content-type", "application/x-www-form-urlencoded"
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & sToken
        '@Ignore UnassignedVariableUsage
        json5 = json1 & json2 & json3
        .send json5
        apiWaxLRS = .responseText
        sToken = ""
        .abort
        Debug.Print apiWaxLRS
        Debug.Print "--------------------------------------------"
    End With
    Set parsed = JsonConverter.ParseJson(apiWaxLRS)
    sInvoiceNumber = parsed("number")            'third level array
    vInvoiceID = parsed("id")                    'third level array
    vStatus = parsed("status")                   'third level array
    vTotal = parsed("total_amount")("value")     'second level array
    vErrorName = parsed("name")                  '("value") 'second level array
    vErrorMessage = parsed("message")            '("value") 'second level array
    vErrorILink = parsed("information_link")     '("value") 'second level array
    'vDetails = Parsed("details") 'second level array
    'For Each rep In vDetails ' third level objects
    '    vErrorIssue = rep("field")
    '    vErrorDetails = rep("issue")
    'Next
    Debug.Print "--------------------------------------------"
    Debug.Print "Error Name:  " & vErrorName
    Debug.Print "Error Message:  " & vErrorMessage
    Debug.Print "Error Info Link:  " & vErrorILink
    'Debug.Print "Error Field:  " & vErrorIssue
    'Debug.Print "Error Details:  " & vErrorDetails
    Debug.Print "--------------------------------------------"
    Debug.Print "Invoice No.:  " & sInvoiceNumber & "   |   Invoice ID:  " & vInvoiceID
    Debug.Print "Status:  " & vStatus & "   |   Total:  " & vTotal
    Debug.Print "--------------------------------------------"

    'update PPID & PPStatus
    Dim sUpdatePPStatus As String
    Dim sUpdatePPID As String
    '@Ignore UnassignedVariableUsage
    sUpdatePPStatus = "UPDATE CourtDates SET PPStatus = " & Chr(34) & vStatus & Chr(34) & " WHERE [ID] = " & sCourtDatesID & ";"
    CurrentDb.Execute sUpdatePPStatus
    '@Ignore UnassignedVariableUsage
    sUpdatePPID = "UPDATE CourtDates SET PPID = " & Chr(34) & vInvoiceID & Chr(34) & " WHERE [ID] = " & sCourtDatesID & ";"
    CurrentDb.Execute sUpdatePPID


End Sub


