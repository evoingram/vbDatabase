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


Sub fSendPPEmailFactored()
'generates factored invoice email for pp
Dim sFile1 As String, sInvoicePathPDF As String, sLine1 As String
Dim sInvoiceNumber As String, sName As String, vPPInvoiceNo As String
Dim sExportedTemplatePath As String, sTemplatePath As String, sOutputPDF As String
Dim sExportInfoCSVPath As String, sQueryName As String, sHTMLPPB As String
Dim sInvoicePathDocX As String, vPPLink As String, sFilePath As String
Dim sFilePathHTML As String, sQuestion As String, sAnswer As String
Dim sToEmail As String, sLine2 As String, sFileNameOut As String
Dim sBuf As String, sTemp As String, sFileName As String


Dim oOutlookApp As Outlook.Application, oOutlookMail As Outlook.MailItem, oWordEditor As Word.editor
Dim oWordApp As New Word.Application, oWordDoc As New Word.Document

Dim qdf As QueryDef
Dim rstQuery As DAO.Recordset
Dim iFileNum As Integer


Call fPPGenerateJSONInfo 'refreshes some necessary info

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

sQueryName = "TRInvoiQPlusCases"
sExportedTemplatePath = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-PP-FactoredInvoiceEmail.docx"
sExportInfoCSVPath = "I:\" & sCourtDatesID & "\WorkingFiles\" & sCourtDatesID & "-InvoiceInfo.xls"
sTemplatePath = "T:\Database\Templates\Stage4s\PP-FactoredInvoiceEmail-Template.docx"

Call pfGetOrderingAttorneyInfo 'refreshes some necessary info
Call pfCurrentCaseInfo 'refreshes some necessary info

Set qdf = CurrentDb.QueryDefs(sQueryName)
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

sFile1 = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-PP-FactoredInvoiceEmail.html"
sInvoicePathPDF = "I:\" & sCourtDatesID & "\Generated\" & sInvoiceNumber & ".pdf"
sInvoicePathDocX = "I:\" & sCourtDatesID & "\Generated\" & sInvoiceNumber & ".docx"

'create pp invoice link
vPPLink = "https://www.paypal.com/invoice/p/#" & vPPInvoiceNo

'create PPButton html (construct string, save it into text file in job folder)
sHTMLPPB = "<html><body><br><br><a href =" & Chr(34) & vPPLink & Chr(34) & "><img src=" & Chr(34) _
    & "https://www.paypalobjects.com/webstatic/en_US/i/buttons/checkout-logo-large.png" & _
        Chr(34) & "alt=" & Chr(34) & "Check out with PayPal" & Chr(34) & "/></a></body></html>"


sFilePath = "I:\" & sCourtDatesID & "\WorkingFiles\" & "PPButton.txt"
sFilePathHTML = "I:\" & sCourtDatesID & "\WorkingFiles\" & "PPButton.html"

'open txt file, save as html file

Open sFilePath For Output As #1
Write #1, sHTMLPPB
Close #1

iFileNum = FreeFile
Open sFilePath For Input As iFileNum

Do Until EOF(iFileNum)
    Line Input #iFileNum, sBuf
    sTemp = sTemp & sBuf & vbCrLf
Loop

Close iFileNum

sTemp = Replace(sTemp, Chr(34) & "<html>", "<html>") 'doing it this way makes weird things happen to the text file
sTemp = Replace(sTemp, "</html>" & Chr(34), "</html>") 'so these 3 sets of changes are necessary to form correct html
sTemp = Replace(sTemp, Chr(34) & Chr(34), Chr(34))


iFileNum = FreeFile
Open sFilePath For Output As iFileNum

Print #iFileNum, sTemp

Close iFileNum

Name sFilePath As sFilePathHTML 'Save txt file as (if possible)

'paste PPButton html into mail merged invoice/PQ at both bookmarks "PPButton" bookmark or #PPB1# AND "PPButton2" or #PPB2#
Set oWordApp = CreateObject("Word.Application")
oWordApp.Visible = False

Set oWordDoc = oWordApp.Documents.Open(sFilePathHTML)

oWordDoc.Content.Copy

Set oWordApp = CreateObject("Word.Application")
oWordApp.Visible = True

sExportedTemplatePath = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-PP-FactoredInvoiceEmail.docx"
Set oWordDoc = oWordApp.Documents.Open(sExportedTemplatePath)

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
   .ActiveDocument.SaveAs2 FileName:=sExportedTemplatePath
    
End With

    oWordDoc.Close
    oWordApp.Quit
    
    Set oWordApp = Nothing
    Set oWordDoc = Nothing
    

'prompt to ask when pdf is made
sQuestion = "Click yes after you have created your final invoice PDF at " & sInvoicePathPDF
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")


If sAnswer = vbNo Then 'IF NO THEN THIS HAPPENS
    
    MsgBox "No invoice will be sent."
    
Else 'if yes then this happens
    
    DoCmd.OutputTo acOutputQuery, sQueryName, acFormatXLS, sExportInfoCSVPath, False
    

    sTemplatePath = "T:\Database\Templates\Stage4s\PP-FactoredInvoiceEmail-Template.docx"
    'Set oWordApp = GetObject(sTemplatePath, "Word.Document")
    Set oWordDoc = oWordApp.Documents.Add(sTemplatePath)
    oWordApp.Application.Visible = False
    
    oWordDoc.MailMerge.OpenDataSource Name:=sExportInfoCSVPath, ReadOnly:=True
    oWordDoc.MailMerge.Execute
    oWordDoc.MailMerge.MainDocumentType = wdNotAMergeDocument
    oWordDoc.Application.ActiveDocument.SaveAs2 FileName:=sExportedTemplatePath

    
    oWordApp.Quit
    Set oWordApp = Nothing
    Set oWordDoc = Nothing
    
    Set oWordApp = CreateObject("Word.Application")
    oWordApp.Visible = False
    
    Set oWordDoc = oWordApp.Documents.Open(sFilePathHTML) 'copy button html file
    
    oWordDoc.Content.Copy
    oWordDoc.Close
    oWordApp.Quit
    
    Set oWordApp = Nothing
    Set oWordDoc = Nothing

    Set oWordApp = CreateObject("Word.Application")
    Set oWordDoc = oWordApp.Documents.Open(sExportedTemplatePath)


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
    oWordDoc.SaveAs FileName:=sExportedTemplatePath
    
    
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
    Set oWordDoc = oWordApp.Documents.Open(sExportedTemplatePath)
    oWordDoc.Content.Copy
    
    With oOutlookMail 'email that will also include pp button
        .To = sToEmail
        .CC = "inquiries@aquoco.co"
        .Subject = "Transcript Delivery & Invoice for " & sName & ", " & sParty1 & " v. " & sParty2
        .BodyFormat = olFormatRichText
        Set oWordEditor = .GetInspector.WordEditor
        .GetInspector.WordEditor.Content.Paste
        .Display
        .Attachments.Add (sInvoicePathPDF)
     End With
    
    oWordDoc.Close
    oWordApp.Quit
    Set oWordApp = Nothing
    Set oWordDoc = Nothing
End If
Call pfCommunicationHistoryAdd("PP Invoice Sent")
Call pfClearGlobals
End Sub





Sub fPPDraft()
'============================================================================
' Name        : fPPDraft
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fPPDraft()
' Description : creates PayPal draft invoice on PayPal website
'============================================================================

Dim sURL As String, sUserName As String, sPassword As String, sAuth As String, stringJSON As String, sEmail As String
Dim vInvoiceID As String, apiWaxLRS As String, vErrorIssue As String, sInvoiceTime As String
Dim oRequest As Object, Json As Object
Dim vStatus As String, vTotal As String

Dim rstRates As DAO.Recordset

Dim resp, response, rep, vDetails As Object
Dim sToken As String, json1 As String, json2 As String, json3 As String, json4 As String, json5 As String
Dim Parsed As Dictionary
Dim vErrorName As String, vErrorMessage As String, vErrorILink As String, vErrorDetails As String
Dim sFile1 As String, sFile2 As String, sText As String, sLine1 As String, sLine2 As String


Call fPPGenerateJSONInfo

Call pfGetOrderingAttorneyInfo

sIRC = 8
Set rstRates = CurrentDb.OpenRecordset("SELECT * FROM Rates WHERE [ID] = " & sIRC & ";")
vInventoryRateCode = rstRates.Fields("Code").Value
rstRates.Close


Debug.Print sCourtDatesID & " " & sInvoiceNumber

'Note: fPPDraft can delete following lines when known safe come back
'sFile1 = "C:\other\1.txt"
'sFile2 = "C:\other\2.txt"

'Open sFile1 For Input As #1
'Line Input #1, sLine1
'Close #1

'Open sFile2 For Input As #2
'Line Input #2, sLine2
'Close #2
'sUserName = sLine1
'sPassword = sLine2

sURL = "https://api.paypal.com/v1/oauth2/token/"
sEmail = "inquiries@aquoco.co"
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
        Set Parsed = JsonConverter.ParseJson(apiWaxLRS)
        sToken = Parsed("access_token") 'third level array
        .abort
Debug.Print "--------------------------------------------"
    End With
      
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
    "due_date" & Chr(34) & ": " & Chr(34) & sInvoiceDate & " " & sInvoiceTime & Chr(34) & "}," & Chr(34) & _
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
'Debug.Print "JSON1--------------------------------------------"
'Debug.Print json1
'Debug.Print "JSON2--------------------------------------------"
'Debug.Print json2
'Debug.Print "JSON3--------------------------------------------"
'Debug.Print json3
'Debug.Print "JSON4--------------------------------------------"
'Debug.Print json4
'Debug.Print "RESPONSETEXT--------------------------------------------"
    sURL = "https://api.paypal.com/v1/invoicing/invoices"
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        '.Visible = True
        .Open "POST", sURL, False
        '.setRequestHeader "Accept", "application/json" 'application/x-www-form-urlencoded
        '.setRequestHeader "content-type", "application/x-www-form-urlencoded"
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & sToken
        json5 = json1 & json2 & json3 & json4
        .send json5
        apiWaxLRS = .responseText
        sToken = ""
        .abort
        Debug.Print apiWaxLRS
'Debug.Print "--------------------------------------------"
    End With
Set Parsed = JsonConverter.ParseJson(apiWaxLRS)
sInvoiceNumber = Parsed("number") 'third level array
vInvoiceID = Parsed("id") 'third level array
vStatus = Parsed("status") 'third level array
'vTotal = Parsed("total_amount")("value") 'second level array
vErrorName = Parsed("name") '("value") 'second level array
vErrorMessage = Parsed("message") '("value") 'second level array
vErrorILink = Parsed("information_link") '("value") 'second level array
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
Debug.Print "Error Field:  " & vErrorIssue
Debug.Print "Error Details:  " & vErrorDetails
Debug.Print "--------------------------------------------"
'Next
Debug.Print "Invoice No.:  " & sInvoiceNumber & "   |   Invoice ID:  " & vInvoiceID
Debug.Print "Status:  " & vStatus & "   |   Total:  " & vTotal
Debug.Print "--------------------------------------------"

'update PPID & PPStatus
Dim sUpdatePPStatus As String, sUpdatePPID As String
sUpdatePPStatus = "UPDATE CourtDates SET PPStatus = " & Chr(34) & vStatus & Chr(34) & " WHERE [ID] = " & sCourtDatesID & ";"
CurrentDb.Execute sUpdatePPStatus
sUpdatePPID = "UPDATE CourtDates SET PPID = " & Chr(34) & vInvoiceID & Chr(34) & " WHERE [ID] = " & sCourtDatesID & ";"
CurrentDb.Execute sUpdatePPID
Call pfClearGlobals
End Sub


Sub fPayPalUpdateCheck()
'============================================================================
' Name        : fPayPalUpdateCheck
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fPayPalUpdateCheck()
' Description : check PayPal for update on invoice
'============================================================================

Dim sQueryName As String
Dim qdf As QueryDef, qdf1 As QueryDef
Dim rstQuery As DAO.Recordset, rstQuery1 As DAO.Recordset

Dim sURL As String, sUserName As String, sPassword As String, sAuth As String, stringJSON As String
Dim apiWaxLRS As String, vErrorIssue As String, sInvoiceTime As String
Dim oRequest As Object, Json As Object
Dim sInvoiceNo As String, sEmail As String, sFirstName As String, sLastName As String
Dim sDescription As String, sInvoiceDate As String, sPaymentTerms As String
Dim sCourtDatesID As String, sNote As String, sTerms As String, sMinimumAmount As String
Dim vmMemo As String, vlURL As String, sTemplateID As String, vTotal As String
Dim sCity As String, sState As String, sZIP As String, sQuantity As String
Dim sValue As String, vInvoiceID As String, sInvoiceNumber As String, vStatus As String
Dim resp As Object, response As Object, rep As Object
Dim sToken As String, json1 As String, json2 As String, json3 As String, json4 As String, json5 As String
Dim Parsed As Dictionary
Dim vErrorName As String, vErrorMessage As String, vErrorILink As String, vErrorDetails As String
Dim vTermDays As String, vDetails As String
Dim sFile1 As String, sFile2 As String, sText As String, sLine1 As String, sLine2 As String

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
            
            'note: fPayPalUpdateCheck can delete following lines when known safe come back
            'sFile1 = "C:\other\1.txt"
            'sFile2 = "C:\other\2.txt"
            
            'Open sFile1 For Input As #1
            'Line Input #1, sLine1
            'Close #1
            
            'Open sFile2 For Input As #2
            'Line Input #2, sLine2
            'Close #2
            'sUserName = sLine1
            'sPassword = sLine2
            
            sURL = "https://api.paypal.com/v1/oauth2/token/"
            sEmail = "inquiries@aquoco.co"
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
                    Set Parsed = JsonConverter.ParseJson(apiWaxLRS)
                    sToken = Parsed("access_token") 'third level array
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
                    .send ' json5
                    apiWaxLRS = .responseText
                    sToken = ""
                    .abort
                    Debug.Print apiWaxLRS
            Debug.Print "--------------------------------------------"
                End With
            Set Parsed = JsonConverter.ParseJson(apiWaxLRS)
            sInvoiceNumber = Parsed("number") 'third level array
            vInvoiceID = Parsed("id") 'third level array
            vStatus = Parsed("status") 'third level array 'can be either DRAFT, UNPAID, SENT, SCHEDULED, PARTIALLY_PAID, PAYMENT_PENDING, PAID, MARKED_AS_PAID,
                                                                        'CANCELLED, REFUNDED, PARTIALLY_REFUNDED, MARKED AS REFUNDED
                                                                        
            Debug.Print "--------------------------------------------"
            
            
            If vStatus <> vPPStatus And vStatus = "PAID" Or vStatus = "MARKED_AS_PAID" Then
            
                Call pfSendWordDocAsEmail("PP-PaymentMadeEmail", "Payment Received")  'send/queue payment received receipt automatically
                
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


Sub fSendPPEmailBalanceDue()
'============================================================================
' Name        : fPayPalUpdateCheck
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fPayPalUpdateCheck()
' Description : sends PayPal email for balance due
'============================================================================
'
Dim sFile1 As String, sInvoicePathPDF As String, sLine1 As String
Dim sInvoiceNumber As String, sName As String, vPPInvoiceNo As String
Dim sExportedTemplatePath As String, sTemplatePath As String, sOutputPDF As String
Dim sExportInfoCSVPath As String, sQueryName As String, sHTMLPPB As String
Dim sInvoicePathDocX As String, vPPLink As String, sFilePath As String
Dim sFilePathHTML As String, sQuestion As String, sAnswer As String
Dim sToEmail As String, sLine2 As String, sFileNameOut As String
Dim sBuf As String, sTemp As String, sFileName As String


Dim oOutlookApp As Outlook.Application, oOutlookMail As Outlook.MailItem, oWordEditor As Word.editor
Dim oWordApp As New Word.Application, oWordDoc As New Word.Document

Dim qdf As QueryDef
Dim rstQuery As DAO.Recordset
Dim iFileNum As Integer


Call fPPGenerateJSONInfo

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

sQueryName = "TRInvoiQPlusCases"
sExportedTemplatePath = "I:\" & sCourtDatesID & "\WorkingFiles\" & sCourtDatesID & "-PP-BalanceDueInvoice.docx"
sExportInfoCSVPath = "I:\" & sCourtDatesID & "\WorkingFiles\" & sCourtDatesID & "-InvoiceInfo.xls"
sTemplatePath = "T:\Database\Templates\Stage4s\PP-BalanceDueInvoiceEmail.docx"

Call pfGetOrderingAttorneyInfo
Call pfCurrentCaseInfo

Set qdf = CurrentDb.QueryDefs(sQueryName)
qdf.Parameters(0) = sCourtDatesID
Set rstQuery = qdf.OpenRecordset

sInvoiceNumber = rstQuery.Fields("TRInv.CourtDates.InvoiceNo").Value
sParty1 = rstQuery.Fields("Party1").Value
sParty2 = rstQuery.Fields("Party2").Value
sName = sFirstName & " " & sLastName
vPPInvoiceNo = rstQuery.Fields("TRInv.PPID").Value
vPPInvoiceNo = Right(vPPInvoiceNo, 20)
vPPInvoiceNo = Replace(Replace(vPPInvoiceNo, " ", ""), "-", "")

sFile1 = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-PPDepositEmail.html"
sInvoicePathPDF = "I:\" & sCourtDatesID & "\Generated\" & sInvoiceNumber & ".pdf"
sInvoicePathDocX = "I:\" & sCourtDatesID & "\WorkingFiles\" & sInvoiceNumber & ".docx"


'create pp invoice link
vPPLink = "https://www.paypal.com/invoice/p/#" & vPPInvoiceNo

'create PPButton html (construct string, save it into text file in job folder)
sHTMLPPB = "<html><body><br><br><a href =" & Chr(34) & vPPLink & Chr(34) & "><img src=" & Chr(34) & "https://www.paypalobjects.com/webstatic/en_US/i/buttons/checkout-logo-large.png" & _
        Chr(34) & "alt=" & Chr(34) & "Check out with PayPal" & Chr(34) & "/></a></body></html>"

sFilePath = "I:\" & sCourtDatesID & "\WorkingFiles\" & "PPButton.txt"
sFilePathHTML = "I:\" & sCourtDatesID & "\WorkingFiles\" & "PPButton.html"

'open txt file, save as html file

Open sFilePath For Output As #1
Write #1, sHTMLPPB
Close #1

iFileNum = FreeFile
Open sFilePath For Input As iFileNum

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
Open sFilePath For Output As iFileNum

Print #iFileNum, sTemp

Close iFileNum

Name sFilePath As sFilePathHTML

'paste PPButton html into mail merged invoice/PQ at both bookmarks "PPButton" bookmark or #PPB1# AND "PPButton2" or #PPB2#

Set oWordApp = CreateObject("Word.Application")
oWordApp.Visible = False

Set oWordDoc = oWordApp.Documents.Open(sFilePathHTML)

oWordDoc.Content.Copy

Set oWordApp = CreateObject("Word.Application")
oWordApp.Visible = True

Set oWordDoc = oWordApp.Documents.Open(sExportedTemplatePath)

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
   .ActiveDocument.SaveAs2 FileName:=sExportedTemplatePath
    
End With

    oWordDoc.Close
    oWordApp.Quit
    
    Set oWordApp = Nothing
    Set oWordDoc = Nothing
    

'prompt to ask when pdf is made
sQuestion = "Click yes after you have created your final invoice PDF at " & sInvoicePathPDF
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

If sAnswer = vbNo Then 'IF NO THEN THIS HAPPENS

    
    DoCmd.OutputTo acOutputQuery, sQueryName, acFormatXLS, sExportInfoCSVPath, False
    
    Set oWordApp = GetObject(sTemplatePath, "Word.Document")
    Set oWordDoc = oWordApp.Documents.Open(sTemplatePath)
    
    oWordApp.Application.Visible = False
    
    oWordDoc.MailMerge.OpenDataSource Name:=sExportInfoCSVPath, ReadOnly:=True
    oWordDoc.MailMerge.Execute
    oWordDoc.MailMerge.MainDocumentType = wdNotAMergeDocument
    oWordDoc.Application.ActiveDocument.SaveAs FileName:=sExportedTemplatePath, FileFormat:=wdFormatDocument
    
    oWordApp.Quit
    Set oWordApp = Nothing
    
    Set oWordApp = CreateObject("Word.Application")
    oWordApp.Visible = False
    
    Set oWordDoc = oWordApp.Documents.Open(sFilePathHTML)
    
    oWordDoc.Content.Copy
    oWordDoc.Close
    oWordApp.Quit
    
    Set oWordApp = Nothing
    Set oWordDoc = Nothing

    Set oWordApp = CreateObject("Word.Application")
    Set oWordDoc = oWordApp.Documents.Open(sExportedTemplatePath)


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
    oWordDoc.SaveAs2 FileName:=sExportedTemplatePath
    
End With

    oWordDoc.Close
    oWordApp.Quit
    
    Set oWordApp = Nothing
    Set oWordDoc = Nothing
    
    
    'generate PP-DraftInvoiceEmail in html format
    'oWordApp.Application.ActiveDocument.SaveAs filename:=sTemplatePath, FileFormat:=wdFormatHTML
    
    rstQuery.Close
    Set rstQuery = Nothing
    qdf.Close
    Set qdf = Nothing
    Set oOutlookApp = CreateObject("Outlook.Application")
    Set oOutlookMail = oOutlookApp.CreateItem(0)
    On Error Resume Next
    
    Set oWordApp = CreateObject("Word.Application")
    Set oWordDoc = oWordApp.Documents.Open(sExportedTemplatePath)
    oWordDoc.Content.Copy
    
    With oOutlookMail
        .To = sToEmail
        .CC = "inquiries@aquoco.co"
        .Subject = "Balance Due for " & sName & ", " & sParty1 & " v. " & sParty2
        .BodyFormat = olFormatRichText
        Set oWordEditor = .GetInspector.WordEditor
        .GetInspector.WordEditor.Content.Paste
        .Display
        .Attachments.Add (sInvoicePathPDF)
     End With
    oWordDoc.Close
    oWordApp.Quit
    Set oWordApp = Nothing
    Set oWordDoc = Nothing
    
    
Else 'if yes then this happens
    
    
    DoCmd.OutputTo acOutputQuery, sQueryName, acFormatXLS, sExportInfoCSVPath, False
    
    '
    Set oWordApp = CreateObject("Word.Application")
    Set oWordDoc = oWordApp.Documents.Open(sTemplatePath)
    oWordApp.Application.Visible = False
    
    oWordDoc.MailMerge.OpenDataSource Name:=sExportInfoCSVPath, ReadOnly:=True
    oWordDoc.MailMerge.Execute
    oWordDoc.MailMerge.MainDocumentType = wdNotAMergeDocument
    oWordDoc.Application.ActiveDocument.SaveAs2 FileName:=sExportedTemplatePath

    
    oWordDoc.Close
    Set oWordApp = Nothing
    Set oWordDoc = Nothing
    
    Set oWordApp = CreateObject("Word.Application")
    oWordApp.Visible = False
    
    Set oWordDoc = oWordApp.Documents.Open(sFilePathHTML)
    
    oWordDoc.Content.Copy
    oWordDoc.Close
    oWordApp.Quit
    
    Set oWordApp = Nothing
    Set oWordDoc = Nothing

    Set oWordApp = CreateObject("Word.Application")
    Set oWordDoc = oWordApp.Documents.Open(sExportedTemplatePath)


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
    .ActiveDocument.SaveAs2 FileName:=sExportedTemplatePath
    
    
End With
     oWordDoc.Close
     oWordApp.Quit
    
    Set oWordApp = Nothing
    Set oWordDoc = Nothing
    
    
    'generate PP-DraftInvoiceEmail in html format
    'oWordApp.Application.ActiveDocument.SaveAs filename:=sTemplatePath, FileFormat:=wdFormatHTML
    
    rstQuery.Close
    Set rstQuery = Nothing
    qdf.Close
    Set qdf = Nothing
    Set oOutlookApp = CreateObject("Outlook.Application")
    Set oOutlookMail = oOutlookApp.CreateItem(0)
    On Error Resume Next
    
    Set oWordApp = CreateObject("Word.Application")
    Set oWordDoc = oWordApp.Documents.Open(sExportedTemplatePath)
    oWordDoc.Content.Copy
    
    With oOutlookMail
        .To = sToEmail
        .CC = "inquiries@aquoco.co"
        .Subject = "Balance Due for " & sName & ", " & sParty1 & " v. " & sParty2
        .BodyFormat = olFormatRichText
        Set oWordEditor = .GetInspector.WordEditor
        .GetInspector.WordEditor.Content.Paste
        .Display
        .Attachments.Add (sInvoicePathPDF)
     End With
    oWordDoc.Close
    oWordApp.Quit
    Set oWordApp = Nothing
    Set oWordDoc = Nothing
    
End If
Call pfCommunicationHistoryAdd("PP Invoice Sent")

End Sub

Sub fSendPPEmailDeposit()
'============================================================================
' Name        : fPayPalUpdateCheck
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fPayPalUpdateCheck()
' Description : generates PayPal email for deposit
'============================================================================
'
Dim sFile1 As String, sInvoicePathPDF As String, sLine1 As String
Dim sInvoiceNumber As String, sName As String, vPPInvoiceNo As String
Dim sExportedTemplatePath As String, sTemplatePath As String, sOutputPDF As String
Dim sExportInfoCSVPath As String, sQueryName As String, sHTMLPPB As String
Dim sInvoicePathDocX As String, vPPLink As String, sFilePath As String
Dim sFilePathHTML As String, sQuestion As String, sAnswer As String
Dim sToEmail As String, sLine2 As String, sFileNameOut As String

Dim oOutlookApp As Outlook.Application, oOutlookMail As Outlook.MailItem, oWordEditor As Word.editor
Dim oWordApp As New Word.Application, oWordDoc As New Word.Document, oWordDoc1 As New Word.Document

Dim qdf As QueryDef
Dim rstQuery As DAO.Recordset
Dim iFileNumIn As Integer, iFileNumOut As Integer

'your invoice docx template MUST contain the phrase "#PPB1#" AND "#PPB2#" without the quotes somewhere on it.
'your e-mail docx template MUST contain the phrase "#PPB1#" without the quotes somewhere on it.

Call fPPGenerateJSONInfo 'refreshes some info, not relevant for purposes of this code being on GH

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

sQueryName = "TRInvoiQPlusCases"
sExportedTemplatePath = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-PP-DraftInvoiceEmail.docx"
sExportInfoCSVPath = "I:\" & sCourtDatesID & "\WorkingFiles\" & sCourtDatesID & "-InvoiceInfo.xls"
sTemplatePath = "T:\Database\Templates\Stage1s\PP-DraftInvoiceEmail-Template.docx"
sOutputPDF = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-" & "PP-DraftInvoiceEmail" & ".html"

Call pfGetOrderingAttorneyInfo 'refreshes some info, not relevant for purposes of this code being on GH





'paste PPButton html into mail merged invoice/PQ at both bookmarks "PPButton" bookmark or #PPB1# AND "PPButton2" or #PPB2#


Set qdf = CurrentDb.QueryDefs(sQueryName)

qdf.Parameters(0) = sCourtDatesID
Set rstQuery = qdf.OpenRecordset
sInvoiceNumber = rstQuery.Fields("TRInvoiceCasesQ.InvoiceNo").Value
sParty1 = rstQuery.Fields("Party1").Value
sParty2 = rstQuery.Fields("Party2").Value
sPPStatus = rstQuery.Fields("TRInv.PPStatus").Value
vPPInvoiceNo = rstQuery.Fields("TRInv.PPID").Value
vPPInvoiceNo = Right(vPPInvoiceNo, 20)
vPPInvoiceNo = Replace(Replace(vPPInvoiceNo, " ", ""), "-", "")
sName = sFirstName & " " & sLastName
sToEmail = sNotes

sFile1 = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-PPDepositEmail.html"
sInvoicePathPDF = "I:\" & sCourtDatesID & "\Generated\" & sInvoiceNumber & ".pdf"
sInvoicePathDocX = "I:\" & sCourtDatesID & "\Generated\" & sInvoiceNumber & ".docx"

'create pp invoice link
vPPLink = "https://www.paypal.com/invoice/p/#" & vPPInvoiceNo
vPPInvoiceNo = ""

'create PPButton html (construct string, save it into text file in job folder)
sHTMLPPB = "<html><body><br><br><a href =" & Chr(34) & vPPLink & Chr(34) & _
    "><img src=" & Chr(34) & "https://www.paypalobjects.com/webstatic/en_US/i/buttons/checkout-logo-large.png" & _
        Chr(34) & "alt=" & Chr(34) & "Check out with PayPal" & Chr(34) & "/></a></body></html>"


sFilePath = "I:\" & sCourtDatesID & "\WorkingFiles\" & "PPButton.txt"
sFilePathHTML = "I:\" & sCourtDatesID & "\WorkingFiles\" & "PPButton.html"


'open txt file, save as html file
Open sFilePath For Output As #1


Write #1, sHTMLPPB
Close #1

sFileNameOut = Replace(sHTMLPPB, Chr(34) & "<html>", "<html>")
iFileNumOut = FreeFile
Open sFilePath For Output As #1
Print #1, sFileNameOut
Close #1

sFileNameOut = Replace(sHTMLPPB, "</html>" & Chr(34), "</html>")
iFileNumOut = FreeFile
Open sFilePath For Output As #1
Print #1, sFileNameOut
Close #1

sFileNameOut = Replace(sHTMLPPB, Chr(34) & Chr(34), Chr(34))
iFileNumOut = FreeFile
Open sFilePath For Output As #1
Print #1, sFileNameOut
Close #1

Name sFilePath As sFilePathHTML

Set oWordApp = CreateObject("Word.Application")
oWordApp.Visible = False
        
        Set oWordDoc = oWordApp.Documents.Open(sFilePathHTML) 'open button in word
        oWordDoc.Content.Copy
        Set oWordApp = CreateObject("Word.Application")
        oWordApp.Visible = False
        
        Set oWordDoc1 = oWordApp.Documents.Open(sInvoicePathDocX) 'open invoice docx
        

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
    .ActiveDocument.SaveAs2 FileName:=sInvoicePathDocX
    
End With

    oWordDoc.Close
    oWordApp.Quit
    
    Set oWordApp = Nothing
    Set oWordDoc = Nothing
    

'prompt to ask when pdf of invoice is made
sQuestion = "Click yes after you have created your final invoice PDF at " & sInvoicePathPDF
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

If sAnswer = vbNo Then 'IF NO THEN THIS HAPPENS

    
    MsgBox "No invoice will be sent at this time."
    
    
Else 'if yes then this happens
    
    
    DoCmd.OutputTo acOutputQuery, sQueryName, acFormatXLS, sExportInfoCSVPath, False
    
    
    
    
    sTemplatePath = "T:\Database\Templates\Stage4s\PP-DraftInvoiceEmail-Template.docx"
    
    Set oWordApp = CreateObject("Word.Application")
    Set oWordDoc = oWordApp.Documents.Open(sTemplatePath)

    oWordApp.Application.Visible = False
    oWordDoc.MailMerge.OpenDataSource Name:=sExportInfoCSVPath, ReadOnly:=True
    oWordDoc.MailMerge.Execute
    oWordDoc.MailMerge.MainDocumentType = wdNotAMergeDocument
    oWordDoc.Application.ActiveDocument.SaveAs2 FileName:=sExportedTemplatePath

    
    
    Set oWordApp = CreateObject("Word.Application")
    oWordApp.Visible = False
    
    Set oWordDoc = oWordApp.Documents.Open(sFilePathHTML)
    
    oWordDoc.Content.Copy
    
    
    
    
    
    
    
    
    Set oWordApp = CreateObject("Word.Application")
    Set oWordDoc = oWordApp.Documents.Open(sExportedTemplatePath)


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
    .ActiveDocument.SaveAs2 FileName:=sExportedTemplatePath
    
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
    Set oWordDoc = oWordApp.Documents.Open(sExportedTemplatePath) 'invoice email file in docx format
    oWordDoc.Content.Copy
    
    With oOutlookMail 'now, you should have an e-mail with a PP button as well as an invoice with two PP buttons on it.
        .To = sToEmail
        .CC = "inquiries@aquoco.co"
        .Subject = "Deposit Invoice for " & sName & ", " & sParty1 & " v. " & sParty2
        .BodyFormat = olFormatRichText
        Set oWordEditor = .GetInspector.WordEditor
        .GetInspector.WordEditor.Content.Paste
        .Display
        .Attachments.Add (sInvoicePathPDF)
     End With
    
End If
    oWordDoc.Close
    oWordApp.Quit

Set oOutlookApp = Nothing
Set oOutlookMail = Nothing
Set oWordEditor = Nothing
Set oWordApp = Nothing
Set oWordDoc = Nothing

Call pfCommunicationHistoryAdd("PP Invoice Sent") 'record entry in comm history table for logs
Call pfClearGlobals
End Sub




Sub fPPGetInvoiceInfo()
'============================================================================
' Name        : fPPGetInvoiceInfo
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fPPGetInvoiceInfo
' Description : gets status of invoice
'============================================================================

Dim sURL As String, sUserName As String, sPassword As String, sAuth As String, stringJSON As String
Dim apiWaxLRS As String, vErrorIssue As String, sInvoiceTime As String
Dim oRequest As Object, Json As Object
Dim sInvoiceNo As String, sEmail As String, sFirstName As String, sLastName As String
Dim sDescription As String, sInvoiceDate As String, sPaymentTerms As String
Dim sCourtDatesID As String, sNote As String, sTerms As String, sMinimumAmount As String
Dim vmMemo As String, vlURL As String, sTemplateID As String, vTotal As String
Dim sCity As String, sState As String, sZIP As String, sQuantity As String
Dim sValue As String, vInvoiceID As String, sInvoiceNumber As String, vStatus As String
Dim resp As Object, response As Object, rep As Object
Dim sToken As String, json1 As String, json2 As String, json3 As String, json4 As String, json5 As String
Dim Parsed As Dictionary
Dim vErrorName As String, vErrorMessage As String, vErrorILink As String, vErrorDetails As String
Dim vTermDays As String, vDetails As String
Dim sFile1 As String, sFile2 As String, sText As String, sLine1 As String, sLine2 As String



Call fPPGenerateJSONInfo

'note: fPPGetInvoiceInfo can delete following lines when known safe come back
'sFile1 = "C:\other\1.txt"
'sFile2 = "C:\other\2.txt"

'Open sFile1 For Input As #1
'Line Input #1, sLine1
'Close #1

'Open sFile2 For Input As #2
'Line Input #2, sLine2
'Close #2
'sUserName = sLine1
'sPassword = sLine2

sURL = "https://api.paypal.com/v1/oauth2/token/"
sEmail = "inquiries@aquoco.co"
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
        Set Parsed = JsonConverter.ParseJson(apiWaxLRS)
        sToken = Parsed("access_token") 'third level array
        .abort
'Debug.Print "--------------------------------------------"
    End With
    'get info for invoice, call separate function for it maybe 'come back
 vInvoiceID = rstTRQPlusCases.Fields("TRInv.PPID").Value ' "INV2-C8EE-ZVC5-5U36-MF27" 'INV2-K8L5-ML2R-2GLL-7KW6
  
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
        .send ' json5
        apiWaxLRS = .responseText
        sToken = ""
        .abort
        Debug.Print apiWaxLRS
Debug.Print "--------------------------------------------"
    End With
Set Parsed = JsonConverter.ParseJson(apiWaxLRS)
sInvoiceNumber = Parsed("number") 'third level array
vInvoiceID = Parsed("id") 'third level array
vStatus = Parsed("status") 'third level array 'can be either DRAFT, UNPAID, SENT, SCHEDULED, PARTIALLY_PAID, PAYMENT_PENDING, PAID, MARKED_AS_PAID,
                                                            'CANCELLED, REFUNDED, PARTIALLY_REFUNDED, MARKED AS REFUNDED
vTotal = Parsed("total_amount")("value") 'second level array
vErrorName = Parsed("name") '("value") 'second level array
vErrorMessage = Parsed("message") '("value") 'second level array
vErrorILink = Parsed("information_link") '("value") 'second level array
vDetails = Parsed("details") 'second level array
'For Each rep In vDetails ' third level objects
vErrorIssue = Parsed("field")
vErrorDetails = Parsed("issue")
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




Sub fPPRefund()
'============================================================================
' Name        : fPPRefund
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fPPRefund()
' Description : refund with PayPal
'============================================================================
'
Dim vAmount As String, sQueryName As String, sExportedTemplatePath As String, sExportInfoCSVPath As String
Dim sTemplatePath As String
Dim qdf As QueryDef
Dim rstQuery As DAO.Recordset
Dim vPPInvoiceNo As String, sFile1 As String, sInvoicePathPDF As String, sInvoicePathDocX As String
Dim sURL As String, sUserName As String, sPassword As String, sAuth As String, stringJSON As String, sEmail As String
Dim apiWaxLRS As String, vErrorIssue As String, sInvoiceTime As String
Dim oRequest As Object, Json As Object
Dim sInvoiceNo As String, sFirstName As String, sLastName As String
Dim sDescription As String, sInvoiceDate As String, sPaymentTerms As String
Dim sCourtDatesID As String, sNote As String, sTerms As String, sMinimumAmount As String
Dim vmMemo As String, vlURL As String, sTemplateID As String, vTotal As String
Dim sLine1 As String, sCity As String, sState As String, sZIP As String, sQuantity As String
Dim sValue As String, vInvoiceID As String, sInvoiceNumber As String, vStatus As String
Dim resp As Object, response As Object, rep As Object
Dim sToken As String, json1 As String, json2 As String, json3 As String, json4 As String, json5 As String
Dim Parsed As Dictionary
Dim vErrorName As String, vErrorMessage As String, vErrorILink As String, vErrorDetails As String
Dim vTermDays As String, vDetails As String
Dim sFile2 As String, sText As String, sLine2 As String


Call fPPGenerateJSONInfo

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

sQueryName = "TRInvoiQPlusCases"
sExportedTemplatePath = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-PP-RefundMadeEmail.docx"
sExportInfoCSVPath = "I:\" & sCourtDatesID & "\WorkingFiles\" & sCourtDatesID & "-InvoiceInfo.xls"
sTemplatePath = "T:\Database\Templates\Stage4s\PP-RefundMadeEmail.docx"

Call pfGetOrderingAttorneyInfo
Call pfCurrentCaseInfo

Set qdf = CurrentDb.QueryDefs(sQueryName)
qdf.Parameters(0) = sCourtDatesID
Set rstQuery = qdf.OpenRecordset

sInvoiceNumber = rstQuery.Fields("TRInv.CourtDates.InvoiceNo").Value
'get pp invoice ID
vPPInvoiceNo = rstQuery.Fields("TRInv.PPID").Value
sPaymentSum = rstQuery.Fields("TRInv.PaymentSum").Value
sFinalPrice = rstQuery.Fields("TRInv.FinalPrice").Value
vPPInvoiceNo = Right(vPPInvoiceNo, 20)
vPPInvoiceNo = Replace(Replace(vPPInvoiceNo, " ", ""), "-", "")

sFile1 = "I:\" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-PP-RefundMadeEmail.html"
sInvoicePathPDF = "I:\" & sCourtDatesID & "\Generated\" & sInvoiceNumber & ".pdf"
sInvoicePathDocX = "I:\" & sCourtDatesID & "\Generated\" & sInvoiceNumber & ".docx"

'create pp invoice link
vPPLink = "https://www.paypal.com/invoice/p/#" & vPPInvoiceNo

'calculate refund
vAmount = InputBox("How much do you want to refund?  Their payments total up to " & sPaymentSum & " their final bill came to " & sFinalPrice & ".  This leaves a difference of " & (sPaymentSum - sFinalPrice) & " left owing to them.")
    
sQuestion = "Are you sure that's correct?  You want to refund $" & vAmount & "?"
sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")

If sAnswer = vbNo Then 'Code for No

    sQuestion = "Do you want to refund anything?"
    sAnswer = MsgBox(sQuestion, vbQuestion + vbYesNo, "???")
    
    If sAnswer = vbNo Then 'Code for No

        GoTo Exitif
    Else

        vAmount = InputBox("How much do you want to refund?  Their payments total up to " & sPaymentSum & " their final bill came to " & sFinalPrice & ".  This leaves a difference of " & (sPaymentSum - sFinalPrice) & " left owing to them.")
    
    End If
    
Else 'Code for yes
End If

'note: fPPRefund can delete following lines when known safe come back
'sFile1 = "C:\other\1.txt"
'sFile2 = "C:\other\2.txt"

'Open sFile1 For Input As #1
'Line Input #1, sLine1
'Close #1

'Open sFile2 For Input As #2
'Line Input #2, sLine2
'Close #2
'sUserName = sLine1
'sPassword = sLine2
'sLine1 = ""
'sLine2 = ""

sInvoiceDate = (Format((Date + 28), "yyyy-mm-dd")) & " PST"
sInvoiceTime = (Format(Now(), "hh:mm:ss"))
sURL = "https://api.paypal.com/v1/oauth2/token/"
sEmail = "inquiries@aquoco.co"
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
    Set Parsed = JsonConverter.ParseJson(apiWaxLRS)
    sToken = Parsed("access_token") 'third level array
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
Set Parsed = JsonConverter.ParseJson(apiWaxLRS)
sInvoiceNumber = Parsed("number") 'third level array
vInvoiceID = Parsed("id") 'third level array
vStatus = Parsed("status") 'third level array
'vTotal = Parsed("total_amount")("value") 'second level array
vErrorName = Parsed("name") '("value") 'second level array
vErrorMessage = Parsed("message") '("value") 'second level array
vErrorILink = Parsed("information_link") '("value") 'second level array
vDetails = Parsed("details") 'second level array
'For Each rep In vDetails ' third level objects
 '   vErrorIssue = rep("field")
  '  vErrorDetails = rep("issue")
'Next
Debug.Print "--------------------------------------------"
Debug.Print "Error Name:  " & vErrorName
Debug.Print "Error Message:  " & vErrorMessage
Debug.Print "Error Info Link:  " & vErrorILink
Debug.Print "Error Field:  " & vErrorIssue
Debug.Print "Error Details:  " & vErrorDetails
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



Function TextBase64Encode(sText, sCharset)
'============================================================================
' Name        : TextBase64Encode
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call TextBase64Encode(sText, sCharset)
' Description : encodes to base64
'============================================================================
'
    Dim aBinary

    With CreateObject("ADODB.Stream")
        .Type = 2 ' adTypeText
        .Open
        .Charset = sCharset
        .WriteText sText
        .Position = 0
        .Type = 1 ' adTypeBinary
        aBinary = .Read
        .Close
    End With
    With CreateObject("Microsoft.XMLDOM").createElement("objNode")
        .DataType = "bin.base64"
        .nodeTypedValue = aBinary
        TextBase64Encode = Replace(Replace(.Text, vbCr, ""), vbLf, "")
    End With

End Function


Sub fPPUpdate()
'============================================================================
' Name        : fPPUpdate
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fPPUpdate
' Description : updates PayPal invoice on PayPal website
'============================================================================

Dim sURL As String, sUserName As String, sPassword As String, sAuth As String, stringJSON As String, sEmail As String
Dim apiWaxLRS As String, vErrorIssue As String, sInvoiceTime As String
Dim oRequest As Object, Json As Object
Dim sInvoiceNo As String, sFirstName As String, sLastName As String
Dim sDescription As String, sInvoiceDate As String, sPaymentTerms As String
Dim sCourtDatesID As String, sNote As String, sTerms As String, sMinimumAmount As String
Dim vmMemo As String, vlURL As String, sTemplateID As String, vTotal As String
Dim sLine1 As String, sCity As String, sState As String, sZIP As String, sQuantity As String
Dim sValue As String, vInvoiceID As String, sInvoiceNumber As String, vStatus As String
Dim resp As Object, response As Object, rep As Object
Dim sToken As String, json1 As String, json2 As String, json3 As String, json4 As String, json5 As String
Dim Parsed As Dictionary
Dim vErrorName As String, vErrorMessage As String, vErrorILink As String, vErrorDetails As String
Dim vTermDays As String, vDetails As String
Dim sFile1 As String, sFile2 As String, sText As String, sLine2 As String


sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]

Call fPPGenerateJSONInfo
Call pfGetOrderingAttorneyInfo


'note: fPPUpdate can delete following lines when known safe come back
'sFile1 = "C:\other\1.txt"
'sFile2 = "C:\other\2.txt"

'Open sFile1 For Input As #1
'Line Input #1, sLine1
'Close #1

'Open sFile2 For Input As #2
'Line Input #2, sLine2
'Close #2

sURL = "https://api.paypal.com/v1/oauth2/token/"
sEmail = "inquiries@aquoco.co"
'sUserName = sLine1
'sPassword = sLine2
'sLine1 = ""
'sLine2 = ""
'sPassword = ""
'sUserName = ""
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
        Set Parsed = JsonConverter.ParseJson(apiWaxLRS)
        sToken = Parsed("access_token") 'third level array
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
    "country_code" & Chr(34) & ": " & Chr(34) & vPCountryCode & Chr(34) & "," & Chr(34) & _
    "national_number" & Chr(34) & ": " & Chr(34) & sCompanyNationalNumber & Chr(34) & "}," & Chr(34) & _
    "address" & Chr(34) & ": {" & Chr(34) & _
    "line1" & Chr(34) & ": " & Chr(34) & sCompanyAddress & Chr(34) & "," & Chr(34) & _
    "city" & Chr(34) & ": " & Chr(34) & sCompanyCity & Chr(34) & "," & Chr(34) & _
    "state" & Chr(34) & ": " & Chr(34) & sCompanyState & Chr(34) & "," & Chr(34) & _
    "postal_code" & Chr(34) & ": " & Chr(34) & sCompanyZIP & Chr(34) & "," & Chr(34) & _
    "country_code" & Chr(34) & ": " & Chr(34) & vZCountryCode & Chr(34) & "}},"
    
    json2 = Chr(34) & "billing_info" & Chr(34) & ": [{" & Chr(34) & _
    "email" & Chr(34) & ": " & Chr(34) & sEmail & Chr(34) & "}],"
    
    json3 = Chr(34) & "items" & Chr(34) & ": [" & _
    "{" & Chr(34) & _
    "name" & Chr(34) & ": " & Chr(34) & sDescription & Chr(34) & "," & Chr(34) & _
    "quantity" & Chr(34) & ": " & Chr(34) & sQuantity & Chr(34) & "," & Chr(34) & _
    "unit_price" & Chr(34) & ": {" & Chr(34) & _
    "currency" & Chr(34) & ": " & Chr(34) & "USD" & Chr(34) & "," & Chr(34) & _
    "value" & Chr(34) & ": " & Chr(34) & sUnitPrice & Chr(34) & "}}]," & Chr(34) & _
    "note" & Chr(34) & ": " & Chr(34) & sNote & Chr(34) & "," & Chr(34) & _
    "payment_term" & Chr(34) & ": {" & Chr(34) & _
    "term_type" & Chr(34) & ": " & Chr(34) & "NET_" & vTermDays & Chr(34) & "}," & Chr(34) & _
    "shipping_info" & Chr(34) & ": {" & Chr(34) & _
    "first_name" & Chr(34) & ": " & Chr(34) & sFirstName & Chr(34) & "," & Chr(34) & _
    "last_name" & Chr(34) & ": " & Chr(34) & sLastName & Chr(34) & "," & Chr(34) & _
    "business_name" & Chr(34) & ": " & Chr(34) & "WRTS Sample" & Chr(34) & "," & Chr(34) & _
    "address" & Chr(34) & ": {" & Chr(34) & _
    "line1" & Chr(34) & ": " & Chr(34) & "320 West Republican Street Suite 207" & Chr(34) & "," & Chr(34) & _
    "city" & Chr(34) & ": " & Chr(34) & "Seattle" & Chr(34) & "," & Chr(34) & _
    "state" & Chr(34) & ": " & Chr(34) & "WA" & Chr(34) & "," & Chr(34) & _
    "postal_code" & Chr(34) & ": " & Chr(34) & "98119" & Chr(34) & "," & Chr(34) & _
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
Set Parsed = JsonConverter.ParseJson(apiWaxLRS)
sInvoiceNumber = Parsed("number") 'third level array
vInvoiceID = Parsed("id") 'third level array
vStatus = Parsed("status") 'third level array
vTotal = Parsed("total_amount")("value") 'second level array
vErrorName = Parsed("name") '("value") 'second level array
vErrorMessage = Parsed("message") '("value") 'second level array
vErrorILink = Parsed("information_link") '("value") 'second level array
'vDetails = Parsed("details") 'second level array
'For Each rep In vDetails ' third level objects
'    vErrorIssue = rep("field")
'    vErrorDetails = rep("issue")
'Next
Debug.Print "--------------------------------------------"
Debug.Print "Error Name:  " & vErrorName
Debug.Print "Error Message:  " & vErrorMessage
Debug.Print "Error Info Link:  " & vErrorILink
Debug.Print "Error Field:  " & vErrorIssue
Debug.Print "Error Details:  " & vErrorDetails
Debug.Print "--------------------------------------------"
Debug.Print "Invoice No.:  " & sInvoiceNumber & "   |   Invoice ID:  " & vInvoiceID
Debug.Print "Status:  " & vStatus & "   |   Total:  " & vTotal
Debug.Print "--------------------------------------------"

'update PPID & PPStatus
Dim sUpdatePPStatus As String, sUpdatePPID As String


sUpdatePPStatus = "UPDATE CourtDates SET PPStatus = " & Chr(34) & vStatus & Chr(34) & " WHERE [ID] = " & sCourtDatesID & ";"


CurrentDb.Execute sUpdatePPStatus
sUpdatePPID = "UPDATE CourtDates SET PPID = " & Chr(34) & vInvoiceID & Chr(34) & " WHERE [ID] = " & sCourtDatesID & ";"
CurrentDb.Execute sUpdatePPID





End Sub


Sub fManualPPPayment()
'============================================================================
' Name        : fManualPPPayment
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call fManualPPPayment
' Description : marks invoice as paid with manual payment, like with check/cash
'============================================================================

'curl -v -X POST https://api.sandbox.paypal.com/v1/invoicing/invoices/INV2-T4UQ-VW4W-K7N7-XM2R/record-payment \
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
Dim sURL As String, sUserName As String, sPassword As String, sAuth As String, stringJSON As String, sEmail As String
Dim apiWaxLRS As String, vErrorIssue As String, sInvoiceTime As String
Dim oRequest As Object, Json As Object
Dim sInvoiceNo As String, sFirstName As String, sLastName As String
Dim sDescription As String, sInvoiceDate As String, sPaymentTerms As String
Dim sCourtDatesID As String, sNote As String, sTerms As String, sMinimumAmount As String
Dim vmMemo As String, vlURL As String, sTemplateID As String, vTotal As String
Dim sLine1 As String, sCity As String, sState As String, sZIP As String, sQuantity As String
Dim sValue As String, vInvoiceID As String, sInvoiceNumber As String, vStatus As String
Dim resp As Object, response As Object, rep As Object
Dim sToken As String, json1 As String, json2 As String, json3 As String, json4 As String, json5 As String
Dim Parsed As Dictionary
Dim vErrorName As String, vErrorMessage As String, vErrorILink As String, vErrorDetails As String
Dim vTermDays As String, vDetails As String
Dim sFile1 As String, sFile2 As String, sText As String, sLine2 As String



Call fPPGenerateJSONInfo
Call pfGetOrderingAttorneyInfo


Set db = CurrentDb
Set qdf = db.QueryDefs("QInfobyInvoiceNumber")
qdf.Parameters(0) = sCourtDatesID
Set rstQInfoInvNo = qdf.OpenRecordset

sInvoiceNumber = rstQInfoInvNo.Fields("InvoiceNo").Value
sFinalPrice = rstQInfoInvNo.Fields("FinalPrice").Value

'note: fManualPPPayment can delete following lines when known safe come back
'sFile1 = "C:\other\1.txt"
'sFile2 = "C:\other\2.txt"

'Open sFile1 For Input As #1
'Line Input #1, sLine1
'Close #1

'Open sFile2 For Input As #2
'Line Input #2, sLine2
'Close #2

sURL = "https://api.paypal.com/v1/oauth2/token/"
sEmail = "inquiries@aquoco.co"
'sUserName = sLine1
'sPassword = sLine2
'sLine1 = ""
'sLine2 = ""
'sPassword = ""
'sUserName = ""
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
        Set Parsed = JsonConverter.ParseJson(apiWaxLRS)
        sToken = Parsed("access_token") 'third level array
        sAuth = ""
        .abort
Debug.Print "--------------------------------------------"
    End With
    vMethod = InputBox("What method was used to pay?  Select/type in either BANK_TRANSFER, CASH, CHECK, CREDIT_CARD, DEBIT_CARD, PAYPAL, WIRE_TRANSFER, or OTHER.")
    sInvoiceDate = (Format((Date + 28), "yyyy-mm-dd")) & " PST"
    sInvoiceTime = (Format(Now(), "hh:mm:ss"))
    vAmount = InputBox("How much was the payment?  Their invoice totals up to $" & sFinalPrice & ".")

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
        json5 = json1 & json2 & json3
        .send json5
        apiWaxLRS = .responseText
        sToken = ""
        .abort
        Debug.Print apiWaxLRS
Debug.Print "--------------------------------------------"
    End With
Set Parsed = JsonConverter.ParseJson(apiWaxLRS)
sInvoiceNumber = Parsed("number") 'third level array
vInvoiceID = Parsed("id") 'third level array
vStatus = Parsed("status") 'third level array
vTotal = Parsed("total_amount")("value") 'second level array
vErrorName = Parsed("name") '("value") 'second level array
vErrorMessage = Parsed("message") '("value") 'second level array
vErrorILink = Parsed("information_link") '("value") 'second level array
'vDetails = Parsed("details") 'second level array
'For Each rep In vDetails ' third level objects
'    vErrorIssue = rep("field")
'    vErrorDetails = rep("issue")
'Next
Debug.Print "--------------------------------------------"
Debug.Print "Error Name:  " & vErrorName
Debug.Print "Error Message:  " & vErrorMessage
Debug.Print "Error Info Link:  " & vErrorILink
Debug.Print "Error Field:  " & vErrorIssue
Debug.Print "Error Details:  " & vErrorDetails
Debug.Print "--------------------------------------------"
Debug.Print "Invoice No.:  " & sInvoiceNumber & "   |   Invoice ID:  " & vInvoiceID
Debug.Print "Status:  " & vStatus & "   |   Total:  " & vTotal
Debug.Print "--------------------------------------------"

'update PPID & PPStatus
Dim sUpdatePPStatus As String, sUpdatePPID As String
sUpdatePPStatus = "UPDATE CourtDates SET PPStatus = " & Chr(34) & vStatus & Chr(34) & " WHERE [ID] = " & sCourtDatesID & ";"
CurrentDb.Execute sUpdatePPStatus
sUpdatePPID = "UPDATE CourtDates SET PPID = " & Chr(34) & vInvoiceID & Chr(34) & " WHERE [ID] = " & sCourtDatesID & ";"
CurrentDb.Execute sUpdatePPID


End Sub




