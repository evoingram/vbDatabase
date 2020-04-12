VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocPaths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Database.Production.Classes")
Option Compare Database



Private pCourtCover As String
Private pRoughDraft As String
Private pCaseInfo As String
Private pInvoicePDF As Boolean
Private sApp0 As Long
Public sDocPaths As Long
Private sDrive As String

Private mrstRecordset0 As Recordset
Private mbooLoaded As Boolean


Public Property Get InvoiceP() As String
    InvoiceP = pInvoicePDF
End Property

Public Property Let InvoiceP(rData As String)
    pInvoicePDF = rData
End Property

Public Property Get InvoiceD() As String
    InvoiceD = pInvoiceDocx
End Property

Public Property Let InvoiceD(rData As String)
    pInvoiceDocx = rData
End Property

Public Property Get ShippingXML() As String
    ShippingXML = pShippingXML
End Property

Public Property Let ShippingXML(rData As String)
    pShippingXML = rData
End Property

Public Property Get ShippingCOAXML() As String
    ShippingCOAXML = pShippingCOAXML
End Property

Public Property Let ShippingCOAXML(rData As String)
    pShippingCOAXML = rData
End Property

Public Property Get RoughDraft() As String
    RoughDraft = pRoughDraft
End Property

Public Property Let RoughDraft(rData As String)
    pRoughDraft = rData
End Property

Public Property Get CaseInfo() As String
    CaseInfo = pCaseInfo
End Property

Public Property Let CaseInfo(rData As String)
    pCaseInfo = rData
End Property

Public Property Get InvoiceInfo() As String
    InvoiceInfo = pInvoiceInfo
End Property

Public Property Let InvoiceInfo(rData As String)
    pInvoiceInfo = rData
End Property


Public Property Get XeroCSV() As String
    XeroCSV = pXeroCSV
End Property

Public Property Let XeroCSV(rData As String)
    pXeroCSV = rData
End Property

Public Property Get PPCSV() As String
    PPCSV = pPPCSV
End Property

Public Property Let PPCSV(rData As String)
    pPPCSV = rData
End Property

Public Property Get PPButton() As String
    PPButton = pPPButton
End Property

Public Property Let PPButton(rData As String)
    pPPButton = rData
End Property

Public Property Get T2upPS() As String
    T2upPS = pT2upPS
End Property

Public Property Let T2upPS(rData As String)
    pT2upPS = rData
End Property

Public Property Get T4upPS() As String
    T4upPS = pT4upPS
End Property

Public Property Let T4upPS(rData As String)
    pT4upPS = rData
End Property

Public Property Get T4upLog() As String
    T4upLog = pT4upLog
End Property

Public Property Let T4upLog(rData As String)
    pT4upLog = rData
End Property

Public Property Get T2upLog() As String
    T2upLog = pT2upLog
End Property

Public Property Let T2upLog(rData As String)
    pT2upLog = rData
End Property

Public Property Get Notes() As String
    Notes = pNotes
End Property

Public Property Let Notes(rData As String)
    pNotes = rData
End Property

Public Property Get CourtCover() As String
    CourtCover = pCourtCover
End Property

Public Property Let CourtCover(rData As String)
    pCourtCover = rData
End Property

Public Property Get CIDIncomeReportD() As String
    CIDIncomeReportD = pCIDIncomeReportD
End Property

Public Property Let CIDIncomeReportD(rData As String)
    pCIDIncomeReportD = rData
End Property

Public Property Get CIDIncomeReportP() As String
    CIDIncomeReportP = pCIDIncomeReportP
End Property

Public Property Let CIDIncomeReportP(rData As String)
    pCIDIncomeReportP = rData
End Property

Public Property Get DepositInvoiceP() As String
    DepositInvoiceP = pDepositInvoiceP
End Property

Public Property Let DepositInvoiceP(rData As String)
    pDepositInvoiceP = rData
End Property

Public Property Get DepositInvoiceD() As String
    DepositInvoiceD = pDepositInvoiceD
End Property

Public Property Let DepositInvoiceD(rData As String)
    pDepositInvoiceD = rData
End Property

Public Property Get OrderConfirmationP() As String
    OrderConfirmationP = pOrderConfirmationP
End Property

Public Property Let OrderConfirmationP(rData As String)
    pOrderConfirmationP = rData
End Property

Public Property Get OrderConfirmationD() As String
    OrderConfirmationD = pOrderConfirmationD
End Property

Public Property Let OrderConfirmationD(rData As String)
    pOrderConfirmationD = rData
End Property

Public Property Get PPDraftInvoiceEmail() As String
    PPDraftInvoiceEmail = pPPDraftInvoiceEmail
End Property

Public Property Let PPDraftInvoiceEmail(rData As String)
    pPPDraftInvoiceEmail = rData
End Property

Public Property Get TranscriptsReadyP() As String
    TranscriptsReadyP = pTranscriptsReadyP
End Property

Public Property Let TranscriptsReadyP(rData As String)
    pTranscriptsReadyP = rData
End Property

Public Property Get TranscriptsReadyD() As String
    TranscriptsReadyD = pTranscriptsReadyD
End Property

Public Property Let TranscriptsReadyD(rData As String)
    pTranscriptsReadyD = rData
End Property

Public Property Get CDLabel() As String
    CDLabel = pCDLabel
End Property

Public Property Let CDLabel(rData As String)
    pCDLabel = rData
End Property

Public Property Get TranscriptsReadyCP() As String
    TranscriptsReadyCP = pTranscriptsReadyCP
End Property

Public Property Let TranscriptsReadyCP(rData As String)
    pTranscriptsReadyCP = rData
End Property

Public Property Get TranscriptsReadyCD() As String
    TranscriptsReadyCD = pTranscriptsReadyCD
End Property

Public Property Let TranscriptsReadyCD(rData As String)
    pTranscriptsReadyCD = rData
End Property

Public Property Get KCIEmpty() As String
    KCIEmpty = pKCIEmpty
End Property

Public Property Let KCIEmpty(rData As String)
    pKCIEmpty = rData
End Property

Public Property Get KCIFilled() As String
    KCIFilled = pKCIFilled
End Property

Public Property Let KCIFilled(rData As String)
    pKCIFilled = rData
End Property

Public Property Get WACoverP() As String
    WACoverP = pWACoverP
End Property

Public Property Let WACoverP(rData As String)
    pWACoverP = rData
End Property

Public Property Get WACoverD() As String
    WACoverD = pWACoverD
End Property

Public Property Let WACoverD(rData As String)
    pWACoverD = rData
End Property

Public Property Get WAConsolidatedP() As String
    WAConsolidatedP = pWAConsolidatedP
End Property

Public Property Let WAConsolidatedP(rData As String)
    pWAConsolidatedPB = rData
End Property
Private Sub Load()
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
        sDocPaths = ""
    'paths so far: transcribing manuals, samples, jurisdiction references, courtrules
        'shipping xml templates maybe
        
        'TODO: Refresh invoice number on this function
        Me.InvoiceP = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + sInvoiceNumber + ".pdf"
        Me.InvoiceD = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + sInvoiceNumber + ".docx"
        'TODO: Check COA shipping xml file name
        Me.ShippingXML = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + "-Shipping.xml"
        Me.ShippingCOAXML = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + "-COA-Shipping.xml"
        
        
        Me.RoughDraft = sDrive + ":/" + sCourtDatesID + "/RoughDraft.docx"
        Me.CaseInfo = sDrive + ":/" + sCourtDatesID + "/WorkingCopies/" + sCourtDatesID + "-CaseInfo.xls"
        Me.InvoiceInfo = sDrive + ":/" + sCourtDatesID + "/WorkingCopies/" + sCourtDatesID + "-InvoiceInfo.xls"
        Me.XeroCSV = sDrive + ":/" + sCourtDatesID + "/WorkingCopies/" + sCourtDatesID + "--XeroInvoiceCSV.csv"
        Me.PPCSV = sDrive + ":/" + sCourtDatesID + "/WorkingCopies/" + sCourtDatesID + "-PayPalInvoiceCSV.csv"
        Me.PPButton = sDrive + ":/" + sCourtDatesID + "/WorkingCopies/" + sCourtDatesID + "-PPButton.html"
        Me.T2upPS = sDrive + ":/" + sCourtDatesID + "/WorkingCopies/" + sCourtDatesID + "-Transcript-FINAL-2up.ps"
        Me.T4upPS = sDrive + ":/" + sCourtDatesID + "/WorkingCopies/" + sCourtDatesID + "-Transcript-FINAL-4up.ps"
        Me.T2upLog = sDrive + ":/" + sCourtDatesID + "/WorkingCopies/" + sCourtDatesID + "-Transcript-FINAL-2up.log"
        Me.T4upLog = sDrive + ":/" + sCourtDatesID + "/WorkingCopies/" + sCourtDatesID + "-Transcript-FINAL-4up.log"
        
        Me.Notes = sDrive + ":/" + sCourtDatesID + "/Notes/" + sCourtDatesID + "-Notes.pdf"
        
        Me.CourtCover = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + "-CourtCover.docx"
        Me.CIDIncomeReportD = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + "-CIDIncomeReport.docx"
        Me.CIDIncomeReportP = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + "-CIDIncomeReport.pdf"
        Me.DepositInvoiceP = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + "-DepositInvoice.pdf"
        Me.DepositInvoiceD = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + "-DepositInvoice.docx"
        Me.OrderConfirmationP = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + "-OrderConfirmation.pdf"
        Me.OrderConfirmationD = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + "-OrderConfirmation.docx"
        Me.PPDraftInvoiceEmail = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + "-PP-DraftInvoiceEmail.docx"
        Me.TranscriptsReadyP = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + "-TranscriptsReady.pdf"
        Me.TranscriptsReadyD = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + "-TranscriptsReady.docx"
        Me.CDLabel = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + "-CDLabel.pub"
        Me.TranscriptsReadyCP = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + "-ContractorTranscriptsReady.pdf"
        Me.TranscriptsReadyCD = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + "-ContractorTranscriptsReady.docx"
        Me.KCIEmpty = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + "-KCICompleted.pdf"
        Me.KCIFilled = sDrive + ":/" + sCourtDatesID + "/Generated/" + sCourtDatesID + "-KCICompleted1.pdf"
        
        Me.WACoverP = sDrive + ":/" + sCourtDatesID + "/Transcripts/" + sCourtDatesID + "-Cover.pdf"
        Me.WACoverD = sDrive + ":/" + sCourtDatesID + "/Transcripts/" + sCourtDatesID + "-Cover.docx"
        Me.WAConsolidatedP = sDrive + ":/" + sCourtDatesID + "/Transcripts/" + sCourtDatesID + "-Transcripts-All.pdf"
        Me.WAConsolidatedP = sDrive + ":/" + sCourtDatesID + "/Transcripts/" + sCourtDatesID + "-Transcripts-All.docx"
        Me.TranscriptFP = sDrive + ":/" + sCourtDatesID + "/Transcripts/" + sCourtDatesID + "-Transcript-FINAL.pdf"
        Me.TranscriptFD = sDrive + ":/" + sCourtDatesID + "/Transcripts/" + sCourtDatesID + "-Transcript-FINAL.docx"
        Me.Transcript2up = sDrive + ":/" + sCourtDatesID + "/Transcripts/" + sCourtDatesID + "-Transcript-FINAL-2up.pdf"
        Me.Transcript4up = sDrive + ":/" + sCourtDatesID + "/Transcripts/" + sCourtDatesID + "-Transcript-FINAL-4up.pdf"
        Me.TranscriptWC = sDrive + ":/" + sCourtDatesID + "/Transcripts/" + sCourtDatesID + "-Transcript-WorkingCopy.docx"
        Me.WordIndexP = sDrive + ":/" + sCourtDatesID + "/Transcripts/" + sCourtDatesID + "-WordIndex.pdf"
        Me.WordIndexD = sDrive + ":/" + sCourtDatesID + "/Transcripts/" + sCourtDatesID + "-WordIndex.docx"
        
        Me.AudioFolder = sDrive + ":/" + sCourtDatesID + "/Audio/"
        
        Me.KCIEnvelope = "T:\Database\Templates\Stage4s\Envelope-KCI.pdf"
        Me.PQTemplate = "T:\Database\Templates\Stage1s\PriceQuoteEmail-Template.docx"
        Me.PQXLS = "T:\Database\Scripts\InProgressExcels\Temp-Export-PQE.xlsx"
        Me.PQEmail = "T:\Database\Templates\Stage1s\PriceQuoteEmail.docx"
        Me.OrderFormCustomersXLSM = "T:\Database\Scripts\InProgressExcels\JotformCustomers.xlsm"
        Me.OrderFormCustomersXLSM = "T:\Database\Scripts\InProgressExcels\JotformCustomers.xlsx"
        Me.OrderFormCustomersCSV = "T:\Database\Scripts\InProgressExcels\JotformCustomers.csv"
        Me.OrderFormXLSM = "T:\Database\Scripts\InProgressExcels\Jotform.xlsm"
        Me.OrderFormXLSX = "T:\Database\Scripts\InProgressExcels\Jotform.xlsx"
        Me.OrderFormCSV = "T:\Database\Scripts\InProgressExcels\Jotform.csv"
        
        Me.ZAudioF = sDrive + ":/" + sCourtDatesID + "/FTP/" + sCourtDatesID + "-Audio.zip"
        Me.ZTranscriptsF = sDrive + ":/" + sCourtDatesID + "/FTP/" + sCourtDatesID + "-Transcripts.zip"
        Me.ZAudioTranscriptsF = sDrive + ":/" + sCourtDatesID + "/FTP/" + sCourtDatesID + "-AudioTranscripts.zip"
        
        Me.ZAudioB = sDrive + ":/" + sCourtDatesID + "/Backups/" + sCourtDatesID + "-Audio.zip"
        Me.ZTranscriptsB = sDrive + ":/" + sCourtDatesID + "/Backups/" + sCourtDatesID + "-Transcripts.zip"
        Me.ZAudioTranscriptsB = sDrive + ":/" + sCourtDatesID + "/Backups/" + sCourtDatesID + "-AudioTranscripts.zip"
        Me.TranscriptFPB = sDrive + ":/" + sCourtDatesID + "/Backups/" + sCourtDatesID + "-Transcript-FINAL.pdf"
        Me.TranscriptFDB = sDrive + ":/" + sCourtDatesID + "/Backups/" + sCourtDatesID + "-Transcript-FINAL.docx"
        Me.Transcript2upB = sDrive + ":/" + sCourtDatesID + "/Backups/" + sCourtDatesID + "-Transcript-FINAL-2up.pdf"
        Me.Transcript4upB = sDrive + ":/" + sCourtDatesID + "/Backups/" + sCourtDatesID + "-Transcript-FINAL-4up.pdf"
        Me.TranscriptWCB = sDrive + ":/" + sCourtDatesID + "/Backups/" + sCourtDatesID + "-Transcript-WorkingCopy.docx"
        Me.WordIndexPB = sDrive + ":/" + sCourtDatesID + "/Backups/" + sCourtDatesID + "-WordIndex.pdf"
        Me.WordIndexDB = sDrive + ":/" + sCourtDatesID + "/Backups/" + sCourtDatesID + "-WordIndex.docx"

    mbooLoaded = True
End Property

Public Sub Update()
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    With RecordsetA0
        If mbooLoaded = True Then
            .Edit
        Else
            .AddNew
        End If
        sApp0 = Nz(.Fields("ID").Value)
        .Fields("Company").Value = Me.Company
        .Fields("MrMs").Value = Me.MrMs
        .Update
    End With
    mbooLoaded = True
End Sub

Public Sub AddNew()
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    mbooLoaded = False
End Sub

'@Ignore FunctionReturnValueNotUsed
Public Function FindFirst(Optional Criteria As Variant) As Boolean
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    If IsMissing(Criteria) Then
        RecordsetA0.MoveFirst
        FindFirst = Not RecordsetA0.EOF
    Else
        RecordsetA0.FindFirst Criteria
        FindFirst = Not RecordsetA0.NoMatch
    End If
    If FindFirst Then Load
End Function

Private Sub Class_Initialize()
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    Set RecordsetA0 = CurrentDb.OpenRecordset("Customers", dbOpenDynaset)
    
End Sub

Private Sub Class_Terminate()
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    RecordsetA0.Close
    Set RecordsetA0 = Nothing
End Sub

Private Function NullIfEmptyString(str As String) As Variant
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    Dim strTrimmed As String: strTrimmed = Trim(str)
    If Len(strTrimmed) = 0 Then
        NullIfEmptyString = Null
    Else
        NullIfEmptyString = strTrimmed
    End If
End Function
