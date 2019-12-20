VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Database.Production.Classes")
Option Compare Database

Dim pRoughDraft As String
Dim psDrive As String
Dim pNetWorkFolder As String

Private pCourtCover As String
Private pCaseInfo As String
Private pInvoiceP As String
Private pInvoiceD As String
Private pShippingXML As String
Private pShippingCOAXML As String
Private pInvoiceInfo As String
Private pXeroCSV As String
Private pPPCSV As String
Private pPPButton As String
Private pT2upPS As String
Private pT4upPS As String
Private pT4upLog As String
Private pT2upLog As String
Private pNotes As String
Private pCIDIncomeReportD As String
Private pCIDIncomeReportP As String
Private pDepositInvoiceP As String
Private pDepositInvoiceD As String
Private pOrderConfirmationP As String
Private pOrderConfirmationD As String
Private pPPDraftInvoiceEmail As String
Private pTranscriptsReadyP As String
Private pTranscriptsReadyD As String
Private pCDLabel As String
Private pTranscriptsReadyCP As String
Private pTranscriptsReadyCD As String
Private pKCIEmpty As String
Private pKCIFilled As String
Private pWACoverP As String
Private pWACoverD As String
Private pWAConsolidatedP As String
Private pWAConsolidatedD As String
Private pTranscriptFP As String
Private pTranscriptFD As String
Private pTranscript2up As String
Private pTranscript4up As String
Private pTranscriptWC As String
Private pWordIndexP As String
Private pWordIndexD As String
Private pAudioFolder As String
Private pZAudioF As String
Private pZTranscriptsF As String
Private pZAudioTranscriptsF As String
Private pZAudioB As String
Private pZTranscriptsB As String
Private pZAudioTranscriptsB As String
Private pTranscriptFPB As String
Private pTranscriptFDB As String
Private pTranscript2upB As String
Private pTranscript4upB As String
Private pTranscriptWCB As String
Private pWordIndexPB As String
Private pWordIndexDB As String
Private pKCIEnvelope As String
Private pPQTemplate As String
Private pPQXLS As String
Private pPQEmail As String
Private pOrderFormCustomersXLSM As String
Private pOrderFormCustomersXLSX As String
Private pOrderFormCustomersCSV As String
Private pOrderFormXLSM As String
Private pOrderFormXLSX As String
Private pOrderFormCSV As String

Public Property Get InvoiceP() As String
    InvoiceP = pInvoiceP
End Property

Public Property Get InvoiceD() As String
    InvoiceD = pInvoiceDocx
End Property

Public Property Get ShippingXML() As String
    ShippingXML = pShippingXML
End Property
Public Property Get ShippingCOAXML() As String
    ShippingCOAXML = pShippingCOAXML
End Property


Public Property Get ShippingXML4() As String
    ShippingXML4 = pShippingXML4
End Property
Public Property Get ShippingCOAXML4() As String
    ShippingCOAXML4 = pShippingCOAXML4
End Property



Public Property Get CaseInfo() As String
    CaseInfo = pCaseInfo
End Property

Public Property Get InvoiceInfo() As String
    InvoiceInfo = pInvoiceInfo
End Property

Public Property Get XeroCSV() As String
    XeroCSV = pXeroCSV
End Property

Public Property Get PPCSV() As String
    PPCSV = pPPCSV
End Property

Public Property Get PPButton() As String
    PPButton = pPPButton
End Property

Public Property Get T2upPS() As String
    T2upPS = pT2upPS
End Property

Public Property Get T4upPS() As String
    T4upPS = pT4upPS
End Property

Public Property Get T4upLog() As String
    T4upLog = pT4upLog
End Property

Public Property Get T2upLog() As String
    T2upLog = pT2upLog
End Property

Public Property Get Notes() As String
    Notes = pNotes
End Property

Public Property Get CourtCover() As String
    CourtCover = pCourtCover
End Property

Public Property Get CIDIncomeReportD() As String
    CIDIncomeReportD = pCIDIncomeReportD
End Property

Public Property Get CIDIncomeReportP() As String
    CIDIncomeReportP = pCIDIncomeReportP
End Property

Public Property Get DepositInvoiceP() As String
    DepositInvoiceP = pDepositInvoiceP
End Property

Public Property Get DepositInvoiceD() As String
    DepositInvoiceD = pDepositInvoiceD
End Property

Public Property Get OrderConfirmationP() As String
    OrderConfirmationP = pOrderConfirmationP
End Property

Public Property Get OrderConfirmationD() As String
    OrderConfirmationD = pOrderConfirmationD
End Property

Public Property Get PPDraftInvoiceEmail() As String
    PPDraftInvoiceEmail = pPPDraftInvoiceEmail
End Property

Public Property Get TranscriptsReadyP() As String
    TranscriptsReadyP = pTranscriptsReadyP
End Property

Public Property Get TranscriptsReadyD() As String
    TranscriptsReadyD = pTranscriptsReadyD
End Property

Public Property Get CDLabel() As String
    CDLabel = pCDLabel
End Property

Public Property Get TranscriptsReadyCP() As String
    TranscriptsReadyCP = pTranscriptsReadyCP
End Property

Public Property Get TranscriptsReadyCD() As String
    TranscriptsReadyCD = pTranscriptsReadyCD
End Property

Public Property Get KCIEmpty() As String
    KCIEmpty = pKCIEmpty
End Property

Public Property Get KCIFilled() As String
    KCIFilled = pKCIFilled
End Property

Public Property Get WACoverP() As String
    WACoverP = pWACoverP
End Property

Public Property Get WACoverD() As String
    WACoverD = pWACoverD
End Property

Public Property Get WAConsolidatedP() As String
    WAConsolidatedP = pWAConsolidatedP
End Property

Public Property Get WAConsolidatedD() As String
    WAConsolidatedD = pWAConsolidatedD
End Property

Public Property Get TranscriptFP() As String
    TranscriptFP = pTranscriptFP
End Property

Public Property Get TranscriptFD() As String
    TranscriptFD = pTranscriptFD
End Property

Public Property Get Transcript2up() As String
    Transcript2up = pTranscript2up
End Property

Public Property Get Transcript4up() As String
    Transcript4up = pTranscript4up
End Property

Public Property Get TranscriptWC() As String
    TranscriptWC = pTranscriptWC
End Property

Public Property Get WordIndexP() As String
    WordIndexP = pWordIndexP
End Property

Public Property Get WordIndexD() As String
    WordIndexD = pWordIndexD
End Property

Public Property Get AudioFolder() As String
    AudioFolder = pAudioFolder
End Property

Public Property Get KCIEnvelope() As String
    KCIEnvelope = pKCIEnvelope
End Property

Public Property Get PQTemplate() As String
    PQTemplate = pPQTemplate
End Property

Public Property Get PQXLS() As String
    PQXLS = pPQXLS
End Property

Public Property Get PQEmail() As String
    PQEmail = pPQEmail
End Property

Public Property Get OrderFormCustomersXLSM() As String
    OrderFormCustomersXLSM = pOrderFormCustomersXLSM
End Property

Public Property Get OrderFormCustomersXLSX() As String
    OrderFormCustomersXLSX = pOrderFormCustomersXLSX
End Property

Public Property Get OrderFormCustomersCSV() As String
    OrderFormCustomersCSV = pOrderFormCustomersCSV
End Property

Public Property Get OrderFormXLSM() As String
    OrderFormXLSM = pOrderFormXLSM
End Property

Public Property Get OrderFormXLSX() As String
    OrderFormXLSX = pOrderFormXLSX
End Property

Public Property Get OrderFormCSV() As String
    OrderFormCSV = pOrderFormCSV
End Property

Public Property Get ZAudioF() As String
    ZAudioF = pZAudioF
End Property

Public Property Get ZTranscriptsF() As String
    ZTranscriptsF = pZTranscriptsF
End Property

Public Property Get ZAudioTranscriptsF() As String
    ZAudioTranscriptsF = pZAudioTranscriptsF
End Property

Public Property Get ZAudioB() As String
    ZAudioB = pZAudioB
End Property
Public Property Get ZTranscriptsB() As String
    ZTranscriptsB = pZTranscriptsB
End Property

Public Property Get ZAudioTranscriptsB() As String
    ZAudioTranscriptsB = pZAudioTranscriptsB
End Property

Public Property Get TranscriptFPB() As String
    TranscriptFPB = pTranscriptFPB
End Property

Public Property Get TranscriptFDB() As String
    TranscriptFDB = pTranscriptFDB
End Property

Public Property Get Transcript2upB() As String
    Transcript2upB = pTranscript2upB
End Property

Public Property Get Transcript4upB() As String
    Transcript4upB = pTranscript4upB
End Property

Public Property Get TranscriptWCB() As String
    TranscriptWCB = pTranscriptWCB
End Property

Public Property Get WordIndexPB() As String
    WordIndexPB = pWordIndexPB
End Property

Public Property Get RoughDraft() As String
    RoughDraft = pRoughDraft
End Property

Public Property Let RoughDraft(sInput As String)
    pRoughDraft = sInput
End Property

Public Property Get sDrive() As String
    sDrive = psDrive
End Property

Public Property Let sDrive(sInput As String)
    psDrive = sInput
End Property

Public Property Get JobDirectory() As String
    JobDirectory = pJobDirectory
End Property

Public Property Let JobDirectory(sInput As String)
    pJobDirectory = JobDirectory
End Property

Public Property Get JobDirectoryG() As String
    JobDirectoryG = pJobDirectoryG
End Property

Public Property Let JobDirectoryG(sInput As String)
    pJobDirectoryG = JobDirectoryG
End Property

Public Property Get JobDirectoryW() As String
    JobDirectoryW = pJobDirectoryW
End Property

Public Property Let JobDirectoryW(sInput As String)
    pJobDirectoryW = JobDirectoryW
End Property

Public Property Get JobDirectoryA() As String
    JobDirectoryA = pJobDirectoryA
End Property

Public Property Let JobDirectoryA(sInput As String)
    pJobDirectoryA = JobDirectoryA
End Property

Public Property Get JobDirectoryF() As String
    JobDirectoryF = pJobDirectoryF
End Property

Public Property Let JobDirectoryF(sInput As String)
    pJobDirectoryF = JobDirectoryF
End Property

Public Property Get JobDirectoryB() As String
    JobDirectoryB = pJobDirectoryB
End Property

Public Property Let JobDirectoryB(sInput As String)
    pJobDirectoryB = JobDirectoryB
End Property

Public Property Get JobDirectoryT() As String
    JobDirectoryT = pJobDirectoryT
End Property

Public Property Let JobDirectoryT(sInput As String)
    pJobDirectoryT = JobDirectoryT
End Property

Public Property Get JobDirectoryGN() As String
    JobDirectoryGN = pJobDirectoryGN
End Property

Public Property Let JobDirectoryGN(sInput As String)
    pJobDirectoryGN = JobDirectoryGN
End Property

Public Property Get JobDirectoryWN() As String
    JobDirectoryWN = pJobDirectoryWN
End Property

Public Property Let JobDirectoryWN(sInput As String)
    pJobDirectoryWN = JobDirectoryWN
End Property

Public Property Get JobDirectoryAN() As String
    JobDirectoryAN = pJobDirectoryAN
End Property

Public Property Let JobDirectoryAN(sInput As String)
    pJobDirectoryAN = JobDirectoryAN
End Property

Public Property Get JobDirectoryFN() As String
    JobDirectoryFN = pJobDirectoryFN
End Property

Public Property Let JobDirectoryFN(sInput As String)
    pJobDirectoryFN = JobDirectoryFN
End Property

Public Property Get JobDirectoryBN() As String
    JobDirectoryBN = pJobDirectoryBN
End Property

Public Property Let JobDirectoryBN(sInput As String)
    pJobDirectoryBN = JobDirectoryBN
End Property

Public Property Get JobDirectoryTN() As String
    JobDirectoryTN = pJobDirectoryTN
End Property

Public Property Let JobDirectoryTN(sInput As String)
    pJobDirectoryTN = JobDirectoryTN
End Property

Public Property Get JobDirectoryNN() As String
    JobDirectoryNN = pJobDirectoryNN
End Property

Public Property Let JobDirectoryNN(sInput As String)
    pJobDirectoryNN = JobDirectoryNN
End Property

Public Property Get JobDirectoryN() As String
    JobDirectoryN = pJobDirectoryN
End Property

Public Property Let JobDirectoryN(sInput As String)
    pJobDirectoryN = JobDirectoryN
End Property

Public Property Get TemplateFolder() As String
    TemplateFolder = pTemplateFolder
End Property

Public Property Let TemplateFolder(sInput As String)
    pTemplateFolder = TemplateFolder
End Property


Public Property Get TemplateFolder1() As String
    TemplateFolder1 = pTemplateFolder1
End Property

Public Property Let TemplateFolder1(sInput As String)
    pTemplateFolder1 = TemplateFolder1
End Property

Public Property Get TemplateFolder2() As String
    TemplateFolder2 = pTemplateFolder2
End Property

Public Property Let TemplateFolder2(sInput As String)
    pTemplateFolder2 = TemplateFolder2
End Property

Public Property Get TemplateFolder3() As String
    TemplateFolder3 = pTemplateFolder3
End Property

Public Property Let TemplateFolder3(sInput As String)
    pTemplateFolder3 = TemplateFolder3
End Property

Public Property Get TemplateFolder4() As String
    TemplateFolder4 = pTemplateFolder4
End Property

Public Property Let TemplateFolder4(sInput As String)
    pTemplateFolder4 = TemplateFolder4
End Property

Public Property Get PELD() As String
    PELD = pPELD
End Property

Public Property Let PELD(sInput As String)
    pPELD = TemplateFolder4
End Property

Public Property Get PELP() As String
    PELP = pPELP
End Property

Public Property Let PELP(sInput As String)
    pPELP = PELP
End Property

Public Property Get TempShipOptionsQ1XLSM() As String
    TempShipOptionsQ1XLSM = pTempShipOptionsQ1XLSM
End Property

Public Property Let TempShipOptionsQ1XLSM(sInput As String)
    pTempShipOptionsQ1XLSM = TempShipOptionsQ1XLSM
End Property

Public Property Get PaymentMade() As String
    PaymentMade = pPaymentMade
End Property

Public Property Let PaymentMade(sInput As String)
    pPaymentMade = PaymentMade
End Property

Public Property Get APFTR() As String
    APFTR = pAPFTR
End Property

Public Property Let APFTR(sInput As String)
    pAPFTR = APFTR
End Property

Public Property Get APCS() As String
    APCS = pAPCS
End Property

Public Property Let APCS(sInput As String)
    pAPCS = APCS
End Property

Public Property Get APCV() As String
    APCV = pAPCV
End Property

Public Property Let APCV(sInput As String)
    pAPCV = APCV
End Property

Public Property Get APES() As String
    APES = pAPES
End Property

Public Property Let APES(sInput As String)
    pAPES = APES
End Property

Public Property Get FileInbox() As String
    FileInbox = pFileInbox
End Property

Public Property Let FileInbox(sInput As String)
    pFileInbox = FileInbox
End Property

Public Property Get UNFileInbox() As String
    UNFileInbox = pUNFileInbox
End Property

Public Property Let UNFileInbox(sInput As String)
    pUNFileInbox = UNFileInbox
End Property

Public Property Get EmailDirectory() As String
    EmailDirectory = pEmailDirectory
End Property

Public Property Let EmailDirectory(sInput As String)
    pEmailDirectory = EmailDirectory
End Property

Public Property Get ShippingOutputFolder() As String
    ShippingOutputFolder = pShippingOutputFolder
End Property

Public Property Let ShippingOutputFolder(sInput As String)
    pShippingOutputFolder = ShippingOutputFolder
End Property

Public Property Get ShippingFolder() As String
    ShippingFolder = pShippingFolder
End Property

Public Property Let ShippingFolder(sInput As String)
    pShippingFolder = ShippingFolder
End Property

Public Property Get CompleteFolder() As String
    CompleteFolder = pCompleteFolder
End Property

Public Property Let CompleteFolder(sInput As String)
    pCompleteFolder = CompleteFolder
End Property

Public Property Get DistillerSettings() As String
    DistillerSettings = pDistillerSettings
End Property

Public Property Let DistillerSettings(sInput As String)
    pDistillerSettings = DistillerSettings
End Property

Public Property Get DistillerSettings1() As String
    DistillerSettings1 = pDistillerSettings1
End Property

Public Property Let DistillerSettings1(sInput As String)
    pDistillerSettings1 = DistillerSettings1
End Property

Public Property Get KCITemplate() As String
    KCITemplate = pKCITemplate
End Property

Public Property Let KCITemplate(sInput As String)
    pKCITemplate = KCITemplate
End Property

Public Property Get CDLabelP() As String
    CDLabelP = pCDLabelP
End Property

Public Property Let CDLabelP(sInput As String)
    pCDLabelP = CDLabelP
End Property

Public Property Get XMLAfter() As String
    XMLAfter = pXMLAfter
End Property

Public Property Let XMLAfter(sInput As String)
    pXMLAfter = XMLAfter
End Property

Public Property Get XMLBefore() As String
    XMLBefore = pXMLBefore
End Property

Public Property Let XMLBefore(sInput As String)
    pXMLBefore = XMLBefore
End Property



Public Sub Class_Initialize()
        sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
        psDrive = "T"
        pNetWorkFolder = "\\hubcloud\evoingram\"
        
    'paths needed: transcribing manuals, samples, jurisdiction references, courtrules
        'shipping xml templates maybe
        
        'TODO: Refresh invoice number on this function
        pInvoiceP = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & sInvoiceNumber & ".pdf"
        pInvoiceD = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & sInvoiceNumber & ".docx"
        'TODO: Check COA shipping xml file name
        pShippingXML = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-Shipping.xml"
        pShippingCOAXML = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-CourtofAppealsDivI-Shipping.xml"
        pShippingXML4 = psDrive & ":\Production\4ShippingXMLs\" & sCourtDatesID & "-shipping.xml"
        pShippingCOAXML4 = psDrive & ":/Production/4ShippingXMLs/" & sCourtDatesID & "-CourtofAppealsDivI-Shipping.xml"
        pShippingOutputFolder = "T:\Production\4ShippingXMLs\Output\"
        pRoughDraft = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/RoughDraft.docx"
        pShippingFolder = "T:\Production\4ShippingXMLs\"
        pCompleteFolder = "T:\Production\3Complete\"
        
        pCaseInfo = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-CaseInfo.xls"
        pInvoiceInfo = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-InvoiceInfo.xls"
        pXeroCSV = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "--XeroInvoiceCSV.csv"
        pPPCSV = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-PayPalInvoiceCSV.csv"
        pPPButton = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-PPButton.html"
        pCDLabel = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-CDLabel.pub"
        pCDLabelP = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-CDLabel.pdf"
        pT2upPS = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-Transcript-FINAL-2up.ps"
        pT4upPS = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-Transcript-FINAL-4up.ps"
        pT2upLog = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-Transcript-FINAL-2up.log"
        pT4upLog = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-Transcript-FINAL-4up.log"
        
        pNotes = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Notes/" & sCourtDatesID & "-Notes.pdf"
        
        pCourtCover = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-CourtCover.docx"
        pCIDIncomeReportD = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-CIDIncomeReport.docx"
        pCIDIncomeReportP = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-CIDIncomeReport.pdf"
        pDepositInvoiceP = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-DepositInvoice.pdf"
        pDepositInvoiceD = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-DepositInvoice.docx"
        pOrderConfirmationP = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-OrderConfirmation.pdf"
        pOrderConfirmationD = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-OrderConfirmation.docx"
        pPPDraftInvoiceEmail = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-PP-DraftInvoiceEmail.docx"
        pTranscriptsReadyP = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-TranscriptsReady.pdf"
        pTranscriptsReadyD = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-TranscriptsReady.docx"
        pTranscriptsReadyCP = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-ContractorTranscriptsReady.pdf"
        pTranscriptsReadyCD = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-ContractorTranscriptsReady.docx"
        pKCIEmpty = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-KCICompleted.pdf"
        pKCIFilled = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-KCICompleted1.pdf"
        pPELP = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-PackageEnclosedLetter.pdf"
        pPELD = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-PackageEnclosedLetter.docx"
        pPaymentMade = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "\Generated\" & sCourtDatesID & "-PaymentMade.docx"
        
        pWACoverP = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-Cover.pdf"
        pWACoverD = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-Cover.docx"
        pWAConsolidatedP = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-Transcripts-All.pdf"
        pWAConsolidatedD = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-Transcripts-All.docx"
        pTranscriptFP = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-Transcript-FINAL.pdf"
        pTranscriptFD = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-Transcript-FINAL.docx"
        pTranscript2up = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-Transcript-FINAL-2up.pdf"
        pTranscript4up = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-Transcript-FINAL-4up.pdf"
        pTranscriptWC = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-Transcript-WorkingCopy.docx"
        pWordIndexP = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-WordIndex.pdf"
        pWordIndexD = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-WordIndex.docx"
        
        pAudioFolder = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Audio/"
        
        pZAudioF = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/FTP/" & sCourtDatesID & "-Audio.zip"
        pZTranscriptsF = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/FTP/" & sCourtDatesID & "-Transcripts.zip"
        pZAudioTranscriptsF = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/FTP/" & sCourtDatesID & "-AudioTranscripts.zip"
        
        pZAudioB = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-Audio.zip"
        pZTranscriptsB = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-Transcripts.zip"
        pZAudioTranscriptsB = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-AudioTranscripts.zip"
        pTranscriptFPB = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-Transcript-FINAL.pdf"
        pTranscriptFDB = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-Transcript-FINAL.docx"
        pTranscript2upB = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-Transcript-FINAL-2up.pdf"
        pTranscript4upB = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-Transcript-FINAL-4up.pdf"
        pTranscriptWCB = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-Transcript-WorkingCopy.docx"
        pWordIndexPB = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-WordIndex.pdf"
        pWordIndexDB = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-WordIndex.docx"
        
        pKCIEnvelope = psDrive & ":/Database\Templates\Stage4s\Envelope-KCI.pdf"
        pKCITemplate = psDrive & ":\Database\Templates\Stage4s\KCICompleted.pdf"
        pPQTemplate = psDrive & ":/Database\Templates\Stage1s\PriceQuoteEmail-Template.docx"
        pPQXLS = psDrive & ":/Database\Scripts\InProgressExcels\Temp-Export-PQE.xlsx"
        pPQEmail = psDrive & ":/Database\Templates\Stage1s\PriceQuoteEmail.docx"
        pOrderFormCustomersXLSM = psDrive & ":/Database\Scripts\InProgressExcels\JotformCustomers.xlsm"
        pOrderFormCustomersXLSX = psDrive & ":/Database\Scripts\InProgressExcels\JotformCustomers.xlsx"
        pOrderFormCustomersCSV = psDrive & ":/Database\Scripts\InProgressExcels\JotformCustomers.csv"
        pOrderFormXLSM = psDrive & ":/Database\Scripts\InProgressExcels\Jotform.xlsm"
        pOrderFormXLSX = psDrive & ":/Database\Scripts\InProgressExcels\Jotform.xlsx"
        pOrderFormCSV = psDrive & ":/Database\Scripts\InProgressExcels\Jotform.csv"
        pTemplateFolder = psDrive & ":/Database/Templates/"
        pTemplateFolder1 = psDrive & ":/Database/Templates/Stage1s/"
        pTemplateFolder2 = psDrive & ":/Database/Templates/Stage2s/"
        pTemplateFolder3 = psDrive & ":/Database/Templates/Stage3s/"
        pTemplateFolder4 = psDrive & ":/Database/Templates/Stage4s/"
        pTempShipOptionsQ1XLSM = psDrive & ":\Database\Scripts\InProgressExcels\TempShippingOptionsQ1.xlsm"
        
        pXMLAfter = psDrive & ":\Database\Scripts\InProgressExcels\AfterXML.xml"
        pXMLBefore = psDrive & ":\Database\Scripts\InProgressExcels\BeforeXML.xml"
        
        pJobDirectory = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/"
        pJobDirectoryG = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/"
        pJobDirectoryW = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/WorkingFiles"
        pJobDirectoryA = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Audio/"
        pJobDirectoryF = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/FTP/"
        pJobDirectoryB = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Backups/"
        pJobDirectoryT = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Transcripts/"
        pJobDirectoryN = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Notes/"
        
        pJobDirectoryGN = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-"
        pJobDirectoryWN = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-"
        pJobDirectoryAN = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Audio/" & sCourtDatesID & "-"
        pJobDirectoryFN = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/FTP/" & sCourtDatesID & "-"
        pJobDirectoryBN = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-"
        pJobDirectoryTN = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-"
        pJobDirectoryNN = psDrive & ":/Production/2InProgress/" & sCourtDatesID & "/Notes/" & sCourtDatesID & "-"
        
        pAPFTR = "C:\Program Files (x86)\FTR\ForTheRecord\TheRecordPlayer.exe"
        pAPCS = "T:\Production\2InProgress\" & sCourtDatesID & "\Audio\bin\cplayer.exe"
        pAPCV = "C:\Program Files (x86)\JAVS\CaseViewer\CaseViewer3.exe"
        pAPES = "C:\Program Files (x86)\NCH Software\Scribe\scribe.exe"
        
        pFileInbox = psDrive & ":/Production/1ToBeEntered/"
        pUNFileInbox = pNetWorkFolder & "Production/1ToBeEntered/"
        
        pEmailDirectory = psDrive & ":\Database\Emails\"
        
        pDistillerSettings = "C:\Program Files (x86)\Adobe\Acrobat 9.0\Acrobat\Settings\Standard.joboptions"
        pDistillerSettings1 = "C:\users\inqui\Standard1.joboptions"
        
End Sub


Private Function NullIfEmptyString(str As String) As Variant
    Dim strTrimmed As String: strTrimmed = Trim(str)
    If Len(strTrimmed) = 0 Then
        NullIfEmptyString = Null
    Else
        NullIfEmptyString = strTrimmed
    End If
End Function
