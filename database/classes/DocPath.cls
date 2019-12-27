VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Ignore OptionExplicit
'@Folder("Database.Production.Classes")
Option Compare Database

Dim psDrive As String
Dim pNetWorkFolder As String

Private pAPCS As String
Private pAPCV As String
Private pAPES As String
Private pAPFTR As String
Private pAudioFolder As String
Private pCaseInfo As String
Private pCDLabel As String
Private pCDLabelP As String
Private pCIDIncomeReportD As String
Private pCIDIncomeReportP As String
Private pCompleteFolder As String
Private pCourtCover As String
Private pDBScripts As String
Private pDepositInvoiceD As String
Private pDepositInvoiceP As String
Private pDistillerSettings As String
Private pDistillerSettings1 As String
Private pEmailDirectory As String
Private pFileInbox As String
Private pInProgressFolder As String
Private pInvoiceD As String
Private pInvoiceInfo As String
Private pInvoiceP As String
Private pJobDirectory As String
Private pJobDirectoryA As String
Private pJobDirectoryAN As String
Private pJobDirectoryB As String
Private pJobDirectoryBN As String
Private pJobDirectoryF As String
Private pJobDirectoryFN As String
Private pJobDirectoryG As String
Private pJobDirectoryGN As String
Private pJobDirectoryN As String
Private pJobDirectoryNN As String
Private pJobDirectoryT As String
Private pJobDirectoryTN As String
Private pJobDirectoryW As String
Private pJobDirectoryWN As String
Private pJurisdictionRefs As String
Private pKCIEmpty As String
Private pKCIEnvelope As String
Private pKCIFilled As String
Private pKCITemplate As String
Private pNotes As String
Private pOrderConfirmationD As String
Private pOrderConfirmationP As String
Private pOrderFormCSV As String
Private pOrderFormCustomersCSV As String
Private pOrderFormCustomersXLSM As String
Private pOrderFormCustomersXLSX As String
Private pOrderFormXLSM As String
Private pOrderFormXLSX As String
Private pPaymentMade As String
Private pPELD As String
Private pPELP As String
Private pPPBalanceDue As String
Private pPPBalanceDueT As String
Private pPPButton As String
Private pPPButtonT As String
Private pPPCSV As String
Private pPPDepositE As String
Private pPPDIET As String
Private pPPDraftInvoiceEmail As String
Private pPPDraftInvoiceH As String
Private pPPFactoredInvoiceEmail As String
Private pPPFactoredInvoiceH As String
Private pPPFIET As String
Private pPQEmail As String
Private pPQTemplate As String
Private pPQXLS As String
Private pRoughDraft As String
Private pShippingCOAXML As String
Private pShippingCOAXML4 As String
Private pShippingFolder As String
Private pShippingOutputFolder As String
Private pShippingXML As String
Private pShippingXML4 As String
Private pSpeech As String
Private pSpeechUA As String
Private pSpeechUP As String
Private pT2upLog As String
Private pT2upPS As String
Private pT4upLog As String
Private pT4upPS As String
Private pTemplateFolder As String
Private pTemplateFolder1 As String
Private pTemplateFolder2 As String
Private pTemplateFolder3 As String
Private pTemplateFolder4 As String
Private pTempShipOptionsQ1XLSM As String
Private pToBeEnteredFolder As String
Private pTranscript2up As String
Private pTranscript2upB As String
Private pTranscript4up As String
Private pTranscript4upB As String
Private pTranscriptFD As String
Private pTranscriptFDB As String
Private pTranscriptFP As String
Private pTranscriptFPB As String
Private pTranscriptsReadyCD As String
Private pTranscriptsReadyCP As String
Private pTranscriptsReadyD As String
Private pTranscriptsReadyP As String
Private pTranscriptWC As String
Private pTranscriptWCB As String
Private pUNFileInbox As String
Private pWAConsolidatedD As String
Private pWAConsolidatedP As String
Private pWACoverD As String
Private pWACoverP As String
Private pWordIndexD As String
Private pWordIndexDB As String
Private pWordIndexP As String
Private pWordIndexPB As String
Private pWordIndexT As String
Private pXeroCSV As String
Private pXMLAfter As String
Private pXMLBefore As String
Private pZAudioB As String
Private pZAudioF As String
Private pZAudioTranscriptsB As String
Private pZAudioTranscriptsF As String
Private pZTranscriptsB As String
Private pZTranscriptsF As String
Public Property Get InvoiceP() As String
    InvoiceP = pInvoiceP
End Property

Public Property Let InvoiceP(sInput As String)
    pInvoiceP = InvoiceP
End Property

Public Property Get InvoiceD() As String
    InvoiceD = pInvoiceD
End Property

Public Property Let InvoiceD(sInput As String)
    pInvoiceD = InvoiceD
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

Public Property Get PPFactoredInvoiceEmail() As String
    PPFactoredInvoiceEmail = pPPFactoredInvoiceEmail
End Property

Public Property Let PPFactoredInvoiceEmail(sInput As String)
    pPPFactoredInvoiceEmail = PPFactoredInvoiceEmail
End Property

Public Property Get PPFIET() As String
    PPFIET = pPPFIET
End Property

Public Property Let PPFIET(sInput As String)
    pPPFIET = PPFIET
End Property

Public Property Get PPFactoredInvoiceH() As String
    PPFactoredInvoiceH = pPPFactoredInvoiceH
End Property

Public Property Let PPFactoredInvoiceH(sInput As String)
    pPPFactoredInvoiceH = PPFactoredInvoiceH
End Property

Public Property Get PPButtonT() As String
    PPButtonT = pPPButtonT
End Property

Public Property Let PPButtonT(sInput As String)
    pPPButtonT = PPButtonT
End Property

Public Property Get PPBalanceDue() As String
    PPBalanceDue = pPPBalanceDue
End Property

Public Property Let PPBalanceDue(sInput As String)
    pPPBalanceDue = PPBalanceDue
End Property

Public Property Get PPBalanceDueT() As String
    PPBalanceDueT = pPPBalanceDueT
End Property

Public Property Let PPBalanceDueT(sInput As String)
    pPPBalanceDueT = PPBalanceDueT
End Property

Public Property Get PPDraftInvoiceH() As String
    PPDraftInvoiceH = pPPDraftInvoiceH
End Property

Public Property Let PPDraftInvoiceH(sInput As String)
    pPPDraftInvoiceH = PPDraftInvoiceH
End Property

Public Property Get PPDIET() As String
    PPDIET = pPPDIET
End Property

Public Property Let PPDIET(sInput As String)
    pPPDIET = PPDIET
End Property

Public Property Get PPDepositE() As String
    PPDepositE = pPPDepositE
End Property

Public Property Let PPDepositE(sInput As String)
    pPPDepositE = PPDepositE
End Property
Public Property Get WordIndexT() As String
    WordIndexT = pWordIndexT
End Property

Public Property Let WordIndexT(sInput As String)
    pWordIndexT = WordIndexT
End Property
Public Property Get WordIndexDBT() As String
    WordIndexDBT = pWordIndexDBT
End Property

Public Property Let WordIndexDBT(sInput As String)
    pWordIndexDBT = WordIndexDBT
End Property

Public Property Get WordIndexDB() As String
    WordIndexDB = pWordIndexDB
End Property

Public Property Let WordIndexDB(sInput As String)
    pWordIndexDB = WordIndexDB
End Property

Public Property Get JurisdictionRefs() As String
    JurisdictionRefs = pJurisdictionRefs
End Property

Public Property Let JurisdictionRefs(sInput As String)
    pJurisdictionRefs = JurisdictionRefs
End Property

Public Property Get DBScripts() As String
    DBScripts = pDBScripts
End Property

Public Property Let DBScripts(sInput As String)
    pDBScripts = DBScripts
End Property

Public Property Get SpeechUP() As String
    SpeechUP = pSpeechUP
End Property

Public Property Let SpeechUP(sInput As String)
    pSpeechUP = SpeechUP
End Property

Public Property Get SpeechUA() As String
    SpeechUA = pSpeechUA
End Property

Public Property Let SpeechUA(sInput As String)
    pSpeechUA = SpeechUA
End Property


Public Property Get InProgressFolder() As String
    InProgressFolder = pInProgressFolder
End Property

Public Property Let InProgressFolder(sInput As String)
    pInProgressFolder = InProgressFolder
End Property

Public Property Get ToBeEnteredFolder() As String
    ToBeEnteredFolder = pToBeEnteredFolder
End Property

Public Property Let ToBeEnteredFolder(sInput As String)
    pToBeEnteredFolder = ToBeEnteredFolder
End Property

Public Sub Class_Initialize()
        psDrive = "T"
        pNetWorkFolder = "E:/P/"
        pShippingOutputFolder = pNetWorkFolder & "Production/4ShippingXMLs/Output/"
        pShippingFolder = pNetWorkFolder & "Production/4ShippingXMLs/"
        pCompleteFolder = pNetWorkFolder & "Production/3Complete/"
        pInProgressFolder = pNetWorkFolder & "Production/2InProgress/"
        pToBeEnteredFolder = pNetWorkFolder & "Production/1ToBeEntered/"
        pKCIEnvelope = pNetWorkFolder & "Database/Templates/Stage4s/Envelope-KCI.pdf"
        pKCITemplate = pNetWorkFolder & "Database/Templates/Stage4s/KCICompleted.pdf"
        pPQTemplate = pNetWorkFolder & "Database/Templates/Stage1s/PriceQuoteEmail-Template.docx"
        pPQXLS = pNetWorkFolder & "Database/Scripts/InProgressExcels/Temp-Export-PQE.xlsx"
        pPQEmail = pNetWorkFolder & "Database/Templates/Stage1s/PriceQuoteEmail.docx"
        pOrderFormCustomersXLSM = pNetWorkFolder & "Database/Scripts/InProgressExcels/JotformCustomers.xlsm"
        pOrderFormCustomersXLSX = pNetWorkFolder & "Database/Scripts/InProgressExcels/JotformCustomers.xlsx"
        pOrderFormCustomersCSV = pNetWorkFolder & "Database/Scripts/InProgressExcels/OrderFormCustomers.csv"
        pOrderFormXLSM = pNetWorkFolder & "Database/Scripts/InProgressExcels/Jotform.xlsm"
        pOrderFormXLSX = pNetWorkFolder & "Database/Scripts/InProgressExcels/Jotform.xlsx"
        pOrderFormCSV = pNetWorkFolder & "Database/Scripts/InProgressExcels/OrderForm.csv"
        pTemplateFolder = pNetWorkFolder & "Database/Templates/"
        pTemplateFolder1 = pNetWorkFolder & "Database/Templates/Stage1s/"
        pTemplateFolder2 = pNetWorkFolder & "Database/Templates/Stage2s/"
        pTemplateFolder3 = pNetWorkFolder & "Database/Templates/Stage3s/"
        pTemplateFolder4 = pNetWorkFolder & "Database/Templates/Stage4s/"
        pTempShipOptionsQ1XLSM = pNetWorkFolder & "Database/Scripts/InProgressExcels/TempShippingOptionsQ1.xlsm"
        pPPFIET = pNetWorkFolder & "Database/Templates/Stage4s/PP-FactoredInvoiceEmail-Template.docx"
        pPPBalanceDueT = pNetWorkFolder & "Database/Templates/Stage4s/PP-BalanceDueInvoiceEmail.docx"
        pPPDIET = pNetWorkFolder & "Database/Templates/Stage1s/PP-DraftInvoiceEmail-Template.htm"
        pWordIndexT = pNetWorkFolder & "Database/Templates/Stage2s/TR-WordIndex.dotm"
        pXMLAfter = pNetWorkFolder & "Database/Scripts/InProgressExcels/AfterXML.xml"
        pXMLBefore = pNetWorkFolder & "Database/Scripts/InProgressExcels/BeforeXML.xml"
        pDBScripts = pNetWorkFolder & "Database/Scripts/"
        pAPCV = "C:/Program Files (x86)/JAVS/CaseViewer/CaseViewer3.exe"
        pAPES = "C:/Program Files (x86)/NCH Software/Scribe/scribe.exe"
        
        pFileInbox = pNetWorkFolder & "Production/1ToBeEntered/"
        pUNFileInbox = pNetWorkFolder & "Production/1ToBeEntered/"
        
        pEmailDirectory = pNetWorkFolder & "Database/Emails/"
        
        pDistillerSettings = "C:/Program Files (x86)/Adobe/Acrobat 9.0/Acrobat/Settings/Standard.joboptions"
        pDistillerSettings1 = "C:/users/inqui/Standard1.joboptions"
        
        pJurisdictionRefs = pNetWorkFolder & "Administration/Jurisdiction References/"
        
        pSpeech = pNetWorkFolder & "speech/UnprocessedAudio/Prepared/"
        pSpeechUA = pNetWorkFolder & "speech/UnprocessedAudio/"
        
        pAPFTR = "C:/Program Files (x86)/FTR/ForTheRecord/TheRecordPlayer.exe"
        
        If Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField] <> "" Then
            
            sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
            Dim rstTableName As DAO.Recordset
            Set rstTableName = CurrentDb.OpenRecordset("SELECT * FROM CourtDates WHERE [ID] = " & sCourtDatesID & ";")
            sInvoiceNumber = rstTableName.Fields("InvoiceNo").Value
            rstTableName.Close
            Set rstTableName = Nothing
            
        End If
        
            pAPCS = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Audio/bin/cplayer.exe"
            
            pInvoiceP = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sInvoiceNumber & ".pdf"
            pInvoiceD = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sInvoiceNumber & ".docx"
            pShippingXML = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-Shipping.xml"
            pShippingCOAXML = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-CourtofAppealsDivI-Shipping.xml"
            pShippingXML4 = pNetWorkFolder & "Production/4ShippingXMLs/" & sCourtDatesID & "-shipping.xml"
            pShippingCOAXML4 = pNetWorkFolder & "Production/4ShippingXMLs/" & sCourtDatesID & "-CourtofAppealsDivI-Shipping.xml"
            pRoughDraft = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/RoughDraft.docx"
            
            pCaseInfo = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-CaseInfo.xls"
            pInvoiceInfo = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-InvoiceInfo.xls"
            pXeroCSV = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "--XeroInvoiceCSV.csv"
            pPPCSV = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-PayPalInvoiceCSV.csv"
            pPPButton = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-PPButton.html"
            pPPButtonT = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-PPButton.txt"
            pCDLabel = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-CDLabel.pub"
            pCDLabelP = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-CDLabel.pdf"
            pT2upPS = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-Transcript-FINAL-2up.ps"
            pT4upPS = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-Transcript-FINAL-4up.ps"
            pT2upLog = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-Transcript-FINAL-2up.log"
            pT4upLog = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-Transcript-FINAL-4up.log"
            
            pNotes = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Notes/" & sCourtDatesID & "-Notes.pdf"
            
            pCourtCover = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-CourtCover.docx"
            pCIDIncomeReportD = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-CIDIncomeReport.docx"
            pCIDIncomeReportP = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-CIDIncomeReport.pdf"
            pDepositInvoiceP = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-DepositInvoice.pdf"
            pDepositInvoiceD = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-DepositInvoice.docx"
            pOrderConfirmationP = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-OrderConfirmation.pdf"
            pOrderConfirmationD = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-OrderConfirmation.docx"
            pPPDraftInvoiceEmail = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-PP-DraftInvoiceEmail.docx"
            pPPDraftInvoiceH = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-PP-DraftInvoiceEmail.html"
            pPPFactoredInvoiceEmail = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-PP-FactoredInvoiceEmail.docx"
            pPPFactoredInvoiceH = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-PP-FactoredInvoiceEmail.html"
            pPPBalanceDue = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-PP-BalanceDueInvoice.docx"
            pPPDepositE = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-PPDepositEmail.html"
            pTranscriptsReadyP = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-TranscriptsReady.pdf"
            pTranscriptsReadyD = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-TranscriptsReady.docx"
            pTranscriptsReadyCP = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-ContractorTranscriptsReady.pdf"
            pTranscriptsReadyCD = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-ContractorTranscriptsReady.docx"
            pKCIEmpty = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-KCICompleted.pdf"
            pKCIFilled = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-KCICompleted1.pdf"
            pPELP = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-PackageEnclosedLetter.pdf"
            pPELD = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-PackageEnclosedLetter.docx"
            pPaymentMade = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-PaymentMade.docx"
            
            pWACoverP = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-Cover.pdf"
            pWACoverD = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-Cover.docx"
            pWAConsolidatedP = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-Transcripts-All.pdf"
            pWAConsolidatedD = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-Transcripts-All.docx"
            pTranscriptFP = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-Transcript-FINAL.pdf"
            pTranscriptFD = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-Transcript-FINAL.docx"
            pTranscript2up = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-Transcript-FINAL-2up.pdf"
            pTranscript4up = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-Transcript-FINAL-4up.pdf"
            pTranscriptWC = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-Transcript-WorkingCopy.docx"
            pWordIndexP = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-WordIndex.pdf"
            pWordIndexD = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-WordIndex.docx"
            
            pAudioFolder = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Audio/"
            
            pZAudioF = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/FTP/" & sCourtDatesID & "-Audio.zip"
            pZTranscriptsF = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/FTP/" & sCourtDatesID & "-Transcripts.zip"
            pZAudioTranscriptsF = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/FTP/" & sCourtDatesID & "-AudioTranscripts.zip"
            
            pZAudioB = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-Audio.zip"
            pZTranscriptsB = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-Transcripts.zip"
            pZAudioTranscriptsB = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-AudioTranscripts.zip"
            pTranscriptFPB = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-Transcript-FINAL.pdf"
            pTranscriptFDB = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-Transcript-FINAL.docx"
            pTranscript2upB = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-Transcript-FINAL-2up.pdf"
            pTranscript4upB = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-Transcript-FINAL-4up.pdf"
            pTranscriptWCB = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-Transcript-WorkingCopy.docx"
            pWordIndexPB = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-WordIndex.pdf"
            pWordIndexDB = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-WordIndex.docx"
            
            pJobDirectory = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/"
            pJobDirectoryG = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/"
            pJobDirectoryW = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/WorkingFiles"
            pJobDirectoryA = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Audio/"
            pJobDirectoryF = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/FTP/"
            pJobDirectoryB = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Backups/"
            pJobDirectoryT = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Transcripts/"
            pJobDirectoryN = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Notes/"
            
            pJobDirectoryGN = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Generated/" & sCourtDatesID & "-"
            pJobDirectoryWN = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/WorkingFiles/" & sCourtDatesID & "-"
            pJobDirectoryAN = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Audio/" & sCourtDatesID & "-"
            pJobDirectoryFN = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/FTP/" & sCourtDatesID & "-"
            pJobDirectoryBN = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Backups/" & sCourtDatesID & "-"
            pJobDirectoryTN = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Transcripts/" & sCourtDatesID & "-"
            pJobDirectoryNN = pNetWorkFolder & "Production/2InProgress/" & sCourtDatesID & "/Notes/" & sCourtDatesID & "-"
End Sub


Private Function NullIfEmptyString(str As String) As Variant
    Dim strTrimmed As String: strTrimmed = Trim(str)
    If Len(strTrimmed) = 0 Then
        NullIfEmptyString = Null
    Else
        NullIfEmptyString = strTrimmed
    End If
End Function
