Attribute VB_Name = "AboutDB"
'@IgnoreModule EmptyModule
'@Folder("Database.General.Modules")
Option Compare Database
Option Explicit

 
'folder organization in 'T:\Production\2InProgress' or 'I:\'
'####
'Audio:         Audio files
'Notes:         Paperwork provided
'FTP:           ZIPs for distribution
'####-Audio
'####-AudioTranscripts
'####-Transcripts

'Transcripts:   Final versions of transcripts
'####-Transcript-FINAL.docx
'####-Transcript-FINAL.pdf
'####-Transcript-FINAL-2up.pdf
'####-Transcript-FINAL-4up.pdf
'####-WordIndex.docx
'####-Transcript-WorkingCopy.docx

'WorkingFiles:  Various CSVs for jobs
'XeroInvoiceCSV
'PayPalInvoiceCSV
'CaseInfo.xls
'InvoiceInfo.xls
'ps files for distiller printing
'pp button html generated files

'Generated:     Generated client-facing paperwork for job
'CourtCover
'CIDIncomeReport/final
'PP emails
'deposit/factoring invoices
'order confirmation
'PEL letter
'transcripts ready letter
'cd label
            
'Backups:       Extra back-up copy of final files instead of in main folder
'####-Transcript-FINAL.docx
'####-Transcript-FINAL.pdf
'####-Transcript-FINAL-2up.pdf
'####-Transcript-FINAL-4up.pdf
'####-WordIndex.docx
'####-Transcript-WorkingCopy.docx






'module AboutDB

'============================================================================
'module AdminFunctions

'variables:
'   NONE

'functions:
'pfUpdateCheckboxStatus:         Description:  updates Statuses checkbox field you specify
'                                Arguments:    sStatusesField
'pfDownloadfromFTP:              Description:  downloads files
'                            Arguments:    NONE
'pfDownloadFTPsite:              Description:  downloads files modified today (a.k.a. new files on FTP)
'                            Arguments:    mySession
'pfCheckFolderExistence:         Description:  checks for Audio, Transcripts, FTP, WorkingFiles, Notes subfolders and RoughDraft and creates if not exists
'                            Arguments:    NONE
'pfCommunicationHistoryAdd:      Description:  adds entry to CommunicationHistory
'                            Arguments:    CHTopic
'pfStripIllegalChar:             Description:  strips illegal characters from input
'                            Arguments:    StrInput
'pfGetFolder:                    Description:  gets folder
'                            Arguments:    Folders, EntryID, StoreID, fld
'pfBrowseForFolder:              Description:  browses for folder
'                            Arguments:    StrSavePath, optional OpenAt
'pfSingleBAScrapeSpecificBarNo:  Description:  gets one bar number's info from the WA bar website (pick any range from 1 to 55000)
'                            Arguments:    sWebSiteBarNo
'pfScrapingBALoop:               Description:  gets a range you specify of bar numbers' info from the WA bar website (pick any range from 1 to 55000)
'                            Arguments:    vWebSiteBarNo, vWebSiteBarNoGoal
'pfReformatTable:                Description:  reformats scraped Bar addresses to useable format for table
'                            Arguments:    NONE
'pfUpdateCheckboxStatus:         Description:  updates Statuses checkbox field you specify
'                            Arguments:    sStatusesField
'pfDelay:                        Description:  sleep function
'                            Arguments:    lSeconds
'pfPriorityPointsAlgorithm:      Description:  assigns priority points to various tasks in Tasks table and inserts it into the PriorityPoints field
'                                          priority scale 1 to 100
'                            Arguments:    NONE
'pfDebugSQLStatement:            Description:  debug.prints data source query string
'                            Arguments:    NONE
'pfGenerateJobTasks:             Description:  generates job tasks in the Tasks table
'                            Arguments:    NONE
'pfDownloadfromFTP:              Description:  downloads files
'                            Arguments:    NONE
'pfDownloadFTPsite:              Description:  downloads files modified today (a.k.a. new files on FTP)
'                            Arguments:    mySession
'pfProcessFolder:                Description:  process emails in Outlook folder named AccessTest and places them in db as UnprocessedCommunication
'                            Arguments:    oOutlookMAPIFolder
'pfFileExists:                   Description:  check if file exists
'                            Arguments:    path
'pfAcrobatGetNumPages:           Description:  gets number of pages from PDF and confirms with you
'                                          IS TOA ON SECOND PAGE?  IF YES, -2 pgs; IF NO, -1 pg
'                            Arguments:    sCourtDatesID
'pfReadXML:                      Description:  reads shipping XML and sends "Shipped" email to client
'                            Arguments:    NONE
'pfFileRenamePrompt:             Description:  renames transcript to specified name, mainly for contractors
'                            Arguments:    NONE
'pfWaitSeconds:                  Description:  waits for a specified number of seconds
'                            Arguments:    iSeconds
'pfDailyTaskAddFunction:         Description:  adds static daily tasks to Tasks table
'                            Arguments:    NONE
'pfAvailabilitySchedule:         Description:  opens availability calculator
'                            Arguments:    NONE
'pfWeeklyTaskAddFunction:        Description:  adds static weekly tasks to Tasks table
'                            Arguments:    NONE
'pfMonthlyTaskAddFunction:       Description:  adds static monthly tasks to Tasks table
'                            Arguments:    NONE
'pfMoveSelectedMessages:         Description:  move selected messages to network drive
'                            Arguments:    NONE
'pfEmailsExport1:                Description:  export specified fields from each mail / item in selected folder
'                            Arguments:    NONE
'pfCommHistoryExportSub:         Description:  exports emails to CommunicationsHistory table
'                            Arguments:    NONE
'pfAskforNotes:                  Description:  file dialog picker to select notes and copy them to notes folder for job
'                            Arguments:    NONE
'pfAskforAudio:                  Description:  file dialog picker to select audio and copy them to audio folder for job
'                            Arguments:    NONE
'fWunderlistGetFolders()         Description:  gets list of Wunderlist folders or folder revisions
'                            Arguments:    NONE
'fWunderlistGetTasksOnList()     Description:  gets tasks on Wunderlist list
'                            Arguments:    NONE
'fWunderlistAdd()                Description:  adds task to Wunderlist
'                            Arguments:    NONE
'fWLGenerateJSONInfo             Description:  get info for WL API
'                            Arguments:    NONE
'fWunderlistGetLists()           Description:  gets all Wunderlist lists
'                            Arguments:    NONE
'pfRCWRuleScraper1()             Description:  builds RCW rule links and citations
'                            Arguments:    NONE
'GetLevel()                      Description:  gets header level in word
'                            Arguments:    NONE
'============================================================================

'module CurrentCaseInfo

'variables:
'    Public theRecognizers As ISpeechObjectTokens
'    Public SharedRecognizer As SpSharedRecognizer
'    Public i As Long
'    Public sParty1 As String, sCompany As String, sParty2 As String, sCourtDatesID As String, sInvoiceNumber As String
'    Public sParty1Name As String, sParty2Name As String, sInvoiceNo As String, sEmail As String, sDescription As String
'    Public sSubtotal As String, sInvoiceDate As String, sInvoiceTime As String, sPaymentTerms As String, sNote As String
'    Public sTerms As String, sMinimumAmount As String, vmMemo As String, vlURL As String, sTemplateID As String
'    Public sLine1 As String, sCity As String, sState As String, sZIP As String, sQuantity As String, sValue As String
'    Public sActualQuantity As String, sJurisdiction As String, sTurnaroundTime As String, sCaseNumber1 As String, sCaseNumber2 As String
'    Public sCustomerID As String, sAudioLength As String, sEstimatedPageCount As String, sStatusesID As String, sFinalPrice As String
'    Public dDueDate As Date, dExpectedAdvanceDate As Date, dExpectedRebateDate As Date, sPaymentSum As String
'    Public sBalanceDue As String, sFactoringCost As String, svURL As String, sLinkToCSV As String, sFactoringApproved As String
'    Public sFirstName As String, sLastName As String, dHearingDate As Date, sMrMs As String
'    Public sName As String, sAddress1 As String, sAddress2 As String, sNotes As String
'    Public qnTRCourtQ As String, qnTRCourtUnionAppAddrQ As String, qnShippingOptionsQ As String, qnViewJobFormAppearancesQ As String
'    HyperlinkString As String, rtfstringbody As String, sTime As String, sTime1 As String
'    Public lngNumOfHrs As Long, lngNumOfMins As Long, lngNumOfSecsRem As Long, lngNumOfSecs As Long
'    Public lngNumOfHrs1 As Long, lngNumOfMins1 As Long, lngNumOfSecsRem1 As Long, lngNumOfSecs1 As Long
'    Public sClientTranscriptName As String, sCurrentTranscriptName As String
 
'functions:

'pfCurrentCaseInfo:          Description:  refreshes global variables for current transcript
'                            Arguments:    NONE

'pfGetOrderingAttorneyInfo:  Description:  refreshes ordering attorney info for transcript
'                            Arguments:    NONE

'pfClearGlobals:             Description:  clears all global variables
'                            Arguments:    NONE

'fPPGenerateJSONInfo:        Description:  get info for invoice
'                            Arguments:    NONE


'============================================================================
'module DocGen

'variables:
'   NONE

'functions:

'pfGenericExportandMailMerge:  Description:  exports to specified template from T:\Database\Templates\ and saves in I:\####\
'                          Arguments:    sQueryName, sExportTopic
'pfSendWordDocAsEmail:         Description:  sends Word document as an e-mail body
'                          Arguments:    vCHTopic, vSubject, Optional sAttachment1, sAttachment2, sAttachment3, sAttachment4
'pfCreateCDLabel:               Description:  makes CD label and prompts for print or no
'                          Arguments:    NONE
'pfSelectCoverTemplate:        Description:  parent function to create correct transcript cover/skeleton from template
'                          Arguments:    NONE
'pfCreateCover:                Description:  creates transcript cover/skeleton from template
'                          Arguments:    sTemplatePath
'fCreatePELLetter:             Description:  creates package enclosed letter
'                          Arguments:    NONE
'fFactorInvoicEmailF:          Description:  creates e-mail to submit invoice to factoring
'                          Arguments:    NONE
'fInfoNeededEmailF:            Description:  creates info needed e-mail
'                          Arguments:    NONE
'pfInvoicesCSV:                Description:  creates CSVs used for invoicing
'                          Arguments:    NONE
'fCreateWorkingCopy:           Description:  creates "working copy" sent to client
'                          Arguments:    NONE
'fSendShippingTrackingEmail:   Description:  creates shipping confirmation e-mail sent to client
'                          Arguments:    NONE
        
'============================================================================

'============================================================================
'module Invoice

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
'AutoCalculateFactorInterest:   Description:  add 1% after every 7 days payment not made
'                           Arguments:    NONE
'DepositPaymentReceived:        Description:  does some things after a deposit is paid
'                           Arguments:    NONE
'IsFactoringApproved:           Description:  checks if factoring is approved for customer
'                           Arguments:    NONE
'GenerateInvoiceNumber:         Description:  generates invoice number
'                           Arguments:    NONE
        
'============================================================================

'============================================================================
'module Stage1:
'variables:
'   NONE

'functions:
'fAssignPS:                                 Description:  prompts to assign file in ProjectSend
'                                       Arguments:    NONE
'pfEnterNewJob:                             Description:  import job info to db from xlsm file
'                                       Arguments:    NONE
'fCheckTempCustomersCustomers:              Description:  retrieve info from TempCustomers/Customers
'                                       Arguments:    NONE
'fCheckTempCasesCases:                      Description:  retrieve info from TempCases/Cases
'                                       Arguments:    NONE
'fInsertCalculatedFieldintoTempCourtDates:  Description:  insert several calculated fields into tempcourtdates
'                                       Arguments:    NONE
'fAudioPlayPromptTyping:                    Description:  prompt to play audio in /Audio/folder
'                                       Arguments:    NONE
'fProcessAudioParent:                       Description:  process audio in express scribe
'                                       Arguments:    NONE
'fPlayAudioParent:                          Description:  play audio as appropriate
'                                       Arguments:    NONE
'fPlayAudioFolder:                          Description:  plays audio folder
'                                       Arguments:    HostFolder
'fProcessAudioFolder:                       Description:  process audio in /Audio/ folder
'                                       Arguments:    HostFolder
'pfPriceQuoteEmail:                         Description:  generates price quote and sends via e-mail
'                                       Arguments:    NONE
'pfStage1Ppwk:                              Description:  completes all stage 1 tasks
'                                       Arguments:    NONE
'fWunderlistAddNewJob:                      Description:  adds new job task list to wunderlist w/ due dates
'                                       Arguments:    NONE
'autointake:                                Description:  process new job email when received
'                                       Arguments:    NONE
'NewOLEntry:                                Description:  checks outlook folder for new job email
'                                       Arguments:    NONE
'ResetDisplay:                              Description:  part of scrolling marquee notification
'                                       Arguments:    NONE
'ScrollingMarquee:                          Description:  scrolling marquee notification for new job
'                                       Arguments:    NONE
'MinimizeNavigationPane:                    Description:  part of scrolling marquee notification
'                                       Arguments:    NONE
'============================================================================

'============================================================================
'module Stage2

'variables:
'   NONE

'functions:

'pfStage2Ppwk:                               Description:  completes all stage 2 tasks
'                                        Arguments:    NONE
'pfAutoCorrect:                              Description:  adds entries as listed on form to rough draft autocorrect in Word
'                                        Arguments:    NONE
'pfRoughDraftToCoverF:                       Description:  Adds rough draft to courtcover
'                                                      does find/replacements of static speakers 1-17, all dynamic speakers, Q&A, : a-z, various AQC & AVT headings
'                                        Arguments:    NONE
'pfStaticSpeakersFindReplace:                Description:  finds and replaces static speakers in CourtCover after rough draft is inserted
'                                        Arguments:    NONE
'pfReplaceColonUndercasewithColonUppercase:  Description:  replaces : a-z with : A-Z, applies styles to fixed phrases in transcript
'                                        Arguments:    NONE
'pfTypeRoughDraftF:                          Description:  copies correct roughdraft template to job folder
'                                        Arguments:    NONE
'pfReplaceWeberOR:                           Description:  Adds rough draft to courtcover
'                                                      does find/replacements of static speakers 1-17, all dynamic speakers, Q&A, : a-z, various AQC & Weber headings
'                                        Arguments:    NONE
'pfReplaceWeberNV:                           Description:  Adds rough draft to courtcover
'                                                      does find/replacements of static speakers 1-17, all dynamic speakers, Q&A, : a-z, various AQC & Weber headings
'                                        Arguments:    NONE
'pfReplaceWeberBR:                           Description:  Adds rough draft to courtcover
'                                                      does find/replacements of static speakers 1-17, all dynamic speakers, Q&A, : a-z, various AQC & Weber headings
'                                        Arguments:    NONE
'pfReplaceAVT:                               Description:  Adds rough draft to courtcover
'                                                      does find/replacements of static speakers 1-17, all dynamic speakers, Q&A, : a-z, various AQC & AVT headings
'                                        Arguments:    NONE
'pfReplaceAQC:                               Description:  Adds rough draft to courtcover
'                                                      does find/replacements of static speakers 1-17, all dynamic speakers, Q&A, : a-z, various AQC & AVT headings
'                                        Arguments:    NONE
        
'============================================================================

'============================================================================
'module Stage3

'variables:
'   NONE

'functions:

'pfStage3Ppwk:        Description:  completes all stage 3 tasks
'                 Arguments:    NONE
'pfBurnCD:            Description:  burns CD to D drive
'                 Arguments:    NONE
'pfCreateRegularPDF:  Description:  creates final PDF of transcript and saves to main/transcripts folders
'                 Arguments:    NONE
'pfHeaders            Description : add sections and headers programmatically
'                 Arguments:    NONE

'============================================================================

'============================================================================
'module Stage4

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
'                        Arguments:    NONE
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
    
'============================================================================

'============================================================================
'module TranscriptFormat

'variables:
'   Private sFileName As String
'   Private oWordApp As Object, oWordDoc As Object
'   Private qdf As QueryDef
'   Private sQueryName As String
'   Private db As Database
'   Public sBookmarkName As String

'functions:

'pfCreateBookmarks:             Description:  replaces #TOC_# notations in transcript with bookmarks and then places index at bookmarks
'                           Arguments:    NONE
'pfReplaceBMKWwithBookmark:     Description:  replaces #__# notations with bookmarks
'                           Arguments:    NONE
'pfApplyStyle:                  Description:  finds specific phrases in activedocument(transcript) and applies a specific style
'                           Arguments:    sStyleName, sTextToFind, sReplacementText
'pfFindRepCitationLinks:        Description:  adds citations and hyperlinks from CitationHyperlinks table in transcript
'                           Arguments:    NONE
'pfCreateIndexesTOAs:           Description:  creates indexes and indexes certain things
'                           Arguments:    NONE
'pfSingleFindReplace:           Description:  find and replace all of one item
'                           Arguments:    sTextToFind, sReplacementText
'                                         Optional wsyWordStyle = "", bForward = True, bWrap = "wdFindContinue"
'                                         Optional bFormat = False, bMatchCase = True, bMatchWholeWord = False
'                                         Optional bMatchWildcards = False, bMatchSoundsLike = False, bMatchAllWordForms = False
'pfReplaceFDA:                  Description:  doctor speaker name find/replaces for FDA transcripts
'                           Arguments:    NONE
'pfDynamicSpeakersFindReplace:  Description:  gets speaker names from ViewJobFormAppearancesQ query and find/replaces in transcript as appropriate
'                           Arguments:    NONE
'pfSingleTCReplaceAll:          Description:  one replace TC entry function for ones with no field entry
'                           Arguments:    sTexttoSearch, sReplacementText
'pfFieldTCReplaceAll:           Description:  one replace TC entry function for ones with field entry
'                           Arguments:    sTexttoSearch, sReplacementText, sFieldText
'pfWordIndexer:                 Description:  builds word index in separate PDF from transcript
'                           Arguments:    NONE
'FPJurors:                      Description:  does find/replacements of prospective jurors in transcript
'                           Arguments:    NONE
'pfTCEntryReplacement:          Description:  parent function that finds certain entries within a transcript and assigns TC entries to them for indexing purposes
'                           Arguments:    NONE
'pfFindRepCitationLinks:        Description:  'originally named fEfficientCiteSearch 'old one now named pfFindRepCitationLinks1
'find citation markings like phonetic in transcript
'list separately so marking doesn't take so long
'                           Arguments:    NONE
        
'============================================================================


'============================================================================
'module Speech

'variables:
'Private URLDownloadToFile as long
    
'functions:

'pfGetFile:                                Description: read binary file as a string value
'                                      Arguments:  sFileName
'pfDoFolder:                               Description: cycles through subfolders of /Prepared/####/ and makes find/replaces listed below
'                                      Arguments:  Folder
'pfIEPostStringRequest:                    Description: sends URL encoded form data To the URL using IE
'                                      Arguments:  sURL, sFormData, sBoundary
'pfUploadFile:                             Description: uploads corpus file to Sphinx site to get LM/DIC files back
'                                                   upload file using input type=file
'                                      Arguments:  sDestinationURL, sCorpusPath, sFieldName
'                                                  Optional sFieldName = "corpus"
'pfCorpusUpload:                           Description: uploads corpus to lmtool at Sphinx to get compatible LM & DIC files back via download
'                                      Arguments:  NONE
'pfDownloadFile:                           Description: downloads provided file
'                                      Arguments:  sURL, sSaveAs
'pfAddSubfolder:                           Description: add concatenatedaudio folder to /UnprocessedAudio/####
'                                      Arguments:  NONE
'pfTrainEngine:                            Description: HOW TO PREPARE FILES FOR TRAINING
'                                      Arguments:  NONE
'pfPrepareAudio:                           Description: runs batch file audioprep / audioprep1.bat in Prepared folder
'                                      Arguments:  NONE
'pfSplitAudio:                             Description: runs batch file splitaudio / audioprep1.bat in Prepared folder
'                                      Arguments:  NONE
'pfRenameBaseFiles:                        Description: runs batch file FileRename / audioprep1.bat in Prepared folder
'                                      Arguments:  NONE
'pfSRTranscribe:                           Description: runs batch file SRTranscribe
'                                      Arguments:  NONE
'pfTrainAudio:                             Description: HOW TO TRAIN AUDIO / run batch file audiotrain.bat
'                                      Arguments:  NONE
'pfCopyTranscriptFromCompletedToPrepared:  Description: copies transcripts from /completed/ folder in T drive to corresponding /prepared/ folder in S drive
'                                      Arguments:  NONE
'pfMultipleAudioOneTranscript:             Description: HOW TO RUN MULTIPLE AUDIO FILES LIST
'                                      Arguments:  NONE
'pfPrepareTranscript:                      Description: makes changes to each transcript so it fits into speech recognition requirements
'                                      Arguments:  NONE
'pfRunCopyTranscTextBAT:                   Description: runs batch file CopyTranscriptTXT
'                                      Arguments:  NONE

'============================================================================



'============================================================================
'module PayPal

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

