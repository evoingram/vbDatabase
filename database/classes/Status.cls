VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Database.Production.Classes")
Option Compare Database
Option Explicit

Private mFactoringApproved As Boolean
Private mContactsEntered As Boolean
Private mJobEntered As Boolean
Private mFixAppears As Boolean
Private mCorrectCDEntries As Boolean
Private mCoverPage As Boolean
Private mAutoCorrect As Boolean
Private mSchedule As Boolean
Private mInvoice As Boolean
Private mAGShortcuts As Boolean
Private mTranscribe As Boolean
Private mAddRDtoCover As Boolean
Private mFindReplaceRD As Boolean
Private mHyperlinkTranscripts As Boolean
Private mSpellingsEmail As Boolean
Private mAudioProof As Boolean
Private mInvoiceCompleted As Boolean
Private mNoticeofService As Boolean
Private mPackageEnclosedLetter As Boolean
Private mCDLabel As Boolean
Private mGenerateZIPs As Boolean
Private mTranscriptsReady As Boolean
Private mInvoicetoFactorEmail As Boolean
Private mFileTranscript As Boolean
Private mBurnCD As Boolean
Private mShippingXMLs As Boolean
Private mGenerateShippingEM As Boolean
Private mAddTrackingNumber As Boolean

'@Ignore EncapsulatePublicField
Public sStatusCID As Long
Private sStatusID As Long

Private mrstRecordset8 As Recordset
Private mbooLoaded As Boolean

Public Property Get ContactsEntered() As Boolean
    ContactsEntered = mContactsEntered
End Property

Public Property Let ContactsEntered(rData As Boolean)
    mContactsEntered = rData
End Property

Public Property Get JobEntered() As Boolean
    JobEntered = mJobEntered
End Property

Public Property Let JobEntered(rData As Boolean)
    mJobEntered = rData
End Property

Public Property Get FixAppears() As Boolean
    FixAppears = mFixAppears
End Property

Public Property Let FixAppears(rData As Boolean)
    mFixAppears = rData
End Property

Public Property Get CorrectCDEntries() As Boolean
    CorrectCDEntries = mCorrectCDEntries
End Property

Public Property Let CorrectCDEntries(rData As Boolean)
    mCorrectCDEntries = rData
End Property

Public Property Get CoverPage() As Boolean
    CoverPage = mCoverPage
End Property

Public Property Let CoverPage(rData As Boolean)
    mCoverPage = rData
End Property

Public Property Get AutoCorrect() As Boolean
    AutoCorrect = mAutoCorrect
End Property

Public Property Let AutoCorrect(rData As Boolean)
    mAutoCorrect = rData
End Property

Public Property Get Schedule() As Boolean
    Schedule = mSchedule
End Property

Public Property Let Schedule(rData As Boolean)
    mSchedule = rData
End Property

Public Property Get Invoice() As Boolean
    Invoice = mInvoice
End Property

Public Property Let Invoice(rData As Boolean)
    mInvoice = rData
End Property

Public Property Get AGShortcuts() As Boolean
    AGShortcuts = mAGShortcuts
End Property

Public Property Let AGShortcuts(rData As Boolean)
    mAGShortcuts = rData
End Property

Public Property Get Transcribe() As Boolean
    Transcribe = mTranscribe
End Property

Public Property Let Transcribe(rData As Boolean)
    mTranscribe = rData
End Property

Public Property Get AddRDtoCover() As Boolean
    AddRDtoCover = mAddRDtoCover
End Property

Public Property Let AddRDtoCover(rData As Boolean)
    mAddRDtoCover = rData
End Property

Public Property Get FindReplaceRD() As Boolean
    FindReplaceRD = mFindReplaceRD
End Property

Public Property Let FindReplaceRD(rData As Boolean)
    mFindReplaceRD = rData
End Property

Public Property Get HyperlinkTranscripts() As Boolean
    HyperlinkTranscripts = mHyperlinkTranscripts
End Property

Public Property Let HyperlinkTranscripts(rData As Boolean)
    mHyperlinkTranscripts = rData
End Property

Public Property Get SpellingsEmail() As Boolean
    SpellingsEmail = mSpellingsEmail
End Property

Public Property Let SpellingsEmail(rData As Boolean)
    mSpellingsEmail = rData
End Property

Public Property Get AudioProof() As Boolean
    AudioProof = mAudioProof
End Property

Public Property Let AudioProof(rData As Boolean)
    mAudioProof = rData
End Property

Public Property Get InvoiceCompleted() As Boolean
    InvoiceCompleted = mInvoiceCompleted
End Property

Public Property Let InvoiceCompleted(rData As Boolean)
    mInvoiceCompleted = rData
End Property

Public Property Get NoticeofService() As Boolean
    NoticeofService = mNoticeofService
End Property

Public Property Let NoticeofService(rData As Boolean)
    mNoticeofService = rData
End Property

Public Property Get PackageEnclosedLetter() As Boolean
    PackageEnclosedLetter = mPackageEnclosedLetter
End Property

Public Property Let PackageEnclosedLetter(rData As Boolean)
    mPackageEnclosedLetter = rData
End Property

Public Property Get CDLabel() As Boolean
    CDLabel = mCDLabel
End Property

Public Property Let CDLabel(rData As Boolean)
    mCDLabel = rData
End Property

Public Property Get GenerateZIPs() As Boolean
    GenerateZIPs = mGenerateZIPs
End Property

Public Property Let GenerateZIPs(rData As Boolean)
    mGenerateZIPs = rData
End Property

Public Property Get TranscriptsReady() As Boolean
    TranscriptsReady = mTranscriptsReady
End Property

Public Property Let TranscriptsReady(rData As Boolean)
    mTranscriptsReady = rData
End Property

Public Property Get InvoicetoFactorEmail() As Boolean
    InvoicetoFactorEmail = mInvoicetoFactorEmail
End Property

Public Property Let InvoicetoFactorEmail(rData As Boolean)
    mInvoicetoFactorEmail = rData
End Property

Public Property Get FileTranscript() As Boolean
    FileTranscript = mFileTranscript
End Property

Public Property Let FileTranscript(rData As Boolean)
    mFileTranscript = rData
End Property

Public Property Get BurnCD() As Boolean
    BurnCD = mBurnCD
End Property

Public Property Let BurnCD(rData As Boolean)
    mBurnCD = rData
End Property

Public Property Get ShippingXMLs() As Boolean
    ShippingXMLs = mShippingXMLs
End Property

Public Property Let ShippingXMLs(rData As Boolean)
    mShippingXMLs = rData
End Property

Public Property Get GenerateShippingEM() As Boolean
    GenerateShippingEM = mGenerateShippingEM
End Property

Public Property Let GenerateShippingEM(rData As Boolean)
    mGenerateShippingEM = rData
End Property

Public Property Get AddTrackingNumber() As Boolean
    AddTrackingNumber = mAddTrackingNumber
End Property

Public Property Let AddTrackingNumber(rData As Boolean)
    mAddTrackingNumber = rData
End Property

Public Property Get ID() As Long
    ID = sStatusID
End Property

Public Property Get CID() As Long
    CID = sStatusCID
End Property

Private Property Get RecordsetA8() As Recordset
    Set RecordsetA8 = mrstRecordset8
End Property

Private Property Set RecordsetA8(rData As Recordset)
    Set mrstRecordset8 = rData
End Property

Private Sub Load()
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    With RecordsetA8
        sStatusID = Nz(.Fields("ID").Value)
        sStatusCID = Nz(.Fields("CourtDatesID").Value)
        Me.ContactsEntered = Nz(.Fields("ContactsEntered").Value)
        Me.JobEntered = Nz(.Fields("JobEntered").Value)
        Me.FixAppears = Nz(.Fields("FixAppears").Value)
        Me.CorrectCDEntries = Nz(.Fields("CorrectCDEntries").Value)
        Me.CoverPage = Nz(.Fields("CoverPage").Value)
        Me.AutoCorrect = Nz(.Fields("AutoCorrect").Value)
        Me.Schedule = Nz(.Fields("Schedule").Value)
        Me.Invoice = Nz(.Fields("Invoice").Value)
        Me.AGShortcuts = Nz(.Fields("AGShortcuts").Value)
        Me.Transcribe = Nz(.Fields("Transcribe").Value)
        Me.AddRDtoCover = Nz(.Fields("AddRDtoCover").Value)
        Me.FindReplaceRD = Nz(.Fields("FindReplaceRD").Value)
        Me.HyperlinkTranscripts = Nz(.Fields("HyperlinkTranscripts").Value)
        Me.SpellingsEmail = Nz(.Fields("SpellingsEmail").Value)
        Me.AudioProof = Nz(.Fields("AudioProof").Value)
        Me.InvoiceCompleted = Nz(.Fields("InvoiceCompleted").Value)
        Me.NoticeofService = Nz(.Fields("NoticeofService").Value)
        Me.PackageEnclosedLetter = Nz(.Fields("PackageEnclosedLetter").Value)
        Me.CDLabel = Nz(.Fields("CDLabel").Value)
        Me.GenerateZIPs = Nz(.Fields("GenerateZIPs").Value)
        Me.TranscriptsReady = Nz(.Fields("TranscriptsReady").Value)
        Me.InvoicetoFactorEmail = Nz(.Fields("InvoicetoFactorEmail").Value)
        Me.FileTranscript = Nz(.Fields("FileTranscript").Value)
        Me.BurnCD = Nz(.Fields("BurnCD").Value)
        Me.ShippingXMLs = Nz(.Fields("ShippingXMLs").Value)
        Me.GenerateShippingEM = Nz(.Fields("GenerateShippingEM").Value)
        Me.AddTrackingNumber = Nz(.Fields("AddTrackingNumber").Value)
    End With
    mbooLoaded = True
End Sub

Public Sub Update()
    sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
    With RecordsetA8
        If mbooLoaded = True Then
            .Edit
        Else
            .AddNew
        End If
        Nz(.Fields("CourtDatesID").Value) = sCourtDatesID
        sStatusID = Nz(.Fields("ID").Value)
        sStatusCID = Nz(.Fields("CourtDatesID").Value)
        .Fields("AddRDtoCover").Value = Me.AddRDtoCover
        .Fields("ContactsEntered").Value = Me.ContactsEntered
        .Fields("JobEntered").Value = Me.JobEntered
        .Fields("FixAppears").Value = Me.FixAppears
        .Fields("CorrectCDEntries").Value = Me.CorrectCDEntries
        .Fields("CoverPage").Value = Me.CoverPage
        .Fields("AutoCorrect").Value = Me.AutoCorrect
        .Fields("Schedule").Value = Me.Schedule
        .Fields("Invoice").Value = Me.Invoice
        .Fields("AGShortcuts").Value = Me.AGShortcuts
        .Fields("Transcribe").Value = Me.Transcribe
        .Fields("FindReplaceRD").Value = Me.FindReplaceRD
        .Fields("HyperlinkTranscripts").Value = Me.HyperlinkTranscripts
        .Fields("SpellingsEmail").Value = Me.SpellingsEmail
        .Fields("AudioProof").Value = Me.AudioProof
        .Fields("InvoiceCompleted").Value = Me.InvoiceCompleted
        .Fields("NoticeofService").Value = Me.NoticeofService
        .Fields("PackageEnclosedLetter").Value = Me.PackageEnclosedLetter
        .Fields("CDLabel").Value = Me.CDLabel
        .Fields("GenerateZIPs").Value = Me.GenerateZIPs
        .Fields("TranscriptsReady").Value = Me.TranscriptsReady
        .Fields("InvoicetoFactorEmail").Value = Me.InvoicetoFactorEmail
        .Fields("FileTranscript").Value = Me.FileTranscript
        .Fields("BurnCD").Value = Me.BurnCD
        .Fields("ShippingXMLs").Value = Me.ShippingXMLs
        .Fields("GenerateShippingEM").Value = Me.GenerateShippingEM
        .Fields("AddTrackingNumber").Value = Me.AddTrackingNumber
        .Update
    End With
    mbooLoaded = True
End Sub

Public Sub AddNew()
    mbooLoaded = False
End Sub

'@Ignore FunctionReturnValueNotUsed
Public Function FindFirst(Optional Criteria As Variant) As Boolean
    If IsMissing(Criteria) Then
        RecordsetA8.MoveFirst
        FindFirst = Not RecordsetA8.EOF
    Else
        RecordsetA8.FindFirst Criteria
        FindFirst = Not RecordsetA8.NoMatch
    End If
    If FindFirst Then Load
End Function

Private Sub Class_Initialize()
    Set RecordsetA8 = CurrentDb.OpenRecordset("Statuses", dbOpenDynaset)
    
End Sub

Private Sub Class_Terminate()
    RecordsetA8.Close
    Set RecordsetA8 = Nothing
End Sub

Private Function NullIfEmptyString(str As String) As Variant
    Dim strTrimmed As String: strTrimmed = Trim(str)
    If Len(strTrimmed) = 0 Then
        NullIfEmptyString = Null
    Else
        NullIfEmptyString = strTrimmed
    End If
End Function

