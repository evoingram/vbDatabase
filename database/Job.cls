VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Job"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private sCourtDatesID As Long, dHearingStartTime As Date, dHearingEndTime As Date
Private mApp1 As App1
Public sApp1 As Long

Public lngNumOfHrs As Long, lngNumOfMins As Long, lngNumOfSecsRem As Long, lngNumOfSecs As Long
Public lngNumOfHrs1 As Long, lngNumOfMins1 As Long, lngNumOfSecsRem1 As Long, lngNumOfSecs1 As Long
Public i As Long

Public oWordApp As Word.Document, oWordDoc As Word.Application
Public lAssigneeID As Long, sDueDate As String, bStarred As String, bCompleted As String
Public sTitle As String, sWLListID As String

Private mrstRecordset As Recordset
Private mbooLoaded As Boolean

Public Property Get Party1() As String
    Party1 = sParty1
End Property

Public Property Let Party1(rData As String)
    sParty1 = rData
End Property

Public Property Get Party1Name() As String
    Party1Name = sParty1Name
End Property

Public Property Let Party1Name(rData As String)
    sParty1Name = rData
End Property

Public Property Get Party2() As String
    Party2 = sParty2
End Property

Public Property Let Party2(rData As String)
    sParty2 = rData
End Property

Public Property Get Party2Name() As String
    Party2Name = sParty2Name
End Property

Public Property Let Party2Name(rData As String)
    sParty2Name = rData
End Property


Public Property Get InvoiceNo() As String
    InvoiceNo = sInvoiceNo
End Property

Public Property Let InvoiceNo(rData As String)
    sInvoiceNo = rData
End Property

Public Property Get Email() As String
    Email = sEmail
End Property

Public Property Let Email(rData As String)
    sEmail = rData
End Property

Public Property Get Description() As String
    Description = sDescription
End Property

Public Property Let Description(rData As String)
    sDescription = rData
End Property

Public Property Get Subtotal() As String
    Subtotal = sSubtotal
End Property

Public Property Let Subtotal(rData As String)
    sSubtotal = rData
End Property

Public Property Get InvoiceDate() As String
    InvoiceDate = sInvoiceDate
End Property

Public Property Let InvoiceDate(rData As String)
    sInvoiceDate = rData
End Property

Public Property Get InvoiceTime() As String
    InvoiceTime = sInvoiceTime
End Property

Public Property Let InvoiceTime(rData As String)
    sInvoiceTime = rData
End Property

Public Property Get PaymentTerms() As String
    PaymentTerms = sPaymentTerms
End Property

Public Property Let PaymentTerms(rData As String)
    sPaymentTerms = rData
End Property

Public Property Get Note() As String
    Note = sNote
End Property

Public Property Let Note(rData As String)
    sNote = rData
End Property

Public Property Get Terms() As String
    Terms = sTerms
End Property

Public Property Let Terms(rData As String)
    sTerms = rData
End Property

Public Property Get MinimumAmount() As String
    MinimumAmount = sMinimumAmount
End Property

Public Property Let MinimumAmount(rData As String)
    sMinimumAmount = rData
End Property

Public Property Get Memo() As String
    Memo = vmMemo
End Property

Public Property Let Memo(rData As String)
    vmMemo = rData
End Property

Public Property Get URL() As String
    URL = vlURL
End Property

Public Property Let URL(rData As String)
    vlURL = rData
End Property

Public Property Get TemplateID() As String
    TemplateID = sTemplateID
End Property

Public Property Let TemplateID(rData As String)
    sTemplateID = rData
End Property

Public Property Get Line1() As String
    Line1 = sLine1
End Property

Public Property Let Line1(rData As String)
    sLine1 = rData
End Property

Public Property Get City() As String
    City = sCity
End Property

Public Property Let City(rData As String)
    sCity = rData
End Property

Public Property Get State() As String
    State = sState
End Property

Public Property Let State(rData As String)
    sState = rData
End Property

Public Property Get ZIP() As String
    ZIP = sZIP
End Property

Public Property Let ZIP(rData As String)
    sZIP = rData
End Property

Public Property Get Quantity() As String
    Quantity = sQuantity
End Property

Public Property Let Quantity(rData As String)
    sQuantity = rData
End Property

Public Property Get Value() As String
    Value = sValue
End Property

Public Property Let Value(rData As String)
    sValue = rData
End Property

Public Property Get InventoryRateCode() As String
    InventoryRateCode = sInventoryRateCode
End Property

Public Property Let InventoryRateCode(rData As String)
    sInventoryRateCode = rData
End Property

Public Property Get IRC() As String
    IRC = sIRC
End Property

Public Property Let IRC(rData As String)
    sIRC = rData
End Property

Public Property Get CaseNumber2() As String
    CaseNumber2 = sCaseNumber2
End Property

Public Property Let CaseNumber2(rData As String)
    sCaseNumber2 = rData
End Property

Public Property Get CaseNumber1() As String
    CaseNumber1 = sCaseNumber1
End Property

Public Property Let CaseNumber1(rData As String)
    sCaseNumber1 = rData
End Property

Public Property Get ActualQuantity() As String
    ActualQuantity = sActualQuantity
End Property

Public Property Let ActualQuantity(rData As String)
    sActualQuantity = rData
End Property

Public Property Get Jurisdiction() As String
    Jurisdiction = sJurisdiction
End Property

Public Property Let Jurisdiction(rData As String)
    sJurisdiction = rData
End Property

Public Property Get TurnaroundTime() As String
    TurnaroundTime = sTurnaroundTime
End Property

Public Property Let TurnaroundTime(rData As String)
    sTurnaroundTime = rData
End Property


Public Property Get AudioLength() As String
    AudioLength = sAudioLength
End Property

Public Property Let AudioLength(rData As String)
    sAudioLength = rData
End Property

Public Property Get EstimatedPageCount() As String
    EstimatedPageCount = sEstimatedPageCount
End Property

Public Property Let EstimatedPageCount(rData As String)
    sEstimatedPageCount = rData
End Property

Public Property Get StatusesID() As String
    StatusesID = sStatusesID
End Property

Public Property Let StatusesID(rData As String)
    sStatusesID = rData
End Property

Public Property Get DueDate() As Date
    DueDate = dDueDate
End Property

Public Property Let DueDate(rData As Date)
    dDueDate = rData
End Property

Public Property Get ExpectedAdvanceDate() As Date
    ExpectedAdvanceDate = dExpectedAdvanceDate
End Property

Public Property Let ExpectedAdvanceDate(rData As Date)
    dExpectedAdvanceDate = rData
End Property

Public Property Get ExpectedRebateDate() As Date
    ExpectedRebateDate = dExpectedRebateDate
End Property

Public Property Let ExpectedRebateDate(rData As Date)
    dExpectedRebateDate = rData
End Property

Public Property Get PaymentSum() As String
    PaymentSum = sPaymentSum
End Property

Public Property Let PaymentSum(rData As String)
    sPaymentSum = rData
End Property

Public Property Get FactoringApproved() As String
    FactoringApproved = sFactoringApproved
End Property

Public Property Let FactoringApproved(rData As String)
    sFactoringApproved = rData
End Property

Public Property Get BrandingTheme() As String
    BrandingTheme = sBrandingTheme
End Property

Public Property Let BrandingTheme(rData As String)
    sBrandingTheme = rData
End Property

Public Property Get FinalPrice() As String
    FinalPrice = sFinalPrice
End Property

Public Property Let FinalPrice(rData As String)
    sFinalPrice = rData
End Property

Public Property Get ClientTranscriptName() As String
    ClientTranscriptName = sClientTranscriptName
End Property

Public Property Let ClientTranscriptName(rData As String)
    sClientTranscriptName = rData
End Property

Public Property Get CurrentTranscriptName() As String
    CurrentTranscriptName = sCurrentTranscriptName
End Property

Public Property Let CurrentTranscriptName(rData As String)
    sCurrentTranscriptName = rData
End Property

Public Property Get Status() As String
    Status = vStatus
End Property

Public Property Let Status(rData As String)
    vStatus = rData
End Property

Public Property Get UnitPrice() As String
    UnitPrice = sUnitPrice
End Property

Public Property Let UnitPrice(rData As String)
    sUnitPrice = rData
End Property

Public Property Get BalanceDue() As String
    BalanceDue = sBalanceDue
End Property

Public Property Let BalanceDue(rData As String)
    sBalanceDue = rData
End Property

Public Property Get FactoringCost() As String
    FactoringCost = sFactoringCost
End Property

Public Property Let FactoringCost(rData As String)
    sFactoringCost = rData
End Property

Public Property Get vURL() As String
    vURL = svURL
End Property

Public Property Let vURL(rData As String)
    svURL = rData
End Property

Public Property Get LinkToCSV() As String
    LinkToCSV = sLinkToCSV
End Property

Public Property Let LinkToCSV(rData As String)
    sLinkToCSV = rData
End Property

Public Property Get FirstName() As String
    FirstName = sFirstName
End Property

Public Property Let FirstName(rData As String)
    sFirstName = rData
End Property

Public Property Get LastName() As String
    LastName = sLastName
End Property

Public Property Let LastName(rData As String)
    sLastName = rData
End Property

Public Property Get HearingDate() As Date
    HearingDate = dHearingDate
End Property

Public Property Let HearingDate(rData As Date)
    dHearingDate = rData
End Property


Public Property Get HearingStartTime() As Date
    HearingStartTime = dHearingStartTime
End Property

Public Property Let HearingStartTime(rData As Date)
    dHearingStartTime = rData
End Property

Public Property Get HearingEndTime() As Date
    HearingEndTime = dHearingEndTime
End Property

Public Property Let HearingEndTime(rData As Date)
    dHearingEndTime = rData
End Property

Public Property Get MrMs() As String
    MrMs = sMrMs
End Property

Public Property Let MrMs(rData As String)
    sMrMs = rData
End Property


Public Property Get Name() As String
    Name = sName
End Property

Public Property Let Name(rData As String)
    sName = rData
End Property

Public Property Get Address1() As String
    Address1 = sAddress1
End Property

Public Property Let Address1(rData As String)
    sAddress1 = rData
End Property

Public Property Get Address2() As String
    Address2 = sAddress2
End Property

Public Property Let Address2(rData As String)
    sAddress2 = rData
End Property

Public Property Get Notes() As String
    Notes = sNotes
End Property

Public Property Let Notes(rData As String)
    sNotes = rData
End Property

Public Property Get Time() As String
    Time = sTime
End Property

Public Property Let Time(rData As String)
    sTime = rData
End Property

Public Property Get Time1() As String
    Time1 = sTime1
End Property

Public Property Let Time1(rData As String)
    sTime1 = rData
End Property

Public Property Get CasesID() As String
    CasesID = sCasesID
End Property

Public Property Let CasesID(rData As String)
    sCasesID = rData
End Property

Public Property Get CasesID1() As String
    CasesID1 = vCasesID
End Property

Public Property Let CasesID1(rData As String)
    vCasesID = rData
End Property

Public Property Get HearingLocation() As String
    HearingLocation = sHearingLocation
End Property

Public Property Let HearingLocation(rData As String)
    sHearingLocation = rData
End Property

Public Property Get URL1() As String
    URL1 = sURL1
End Property

Public Property Let URL1(rData As String)
    sURL1 = rData
End Property

Public Property Get ID() As Long
    ID = sCourtDatesID
End Property

Public Property Get App1() As App1
    'Private mApp1 As App1
    If (mApp1 Is Nothing) Then Set App1 = New App1
    Set App1 = mApp1
End Property



Private Property Get Recordset() As Recordset
    Set Recordset = mrstRecordset
End Property

Private Property Set Recordset(rData As Recordset)
    Set mrstRecordset = rData
End Property


Private Sub Load()
    With Recordset
        'Me.BirthDate = Nz(.Fields("BirthDate").Value)
        'Me.EmailAddress = Nz(.Fields("EmailAddress").Value)
        'Me.FirstName = Nz(.Fields("FirstName").Value)
        'mlngID = Nz(.Fields("ID").Value)
        'Me.LastName = Nz(.Fields("LastName").Value)
        'Me.Score = Nz(.Fields("Score").Value)
        'sApp0 = Nz(.Fields("OrderingID").Value)
        sApp1 = Nz(.Fields("App1").Value)
        'sApp2 = Nz(.Fields("App2").Value)
        'sApp3 = Nz(.Fields("App3").Value)
        'sApp4 = Nz(.Fields("App4").Value)
        'sApp5 = Nz(.Fields("App5").Value)
        'sApp6 = Nz(.Fields("App6").Value)
        Me.InvoiceNo = .Fields("InvoiceNo").Value
        'Me.Party1 = Nz(.Fields("Party1").Value)
        'Me.Party2 = Nz(.Fields("Party2").Value)
        'Me.Party1Name = Nz(.Fields("Party1Name").Value)
        'Me.Party2Name = Nz(.Fields("Party2Name").Value)
        'Me.Company = Nz(.Fields("Company").Value)
        'Me.Email = Nz(.Fields("Email").Value)
        'Me.Description = Nz(.Fields("Description").Value)
        Me.Subtotal = Nz(.Fields("Subtotal").Value)
        Me.InvoiceDate = Nz(.Fields("InvoiceDate").Value)
        'Me.InvoiceTime = Nz(.Fields("InvoiceTime").Value)
        'Me.PaymentTerms = Nz(.Fields("PaymentTerms").Value)
        'Me.Note = Nz(.Fields("Note").Value)
        'Me.Terms = Nz(.Fields("Terms").Value)
        'Me.MinimumAmount = Nz(.Fields("MinimumAmount").Value)
        'Me.Memo = Nz(.Fields("Memo").Value)
        'Me.URL = Nz(.Fields("URL").Value)
        'Me.TemplateID = Nz(.Fields("TemplateID").Value)
        'Me.Line1 = Nz(.Fields("Line").Value)
        'Me.City = Nz(.Fields("City").Value)
        'Me.State = Nz(.Fields("State").Value)
        'Me.ZIP = Nz(.Fields("ZIP").Value)
        Me.Quantity = Nz(.Fields("Quantity").Value)
        'Me.Value = Nz(.Fields("Value").Value)
        Me.InventoryRateCode = Nz(.Fields("InventoryRateCode").Value)
        'Me.IRC = Nz(.Fields("IRC").Value)
        'Me.CaseNumber1 = Nz(.Fields("CaseNumber1").Value)
        'Me.CaseNumber2 = Nz(.Fields("CaseNumber2").Value)
        Me.ActualQuantity = Nz(.Fields("ActualQuantity").Value)
        'Me.Jurisdiction = Nz(.Fields("Jurisdiction").Value)
        Me.TurnaroundTime = .Fields("TurnaroundTimesCD").Value
        'Me.OrderingID = Nz(.Fields("OrderingID").Value)
        Me.AudioLength = Nz(.Fields("AudioLength").Value)
        Me.EstimatedPageCount = Nz(.Fields("EstimatedPageCount").Value)
        Me.StatusesID = Nz(.Fields("StatusesID").Value)
        Me.DueDate = Nz(.Fields("DueDate").Value)
        Me.ExpectedAdvanceDate = Nz(.Fields("ExpectedAdvanceDate").Value)
        Me.ExpectedRebateDate = Nz(.Fields("ExpectedRebateDate").Value)
        Me.PaymentSum = Nz(.Fields("PaymentSum").Value)
        'Me.FactoringApproved = Nz(.Fields("FactoringApproved").Value)
        Me.BrandingTheme = Nz(.Fields("BrandingTheme").Value)
        Me.FinalPrice = Nz(.Fields("FinalPrice").Value)
        'Me.ClientTranscriptName = Nz(.Fields("ClientTranscriptName").Value)
        'Me.CurrentTranscriptName = Nz(.Fields("CurrentTranscriptName").Value)
        'Me.BalanceDue = Nz(.Fields("BalanceDue").Value)
        Me.FactoringCost = Nz(.Fields("FactoringCost").Value)
        'Me.URL1 = Nz(.Fields("URL1").Value)
        'Me.LinkToCSV = Nz(.Fields("LinkToCSV").Value)
        'Me.FirstName = Nz(.Fields("FirstName").Value)
        'Me.LastName = Nz(.Fields("LastName").Value)
        Me.HearingDate = Nz(.Fields("HearingDate").Value)
        Me.HearingStartTime = .Fields("HearingStartTime").Value
        Me.HearingEndTime = .Fields("HearingEndTime").Value
        'Me.MrMs = Nz(.Fields("MrMs").Value)
        'Me.Name = Nz(.Fields("Name").Value)
        'Me.Address1 = Nz(.Fields("Address1").Value)
        'Me.Address2 = Nz(.Fields("Address2").Value)
        'Me.Notes = Nz(.Fields("Notes").Value)
        'Me.Time = Nz(.Fields("Time").Value)
        'Me.Time1 = Nz(.Fields("Time1").Value)
        Me.CasesID = Nz(.Fields("CasesID").Value)
        'Me.CasesID1 = Nz(.Fields("CasesID1").Value)
        'Me.Status = Nz(.Fields("Status").Value)
        Me.UnitPrice = Nz(.Fields("UnitPrice").Value)
        'Me.HearingLocation = Nz(.Fields("HearingLocation").Value)
        'Me. = Nz(.Fields("").Value)
        sCourtDatesID = .Fields("ID").Value
    End With
    mbooLoaded = True
    mApp1.FindFirst "ID=" & sApp1
    'mApp1.SetAppValues (sApp1)
End Sub

Public Sub Update()
    With Recordset
        If mbooLoaded = True Then
            .Edit
        Else
            .AddNew
        End If
        sCourtDatesID = Nz(.Fields("ID").Value)
        'sApp0 = Nz(.Fields("OrderingID").Value)
        sApp1 = Nz(.Fields("App1").Value)
        'sApp2 = Nz(.Fields("App2").Value)
        'sApp3 = Nz(.Fields("App3").Value)
        'sApp4 = Nz(.Fields("App4").Value)
        'sApp5 = Nz(.Fields("App5").Value)
        'sApp6 = Nz(.Fields("App6").Value)
        '.Fields("BirthDate").Value = Me.BirthDate
        '.Fields("EmailAddress").Value = Me.EmailAddress
        '.Fields("FirstName").Value = Me.FirstName
        '.Fields("LastName").Value = Me.LastName
        .Fields("Party1").Value = Me.Party1
        .Fields("Party2").Value = Me.Party2
        .Fields("Party1Name").Value = Me.Party1Name
        .Fields("Party2Name").Value = Me.Party2Name
        '.Fields("Company").Value = Me.Company
        .Fields("InvoiceNo").Value = Me.InvoiceNo
        '.Fields("Email").Value = Me.Email
        '.Fields("Description").Value = Me.Description
        .Fields("Subtotal").Value = Me.Subtotal
        .Fields("InvoiceDate").Value = Me.InvoiceDate
        '.Fields("InvoiceTime").Value = Me.InvoiceTime
        '.Fields("PaymentTerms").Value = Me.PaymentTerms
        '.Fields("Note").Value = Me.Note
        '.Fields("Terms").Value = Me.Terms
        '.Fields("MinimumAmount").Value = Me.MinimumAmount
        '.Fields("Memo").Value = Me.Memo
        '.Fields("URL").Value = Me.URL
        '.Fields("TemplateID").Value = Me.TemplateID
        '.Fields("Line").Value = Me.Line1
        '.Fields("City").Value = Me.City
        '.Fields("State").Value = Me.State
        '.Fields("ZIP").Value = Me.ZIP
        .Fields("Quantity").Value = Me.Quantity
        '.Fields("Value").Value = Me.Value
        .Fields("InventoryRateCode").Value = Me.InventoryRateCode
        '.Fields("IRC").Value = Me.IRC
        .Fields("CaseNumber1").Value = Me.CaseNumber1
        .Fields("CaseNumber2").Value = Me.CaseNumber2
        .Fields("ActualQuantity").Value = Me.ActualQuantity
        .Fields("Jurisdiction").Value = Me.Jurisdiction
        .Fields("TurnaroundTimesCD").Value = Me.TurnaroundTime
        .Fields("OrderingID").Value = Me.OrderingID
        .Fields("AudioLength").Value = Me.AudioLength
        .Fields("EstimatedPageCount").Value = Me.EstimatedPageCount
        .Fields("StatusesID").Value = Me.StatusesID
        .Fields("DueDate").Value = Me.DueDate
        .Fields("ExpectedAdvanceDate").Value = Me.ExpectedAdvanceDate
        .Fields("ExpectedRebateDate").Value = Me.ExpectedRebateDate
        .Fields("PaymentSum").Value = Me.PaymentSum
        '.Fields("FactoringApproved").Value = Me.FactoringApproved
        .Fields("BrandingTheme").Value = Me.BrandingTheme
        .Fields("FinalPrice").Value = Me.FinalPrice
        '.Fields("ClientTranscriptName").Value = Me.ClientTranscriptName
        '.Fields("CurrentTranscriptName").Value = Me.CurrentTranscriptName
        '.Fields("BalanceDue").Value = Me.BalanceDue
        '.Fields("FactoringCost").Value = Me.FactoringCost
        '.Fields("URL1").Value = Me.URL1
        '.Fields("LinkToCSV").Value = Me.LinkToCSV
        '.Fields("FirstName").Value = Me.FirstName
        '.Fields("LastName").Value = Me.LastName
        .Fields("HearingDate").Value = Me.HearingDate
        '.Fields("MrMs").Value = Me.MrMs
        '.Fields("Name").Value = Me.Name
        '.Fields("Address1").Value = Me.Address1
        '.Fields("Address2").Value = Me.Address2
        '.Fields("Notes").Value = Me.Notes
        '.Fields("Time").Value = Me.Time
        '.Fields("Time1").Value = Me.Time1
        .Fields("CasesID").Value = Me.CasesID
        '.Fields("CasesID1").Value = Me.CasesID1
        '.Fields("Status").Value = Me.Status
        .Fields("UnitPrice").Value = Me.UnitPrice
        '.Fields("HearingLocation").Value = Me.HearingLocation
        .Fields("HearingStartTime").Value = Me.HearingStartTime
        .Fields("HearingEndTime").Value = Me.HearingEndTime
        .Update
    End With
    mbooLoaded = True
End Sub

Public Sub AddNew()
    mbooLoaded = False
End Sub

Public Function FindFirst(Optional Criteria As Variant) As Boolean
    If IsMissing(Criteria) Then
        Recordset.MoveFirst
        FindFirst = Not Recordset.EOF
    Else
        Recordset.FindFirst Criteria
        FindFirst = Not Recordset.NoMatch
    End If
    If FindFirst Then Load
End Function

Private Sub Class_Initialize()
    Set Recordset = CurrentDb.OpenRecordset("CourtDates", dbOpenDynaset)
    'Set mApp0 = New App0
    'sApp1 = Nz(Recordset.Fields("App1").Value)
    'sApp2 = Nz(.Fields("App2").Value)
    'sApp3 = Nz(.Fields("App3").Value)
    'sApp4 = Nz(.Fields("App4").Value)
    'sApp5 = Nz(.Fields("App5").Value)
    'sApp6 = Nz(.Fields("App6").Value)
    Set mApp1 = New App1
    'Set mApp2 = New App2
    'Set mApp3 = New App3
    'Set mApp4 = New App4
    'Set mApp5 = New App5
    'Set mApp6 = New App6
    
    '
'Public Property Get App1() As App1
    'Private mApp1 As App1
    'App1 = mApp1
'End Property
    'mApp1.FindFirst "ID=" & sApp1
End Sub

Private Sub Class_Terminate()
    Recordset.Close
    Set Recordset = Nothing
    Set mApp1 = Nothing
End Sub

Function NullIfEmptyString(str As String) As Variant
    Dim strTrimmed As String: strTrimmed = Trim(str)
    If Len(strTrimmed) = 0 Then
        NullIfEmptyString = Null
    Else
        NullIfEmptyString = strTrimmed
    End If
End Function



Public Sub SetValues(ByVal sApp1 As String)
', ByVal sCompany As String, ByVal sMrMs As String, ByVal sLastName As String, _
ByVal sFirstName As String, ByVal sEmailAddress As String, ByVal sJobTitle As String, _
ByVal sBusinessPhone As String, ByVal sAddress As String, ByVal sCity As String, _
sZIP As String, ByVal sState As String, ByVal sNotes As String
    ', ByVal sFactoringApproved As CheckBox
    mApp1.SetAppValues sApp1 ', sCompany, sMrMs, sLastName, sFirstName, sEmailAddress, sJobTitle, _
    sBusinessPhone, sAddress, sCity, sZIP, sState, sNotes ', sFactoringApproved
End Sub






















