VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Job"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Database.Production.Classes")
Option Compare Database
Option Explicit


Private sCourtDatesID As Long
Private dHearingStartTime As Date
Private dHearingEndTime As Date
Private mApp0 As App0
Private mApp1 As App1
Private mApp2 As App2
Private mApp3 As App3
Private mApp4 As App4
Private mApp5 As App5
Private mApp6 As App6
Private mCaseInfo As CaseInfo
Private mStatus As Status
Public sApp0 As Long
Public sApp1 As Long
Public sApp2 As Long
Public sApp3 As Long
Public sApp4 As Long
Public sApp5 As Long
Public sApp6 As Long
Public sCaseID As Long
Public sStatusID As Long
Public sStatusCID As Long

Public lngNumOfHrs As Long
Public lngNumOfMins As Long
Public lngNumOfSecsRem As Long
Public lngNumOfSecs As Long
Public lngNumOfHrs1 As Long
Public lngNumOfMins1 As Long
Public lngNumOfSecsRem1 As Long
Public lngNumOfSecs1 As Long
Public i As Long

Public oWordApp As Word.Document
Public oWordDoc As Word.Application
Public lAssigneeID As Long
Public sDueDate As String
Public bStarred As String
Public bCompleted As String
Public sTitle As String
Public sWLListID As String
Private sPPID As String

Private mrstRecordset As Recordset
Private mbooLoaded As Boolean

Public Property Get InvoiceNo() As String
    InvoiceNo = sInvoiceNo
End Property

Public Property Let InvoiceNo(rData As String)
    sInvoiceNo = rData
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

Public Property Get Quantity() As String
    Quantity = sQuantity
End Property

Public Property Let Quantity(rData As String)
    sQuantity = rData
End Property

Public Property Get InventoryRateCode() As String
    InventoryRateCode = sInventoryRateCode
End Property

Public Property Let InventoryRateCode(rData As String)
    sInventoryRateCode = rData
End Property

Public Property Get ActualQuantity() As String
    ActualQuantity = sActualQuantity
End Property

Public Property Let ActualQuantity(rData As String)
    sActualQuantity = rData
End Property

Public Property Get TurnaroundTime() As String
    TurnaroundTime = sTurnaroundTime
End Property

Public Property Let TurnaroundTime(rData As String)
    sTurnaroundTime = rData
End Property

Public Property Get Location() As String
    Location = sLocation
End Property

Public Property Let Location(rData As String)
    sLocation = rData
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

Public Property Get UnitPrice() As String
    UnitPrice = sUnitPrice
End Property

Public Property Let UnitPrice(rData As String)
    sUnitPrice = rData
End Property

Public Property Get FactoringCost() As String
    FactoringCost = sFactoringCost
End Property

Public Property Let FactoringCost(rData As String)
    sFactoringCost = rData
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

Public Property Get CaseID() As String
    CaseID = sCaseID
End Property

Public Property Let CaseID(rData As String)
    sCaseID = rData
End Property

Public Property Get PPID() As String
    PPID = sPPID
End Property

Public Property Let PPID(rData As String)
    sPPID = rData
End Property

Public Property Get ID() As Long
    ID = sCourtDatesID
End Property

Public Property Get StatusID() As Long
    StatusID = sStatusID
End Property

Public Property Get Status() As Status
    'Private mApp1 As App1
    If (mStatus Is Nothing) Then Set Status = New Status
    Set Status = mStatus
End Property

Public Property Get CaseInfo() As CaseInfo
    'Private mApp1 As App1
    If (mCaseInfo Is Nothing) Then Set CaseInfo = New CaseInfo
    Set CaseInfo = mCaseInfo
End Property

Public Property Get App0() As App0
    'Private mApp1 As App1
    If (mApp0 Is Nothing) Then Set App0 = New App0
    Set App0 = mApp0
End Property

Public Property Get App1() As App1
    'Private mApp1 As App1
    If (mApp1 Is Nothing) Then Set App1 = New App1
    Set App1 = mApp1
End Property

Public Property Get App2() As App2
    'Private mApp1 As App1
    If (mApp2 Is Nothing) Then Set App2 = New App2
    Set App2 = mApp2
End Property

Public Property Get App3() As App3
    'Private mApp1 As App1
    If (mApp3 Is Nothing) Then Set App3 = New App3
    Set App3 = mApp3
End Property

Public Property Get App4() As App4
    'Private mApp1 As App1
    If (mApp4 Is Nothing) Then Set App4 = New App4
    Set App4 = mApp4
End Property

Public Property Get App5() As App5
    'Private mApp1 As App1
    If (mApp5 Is Nothing) Then Set App5 = New App5
    Set App5 = mApp5
End Property

Public Property Get App6() As App6
    'Private mApp1 As App1
    If (mApp6 Is Nothing) Then Set App6 = New App6
    Set App6 = mApp6
End Property

Private Property Get Recordset() As Recordset
    Set Recordset = mrstRecordset
End Property

Private Property Set Recordset(rData As Recordset)
    Set mrstRecordset = rData
End Property

Private Sub Load()
    With Recordset
        sApp0 = Nz(.Fields("OrderingID").Value)
        sApp1 = Nz(.Fields("App1").Value)
        sApp2 = Nz(.Fields("App2").Value)
        sApp3 = Nz(.Fields("App3").Value)
        sApp4 = Nz(.Fields("App4").Value)
        sApp5 = Nz(.Fields("App5").Value)
        sApp6 = Nz(.Fields("App6").Value)
        sStatusID = Nz(.Fields("StatusesID").Value)
        sStatusCID = Nz(.Fields("ID").Value)
        sCourtDatesID = .Fields("ID").Value
        sCaseID = .Fields("CasesID").Value
        Me.InvoiceNo = .Fields("InvoiceNo").Value
        Me.Subtotal = Nz(.Fields("Subtotal").Value)
        Me.InvoiceDate = Nz(.Fields("InvoiceDate").Value)
        Me.Quantity = Nz(.Fields("Quantity").Value)
        Me.InventoryRateCode = Nz(.Fields("InventoryRateCode").Value)
        Me.ActualQuantity = Nz(.Fields("ActualQuantity").Value)
        Me.TurnaroundTime = .Fields("TurnaroundTimesCD").Value
        Me.AudioLength = Nz(.Fields("AudioLength").Value)
        Me.EstimatedPageCount = Nz(.Fields("EstimatedPageCount").Value)
        Me.StatusesID = Nz(.Fields("StatusesID").Value)
        Me.DueDate = Nz(.Fields("DueDate").Value)
        Me.ExpectedAdvanceDate = Nz(.Fields("ExpectedAdvanceDate").Value)
        Me.ExpectedRebateDate = Nz(.Fields("ExpectedRebateDate").Value)
        Me.PaymentSum = Nz(.Fields("PaymentSum").Value)
        Me.BrandingTheme = Nz(.Fields("BrandingTheme").Value)
        Me.FinalPrice = Nz(.Fields("FinalPrice").Value)
        Me.FactoringCost = Nz(.Fields("FactoringCost").Value)
        Me.HearingDate = Nz(.Fields("HearingDate").Value)
        Me.HearingStartTime = .Fields("HearingStartTime").Value
        Me.HearingEndTime = .Fields("HearingEndTime").Value
        Me.CaseID = Nz(.Fields("CasesID").Value)
        Me.UnitPrice = Nz(.Fields("UnitPrice").Value)
        Me.Location = Nz(.Fields("Location").Value)
        Me.PPID = Nz(.Fields("PPID").Value)
    End With
    mbooLoaded = True
    mApp0.FindFirst "ID=" & sApp0
    mApp1.FindFirst "ID=" & sApp1
    mApp2.FindFirst "ID=" & sApp2
    mApp3.FindFirst "ID=" & sApp3
    mApp4.FindFirst "ID=" & sApp4
    mApp5.FindFirst "ID=" & sApp5
    mApp6.FindFirst "ID=" & sApp6
    mCaseInfo.FindFirst "ID=" & sCaseID
    mStatus.FindFirst "ID=" & sStatusID
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
        sApp0 = Nz(.Fields("OrderingID").Value)
        sApp1 = Nz(.Fields("App1").Value)
        sApp2 = Nz(.Fields("App2").Value)
        sApp3 = Nz(.Fields("App3").Value)
        sApp4 = Nz(.Fields("App4").Value)
        sApp5 = Nz(.Fields("App5").Value)
        sApp6 = Nz(.Fields("App6").Value)
        sCaseID = Nz(.Fields("CasesID").Value)
        sStatusID = Nz(.Fields("StatusesID").Value)
        sStatusCID = Nz(.Fields("ID").Value)
        .Fields("InvoiceNo").Value = Me.InvoiceNo
        .Fields("Subtotal").Value = Me.Subtotal
        .Fields("InvoiceDate").Value = Me.InvoiceDate
        .Fields("Quantity").Value = Me.Quantity
        .Fields("InventoryRateCode").Value = Me.InventoryRateCode
        .Fields("ActualQuantity").Value = Me.ActualQuantity
        .Fields("TurnaroundTimesCD").Value = Me.TurnaroundTime
        .Fields("AudioLength").Value = Me.AudioLength
        .Fields("EstimatedPageCount").Value = Me.EstimatedPageCount
        .Fields("StatusesID").Value = Me.StatusesID
        .Fields("DueDate").Value = Me.DueDate
        .Fields("ExpectedAdvanceDate").Value = Me.ExpectedAdvanceDate
        .Fields("ExpectedRebateDate").Value = Me.ExpectedRebateDate
        .Fields("PaymentSum").Value = Me.PaymentSum
        .Fields("BrandingTheme").Value = Me.BrandingTheme
        .Fields("FinalPrice").Value = Me.FinalPrice
        .Fields("HearingDate").Value = Me.HearingDate
        .Fields("CasesID").Value = Me.CaseID
        .Fields("UnitPrice").Value = Me.UnitPrice
        .Fields("HearingStartTime").Value = Me.HearingStartTime
        .Fields("HearingEndTime").Value = Me.HearingEndTime
        .Fields("HearingDate").Value = Me.HearingDate
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
    Set mApp0 = New App0
    Set mApp1 = New App1
    Set mApp2 = New App2
    Set mApp3 = New App3
    Set mApp4 = New App4
    Set mApp5 = New App5
    Set mApp6 = New App6
    Set mCaseInfo = New CaseInfo
    Set mStatus = New Status
End Sub

Private Sub Class_Terminate()
    Recordset.Close
    Set Recordset = Nothing
    Set mApp1 = Nothing
End Sub

Private Function NullIfEmptyString(str As String) As Variant
    Dim strTrimmed As String: strTrimmed = Trim(str)
    If Len(strTrimmed) = 0 Then
        NullIfEmptyString = Null
    Else
        NullIfEmptyString = strTrimmed
    End If
End Function

