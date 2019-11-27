VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mAppID6 As Long
Private mCompany As String, mMrMs As String, mLastName As String
Private mFirstName As String, mEmailAddress As String, mJobTitle As String
Private mBusinessPhone As String, mAddress As String, mCity As String
Private mState As String
Private mZIP As String, mNotes As String, mFactoringApproved As Boolean
Public sApp6 As Long

Private mrstRecordset6 As Recordset
Private mbooLoaded As Boolean

Public Property Get Company() As String
    Company = mCompany
End Property
Public Property Let Company(rData As String)
    mCompany = rData
End Property

Public Property Get MrMs() As String
    MrMs = mMrMs
End Property
Public Property Let MrMs(rData As String)
    mMrMs = rData
End Property

Public Property Get LastName() As String
    LastName = mLastName
End Property
Public Property Let LastName(rData As String)
    mLastName = rData
End Property

Public Property Get FirstName() As String
    FirstName = mFirstName
End Property
Public Property Let FirstName(rData As String)
    mFirstName = rData
End Property

Public Property Get EmailAddress() As String
    EmailAddress = mEmailAddress
End Property
Public Property Let EmailAddress(rData As String)
    mEmailAddress = rData
End Property

Public Property Get JobTitle() As String
    JobTitle = mJobTitle
End Property
Public Property Let JobTitle(rData As String)
    mJobTitle = rData
End Property

Public Property Get BusinessPhone() As String
    BusinessPhone = mBusinessPhone
End Property
Public Property Let BusinessPhone(rData As String)
    mBusinessPhone = rData
End Property

Public Property Get Address() As String
    Address = mAddress
End Property
Public Property Let Address(rData As String)
    mAddress = rData
End Property

Public Property Get City() As String
    Address = mCity
End Property
Public Property Let City(rData As String)
    mCity = rData
End Property

Public Property Get ZIP() As String
    ZIP = mZIP
End Property
Public Property Let ZIP(rData As String)
    mZIP = rData
End Property

Public Property Get Notes() As String
    Notes = mNotes
End Property
Public Property Let Notes(rData As String)
    mNotes = rData
End Property

Public Property Get State() As String
    State = mState
End Property
Public Property Let State(rData As String)
    mState = rData
End Property

Public Property Get FactoringApproved() As Boolean
    FactoringApproved = mFactoringApproved
End Property
Public Property Let FactoringApproved(rData As Boolean)
    mFactoringApproved = rData
End Property

Public Property Get ID() As Long
    ID = sApp6
End Property


Private Property Get RecordsetA6() As Recordset
    Set RecordsetA6 = mrstRecordset6
End Property

Private Property Set RecordsetA6(rData As Recordset)
    Set mrstRecordset6 = rData
End Property


Private Sub Load()
    With RecordsetA6
        sApp6 = Nz(.Fields("ID").Value)
        Me.Company = Nz(.Fields("Company").Value)
        Me.MrMs = Nz(.Fields("MrMs").Value)
        Me.LastName = Nz(.Fields("LastName").Value)
        Me.FirstName = Nz(.Fields("FirstName").Value)
        Me.JobTitle = Nz(.Fields("JobTitle").Value)
        Me.EmailAddress = Nz(.Fields("EmailAddress").Value)
        Me.BusinessPhone = Nz(.Fields("BusinessPhone").Value)
        Me.Address = Nz(.Fields("Address").Value)
        Me.City = Nz(.Fields("City").Value)
        Me.State = Nz(.Fields("State").Value)
        Me.ZIP = Nz(.Fields("ZIP").Value)
        Me.Notes = Nz(.Fields("Notes").Value)
       Me.FactoringApproved = Nz(.Fields("FactoringApproved").Value)
    End With
    mbooLoaded = True
End Sub

Public Sub Update()
    With RecordsetA6
        If mbooLoaded = True Then
            .Edit
        Else
            .AddNew
        End If
        sApp6 = Nz(.Fields("ID").Value)
        .Fields("Company").Value = Me.Company
        .Fields("MrMs").Value = Me.MrMs
        .Fields("LastName").Value = Me.LastName
        .Fields("FirstName").Value = Me.FirstName
        .Fields("JobTitle").Value = Me.JobTitle
        .Fields("EmailAddress").Value = Me.EmailAddress
        .Fields("BusinessPhone").Value = Me.BusinessPhone
        .Fields("Address").Value = Me.Address
        .Fields("City").Value = Me.City
        .Fields("State").Value = Me.State
        .Fields("ZIP").Value = Me.ZIP
        .Fields("Notes").Value = Me.Notes
        .Fields("FactoringApproved").Value = Me.FactoringApproved
        .Update
    End With
    mbooLoaded = True
End Sub

Public Sub AddNew()
    mbooLoaded = False
End Sub

Public Function FindFirst(Optional Criteria As Variant) As Boolean
    If IsMissing(Criteria) Then
        RecordsetA6.MoveFirst
        FindFirst = Not RecordsetA6.EOF
    Else
        RecordsetA6.FindFirst Criteria
        FindFirst = Not RecordsetA6.NoMatch
    End If
    If FindFirst Then Load
End Function

Private Sub Class_Initialize()
    Set RecordsetA6 = CurrentDb.OpenRecordset("Customers", dbOpenDynaset)
    
End Sub

Private Sub Class_Terminate()
    RecordsetA6.Close
    Set RecordsetA6 = Nothing
End Sub

Function NullIfEmptyString(str As String) As Variant
    Dim strTrimmed As String: strTrimmed = Trim(str)
    If Len(strTrimmed) = 0 Then
        NullIfEmptyString = Null
    Else
        NullIfEmptyString = strTrimmed
    End If
End Function


