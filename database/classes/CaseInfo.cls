VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CaseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Database.Production.Classes")
Option Compare Database
Option Explicit

Private mParty1 As String
Private mParty1Name As String
Private mParty2 As String
Private mParty2Name As String
Private mCaseNumber1 As String
Private mCaseNumber2 As String
Private mJurisdiction As String
Private mHearingTitle As String
Private mJudge As String
Private mJudgeTitle As String
Private mNotes As String
'@Ignore EncapsulatePublicField
Public sCaseID As Long

Private mrstRecordset7 As Recordset
Private mbooLoaded As Boolean

Public Property Get Party1() As String
    Party1 = mParty1
End Property

Public Property Let Party1(rData As String)
    mParty1 = rData
End Property

Public Property Get Party1Name() As String
    Party1Name = mParty1Name
End Property

Public Property Let Party1Name(rData As String)
    mParty1Name = rData
End Property

Public Property Get Party2() As String
    Party2 = mParty2
End Property

Public Property Let Party2(rData As String)
    mParty2 = rData
End Property

Public Property Get Party2Name() As String
    Party2Name = mParty2Name
End Property

Public Property Let Party2Name(rData As String)
    mParty2Name = rData
End Property

Public Property Get CaseNumber2() As String
    CaseNumber2 = mCaseNumber2
End Property

Public Property Let CaseNumber2(rData As String)
    mCaseNumber2 = rData
End Property

Public Property Get CaseNumber1() As String
    CaseNumber1 = mCaseNumber1
End Property

Public Property Let CaseNumber1(rData As String)
    mCaseNumber1 = rData
End Property

Public Property Get Jurisdiction() As String
    Jurisdiction = mJurisdiction
End Property

Public Property Let Jurisdiction(rData As String)
    mJurisdiction = rData
End Property

Public Property Get HearingTitle() As String
    HearingTitle = mHearingTitle
End Property

Public Property Let HearingTitle(rData As String)
    mHearingTitle = rData
End Property

Public Property Get Judge() As String
    Judge = mJudge
End Property

Public Property Let Judge(rData As String)
    mJudge = rData
End Property

Public Property Get JudgeTitle() As String
    JudgeTitle = mJudgeTitle
End Property

Public Property Let JudgeTitle(rData As String)
    mJudgeTitle = rData
End Property

Public Property Get Notes() As String
    Notes = mNotes
End Property

Public Property Let Notes(rData As String)
    mNotes = rData
End Property

Public Property Get ID() As Long
    ID = sCaseID
End Property

Private Property Get RecordsetA7() As Recordset
    Set RecordsetA7 = mrstRecordset7
End Property

Private Property Set RecordsetA7(rData As Recordset)
    Set mrstRecordset7 = rData
End Property

Private Sub Load()
    With RecordsetA7
        sCaseID = Nz(.Fields("ID").Value)
        Me.Party1 = Nz(.Fields("Party1").Value)
        Me.Party2 = Nz(.Fields("Party2").Value)
        Me.Party1Name = Nz(.Fields("Party1Name").Value)
        Me.Party2Name = Nz(.Fields("Party2Name").Value)
        Me.CaseNumber1 = Nz(.Fields("CaseNumber1").Value)
        Me.CaseNumber2 = Nz(.Fields("CaseNumber2").Value)
        Me.Jurisdiction = Nz(.Fields("Jurisdiction").Value)
        Me.HearingTitle = Nz(.Fields("HearingTitle").Value)
        Me.Judge = Nz(.Fields("Judge").Value)
        Me.JudgeTitle = Nz(.Fields("JudgeTitle").Value)
        Me.Notes = Nz(.Fields("Notes").Value)
    End With
    mbooLoaded = True
End Sub

Public Sub Update()
    With RecordsetA7
        If mbooLoaded = True Then
            .Edit
        Else
            .AddNew
        End If
        sCaseID = Nz(.Fields("ID").Value)
        .Fields("Party1").Value = Me.Party1
        .Fields("Party2").Value = Me.Party2
        .Fields("Party1Name").Value = Me.Party1Name
        .Fields("Party2Name").Value = Me.Party2Name
        .Fields("CaseNumber1").Value = Me.CaseNumber1
        .Fields("CaseNumber2").Value = Me.CaseNumber2
        .Fields("Jurisdiction").Value = Me.Jurisdiction
        .Fields("HearingTitle").Value = Me.HearingTitle
        .Fields("Judge").Value = Me.Judge
        .Fields("JudgeTitle").Value = Me.JudgeTitle
        .Fields("Notes").Value = Me.Notes
        .Fields("CasesID").Value = Me.ID
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
        RecordsetA7.MoveFirst
        FindFirst = Not RecordsetA7.EOF
    Else
        RecordsetA7.FindFirst Criteria
        FindFirst = Not RecordsetA7.NoMatch
    End If
    If FindFirst Then Load
End Function

Private Sub Class_Initialize()
    Set RecordsetA7 = CurrentDb.OpenRecordset("Cases", dbOpenDynaset)
    
End Sub

Private Sub Class_Terminate()
    RecordsetA7.Close
    Set RecordsetA7 = Nothing
End Sub

Private Function NullIfEmptyString(str As String) As Variant
    Dim strTrimmed As String: strTrimmed = Trim(str)
    If Len(strTrimmed) = 0 Then
        NullIfEmptyString = Null
    Else
        NullIfEmptyString = strTrimmed
    End If
End Function

