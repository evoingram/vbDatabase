VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MMProcess Jobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub JobInfobtn_Click()
sCourtDatesID = Me.JobNumberField.Value
With Me.ProcessJobSubform
.SourceObject = "PJViewJobForm"
.LinkMasterFields = ""
.LinkChildFields = ""
.LinkMasterFields = "Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]"
.LinkChildFields = "VJFJobID"
End With
Me.ProcessJobSubform.Requery
End Sub

Private Sub StatusChecklistBTN_Click()
With Me.ProcessJobSubform
.SourceObject = "PJStatuses"
.LinkMasterFields = ""
.LinkChildFields = ""
.LinkMasterFields = "Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]"
.LinkChildFields = "CourtDatesID"
End With
Me.ProcessJobSubform.Requery
End Sub
Private Sub OrderingInfoBTN_Click()
With Me.ProcessJobSubform
.SourceObject = "PJOrderingInfoForm"
.LinkMasterFields = ""
.LinkChildFields = ""
.LinkMasterFields = "Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]"
.LinkChildFields = "CourtDatesID"
End With
End Sub

Private Sub ViewDocHistoryBTN_Click()

With Me.ProcessJobSubform
.SourceObject = "SBFMDocHistQ"
.LinkMasterFields = ""
.LinkChildFields = ""
.LinkMasterFields = "Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]"
.LinkChildFields = "CourtDatesID"
End With
Me.ProcessJobSubform.Requery
End Sub

Private Sub btnUSPSTracking_Click()

Forms![NewMainMenu]![ProcessJobSubformNMM].SourceObject = "Speech"
Forms![NewMainMenu]![ProcessJobSubformNMM].Requery
End Sub

Private Sub Command0_Click()

Forms![NewMainMenu]![ProcessJobSubformNMM].SourceObject = "MMProcess Jobs"
Forms![NewMainMenu]![ProcessJobSubformNMM].Requery

End Sub

Private Sub CreateDocsBTN_Click()
sCourtDatesID = Me.JobNumberField.Value
Me.ProcessJobSubform.SourceObject = "PJCreateCourtDocs"
Me.ProcessJobSubform.Requery
End Sub
Private Sub TypeBTN_Click()
Me.ProcessJobSubform.SourceObject = "PJType"
Me.ProcessJobSubform.Requery
Dim TranscriptDraft As String
Dim fs As Object

Call pfCurrentCaseInfo  'refresh transcript info

Set fs = CreateObject("Scripting.FileSystemObject")
If sJurisdiction = "Weber Nevada" Then
If Not fs.FileExists("I:\" & sCourtDatesID & "\" & "RoughDraft.docx") Then
    FileCopy "T:\Database\Templates\Stage2s\RoughDraft-WeberNV.docx", "I:\" & sCourtDatesID & "\" & "RoughDraft.docx"
End If
    If Not fs.FileExists("I:\" & sCourtDatesID & "\Notes\" & "Transcribing Manual.PDF") Then
        FileCopy "T:\Database\Templates\Stage1s\Transcribing Manual.PDF", "I:\" & sCourtDatesID & "\Notes\" & "Transcribing Manual.PDF"
    End If
    If Not fs.FileExists("I:\" & sCourtDatesID & "\Notes\" & "Proofreading Manual - nevada.PDF") Then
        FileCopy "T:\Database\Templates\Stage3s\Proofreading Manual - nevada.PDF", "I:\" & sCourtDatesID & "\Notes\" & "Proofreading Manual - nevada.PDF"
    End If
        If Not fs.FileExists("I:\" & sCourtDatesID & "\Notes\" & "WeberNVSample.docx") Then
        FileCopy "T:\Database\Templates\Stage2s\WeberNVSample.docx", "I:\" & sCourtDatesID & "\Notes\" & "WeberNVSample.docx"
    End If
Else

End If
If sJurisdiction = "Weber Bankruptcy" Then
        If Not fs.FileExists("I:\" & sCourtDatesID & "\Notes\" & "WeberBKSample.docx") Then
        FileCopy "T:\Database\Templates\Stage2s\WeberNVSample.docx", "I:\" & sCourtDatesID & "\Notes\" & "WeberNVSample.docx"
    End If
    If Not fs.FileExists("I:\" & sCourtDatesID & "\" & "RoughDraft.docx") Then
        FileCopy "T:\Database\Templates\Stage2s\RoughDraft.docx", "I:\" & sCourtDatesID & "\" & "RoughDraft.docx"
    Else
    End If
Else

End If
If sJurisdiction = "Weber Oregon" Then
    If Not fs.FileExists("I:\" & sCourtDatesID & "\" & "RoughDraft.docx") Then
        FileCopy "T:\Database\Templates\Stage2s\RoughDraft-WeberOR.docx", "I:\" & sCourtDatesID & "\" & "RoughDraft.docx"
    End If
        If Not fs.FileExists("I:\" & sCourtDatesID & "\Notes\" & "WeberORSample.docx") Then
        FileCopy "T:\Database\Templates\Stage2s\WeberORSample.docx", "I:\" & sCourtDatesID & "\Notes\" & "WeberORSample.docx"
    End If
            If Not fs.FileExists("I:\" & sCourtDatesID & "\Notes\" & "WeberORSample1.docx") Then
        FileCopy "T:\Database\Templates\Stage2s\WeberORSample1.docx", "I:\" & sCourtDatesID & "\Notes\" & "WeberORSample1.docx"
    End If
            If Not fs.FileExists("I:\" & sCourtDatesID & "\Notes\" & "WeberORSampleTM.docx") Then
        FileCopy "T:\Database\Templates\Stage2s\WeberORSampleTM.docx", "I:\" & sCourtDatesID & "\Notes\" & "WeberORSampleTM.docx"
    End If
                If Not fs.FileExists("I:\" & sCourtDatesID & "\Notes\" & "WeberORSample2.docx") Then
        'FileCopy "T:\Database\Templates\Stage2s\WeberORSample2.docx", "I:\" & sCourtDatesID & "\Notes\" & "WeberORSample2.docx"
    End If
Else

End If
If sJurisdiction = "USBC Western Washington" Then
        If Not fs.FileExists("I:\" & sCourtDatesID & "\Notes\" & "BankruptcyWAGuide.pdf") Then
        FileCopy "T:\Database\Templates\Stage2s\BankruptcyWAGuide.pdf", "I:\" & sCourtDatesID & "\Notes\" & "BankruptcyWAGuide.pdf"
    End If
    If Not fs.FileExists("I:\" & sCourtDatesID & "\" & "RoughDraft.docx") Then
        FileCopy "T:\Database\Templates\Stage2s\RoughDraft.docx", "I:\" & sCourtDatesID & "\" & "RoughDraft.docx"
    Else
    End If
Else
End If
If sJurisdiction = "Food and Drug Administration" Then
    If Not fs.FileExists("I:\" & sCourtDatesID & "\" & "RoughDraft.docx") Then
        FileCopy "T:\Database\Templates\Stage2s\RoughDraft-FDA.docx", "I:\" & sCourtDatesID & "\" & "RoughDraft.docx"
    Else
    End If
Else
End If

If sJurisdiction = "*Superior Court*" Then
        If Not fs.FileExists("I:\" & sCourtDatesID & "\Notes\" & "CourtRules-WACounties.pdf") Then
        FileCopy "T:\Database\Templates\Stage2s\CourtRules-WACounties.pdf", "I:\" & sCourtDatesID & "\Notes\" & "CourtRules-WACounties.pdf"
    End If
    If Not fs.FileExists("I:\" & sCourtDatesID & "\" & "RoughDraft.docx") Then
        FileCopy "T:\Database\Templates\Stage2s\RoughDraft.docx", "I:\" & sCourtDatesID & "\" & "RoughDraft.docx"
    Else
    End If
Else
End If

If sJurisdiction = "*USBC*" Then
        If Not fs.FileExists("I:\" & sCourtDatesID & "\Notes\" & "CourtRules-Bankruptcy-TranscriptFormatGuide-1.pdf") Then
        FileCopy "T:\Database\Templates\Stage2s\CourtRules-Bankruptcy-TranscriptFormatGuide-1.pdf", "I:\" & sCourtDatesID & "\Notes\" & "CourtRules-Bankruptcy-TranscriptFormatGuide-1.pdf"
        End If
        If Not fs.FileExists("I:\" & sCourtDatesID & "\Notes\" & "CourtRules-Bankruptcy-TranscriptFormatGuide-2.pdf") Then
            FileCopy "T:\Database\Templates\Stage2s\CourtRules-Bankruptcy-TranscriptFormatGuide-2.pdf", "I:\" & sCourtDatesID & "\Notes\" & "CourtRules-Bankruptcy-TranscriptFormatGuide-2.pdf"
        End If
        If Not fs.FileExists("I:\" & sCourtDatesID & "\" & "RoughDraft.docx") Then
            FileCopy "T:\Database\Templates\Stage2s\RoughDraft.docx", "I:\" & sCourtDatesID & "\" & "RoughDraft.docx"
        Else
        End If
        
Else
    If Not fs.FileExists("I:\" & sCourtDatesID & "\" & "RoughDraft.docx") Then
        FileCopy "T:\Database\Templates\Stage2s\RoughDraft.docx", "I:\" & sCourtDatesID & "\" & "RoughDraft.docx"
    Else
    End If
End If

If sJurisdiction = "Massachusetts" Then
    If Not fs.FileExists("I:\" & sCourtDatesID & "\Notes\" & "uniformtranscriptformat.pdf") Then
    FileCopy "T:\Administration\Jurisdiction References\Massachusetts\uniformtranscriptformat.pdf", "I:\" & sCourtDatesID & "\Notes\" & "uniformtranscriptformat.pdf"
    End If
End If

If Not fs.FileExists("I:\" & sCourtDatesID & "\" & "RoughDraft.docx") Then
    FileCopy "T:\Database\Templates\Stage2s\RoughDraft.docx", "I:\" & sCourtDatesID & "\" & "RoughDraft.docx"
Else
End If

TranscriptDraft = "I:\" & sCourtDatesID & "\" & "RoughDraft.docx"
If sJurisdiction = "Food and Drug Administration" Then
    Me.ProcessJobSubform.SourceObject = "PJDoctors"
    Me.ProcessJobSubform.Requery
ElseIf sJurisdiction = "*FDA*" Then
    Me.ProcessJobSubform.SourceObject = "PJDoctors"
    Me.ProcessJobSubform.Requery
ElseIf sJurisdiction = "*fda*" Then
    Me.ProcessJobSubform.SourceObject = "PJDoctors"
    Me.ProcessJobSubform.Requery
Else
    Me.ProcessJobSubform.SourceObject = "PJType"
    Me.ProcessJobSubform.Requery
End If
'Call PlayAudioFolder
Documents.Open FileName:=TranscriptDraft, ReadOnly:=False
Call pfClearGlobals
End Sub
Private Sub ViewJobFolderBTN_Click()
sCourtDatesID = Me.JobNumberField.Value
Me.ProcessJobSubform.SourceObject = "PJViewJobFolder"
Me.ProcessJobSubform.Requery

End Sub
Private Sub ShippingInfoBTN_Click()
Dim rs As DAO.Recordset
Dim SQLString As String, sMailClassNo As String, sPackageTypeNo As String
Dim sMailClass As String, sPackageType As String
Dim rstShip As DAO.Recordset, rstMailC As DAO.Recordset, rstPkgType As DAO.Recordset

Call pfGetOrderingAttorneyInfo

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
SQLString = "SELECT * FROM ShippingOptions WHERE [CourtDatesIDLK] = " & sCourtDatesID & ";"
Set rs = CurrentDb.OpenRecordset(SQLString)

If rs.RecordCount < 1 Then

    rs.AddNew
    rs("CourtDatesID").Value = sCourtDatesID
    rs("CourtDatesIDLK").Value = sCourtDatesID
    rs("ReferenceID").Value = sCourtDatesID
    rs("MailClass").Value = "3"
    rs("PackageType").Value = "4"
    rs("Width").Value = "11"
    rs("Length").Value = "14"
    rs("Depth").Value = "1"
    rs("PriorityMailExpress1030").Value = "FALSE"
    rs("HolidayDelivery").Value = "FALSE"
    rs("SundayDelivery").Value = "FALSE"
    rs("SaturdayDelivery").Value = "FALSE"
    rs("SignatureRequired").Value = "FALSE"
    rs("Stealth").Value = "TRUE"
    rs("ReplyPostage").Value = "FALSE"
    rs("InsuredMail").Value = "OFF"
    rs("COD").Value = "OFF"
    rs("RestrictedDelivery").Value = "OFF"
    rs("AdultSignatureRestricted").Value = "OFF"
    rs("AdultSignatureRequired").Value = "OFF"
    rs("ReturnReceipt").Value = "OFF"
    rs("CertifiedMail").Value = "OFF"
    rs("SignatureConfirmation").Value = "OFF"
    rs("USPSTracking").Value = "ON"
    rs("ToName").Value = sName
    rs("ToAddress1").Value = sAddress1
    rs("ToAddress2").Value = sAddress2
    rs("ToCity").Value = sCity
    rs("ToState").Value = sState
    rs("ToPostalCode").Value = sZIP
    rs("ToCountry").Value = "United States"
    rs("Description").Value = sCourtDatesID & " Transcripts"
    rs("Value").Value = "100"
    rs("ToEmail").Value = "inquiries@aquoco.co"
    rs("ToPhone").Value = "206-478-5028"
    rs("PriorityMailExpress1030").Value = "FALSE"
    rs.Update
    rs.Close
    With Me.ProcessJobSubform
        .SourceObject = "FShippingOptions"
        .Requery
        .LinkMasterFields = ""
        .LinkChildFields = ""
        .LinkMasterFields = "Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]"
        .LinkChildFields = "CourtDatesID"
        .Form.RecordSource = SQLString
        .Form.Requery
    End With
Else
    With Me.ProcessJobSubform
        .SourceObject = "FShippingOptions"
        .Requery
        .LinkMasterFields = ""
        .LinkChildFields = ""
        .LinkMasterFields = "Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]"
        .LinkChildFields = "CourtDatesID"
        .Form.RecordSource = SQLString
        .Form.Requery
    End With

End If
rs.Close
Set rs = Nothing

End Sub
Private Sub InvoicingBTN_Click()
With Me.ProcessJobSubform
.SourceObject = "INV-F-ViewInvoiceForm"
End With
Me.ProcessJobSubform.Requery
'Call InvoicesCSV
End Sub
Private Sub SRTextBTN_Click()
sCourtDatesID = Me.JobNumberField.Value
Me.ProcessJobSubform.SourceObject = "JurisdictionRef"
Me.ProcessJobSubform.Requery
End Sub
