VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ViewInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub btnViewInvoice_Click()


Dim rstCourtDates As DAO.Recordset
Dim sInvoiceNumber As String, sInvoicePath As String, wbViewInvoice As Object

sCourtDatesID = Me.btnVIJobNumberField.Value
Set rstCourtDates = CurrentDb.OpenRecordset("SELECT * FROM COURTDATES WHERE ID = " & sCourtDatesID & ";")
sInvoiceNumber = rstCourtDates.Fields("InvoiceNo").Value
rstCourtDates.Close
Set rstCourtDates = Nothing
sInvoicePath = "I:\" & sCourtDatesID & "\Generated\" & sInvoiceNumber & ".pdf"
Me.txtInvoiceNo.Value = sInvoiceNumber
If Len(Dir(sInvoicePath, vbDirectory)) <> 0 Then
    With Me.wbViewInvoice
        .Navigate2 sInvoicePath
        .Requery
    End With
Else
    MsgBox "Job number " & sCourtDatesID & ", Invoice number " & sInvoiceNumber & " doesn't have an invoice processed yet."
End If

End Sub
