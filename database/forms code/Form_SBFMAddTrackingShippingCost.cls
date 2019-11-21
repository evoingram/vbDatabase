VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SBFMAddTrackingShippingCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAddUSPSTrackingNumber_Click()


Dim sTrackingNumber As String, sCourtDatesID As String, sShippingCost As String
Dim rstShippingOptions As DAO.Recordset, rstCourtDates As DAO.Recordset

'get info from form
sCourtDatesID = Me.cmbCourtDatesID.Value
sTrackingNumber = Me.tTrackingNo.Value
sShippingCost = Me.tShippingCost.Value

'insert via dao recordset
Set rstCourtDates = CurrentDb.OpenRecordset("SELECT ID, TrackingNumber from CourtDates where ID = " & sCourtDatesID & ";")
rstCourtDates.Edit
rstCourtDates.Fields("TrackingNumber").Value = sTrackingNumber
rstCourtDates.Update
rstCourtDates.Close

Set rstShippingOptions = CurrentDb.OpenRecordset("SELECT CourtDatesID, TrackingNumber, ShippingCost from ShippingOptions where CourtDatesID = " & sCourtDatesID & ";")

rstShippingOptions.Edit
rstShippingOptions.Fields("CourtDatesID").Value = sCourtDatesID
rstShippingOptions.Fields("TrackingNumber").Value = sTrackingNumber
rstShippingOptions.Fields("ShippingCost").Value = sShippingCost
rstShippingOptions.Update
rstShippingOptions.Close

MsgBox ("Tracking number updated")
Me.cmbCourtDatesID.Value = ""
Me.tTrackingNo.Value = ""
Me.tShippingCost.Value = ""

End Sub
