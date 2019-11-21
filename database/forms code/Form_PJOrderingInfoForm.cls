VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PJOrderingInfoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub HDTOrderingInfo_Click()
Me.OrderingAttorneyInfoSBFM.SourceObject = "Query." & "OrderingAttorneyInfo"
Me.OrderingAttorneyInfoSBFM.Requery
Me.OrderingAttorneyInfoSBFM.Form.DefaultView = "Single Form"
End Sub
Private Sub OrderingAttorneyInfoSBFM_AfterUpdate()
Me.OrderingAttorneyInfoSBFM.Form.DefaultView = "Single Form"
End Sub

