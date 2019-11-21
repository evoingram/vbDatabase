VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PJDoctors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnBrowse_Click()

Dim sFilePath As String
Dim fdSelectedFile As FileDialog

Set fdSelectedFile = Application.FileDialog(msoFileDialogFilePicker)
fdSelectedFile.AllowMultiSelect = False

If fdSelectedFile.Show Then
    Me.txtFileLocation = fdSelectedFile.SelectedItems.item(1)
    sFilePath = Me.txtFileLocation
    Me.WebBrowser4.Navigate_self sFilePath, 0
    Me.WebBrowser4.Requery
End If


End Sub


