VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_JurisdictionRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'JurisdictionRef
Private Sub WBCViewJobFolder_OnUpdated()
Dim sTranscriptDir As String, sJobNumberDirectory As String, sAppName As String

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
Call pfCurrentCaseInfo



If sJurisdiction = "*AVT*" Or sJurisdiction = "*eScribers*" Then

    'T:\Administration\Jurisdiction References\eScribers
    sJobNumberDirectory = "T:\Administration\Jurisdiction References\eScribers\"
    
    sTranscriptDir = "file:///" & sJobNumberDirectory
        
    Me.WBCViewJobFolder.Navigate2 sTranscriptDir


ElseIf sJurisdiction = "Food and Drug Administration" Then
    
    'T:\Administration\Jurisdiction References\FDA
    sJobNumberDirectory = "T:\Administration\Jurisdiction References\FDA"
    
    sTranscriptDir = "file:///" & sJobNumberDirectory
        
    Me.WBCViewJobFolder.Navigate2 sTranscriptDir
    
ElseIf sJurisdiction = "*FDA*" Then
     
     'T:\Administration\Jurisdiction References\FDA
    sJobNumberDirectory = "T:\Administration\Jurisdiction References\FDA"
    
    sTranscriptDir = "file:///" & sJobNumberDirectory
        
    Me.WBCViewJobFolder.Navigate2 sTranscriptDir

ElseIf sJurisdiction = "*Weber*" Then
    
    'T:\Administration\Jurisdiction References\Weber
    sJobNumberDirectory = "T:\Administration\Jurisdiction References\Weber"
    
    sTranscriptDir = "file:///" & sJobNumberDirectory
        
    Me.WBCViewJobFolder.Navigate2 sTranscriptDir

ElseIf sJurisdiction = "*County*" Then
    
        'T:\Administration\Jurisdiction References\WACounties
    sJobNumberDirectory = "T:\Administration\Jurisdiction References\WACounties"
    
    sTranscriptDir = "file:///" & sJobNumberDirectory
        
    Me.WBCViewJobFolder.Navigate2 sTranscriptDir
        
ElseIf sJurisdiction = "*USBC*" Then
    
        'T:\Administration\Jurisdiction References\USBC
    sJobNumberDirectory = "T:\Administration\Jurisdiction References\USBC"
    
    sTranscriptDir = "file:///" & sJobNumberDirectory
        
    Me.WBCViewJobFolder.Navigate2 sTranscriptDir
 
ElseIf sJurisdiction = "Massachusetts" Then
    
        'T:\Administration\Jurisdiction References\USBC
    sJobNumberDirectory = "T:\Administration\Jurisdiction References\Massachusetts"
    
    sTranscriptDir = "file:///" & sJobNumberDirectory
        
    Me.WBCViewJobFolder.Navigate2 sTranscriptDir
Else
    'T:\Administration\Jurisdiction References
    sJobNumberDirectory = "T:\Administration\Jurisdiction References"
    
    sTranscriptDir = "file:///" & sJobNumberDirectory
        
    Me.WBCViewJobFolder.Navigate2 sTranscriptDir
    Application.FollowHyperlink ("https://ac.courts.wa.gov/index.cfm?fa=efiling.home")
        
End If
    
End Sub
Private Sub Command5_Click()

Dim sTranscriptDir As String, sJobNumberDirectory As String, sAppName As String

sCourtDatesID = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]
Call pfCurrentCaseInfo



If sJurisdiction = "*AVT*" Or sJurisdiction = "*eScribers*" Then

    'T:\Administration\Jurisdiction References\eScribers
    sJobNumberDirectory = "T:\Administration\Jurisdiction References\eScribers\"
    
    sTranscriptDir = "file:///" & sJobNumberDirectory
        
    Me.WBCViewJobFolder.Navigate2 sTranscriptDir


ElseIf sJurisdiction = "Food and Drug Administration" Then
    
    'T:\Administration\Jurisdiction References\FDA
    sJobNumberDirectory = "T:\Administration\Jurisdiction References\FDA"
    
    sTranscriptDir = "file:///" & sJobNumberDirectory
        
    Me.WBCViewJobFolder.Navigate2 sTranscriptDir
    
ElseIf sJurisdiction = "*FDA*" Then
     
     'T:\Administration\Jurisdiction References\FDA
    sJobNumberDirectory = "T:\Administration\Jurisdiction References\FDA"
    
    sTranscriptDir = "file:///" & sJobNumberDirectory
        
    Me.WBCViewJobFolder.Navigate2 sTranscriptDir

ElseIf sJurisdiction = "*Weber*" Then
    
    'T:\Administration\Jurisdiction References\Weber
    sJobNumberDirectory = "T:\Administration\Jurisdiction References\Weber"
    
    sTranscriptDir = "file:///" & sJobNumberDirectory
        
    Me.WBCViewJobFolder.Navigate2 sTranscriptDir

ElseIf sJurisdiction = "*County*" Then
    
        'T:\Administration\Jurisdiction References\WACounties
    sJobNumberDirectory = "T:\Administration\Jurisdiction References\WACounties"
    
    sTranscriptDir = "file:///" & sJobNumberDirectory
        
    Me.WBCViewJobFolder.Navigate2 sTranscriptDir
        
ElseIf sJurisdiction = "*USBC*" Then
    
        'T:\Administration\Jurisdiction References\USBC
    sJobNumberDirectory = "T:\Administration\Jurisdiction References\USBC"
    
    sTranscriptDir = "file:///" & sJobNumberDirectory
        
    Me.WBCViewJobFolder.Navigate2 sTranscriptDir

ElseIf sJurisdiction = "Massachusetts" Then
    
        'T:\Administration\Jurisdiction References\USBC
    sJobNumberDirectory = "T:\Administration\Jurisdiction References\Massachusetts"
    
    sTranscriptDir = "file:///" & sJobNumberDirectory
        
    Me.WBCViewJobFolder.Navigate2 sTranscriptDir
    
Else
    'T:\Administration\Jurisdiction References
    sJobNumberDirectory = "T:\Administration\Jurisdiction References"
    
    sTranscriptDir = "file:///" & sJobNumberDirectory
        
    Me.WBCViewJobFolder.Navigate2 sTranscriptDir
    Application.FollowHyperlink ("https://ac.courts.wa.gov/index.cfm?fa=efiling.home")
        
End If
    

End Sub

