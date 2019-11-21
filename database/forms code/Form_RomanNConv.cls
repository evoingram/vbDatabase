VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_RomanNConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub RNCButton_Click()
Dim iX As Integer
Dim vRomanFormat

iX = Text6.Value
If iX = 0 Then Exit Sub

Const coRomanSymbols = "IVXLCDM" ' roman symbols
Dim iZ As Integer: iZ = Abs(iX)
Dim sRoman As String, iY As Integer

For iY = 1 To 5 Step 2

    Dim w As Integer: w = iZ Mod 10: iZ = iZ \ 10
    
    Select Case w ' format decimal digit
    
        Case 0 To 3: sRoman = String(w, Mid(coRomanSymbols, iY, 1)) & sRoman
        Case 4:      sRoman = Mid(coRomanSymbols, iY, 2) & sRoman
        Case 5 To 8: sRoman = Mid(coRomanSymbols, iY + 1, 1) & String(w - 5, Mid(coRomanSymbols, iY, 1)) & sRoman
        Case 9:      sRoman = Mid(coRomanSymbols, iY, 1) & Mid(coRomanSymbols, iY + 2, 1) & sRoman
        
    End Select
    
Next

sRoman = String(iZ, "M") & sRoman    ' thousands
If iX < 0 Then sRoman = "-" & sRoman ' sign if negative
vRomanFormat = sRoman
Text8 = sRoman

End Sub
