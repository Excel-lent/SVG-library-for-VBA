VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvgFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mScaling As Double
Private mmFont As Font

' TODO:
'    font-variant="small-caps"
'    font-style="oblique" stroke="black" stroke-width="1"

Public Sub Init(mFont As Font, scaling As Double)
    mScaling = scaling
    Set mmFont = mFont
End Sub

Public Function ToString$()
    Dim mColor As New SvgColor
    
    If mmFont.Size <> 0 Then
        ToString = "font-size=""" & Transform(mmFont.Size / mScaling) & """ "
    End If
    If InStr(1, LCase(mmFont.FontStyle), "bold") <> 0 Then
        ToString = ToString & "font-weight=""bold"" "
    End If
    If InStr(1, LCase(mmFont.FontStyle), "italic") <> 0 Then
        ToString = ToString & "font-style=""italic"" "
    End If
    If mmFont.Name <> "" Then
        ToString = ToString & "font-family=""" & mmFont.Name & """ "
    End If
    mColor.Init (mmFont.Color)
    ToString = ToString & "fill=""" & mColor.ToString & """ "
End Function


