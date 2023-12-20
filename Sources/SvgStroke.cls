VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvgStroke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mStrokeColor As SvgColor
Private mStrokeWidth!

Public Sub Init(strokeColor As SvgColor, strokeWidth!)
    Set mStrokeColor = strokeColor
    mStrokeWidth = strokeWidth
End Sub

Public Function ToString$()
    ToString = "stroke=""" & mStrokeColor.ToString & """ stroke-width=""" & Transform(mStrokeWidth) & """"
End Function
