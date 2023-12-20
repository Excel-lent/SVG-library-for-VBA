VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvgLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mStroke As SvgStroke
Private mX1!
Private mY1!
Private mX2!
Private mY2!
Private mTransform As New SvgTransform

Public Sub Init(x1!, y1!, x2!, y2!, stroke As SvgStroke, Transform As SvgTransform)
    Set mStroke = stroke
    mX1 = x1
    mY1 = y1
    mX2 = x2
    mY2 = y2
    Set mTransform = Transform
End Sub

Public Function ToString$()
    ToString = "<line x1=""" & Transform(mX1) & """ y1=""" & Transform(mY1) & """ x2=""" & Transform(mX2) & """ y2=""" & Transform(mY2) & """ " & _
               mStroke.ToString & " " & mTransform.ToString & " />" & vbCrLf
End Function
