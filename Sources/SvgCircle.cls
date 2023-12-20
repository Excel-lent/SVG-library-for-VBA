VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvgCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This is the class that save a circle,
' That is used in the class that saves scatter charts in SVG,
' That is called from the function that saves one bar chart as SVG,
' That is called from the main function that saves all the charts as SVG.

Private cX!
Private cY!
Private mRadius!
Private mFillColor As New SvgColor
Private mTransform As New SvgTransform
Private mStroke As SvgStroke

Public Sub Init(x!, y!, radius!, fillColor&, Transform As SvgTransform, Optional mmStroke As SvgStroke = Nothing)
    cX = x
    cY = y
    mRadius = radius
    Call mFillColor.Init(fillColor)
    Set mTransform = Transform
    Set mStroke = mmStroke
End Sub

Public Function ToString$()
    If mStroke Is Nothing Then
        ToString = "<circle cx=""" & Transform(cX) & """ cy=""" & Transform(cY) & _
            """ r=""" & Transform(mRadius) & """ fill=""" & mFillColor.ToString & """ " & mTransform.ToString & "/>" & vbCrLf
    Else
        ToString = "<circle cx=""" & Transform(cX) & """ cy=""" & Transform(cY) & _
            """ r=""" & Transform(mRadius) & """ fill=""" & _
            mFillColor.ToString & """ " & mTransform.ToString & " " & mStroke.ToString & "/>" & vbCrLf
    End If
End Function
