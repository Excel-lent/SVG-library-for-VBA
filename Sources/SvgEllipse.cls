VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvgEllipse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This is the class that save an ellipse,
' That is used in the class that saves scatter charts in SVG,
' That is called from the function that saves one bar chart as SVG,
' That is called from the main function that saves all the charts as SVG.

Private cX!
Private cY!
Private rX!
Private rY!
Private mFillColor As New SvgColor
Private mTransform As New SvgTransform
Private mStroke As SvgStroke

Public Sub Init(x!, y!, radiusX!, radiusY!, fillColor&, Transform As SvgTransform, Optional mmStroke As SvgStroke = Nothing)
    cX = x
    cY = y
    rX = radiusX
    rY = radiusY
    Call mFillColor.Init(fillColor)
    Set mTransform = Transform
    Set mStroke = mmStroke
End Sub

'        <ellipse cx="1" cy="20" rx="0.1" ry="40" fill="rgb(0,0,0)" transform="matrix(41.5 0 0 -0.125 29.1 159)"/>
Public Function ToString$()
    If mStroke Is Nothing Then
        ToString = "<ellipse cx=""" & Transform(cX) & """ cy=""" & Transform(cY) & _
            """ rx=""" & Transform(rX) & """ ry=""" & Transform(rY) & """ fill=""" & mFillColor.ToString & """ " & mTransform.ToString & "/>" & vbCrLf
    Else
        ToString = "<ellipse cx=""" & Transform(cX) & """ cy=""" & Transform(cY) & _
            """ rx=""" & Transform(rX) & """ ry=""" & Transform(rY) & """ fill=""" & _
            mFillColor.ToString & """ " & mTransform.ToString & " " & mStroke.ToString & "/>" & vbCrLf
    End If
End Function
