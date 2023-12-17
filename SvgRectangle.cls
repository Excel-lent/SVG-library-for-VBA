VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvgRectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This is the class that save a rectangle,
' That is used in the class that saves a basic chart, plot area and bar charts in SVG,
' That is called from the function that saves one bar chart as SVG,
' That is called from the main function that saves all the charts as SVG.

Private mX!
Private mY!
Private mWidth!
Private mHeight!
Private mFillColor As New SvgColor
Private mTransform As New SvgTransform
Private mStroke As SvgStroke

Public Sub Init(x!, y!, height!, Width!, fillColor&, Transform As SvgTransform, Optional mmStroke As SvgStroke = Nothing)
    mX = x
    mY = y
    mHeight = height
    mWidth = Width
    Call mFillColor.Init(fillColor)
    Set mTransform = Transform
    Set mStroke = mmStroke
End Sub

Public Function ToString$()
    If mStroke Is Nothing Then
        ToString = "<rect x=""" & Transform(mX) & """ y=""" & Transform(mY) & _
            """ width=""" & Transform(mWidth) & """ height=""" & Transform(mHeight) & """ fill=""" & mFillColor.ToString & """ " & mTransform.ToString & "/>" & vbCrLf
    Else
        ToString = "<rect x=""" & Transform(mX) & """ y=""" & Transform(mY) & _
            """ width=""" & Transform(mWidth) & """ height=""" & Transform(mHeight) & """ fill=""" & _
            mFillColor.ToString & """ " & mTransform.ToString & " " & mStroke.ToString & "/>" & vbCrLf
    End If
End Function
