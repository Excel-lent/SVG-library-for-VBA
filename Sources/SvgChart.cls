VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvgChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This is the class that initializes and saves a basic chart,
' That is called from the function that saves one bar chart as SVG,
' That is called from the main function that saves all the charts as SVG.

Private chartRectangle As New SvgRectangle
Private title As SvgText
Private plotArea As New SvgPlotArea

Private Const todoRotationAngle& = 0

Public Sub Init(chrtObject As ChartObject)
    Dim viewBoxMinX!, viewBoxMinY!, viewBoxWidth!, viewBoxHeight!
    Dim mColor As New SvgColor
    Dim mTransform As New SvgTransform
    Dim mStroke As New SvgStroke
    Dim mWidth!
    
    viewBoxMinX = 0
    viewBoxMinY = 0
    viewBoxWidth = chrtObject.Width
    viewBoxHeight = chrtObject.height
    
    Call mTransform.Init(1, 1, 0, 0)
    
    Call mColor.Init(chrtObject.Border.Color)
    mWidth = chrtObject.Border.Weight
    Call mStroke.Init(mColor, mWidth)
    Call chartRectangle.Init(Transform(viewBoxMinX), Transform(viewBoxMinY), _
                             Transform(viewBoxHeight), Transform(viewBoxWidth), chrtObject.Interior.Color, mTransform, mStroke)
    
    ' Title:
    If chrtObject.Chart.HasTitle Then
        Dim mFont As New SvgFont
        Set title = New SvgText
        mFont.Init mFont:=chrtObject.Chart.ChartTitle.Font, scaling:=1
        title.Init text:=chrtObject.Chart.ChartTitle.text, _
                   posX:=chrtObject.Chart.ChartTitle.Left, _
                   posY:=chrtObject.Chart.ChartTitle.Top, _
                   rotationAngle:=todoRotationAngle, _
                   textAnchor:="start", _
                   alignmentBaseline:="hanging", _
                   Font:=mFont
    End If
    
    Call plotArea.Init(chrtObject)
End Sub

Public Function ToString$()
    ToString = chartRectangle.ToString
    ToString = ToString & plotArea.ToString
    
    If Not title Is Nothing Then
        ToString = ToString & title.ToString
    End If
End Function
