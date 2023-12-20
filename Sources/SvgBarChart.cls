VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvgBarChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This is the function that saves one bar chart as SVG,
' That is called from the main function that saves all the charts as SVG.
Public Sub ProcessAndSave(chrtObject As ChartObject, outputPath$)
    Dim mSvgChart As New SvgChart
    Dim viewBoxMinX!, viewBoxMinY!, viewBoxWidth!, viewBoxHeight!
    Dim FileName$, string2save$
    
    FileName = outputPath & chrtObject.Name & ".svg"
    
    ' This is not very beautiful because the same parameters are read in in mSvgChart.Init(chrtObject), but ok.
    viewBoxMinX = 0
    viewBoxMinY = 0
    viewBoxWidth = chrtObject.Width
    viewBoxHeight = chrtObject.height
    
    ' The header of SVG file
    string2save = "<svg viewBox=""" & Transform(viewBoxMinX) & " " & _
        Transform(viewBoxMinY) & " " & _
        Transform(viewBoxWidth) & " " & _
        Transform(viewBoxHeight) & """ xmlns=""http://www.w3.org/2000/svg"">" & vbCrLf
    
    ' Initialization of the chart:
    Call mSvgChart.Init(chrtObject)
    
    ' Saves the chart:
    string2save = string2save & vbTab & mSvgChart.ToString
    
    ' The end of SVG file:
    string2save = string2save & "</svg>"
    
    Open FileName For Output As #1
    Print #1, string2save
    Close #1
End Sub
