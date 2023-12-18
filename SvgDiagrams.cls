VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvgDiagrams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This is the main function that saves the charts as SVG.
Public Sub Init(chrtObject As ChartObject, outputPath$)
    Dim svgBarChrt As New SvgBarChart
    
    Select Case chrtObject.Chart.ChartType
        Case xlColumnClustered
            Call svgBarChrt.ProcessAndSave(chrtObject, outputPath)
        Case xlXYScatter
            Call svgBarChrt.ProcessAndSave(chrtObject, outputPath)
    End Select
End Sub
