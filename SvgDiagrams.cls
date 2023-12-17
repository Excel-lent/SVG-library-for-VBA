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
    Select Case chrtObject.Chart.ChartType
        Case xlColumnClustered
            Dim svgBarChrt As New SvgBarChart
            Call svgBarChrt.ProcessAndSave(chrtObject, outputPath)
        Case xlXYScatter
    End Select
End Sub
