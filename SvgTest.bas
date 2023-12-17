Attribute VB_Name = "SvgTest"
Option Explicit

Public Sub SvgTest()
    Dim svgDiagram As New SvgDiagrams
    Dim chrtObject As ChartObject
    
    For Each chrtObject In ActiveSheet.ChartObjects
        Set svgDiagram = New SvgDiagrams
        Call svgDiagram.Init(chrtObject, ThisWorkbook.Path & "\")
    Next chrtObject
End Sub
