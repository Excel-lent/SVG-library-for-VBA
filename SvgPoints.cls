VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvgPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' These are bars that are shown at plot area,
' That is called from plot area,
' That is shown above the SvgChart,
' That is called from the function that saves one bar chart as SVG,
' That is called from the main function that saves all the charts as SVG.

Private rectangleCollections As New Collection

Public Sub Init(chrtObject As ChartObject)
    Dim i&, j!, nSeries#
    Dim rectangleCollection As Collection
    Dim mX As Variant
    Dim mY As Variant
    Dim mGapWidth!, offsetFirst!, rectangleWidth!
    Dim plotAreaMinX!, plotAreaMinY!, plotAreaWidth!, plotAreaHeight!
    Dim mFillColor&, nBars&
    Dim scaleX!, scaleY!, zeroX!, zeroY!
    Dim mMaximumScale!, mMinimumScale!
    Dim rectangle As SvgRectangle
    Dim barsTransform As SvgTransform
    
    plotAreaMinX = chrtObject.Chart.plotArea.InsideLeft
    plotAreaMinY = chrtObject.Chart.plotArea.InsideTop
    plotAreaWidth = chrtObject.Chart.plotArea.InsideWidth
    plotAreaHeight = chrtObject.Chart.plotArea.InsideHeight
    
    mMaximumScale = chrtObject.Chart.Axes(xlValue).MaximumScale
    mMinimumScale = chrtObject.Chart.Axes(xlValue).MinimumScale
    
    scaleY = plotAreaHeight / (mMaximumScale - mMinimumScale)
    zeroY = scaleY * mMaximumScale + plotAreaMinY
    nBars = (chrtObject.Chart.Axes(xlCategory).MaximumScale - chrtObject.Chart.Axes(xlCategory).MinimumScale) / chrtObject.Chart.Axes(xlCategory).MajorUnit + 1
    
    nSeries = chrtObject.Chart.SeriesCollection.Count
    For i = 1 To nSeries
        Set rectangleCollection = New Collection
        
        mX = chrtObject.Chart.SeriesCollection(i).Values
        mY = chrtObject.Chart.SeriesCollection(i).Values
        
        rectangleWidth = 1 / (mGapWidth + chrtObject.Chart.SeriesCollection.Count)
        
        mFillColor = chrtObject.Chart.SeriesCollection(i).Format.Fill.ForeColor.RGB
        
        zeroX = plotAreaMinX ' - offsetFirst / 2 - spaceBetween / 2 - (nSeries - i) * plotAreaWidth / UBound(mValues) / (mGapWidth + nSeries)
        scaleX = plotAreaWidth / (chrtObject.Chart.Axes(xlCategory).MaximumScale - chrtObject.Chart.Axes(xlCategory).MinimumScale)
        
        Set barsTransform = New SvgTransform
        barsTransform.Init scalingX:=scaleX, scalingY:=-scaleY, offsetX:=zeroX, offsetY:=zeroY
        
        For j = 1 To UBound(mValues)
            Set rectangle = New SvgRectangle
            If mValues(j) > 0 Then
                rectangle.Init x:=j, _
                               y:=0, _
                               height:=CDbl(mValues(j)), _
                               Width:=rectangleWidth, _
                               fillColor:=mFillColor, _
                               Transform:=barsTransform
            Else
                rectangle.Init x:=j, _
                               y:=CDbl(mValues(j)), _
                               height:=-mValues(j), _
                               Width:=rectangleWidth, _
                               fillColor:=mFillColor, _
                               Transform:=barsTransform
            End If
            Call rectangleCollection.Add(rectangle, CStr(j))
        Next j
        Call rectangleCollections.Add(rectangleCollection)
    Next i
End Sub

Public Function ToString$()
    Dim i&, j&
    Dim rectangleCollection As Collection
    
    For i = 1 To rectangleCollections.Count
        Set rectangleCollection = rectangleCollections(i)
        For j = 1 To rectangleCollection.Count
            ToString = ToString & vbTab & vbTab & rectangleCollection(j).ToString
        Next j
    Next i
End Function
