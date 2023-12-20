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

Private ellipseCollections As New Collection

Public Sub Init(chrtObject As ChartObject)
    Dim i&, j!, nSeries#
    Dim ellipseCollection As Collection
    Dim mValues As Variant
    Dim mX As Variant
    Dim mY As Variant
    Dim mRadius!
    Dim plotAreaMinX!, plotAreaMinY!, plotAreaWidth!, plotAreaHeight!
    Dim mFillColor&, nBars&
    Dim scaleX!, scaleY!, zeroX!, zeroY!
    Dim mMaximumScale!, mMinimumScale!
    Dim mEllipse As SvgEllipse
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
        Set ellipseCollection = New Collection
        
        mX = chrtObject.Chart.SeriesCollection(i).XValues
        mY = chrtObject.Chart.SeriesCollection(i).Values
        
        zeroX = plotAreaMinX ' - offsetFirst / 2 - spaceBetween / 2 - (nSeries - i) * plotAreaWidth / UBound(mValues) / (mGapWidth + nSeries)
        scaleX = plotAreaWidth / (chrtObject.Chart.Axes(xlCategory).MaximumScale - chrtObject.Chart.Axes(xlCategory).MinimumScale)
        
        ' Marker size in points. For plot in (mX, mY) coordinate system it has to be scaled!
        mRadius = chrtObject.Chart.SeriesCollection(i).MarkerSize * (chrtObject.Chart.Axes(xlCategory).MaximumScale - chrtObject.Chart.Axes(xlCategory).MinimumScale) / scaleX
        mFillColor = chrtObject.Chart.SeriesCollection(i).MarkerForegroundColor
        
        Set barsTransform = New SvgTransform
        barsTransform.Init scalingX:=scaleX, scalingY:=-scaleY, offsetX:=zeroX, offsetY:=zeroY
        
        ' VERY IMPORTANT NOTE!
        ' Transform is necessary to scale the real coordinates (plotAreaMinX, plotAreaMinY, plotAreaWidth, plotAreaHeight) in points
        ' to plot scale (mX, mY) in arbitrary units. The SVG "transform" function distorts circles to ellipses.
        ' Therefore to show circles in SVG graph, the ellipses with radiusX:=mRadius / scaleX and radiusY:=mRadius / scaleY should be used.
        For j = 1 To UBound(mX)
            Set mEllipse = New SvgEllipse
            mEllipse.Init x:=CDbl(mX(j)), _
                          y:=CDbl(mY(j)), _
                          radiusX:=mRadius / scaleX, _
                          radiusY:=mRadius / scaleY, _
                          fillColor:=mFillColor, _
                          Transform:=barsTransform
            Call ellipseCollection.Add(mEllipse, CStr(j))
        Next j
        Call ellipseCollections.Add(ellipseCollection)
    Next i
End Sub

Public Function ToString$()
    Dim i&, j&
    Dim ellipseCollection As Collection
    
    For i = 1 To ellipseCollections.Count
        Set ellipseCollection = ellipseCollections(i)
        For j = 1 To ellipseCollection.Count
            ToString = ToString & vbTab & vbTab & ellipseCollection(j).ToString
        Next j
    Next i
End Function
