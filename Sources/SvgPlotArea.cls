VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvgPlotArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This is the class that initializes and saves plot area,
' That is shown above the SvgChart,
' That is called from the function that saves one bar chart as SVG,
' That is called from the main function that saves all the charts as SVG.

Private plotAreaRectangle As New SvgRectangle
Private Const todoRotationAngle& = 0
Private mSvgBars As New SvgBars
Private mSvgPoints As New SvgPoints
Private mChartType$

Private horizontalLabelsCollection As New Collection
Private verticalLabelsCollection As New Collection
Private horizontalLinesCollection As New Collection

Public Sub Init(chrtObject As ChartObject)
    Dim plotAreaMinX!, plotAreaMinY!, plotAreaWidth!, plotAreaHeight!
    Dim scaleX!, scaleY!, zeroX!, zeroY!
    Dim mMaximumScale!, mMinimumScale!, mMajorUnits!
    Dim nVerticalLines&, mGapWidth!, spaceBetween!, position!, offsetFirst!
    Dim j!
    Dim mValues As Variant
    Dim hLine As SvgLine
    
    plotAreaMinX = chrtObject.Chart.plotArea.InsideLeft
    plotAreaMinY = chrtObject.Chart.plotArea.InsideTop
    plotAreaWidth = chrtObject.Chart.plotArea.InsideWidth
    plotAreaHeight = chrtObject.Chart.plotArea.InsideHeight
    
    mMaximumScale = chrtObject.Chart.Axes(xlValue).MaximumScale
    mMinimumScale = chrtObject.Chart.Axes(xlValue).MinimumScale
    mMajorUnits = chrtObject.Chart.Axes(xlValue).MajorUnit
    
    If chrtObject.Chart.ChartType = xlXYScatter Then
        nVerticalLines = (chrtObject.Chart.Axes(xlCategory).MaximumScale - chrtObject.Chart.Axes(xlCategory).MinimumScale) / chrtObject.Chart.Axes(xlCategory).MajorUnit + 1
    Else
        nVerticalLines = UBound(chrtObject.Chart.Axes(xlCategory).CategoryNames)
    End If
    
    zeroX = plotAreaMinX
    scaleX = plotAreaWidth / (nVerticalLines + 1)
    scaleY = plotAreaHeight / (mMaximumScale - mMinimumScale)
    zeroY = scaleY * mMaximumScale + plotAreaMinY
    
    mGapWidth = chrtObject.Chart.ChartGroups(1).GapWidth / 100
    
    Dim mFont As New SvgFont
    mFont.Init mFont:=chrtObject.Chart.Axes(xlCategory).TickLabels.Font, scaling:=1
    
    Dim plotAreaTransform As New SvgTransform
    plotAreaTransform.Init scalingX:=scaleX, scalingY:=-scaleY, offsetX:=zeroX, offsetY:=zeroY
    
    Dim mColor As New SvgColor
    mColor.Init vbaColor:=chrtObject.Chart.Axes(xlValue).MajorGridlines.Border.Color
    Dim mStroke As New SvgStroke
    mStroke.Init strokeColor:=mColor, strokeWidth:=chrtObject.Chart.Axes(xlValue).MajorGridlines.Border.Weight / scaleY
    
    ' X axis labels
    If chrtObject.Chart.ChartType = xlXYScatter Then
        Set mValues = New Collection
        For j = 1 To nVerticalLines
            Call mValues.Add(chrtObject.Chart.Axes(xlCategory).MinimumScale + (j - 1) * chrtObject.Chart.Axes(xlCategory).MajorUnit)
        Next j
        mChartType = "xlXYScatter"
    Else
        mValues = chrtObject.Chart.Axes(xlCategory).CategoryNames
        mChartType = "xlColumnClustered"
    End If
    
    spaceBetween = plotAreaWidth * mGapWidth / nVerticalLines / (1 + mGapWidth)
    offsetFirst = spaceBetween / 2
    If mChartType = "xlXYScatter" Then
        position = plotAreaMinX
    Else
        position = offsetFirst + plotAreaMinX
    End If
    Dim mText As SvgText
    For j = 1 To nVerticalLines
        Set mText = New SvgText
        If mChartType = "xlXYScatter" Then
            mText.Init text:=CStr(mValues(j)), _
                       posX:=position, _
                       posY:=plotAreaHeight + plotAreaMinY + chrtObject.Chart.Axes(xlCategory).TickLabelSpacing, _
                       rotationAngle:=todoRotationAngle, _
                       textAnchor:="middle", _
                       alignmentBaseline:="hanging", _
                       Font:=mFont
        Else
            mText.Init text:=CStr(mValues(j)), _
                       posX:=position + plotAreaWidth / nVerticalLines / (1 + mGapWidth) / 2, _
                       posY:=plotAreaHeight + plotAreaMinY + chrtObject.Chart.Axes(xlCategory).TickLabelSpacing * scaleY / 2, _
                       rotationAngle:=todoRotationAngle, _
                       textAnchor:="middle", _
                       alignmentBaseline:="hanging", _
                       Font:=mFont
        End If
        Call horizontalLabelsCollection.Add(mText, CStr(j))
        
        If mChartType = "xlXYScatter" Then
            position = position + plotAreaWidth / (mValues.Count - 1)
        Else
            position = position + plotAreaWidth / UBound(mValues) / (1 + mGapWidth) + spaceBetween
        End If
    Next j
    
        ' TODO:
'        .Axes (xlCategory)
'    .HasTitle = True
'    .AxisTitle.text = "Position(nm)"  'X-axis title
'    .TickLabels.font.Size = 10 'X-axis coordinate number size
'    .AxisTitle.font.Size = 14  'X-axis title word font size
'    .TickMarkSpacing = 3
    
    ' Horizontal lines and labels:
    Dim hLinePosition!: hLinePosition = mMinimumScale
    While hLinePosition <= mMaximumScale
        Set hLine = New SvgLine
        hLine.Init x1:=0, _
                   y1:=hLinePosition, _
                   x2:=nVerticalLines + 1, _
                   y2:=hLinePosition, _
                   stroke:=mStroke, _
                   Transform:=plotAreaTransform
        Call horizontalLinesCollection.Add(hLine, CStr(hLinePosition))
        
        Set mText = New SvgText
        If InStr(1, chrtObject.Chart.Axes(xlValue).TickLabels.NumberFormat, "%") <> 0 Then
            mText.Init text:=CStr(hLinePosition * 100 & "%"), _
                       posX:=plotAreaMinX - chrtObject.Chart.Axes(xlCategory).TickLabelSpacing * scaleX / 2, _
                       posY:=zeroY - hLinePosition * scaleY, _
                       rotationAngle:=todoRotationAngle, _
                       textAnchor:="end", _
                       alignmentBaseline:="middle", _
                       Font:=mFont
        Else
            mText.Init text:=CStr(hLinePosition), _
                       posX:=plotAreaMinX - chrtObject.Chart.Axes(xlCategory).TickLabelSpacing * scaleX / 2, _
                       posY:=zeroY - hLinePosition * scaleY, _
                       rotationAngle:=todoRotationAngle, _
                       textAnchor:="end", _
                       alignmentBaseline:="middle", _
                       Font:=mFont
        End If
        Call verticalLabelsCollection.Add(mText, CStr(hLinePosition))
        
        hLinePosition = hLinePosition + mMajorUnits
    Wend
    
    If mChartType = "xlXYScatter" Then
        Call mSvgPoints.Init(chrtObject)
    Else
        Call mSvgBars.Init(chrtObject)
    End If
End Sub

Public Function ToString$()
    Dim i&
    
    '    string2save = string2save & vbTab & "<rect x=""" & Transform(plotAreaMinX) & """ y=""" & _
'        Transform(plotAreaMinY) & """ width=""" & _
'        Transform(plotAreaWidth) & """ height=""" & _
'        Transform(plotAreaHeight) & """ style=""fill: white; stroke: black; stroke-width: 0.05;"" />" & vbCrLf
    
    For i = 1 To horizontalLinesCollection.Count
        ToString = ToString & vbTab & vbTab & horizontalLinesCollection(i).ToString
    Next i
    For i = 1 To horizontalLabelsCollection.Count
        ToString = ToString & vbTab & vbTab & horizontalLabelsCollection(i).ToString
    Next i
    For i = 1 To verticalLabelsCollection.Count
        ToString = ToString & vbTab & vbTab & verticalLabelsCollection(i).ToString
    Next i
    
    If mChartType = "xlXYScatter" Then
        ToString = ToString & mSvgPoints.ToString
    Else
        ToString = ToString & mSvgBars.ToString
    End If
End Function

