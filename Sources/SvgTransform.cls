VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvgTransform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mScalingX!
Private mScalingY!
Private mOffsetX!
Private mOffsetY!

Public Sub Init(scalingX!, scalingY!, offsetX!, offsetY!)
    mScalingX = scalingX
    mScalingY = scalingY
    mOffsetX = offsetX
    mOffsetY = offsetY
End Sub

Public Function ToString$()
    ToString = "transform=""matrix(" & Transform(mScalingX) & " 0 0 " & Transform(mScalingY) & " " & Transform(mOffsetX) & " " & Transform(mOffsetY) & ")"""
End Function
