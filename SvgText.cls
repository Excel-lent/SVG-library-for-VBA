VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvgText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mText$, mPosX!, mPosY!, mRotationAngle!
Private mTextAnchor$            ' "start", "middle", "end"
Private mAlignmentBaseline$     ' "alphabetic", "middle", "hanging"
Private mFont As New SvgFont

' TODO:
'    letter-spacing="30px"

Public Sub Init(text$, posX!, posY!, rotationAngle!, textAnchor$, alignmentBaseline$, Font As SvgFont)
    mText = text
    mPosX = posX
    mPosY = posY
    mRotationAngle = rotationAngle
    mTextAnchor = textAnchor
    mAlignmentBaseline = alignmentBaseline
    Set mFont = Font
    
    If LCase(textAnchor) <> "start" And LCase(textAnchor) <> "middle" And LCase(textAnchor) <> "end" Then
        Call MsgBox("Text anchor """ & textAnchor & """ is not allowed." & vbCrLf & _
                        "Only ""start"", ""middle"" or ""end"" are allowed.", vbCritical, "Critical error!")
    End If
    
    If LCase(alignmentBaseline) <> "alphabetic" And LCase(alignmentBaseline) <> "middle" And LCase(alignmentBaseline) <> "hanging" Then
        Call MsgBox("Bseline allignment """ & alignmentBaseline & """ is not allowed." & vbCrLf & _
                        "Only ""alphabetic"", ""middle"" or ""hanging"" are allowed.", vbCritical, "Critical error!")
    End If
End Sub

Public Function ToString$()
    If Transform(mRotationAngle) = 0 Then
        ToString = "<text text-anchor=""" & mTextAnchor & _
                   """ alignment-baseline=""" & mAlignmentBaseline & _
                   """ transform=""translate(" & Transform(mPosX) & ", " & _
                   Transform(mPosY) & ")"" " & mFont.ToString & ">" & mText & "</text>" & vbCrLf
    Else
        ToString = "<text text-anchor=""" & mTextAnchor & _
                   """ alignment-baseline=""" & mAlignmentBaseline & _
                   """ transform=""translate(" & Transform(mPosX) & ", " & _
                   Transform(mPosY) & ") rotate(" & Transform(mRotationAngle) & ")"" " & mFont.ToString & ">" & mText & "</text>" & vbCrLf
    End If
End Function
