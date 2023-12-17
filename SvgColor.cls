VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvgColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mR&, mG&, mB&

Public Sub Init(vbaColor&)
    mR = vbaColor Mod 256
    mG = (vbaColor \ 256) Mod 256
    mB = vbaColor \ 65536
End Sub

Public Function ToString$()
    ToString = "rgb(" & mR & "," & mG & "," & mB & ")"
End Function
