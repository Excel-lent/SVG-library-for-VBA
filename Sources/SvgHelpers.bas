Attribute VB_Name = "SvgHelpers"
Option Explicit

Const Accuracy& = 3

Public Function Transform$(x!)
    Transform = Replace(CStr(ReduceAccuracy(x)), ",", ".")
End Function

' Float has too much numbers that actually not necessary (for example -0.1552287 can be reduced to -0.1552 without any visible lost).
Private Function ReduceAccuracy!(x!)
    Dim tmp!
    
    If x = 0 Then
        ReduceAccuracy = 0
        Exit Function
    End If
    
    tmp = x * 10 ^ (Accuracy - Int(WorksheetFunction.Log10(Abs(x))) - 1)
    ReduceAccuracy = Round(tmp) * 10 ^ (Int(WorksheetFunction.Log10(Abs(x))) + 1 - Accuracy)
End Function

' 193.3769
' 3.221054; 20.78495 32.05583 193.3769
' ; 15.99705 -13.12514 43.70559 223.6354
' 256.4482
' -13.12514 31.16245 223.6354

' ReduceAccuracy test cases
Private Sub ReduceAccuracyTestCases()
    Call AssertEqual(ReduceAccuracy(22.86345), 22.86)           ' 22.86345 -> 22.86
    Call AssertEqual(ReduceAccuracy(-22.86345), -22.86)         ' -22.86345 -> -22.86
    Call AssertEqual(ReduceAccuracy(52.06768), 52.07)           ' 52.06768 -> 52.07
    Call AssertEqual(ReduceAccuracy(-52.06768), -52.07)         ' -52.06768 -> -52.07
    Call AssertEqual(ReduceAccuracy(0.1552987), 0.1553)         ' 0.1552987 -> 0.1553
    Call AssertEqual(ReduceAccuracy(-0.1552987), -0.1553)       ' -0.1552987 -> -0.1553
    Call AssertEqual(ReduceAccuracy(0.05714224), 0.05714)       ' 5.714224E-02 -> 5.714E-02
    Call AssertEqual(ReduceAccuracy(-0.05714224), -0.05714)     ' -5.714224E-02 -> -5.714E-02
End Sub

Private Sub AssertEqual(x1!, x2!)
    If x1 <> x2 Then x1 = x2 / 0
End Sub
