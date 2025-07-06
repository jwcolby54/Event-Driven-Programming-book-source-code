Attribute VB_Name = "basTest"
Option Compare Database
Option Explicit

Function TmrTest()
Dim lngCtr1 As Long
Dim lngCtr2 As Long
Dim clsTmr1 As clsTimer
Dim clsTmr2 As clsTimer
Set clsTmr1 = New clsTimer
Set clsTmr2 = New clsTimer

    For lngCtr1 = 1 To 5
        For lngCtr2 = 1 To 1000000
            Pi
        Next lngCtr2
        assDebugPrint clsTmr2.EndTimer
    Next lngCtr1
    assDebugPrint clsTmr1.EndTimer
End Function
Function Pi() As Double
Dim dblPi As Double
    dblPi = 4 * Atn(1)
    Pi = dblPi
End Function

