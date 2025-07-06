VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'
'If your version of Office is 2010 or later leave this true
'else set this constant false
'
#Const VBA7 = True

'This timer class is based on classic Windows API usage patterns adapted for VBA.
'
#If VBA7 Then 'A2010 or later (32-bit/64-bit)
Private Declare PtrSafe Function apiGetTime Lib "winmm.dll" Alias "timeGetTime" () As Long

Private lngStartTime As Long
#Else 'VBA6 (A2007 or earlier
Private Declare Function apiGetTime Lib "winmm.dll" Alias "timeGetTime" () As Long

Private lngStartTime As Long
#End If

Private Sub Class_Initialize()
    assDebugPrint "Class_Initialize"
    StartTimer
End Sub

Private Sub Class_Terminate()
    assDebugPrint "Class_Terminate: " & EndTimer()
End Sub

Function EndTimer()
    EndTimer = apiGetTime() - lngStartTime
End Function

Function StartTimer()
    lngStartTime = apiGetTime()
End Function


