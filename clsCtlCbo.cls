VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCtlCbo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents mctlCbo As ComboBox
Attribute mctlCbo.VB_VarHelpID = -1
Private Const mcstrEvProc As String = "[Event Procedure]"
Private mInitialBackColor As Long
Private Const mcstrModuleName = "clsCtlCbo"
Private Const cBackColorChange As Long = vbCyan

Function fInit(lctlCbo As ComboBox)
    'Store a pointer to the combo control passed in
    Set mctlCbo = lctlCbo
    assDebugPrint mcstrModuleName & ".fInit"
    'Store the initial back color in fInit
    mInitialBackColor = mctlCbo.BackColor
    mctlCbo.BeforeUpdate = mcstrEvProc
    mctlCbo.AfterUpdate = mcstrEvProc
    mctlCbo.OnGotFocus = mcstrEvProc
    mctlCbo.OnLostFocus = mcstrEvProc

End Function

Private Sub Class_Initialize()
    assDebugPrint "clsCbo.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    assDebugPrint "clsCbo.Class_Terminate"
    Set mctlCbo = Nothing
End Sub

Private Sub mctlCbo_AfterUpdate()
    assDebugPrint "AfterUpdate: " & mctlCbo.Name
End Sub

Private Sub mctlCbo_BeforeUpdate(Cancel As Integer)
    assDebugPrint "BeforeUpdate: " & mctlCbo.Name
End Sub

Private Sub mctlCbo_GotFocus()
    assDebugPrint "GotFocus: " & mctlCbo.Name
    mctlCbo.BackColor = cBackColorChange
End Sub

Private Sub mctlCbo_LostFocus()
    assDebugPrint "LostFocus: " & mctlCbo.Name
    mctlCbo.BackColor = mInitialBackColor
End Sub
