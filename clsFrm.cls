VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents mFrm As Form
Attribute mFrm.VB_VarHelpID = -1
Private Const mcstrEvProc As String = "[Event Procedure]"
Private mcolCtls As Collection
'
'This is our initialization code for this class
'A pointer to the form is passed in and stored in the header
'
Function fInit(lFrm As Form)
    'We are going to print to the debug window
    'to watch code run so clear the debug window before we begin
    'ClearImmediateWindow
    assDebugPrint "clsFrm.fInit"
    Set mFrm = lFrm
    mFrm.BeforeUpdate = mcstrEvProc
    mFrm.OnClose = mcstrEvProc
    fCtlScanner
End Function

Private Sub Class_Initialize()
    assDebugPrint "clsFrm.Class_Initialize"
    Set mcolCtls = New Collection
End Sub

'
'This is our cleanup code for this class.
'In here we set the pointer to the form to nothing
'
Private Sub Class_Terminate()
    assDebugPrint "clsFrm.Class_Terminate"
    Set mFrm = Nothing
    Set mcolCtls = Nothing
End Sub

Private Sub mfrm_BeforeUpdate(Cancel As Integer)
    assDebugPrint "Before Update: " & mFrm.Name
End Sub

Property Get pFrmName() As String
    pFrmName = mFrm.Name
End Property

Private Sub mfrm_Close()
    assDebugPrint "clsFrm.mfrm_Close"
    Set mFrm = Nothing
End Sub

Private Function fCtlScanner()
Dim ctl As Control
    For Each ctl In mFrm.Controls
        assDebugPrint "clsFrm.fCtlScanner: " & ctl.Name
        Select Case ctl.ControlType
            Case acCheckBox
'                Dim lclsCtlChk As clsCtlChk
'                Set lclsCtlChk = New clsCtlChk
'                lclsCtlChk.fInit ctl
'                mcolCtls.Add lclsCtlChk, ctl.Name
            Case acComboBox
                Dim lclsCtlCbo As clsCtlCbo
                Set lclsCtlCbo = New clsCtlCbo
                lclsCtlCbo.fInit ctl
                mcolCtls.Add lclsCtlCbo, ctl.Name
'            Case acCommandButton
'                Dim lclsCtlCmd As clsCtlCmd
'                Set lclsCtlCmd = New clsCtlCmd
'                lclsCtlCmd.fInit Me, mFrm, ctl
'                'Store a pointer to the new class in our collection
'                mcolCtls.Add lclsCtlCmd, ctl.Name
            Case acListBox
'                Dim lclsCtlLst As clsCtlLst
'                Set lclsCtlLst = New clsCtlLst
'                lclsCtlLst.fInit Me, mFrm, ctl
'                'Store a pointer to the new class in our collection
'                mcolCtls.Add lclsCtlLst, ctl.Name
            Case acOptionButton
'                Dim lclsCtlCmd As clsCtlOpt
'                Set lclsCtlOpt = New clsCtlOpt
'                lclsCtlOpt.fInit Me, mFrm, ctl
'                'Store a pointer to the new class in our collection
'                mcolCtls.Add lclsCtlOpt, ctl.Name
            Case acOptionGroup
'                Dim lclsCtlGrp As clsCtlGrp
'                Set lclsCtlGrp = New clsCtlGrp
'                lclsCtlGrp.fInit Me, mFrm, ctl
'                'Store a pointer to the new class in our collection
'                mcolCtls.Add lclsCtlGrp, ctl.Name
            Case acTabCtl 'tab pages are handled in the tab control
'                Dim lclsTab As clsTab
'                Set lclsTab = New clsTab
'                lclsTab.fInit Me, mFrm, ctl
'                'Store a pointer to the new class in our collection
'                mcolCtls.Add lclsTab, ctl.Name
            Case acPage
'                Dim lclsCtlPg As clsCtlPg
'                Set lclsCtlPg = New clsCtlPg
'                lclsCtlPg.fInit Me, mFrm, ctl
'                'Store a pointer to the new class in our collection
'                mcolCtls.Add lclsCtlPg, ctl.Name
            Case acSubform 'subform controls
'                Dim lclsSFrm As clsSFrm
'                Set lclsSFrm = New clsSFrm
'                lclsSFrm.fInit Me, mFrm, ctl
'                'Store a pointer to the new class in our collection
'                mcolCtls.Add lclsSFrm, ctl.Name
            Case acTextBox 'Find all text boxes and load class to change backcolor
'                Dim lclsTxt As clsTxt
'                Set lclsTxt= New clsTxt
'                lclsTxt.fInit Me, mFrm, ctl
'                'Store a pointer to the new class in our collection
'                mcolCtls.Add lclsTxt, ctl.Name
            Case acToggleButton
'                Dim lclsTgl As clsTgl
'                Set lclsTgl= New clsTgl
'                lclsTgl.fInit Me, mFrm, ctl
'                'Store a pointer to the new class in our collection
'                mcolCtls.Add lclsTgl, ctl.Name
         End Select
    Next ctl
End Function
