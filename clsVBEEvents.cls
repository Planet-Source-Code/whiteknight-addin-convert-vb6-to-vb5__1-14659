VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVBEEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public WithEvents cMenu As CommandBarEvents
Attribute cMenu.VB_VarHelpID = -1


Private Sub cMenu_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

    'Do the appropriate action, depending on the menu item's parameter value
    Select Case CommandBarControl.Parameter
        Case "About"
            'frmAbout.show vbmodal
            frmmain.Show
        Case "Check"
            'CheckFolder
            MsgBox "Check For Folder"
        Case "Convert"
            'ConvertPrj
            MsgBox "Convert Project File"
        Case "Options"
            'ShowOptions
            'frmmain.Show vbSystemModal
    End Select

    'We handled it OK
    handled = True

End Sub

