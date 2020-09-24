VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "My Add-In"
Option Explicit
Option Compare Text
Option Base 1

Implements IDTExtensibility

'Define some constants for the text of the menu items
Const msMENU_TEXT As String = "&Six2Five"
Const msMENU_ABOUT As String = "&About_Six2Five"
Const msMENU_CHECK As String = "&Check Folder"
Const msMENU_CONVERT As String = "Con&vert Project File"
Const msMENU_OPTIONS As String = "&Options"
'Const msMENU_ABOUT As String = "About Smart Indenter"

'Define an array to store the menu item click event handlers
Dim vbeMenus(1 To 9) As New clsVBEEvents


'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)

    Dim oCtl As CommandBarControl

    On Error GoTo error_handler

    'save the vb instance
    Set VBInstance = VBInst

    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods
    Debug.Print VBInst.FullName

    'SetUpCodeWindowMenus
    'SetUpEditMenus

    Set oCtl = VBInstance.CommandBars("Add-Ins").Controls.Add(Type:=msoControlButton, Parameter:="About", temporary:=True)
    oCtl.Caption = msMENU_TEXT
    oCtl.FaceId = 135

    Set vbeMenus(9).cMenu = VBInstance.Events.CommandBarEvents(oCtl)

    Exit Sub

error_handler:

    MsgBox Err.Description

End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbext_DisconnectMode, custom() As Variant)

    On Error Resume Next

    Dim i As Integer

    'Ignore errors in case the menu items don't exist
    On Error Resume Next

    'Clear down the class handler objects
    For i = 1 To 8
        Set vbeMenus(i).cMenu = Nothing
    Next

    'Remove our menus from the VBE Edit menu
    VBInstance.CommandBars(1).FindControl(msoControlPopup, 30003).Controls(msMENU_TEXT).Delete
    VBInstance.CommandBars("Add-Ins").Controls(msMENU_TEXT).Delete
    'VBInstance.CommandBars("Code Window").Controls(msMENU_TEXT).Delete

End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)

End Sub
