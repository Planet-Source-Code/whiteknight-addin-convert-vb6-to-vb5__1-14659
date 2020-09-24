VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDialogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Declare Function MessageBoxEx Lib "user32" _
    Alias "MessageBoxExA" _
    (ByVal hwnd As Long, _
    ByVal lpText As String, _
    ByVal lpCaption As String, _
    ByVal uType As Long, _
    ByVal wLanguageID As Long) As Long
Private Declare Function MessageBoxIndirect Lib "user32" _
    Alias "MessageBoxIndirectA" _
    (lpMsgBoxParams As MSGBOXPARAMS) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" _
    Alias "GetOpenFileNameA" _
    (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" _
    Alias "GetSaveFileNameA" _
    (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
    Alias "SHBrowseForFolderA" _
    (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
    Alias "SHGetPathFromIDListA" _
    (ByVal pidl As Long, _
    ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function ChooseColor Lib "comdlg32.dll" _
    Alias "ChooseColorA" _
    (pChooseColor As udtCHOOSECOLOR) As Long
Private Declare Function PrintDlg Lib "comdlg32.dll" _
    Alias "PrintDlgA" _
    (pPrintDlg As udtPRINTDLG) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" _
    Alias "ChooseFontA" _
    (pChooseFont As udtCHOOSEFONT) As Long
Private Declare Function WNetConnectionDialog Lib "mpr.dll" _
    (ByVal hwnd As Long, _
    ByVal dwType As Long) As Long
Private Declare Function WNetDisconnectDialog Lib "mpr.dll" _
    (ByVal hwnd As Long, _
    ByVal dwType As Long) As Long
    
'WNet Dialogs
Public Enum ConnectionType
    RESOURCETYPE_DISK = &H1
    RESOURCETYPE_PRINT = &H2
End Enum

Private Type udtCHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String      'Long in the API Viewer
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Const CC_ANYCOLOR = &H100
Private Const CC_FULLOPEN = &H2

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type udtPRINTDLG
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Private Type MSGBOXPARAMS
    cbSize As Long
    hwndOwner As Long
    hInstance As Long
    lpszText As String
    lpszCaption As String
    dwStyle As Long
    lpszIcon As String
    dwContextHelpId As Long
    lpfnMsgBoxCallback As Long
    dwLanguageId As Long
End Type

'Font Constants
Private Const LF_FACESIZE = 32
Private Const BOLD_FONTTYPE = &H100
Private Const DEVICE_FONTTYPE = &H2
Private Const ITALIC_FONTTYPE = &H200
Private Const PRINTER_FONTTYPE = &H4000
Private Const RASTER_FONTTYPE = &H1
Private Const REGULAR_FONTTYPE = &H400
Private Const SCREEN_FONTTYPE = &H2000
Private Const SIMULATED_FONTTYPE = &H8000
Private Const TRUETYPE_FONTTYPE = &H4
Private Const CF_PRINTERFONTS = &H2
Private Const CF_SCREENFONTS = &H1
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_USESTYLE = &H80&
Private Const CF_EFFECTS = &H100&
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const CF_NOSCRIPTSEL = &H800000

Private Type udtCHOOSEFONT
    lStructSize As Long
    hwndOwner As Long          '  caller's window handle
    hdc As Long                '  printer DC/IC or NULL
    lpLogFont As Long
    iPointSize As Long         '  10 * size in points of selected font
    flags As Long              '  enum. type flags
    rgbColors As Long          '  returned text color
    lCustData As Long          '  data passed to hook fn.
    lpfnHook As Long           '  ptr. to hook function
    lpTemplateName As String     '  custom template name
    hInstance As Long          '  instance handle of.EXE that
                                   '    contains cust. dlg. template
    lpszStyle As String          '  return the style field here
                                   '  must be LF_FACESIZE or bigger
    nFontType As Integer          '  same value reported to the EnumFonts
                                   '    call back with the extra FONTTYPE_
                                   '    bits added
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long           '  minimum pt size allowed &
    nSizeMax As Long           '  max pt size allowed if
                                   '    CF_LIMITSIZE is used
End Type

Private Type udtLOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * LF_FACESIZE
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Function GetFolder(Optional Title As String, Optional hwnd) As String
    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim folder As String
    
    folder = Space$(255)
    
    With bi
        If IsNumeric(hwnd) Then .hOwner = hwnd
        .ulFlags = BIF_RETURNONLYFSDIRS
        .pidlRoot = 0
        If Title <> "" Then
            .lpszTitle = Title & Chr$(0)
        Else
            .lpszTitle = "Select a Folder"
        End If
    End With
    
    pidl = SHBrowseForFolder(bi)
    
    If SHGetPathFromIDList(ByVal pidl, ByVal folder) Then
        GetFolder = Left(folder, InStr(folder, Chr$(0)) - 1)
    Else
        GetFolder = ""
    End If
    
    CoTaskMemFree pidl
End Function
Public Function Connect(Mode As ConnectionType, Optional hwnd As Long) As Long
    Dim rc As Long
    
    If IsNumeric(hwnd) Then
        rc = WNetConnectionDialog(hwnd, Mode)
    Else
        rc = WNetConnectionDialog(0, Mode)
    End If
End Function

Public Function Disconnect(Mode As ConnectionType, Optional hwnd As Long) As Long
    Dim rc As Long
    
    If IsNumeric(hwnd) Then
        rc = WNetDisconnectDialog(hwnd, Mode)
    Else
        rc = WNetDisconnectDialog(0, Mode)
    End If
End Function


Public Function FileOpen(Optional Title As String, _
        Optional StartPath As String, _
        Optional Filter As String, _
        Optional FilterIndex As Long, _
        Optional hwnd) As String
        
    Dim rc As Long
    Dim pOpenfilename As OPENFILENAME
    Const MAX_BUFFER_LENGTH = 256
    
    With pOpenfilename
        If IsNumeric(hwnd) Then .hwndOwner = hwnd
        .hInstance = App.hInstance
        If Title <> "" Then
            .lpstrTitle = Title
        Else
            .lpstrTitle = "Open"
        End If
        If StartPath <> "" Then
            .lpstrInitialDir = StartPath
        Else
            .lpstrInitialDir = App.Path
        End If
        If Filter <> "" Then
            .lpstrFilter = Filter
        Else
            .lpstrFilter = "All Files" & Chr$(0) & "*.*" & Chr$(0)
        End If
        If IsNumeric(FilterIndex) Then .nFilterIndex = FilterIndex
        .lpstrFile = String(MAX_BUFFER_LENGTH, 0)
        .nMaxFile = MAX_BUFFER_LENGTH - 1
        .lpstrFileTitle = .lpstrFile
        .nMaxFileTitle = MAX_BUFFER_LENGTH
        .lStructSize = Len(pOpenfilename)
    End With
    
    rc = GetOpenFileName(pOpenfilename)
    
    If rc <> 0 Then
        'A file selected
        FileOpen = Left$(pOpenfilename.lpstrFile, pOpenfilename.nMaxFile)
    Else
        'The cancel button was pressed
        FileOpen = ""
    End If
End Function

Public Function FilePrint(Copies As Integer, FromPage As Integer, ToPage As Integer, MinPage As Integer, MaxPage As Integer, Optional hwnd As Long) As Long
    Dim rc As Long
    Dim pPrintDlg As udtPRINTDLG
    
    With pPrintDlg
        If IsNumeric(hwnd) Then
            .hwndOwner = hwnd
        Else
            .hwndOwner = 0
        End If
        .flags = 0
        .hInstance = App.hInstance
        .nCopies = Copies
        .nFromPage = FromPage
        .nToPage = ToPage
        .nMinPage = MinPage
        .nMaxPage = MaxPage
        .lStructSize = Len(pPrintDlg)
    End With
    
    'Call the API
    rc = PrintDlg(pPrintDlg)
    If rc = 0 Then
        'Fetch the settings
        With pPrintDlg
            Copies = .nCopies
            FromPage = .nFromPage
            ToPage = .nToPage
            MinPage = .nMinPage
            MaxPage = .nMaxPage
        End With
    Else
    End If
End Function
Public Function FileSave(Optional Title As String, _
        Optional StartPath As String, _
        Optional DefaultExtension As String, _
        Optional Filter As String, _
        Optional FilterIndex As Long, _
        Optional hwnd) As String
    
    Dim rc As Long
    Dim pOpenfilename As OPENFILENAME
    Const MAX_BUFFER_LENGTH = 256
    
    With pOpenfilename
        If IsNumeric(hwnd) Then .hwndOwner = hwnd
        .hInstance = App.hInstance
        If Title <> "" Then
            .lpstrTitle = Title
        Else
            .lpstrTitle = "Save"
        End If
        If StartPath <> "" Then
            .lpstrInitialDir = StartPath
        Else
            .lpstrInitialDir = App.Path
        End If
        If Filter <> "" Then
            .lpstrFilter = Filter
        Else
            .lpstrFilter = "All Files" & Chr$(0) & "*.*" & Chr$(0)
        End If
        If DefaultExtension <> "" Then .lpstrDefExt = DefaultExtension
        If IsNumeric(FilterIndex) Then .nFilterIndex = FilterIndex
        .lpstrFile = String(MAX_BUFFER_LENGTH, 0)
        .nMaxFile = MAX_BUFFER_LENGTH - 1
        .lpstrFileTitle = .lpstrFile
        .nMaxFileTitle = MAX_BUFFER_LENGTH
        .lStructSize = Len(pOpenfilename)
    End With
    
    rc = GetSaveFileName(pOpenfilename)
    
    If rc <> 0 Then
        'A file selected
        FileSave = Left$(pOpenfilename.lpstrFile, pOpenfilename.nMaxFile)
    Else
        'The cancel button was pressed
        FileSave = ""
    End If
End Function


Public Function GetColor(DefaultColor As Long, Optional hwnd As Long) As Long
    Dim rc As Long
    Dim pChooseColor As udtCHOOSECOLOR
    Dim CustomColors() As Byte
    
    'Initailize the UDT for the color dialog
    With pChooseColor
        If IsNumeric(hwnd) Then
            .hwndOwner = hwnd
        Else
            .hwndOwner = 0
        End If
        .hInstance = 0
        .lpCustColors = StrConv(CustomColors, vbUnicode)
        .flags = 0
        .lStructSize = Len(pChooseColor)
    End With
    
    'Call the API
    rc = ChooseColor(pChooseColor)
          
    'Return the RGB value of the color
    If rc Then
        GetColor = pChooseColor.rgbResult
    Else
        GetColor = DefaultColor
    End If
End Function




Public Function GetFont(Optional FontName As String, _
        Optional Size As Integer, _
        Optional Bold As Boolean, _
        Optional Italic As Boolean, _
        Optional Underline As Boolean, _
        Optional Strikeout As Boolean, _
        Optional Color As Long, _
        Optional hwnd) As Long
    
    Dim rc As Long
    Dim pChooseFont As udtCHOOSEFONT
    Dim pLogFont As udtLOGFONT
    
    'Initailize the buffer
    With pLogFont
        .lfFaceName = FontName & Chr$(0)
        .lfItalic = Italic
        .lfUnderline = Underline
        .lfStrikeOut = Strikeout
    End With
    
    'Initialize the structure
    With pChooseFont
        .hInstance = App.hInstance
        If IsNumeric(hwnd) Then .hwndOwner = hwnd
        .flags = CF_BOTH + CF_INITTOLOGFONTSTRUCT + _
            CF_EFFECTS + CF_NOSCRIPTSEL
            
        If IsNumeric(Size) Then .iPointSize = -(Size * 10)
        If Bold Then .nFontType = .nFontType + BOLD_FONTTYPE
        If Italic Then .nFontType = .nFontType + ITALIC_FONTTYPE
        If IsNumeric(Color) Then .rgbColors = Color
        
        .lStructSize = Len(pChooseFont)
        .lpLogFont = VarPtr(pLogFont)
    End With
    
    'Call the API
    rc = ChooseFont(pChooseFont)
    
    If rc <> 0 Then
        'Success!
        FontName = StrConv(pLogFont.lfFaceName, vbUnicode)
        FontName = Left$(FontName, InStr(FontName, vbNullChar) - 1)
        
        'Return it's properties
        With pChooseFont
            Size = .iPointSize / 10
            Bold = (.nFontType And BOLD_FONTTYPE)
            Italic = (.nFontType And ITALIC_FONTTYPE)
            Underline = (pLogFont.lfUnderline)
            Strikeout = (pLogFont.lfStrikeOut)
        End With
        
        'Return the font name
        GetFont = rc
    Else
        'The user clicked cancel
        GetFont = 0
    End If
End Function

Public Function YesNoBox(Message As String, Caption As String, Optional hwndOwner As Long) As Long
    Dim rc As Long
    Dim hwnd As Long
    Dim wLanguageID As Long
    
    If IsNumeric(hwndOwner) Then
        hwnd = hwndOwner
    Else
        hwnd = 0
    End If
    
    YesNoBox = MessageBoxEx(hwnd, _
            Message, _
            Caption, _
            vbYesNo + vbQuestion, _
            wLanguageID)
End Function

Public Function WarningBox(Message As String, Caption As String, Optional hwndOwner As Long) As Long
    Dim rc As Long
    Dim hwnd As Long
    Dim wLanguageID As Long
    
    If IsNumeric(hwndOwner) Then
        hwnd = hwndOwner
    Else
        hwnd = 0
    End If
    
    WarningBox = MessageBoxEx(hwnd, _
            Message, _
            Caption, _
            vbExclamation, _
            wLanguageID)
End Function
Public Function ErrorBox(Message As String, Caption As String, Optional hwndOwner As Long) As Long
    Dim rc As Long
    Dim hwnd As Long
    Dim wLanguageID As Long
    
    If IsNumeric(hwndOwner) Then
        hwnd = hwndOwner
    Else
        hwnd = 0
    End If
    
    ErrorBox = MessageBoxEx(hwnd, _
            Message, _
            Caption, _
            vbOK + vbCritical, _
            wLanguageID)
End Function

