VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Six 2 Five"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4650
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txt_index 
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Auto-Save"
      Height          =   2655
      Left            =   5400
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton bttn_autoOFF 
         Caption         =   "Auto-Save Off"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton bttn_autoON 
         Caption         =   "Auto-Save On"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox auto_interval 
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "10"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Interval (in minutes)"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.ListBox lst_proj 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   3135
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   4680
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton bttn_check 
      Caption         =   "Check Folder For VB 6 Project Files"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Visual Basic 6 Project files found"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public VBInstance As VBIDE.VBE
Public Connect As Connect
Public dlg As clsDialogs
Option Explicit

Private Sub CancelButton_Click()
    'Connect.Hide
End Sub

Private Sub bttn_check_Click()
    Dim x As Integer
    folder = dlg.GetFolder("Six 2 Five : Select A Folder", Me.hWnd)
    File1.path = folder
    lst_proj.Clear
    For x = 0 To File1.ListCount - 1
        If LCase(Right(File1.List(x), 3)) = "vbp" Then
            If IsVB6(folder & "\" & File1.List(x)) = True Then
                lst_proj.AddItem File1.List(x)
                'Text1.Text = Text1.Text & vbCrLf & ReadFile(folder & "\" & File1.List(x))
            End If
        End If
    Next x
    Me.Caption = "Six 2 Five : Found " & lst_proj.ListCount & " Visual Basic 6 Project File(s)"
End Sub

Private Sub Command1_Click()
'VBInstance.ActiveCodePane.CodeModule.InsertLines VBInstance.ActiveCodePane.TopLine, "'HELLO THERE WORLD!"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Set dlg = New clsDialogs
    'TopMost Me
End Sub

Private Sub Form_Resize()
If WindowState = vbMinimized Then Exit Sub
Width = 4770
Height = 1920
End Sub

Private Sub lst_proj_Click()
    Dim ret As VbMsgBoxResult
    ret = MsgBox(InvalidMsg(folder & "\" & lst_proj.List(lst_proj.ListIndex)), vbExclamation + vbApplicationModal + vbYesNo + vbDefaultButton1, App.Title)
    Select Case ret
        Case vbYes
            Screen.MousePointer = vbHourglass
            Convert (folder & "\" & lst_proj.List(lst_proj.ListIndex))
            Screen.MousePointer = vbDefault
    End Select
End Sub

Private Sub OKButton_Click()
    'MsgBox "AddIn operation on: " & VBInstance.FullName
    Unload Me
End Sub
