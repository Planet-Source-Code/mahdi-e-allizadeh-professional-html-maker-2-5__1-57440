VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form11 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   7590
   ClientLeft      =   2955
   ClientTop       =   2205
   ClientWidth     =   5085
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check4 
      Caption         =   "Disable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   21
      Top             =   6600
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      Caption         =   " Do you want to see the hidden file in file list ? "
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Width           =   4815
      Begin VB.OptionButton Option4 
         Caption         =   " Yes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   " No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Height          =   135
      Left            =   0
      TabIndex        =   17
      Top             =   4560
      Width           =   5535
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   0
      TabIndex        =   16
      Top             =   1800
      Width           =   5535
   End
   Begin MSComCtl2.FlatScrollBar HScroll2 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Appearance      =   2
      Arrows          =   65536
      LargeChange     =   30
      Max             =   255
      Orientation     =   1179649
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4440
      Top             =   6600
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   2880
      ScaleHeight     =   1395
      ScaleWidth      =   1995
      TabIndex        =   12
      Top             =   240
      Width           =   2055
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Default Back Grounf Color"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   240
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   " Set this settings "
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   4815
      Begin VB.CheckBox Check2 
         Caption         =   "Show Toolbar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show Status Bar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Select the internet browser "
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   4815
      Begin VB.OptionButton Option2 
         Caption         =   "MS Internet Explorer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "M2A Web  Browser"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   6240
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Enabled         =   0   'False
      Height          =   2340
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   1815
   End
   Begin MSComCtl2.FlatScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Appearance      =   2
      Arrows          =   65536
      LargeChange     =   30
      Max             =   255
      Orientation     =   1179649
   End
   Begin MSComCtl2.FlatScrollBar HScroll3 
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Appearance      =   2
      Arrows          =   65536
      LargeChange     =   30
      Max             =   255
      Orientation     =   1179649
   End
   Begin M2AHTMLMaker.chameleonButton Command2 
      Height          =   375
      Left            =   3600
      TabIndex        =   22
      Top             =   7080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form11.frx":058A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin M2AHTMLMaker.chameleonButton Command1 
      Height          =   375
      Left            =   2280
      TabIndex        =   23
      Top             =   7080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form11.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      Caption         =   $"Form11.frx":05C2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2160
      TabIndex        =   4
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Back Ground Color "
      Height          =   255
      Left            =   3240
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public For_Picture_Color_Change As Variant
Private Sub Check1_Click()
 If Check1.Value = Checked Then Form1.sbStatusBar.Visible = True Else Form1.sbStatusBar.Visible = False
 SaveSetting App.Title, "Settings", "StatuseBar", Check1.Value
End Sub

Private Sub Check2_Click()
 If Check2.Value = Checked Then Form1.tbToolBar.Visible = True Else Form1.tbToolBar.Visible = False
 SaveSetting App.Title, "Settings", "ToolBar", Check2.Value
End Sub

Private Sub Check3_Click()
 If Check3.Value = Checked Then
  Picture1.BackColor = vbButtonFace
  Form1.BackColor = vbButtonFace
  HScroll1.Enabled = False
  HScroll2.Enabled = False
  HScroll3.Enabled = False
  Timer1.Enabled = False
 End If
 If Check3.Value = Unchecked Then
  HScroll1.Value = 255
  HScroll2.Value = 255
  HScroll3.Value = 255
  
  HScroll1.Enabled = True
  HScroll2.Enabled = True
  HScroll3.Enabled = True
  Timer1.Enabled = True
  Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
  Form1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
 End If
End Sub

Private Sub Check4_Click()
If Check4.Value = Unchecked Then
 Text1.Enabled = True
 Drive1.Enabled = True
 Dir1.Enabled = True
Else
 Text1.Enabled = False
 Drive1.Enabled = False
 Dir1.Enabled = False
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
 If Check4.Value = Unchecked Then
  SaveSetting App.Title, "Settings", "ForStartAddress", Text1.Text
 End If
 Me.Hide
 Form1.Show
 If Form1.RichTextBox3.Visible = True Then Form1.RichTextBox3.SetFocus
 Check4.Value = Checked
 Text1.Enabled = False
 Drive1.Enabled = False
 Dir1.Enabled = False
End Sub

Private Sub Command2_Click()
 Me.Hide
 Form1.Show
End Sub

Private Sub Dir1_Change()
 Text1.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
 Dir1.Path = Drive1.Drive
 Text1.Text = Dir1.Path
End Sub

Private Sub Form_Load()
On Error Resume Next
 Me.Left = (Screen.Width - Form1.Width) / 2
 Me.Top = (Screen.Height - Form1.Height) / 2
 Call FormOnTop(Me.hWnd, True)
ForHidden = GetSetting(App.Title, "Settings", "ForHidden")
If ForHidden = "No" Then Option3.Value = True
If ForHidden = "Yes" Then Option4.Value = True
Check4.Value = Checked
Text1.Enabled = False
Drive1.Enabled = False
Dir1.Enabled = False
If YourFavorite = 0 Then Option1.Value = True
If YourFavorite = 1 Then Option2.Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Timer1.Enabled = False
 For_save_The_Text1 = 0
End Sub

Private Sub HScroll1_Change()
 Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll2_Change()
 Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll3_Change()
 Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub Option1_Click()
 YourFavorite = 0
 SaveSetting App.Title, "Settings", "ForInternetBrowser", "0"
End Sub

Private Sub Option2_Click()
 YourFavorite = 1
 SaveSetting App.Title, "Settings", "ForInternetBrowser", "1"
End Sub

Private Sub Option3_Click()
Form1.File1.Hidden = False
SaveSetting App.Title, "Settings", "ForHidden", "No"
End Sub

Private Sub Option4_Click()
Form1.File1.Hidden = True
SaveSetting App.Title, "Settings", "ForHidden", "Yes"
End Sub

Private Sub Text1_Change()
For_save_The_Text1 = 1
End Sub

Private Sub Timer1_Timer()
If HScroll3.Enabled = True Then
 If Picture1.BackColor <> For_Picture_Color_Change Then
  Form1.BackColor = Picture1.BackColor
  For_Picture_Color_Change = Picture1.BackColor
 End If
End If
End Sub

