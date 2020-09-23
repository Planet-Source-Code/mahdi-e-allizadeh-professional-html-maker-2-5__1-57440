VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Form8 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Sound"
   ClientHeight    =   4710
   ClientLeft      =   4140
   ClientTop       =   2580
   ClientWidth     =   3465
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin M2AHTMLMaker.chameleonButton Command1 
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Insert"
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
      MICON           =   "Form8.frx":058A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Hidden"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
   End
   Begin M2AHTMLMaker.chameleonButton Command2 
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Dont Insert"
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
      MICON           =   "Form8.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin M2AHTMLMaker.chameleonButton Command4 
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "g"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
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
      MICON           =   "Form8.frx":05C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin M2AHTMLMaker.chameleonButton Command3 
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "4"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
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
      MICON           =   "Form8.frx":05DE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Hid As Variant
Private Sub Command1_Click()
    If Check1.Value = Checked Then
     Hid = "Hidden=""True"""
    Else
     Hid = "Hidden=""False"""
    End If
    ss = Form1.RichTextBox3.SelLength
    Form1.RichTextBox3.SelStart = Form1.RichTextBox3.SelStart
    Form1.RichTextBox3.SelText = "<EMBED SRC="
    Form1.RichTextBox3.SelText = """"
    Form1.RichTextBox3.SelText = Text1.Text
    Form1.RichTextBox3.SelText = """"
    Form1.RichTextBox3.SelText = " "
    Form1.RichTextBox3.SelText = Hid
    'Form1.RichTextBox3.SelText = """"
    'Form1.RichTextBox3.SelText = "False"
    'Form1.RichTextBox3.SelText = """"
    Form1.RichTextBox3.SelText = ">"
    Form1.RichTextBox3.Visible = False
    Form1.RichTextBox3.Visible = True
    Me.Hide
    Form1.Show
    If Form1.RichTextBox3.Visible = True Then Form1.RichTextBox3.SetFocus
    Command4_Click
End Sub

Private Sub Command2_Click()
Command4_Click
Me.Hide
Form1.Show
If Form1.RichTextBox3.Visible = True Then Form1.RichTextBox3.SetFocus
End Sub

Private Sub Command3_Click()
MediaPlayer1.Play
End Sub

Private Sub Command4_Click()
MediaPlayer1.Stop
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
ChDrive Drive1.Drive
End Sub

Private Sub File1_Click()
If Len(Dir1.Path) > 3 Then Text1.Text = Dir1.Path + "\" + File1.FileName
If Len(Dir1.Path) = 3 Then Text1.Text = Dir1.Path + File1.FileName
MediaPlayer1.FileName = File1.FileName
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
ChDir Dir1.Path
End Sub

Private Sub File1_dblClick()
MediaPlayer1.FileName = File1.FileName
Command3_Click
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Form1.Width) / 2
Me.Top = (Screen.Height - Form1.Height) / 2
File1.Path = Dir1.Path
ChDir Dir1.Path
File1.Pattern = "*.Wav;*.Mid"
Dir1.Path = Drive1.Drive
ChDrive Drive1.Drive
Drive1.Drive = Form1.Drive1.Drive
Dir1.Path = Form1.Dir1.Path
File1.Path = Form1.File1.Path
Call FormOnTop(Me.hWnd, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Command4_Click
Unload Form8
End Sub

