VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1170
   ClientLeft      =   2670
   ClientTop       =   330
   ClientWidth     =   7770
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin M2AHTMLMaker.chameleonButton Command5 
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "<---  Clear"
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
      MICON           =   "Form3.frx":058A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6015
   End
   Begin M2AHTMLMaker.chameleonButton Command1 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Find"
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
      MICON           =   "Form3.frx":05A6
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
      Left            =   2160
      TabIndex        =   5
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Replace"
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
      MICON           =   "Form3.frx":05C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin M2AHTMLMaker.chameleonButton Command2 
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Delete"
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
      MICON           =   "Form3.frx":05DE
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
      Left            =   6240
      TabIndex        =   2
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Close"
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
      MICON           =   "Form3.frx":05FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Enter your word for Find :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
   Begin VB.Menu mnupopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnupopupCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnupopupCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnupopupPaste 
         Caption         =   "Paste"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
          On Error Resume Next
          If Command1.Caption = "Next Find" Then
           kah = InStr(Komak, Form1.RichTextBox3.Text, Text1.Text, vbTextCompare)
           Sha = Len(Text1.Text)
           Form1.RichTextBox3.SelStart = kah - 1
           Form1.RichTextBox3.SelLength = Sha
           Komak = Komak + (kah - 1) + Sha
           If kah = 0 Then Chr$ (13) + Chr$(10)
          Else
           kah = InStr(1, Form1.RichTextBox3.Text, Text1.Text, vbTextCompare)
           Sha = Len(Text1.Text)
           Form1.RichTextBox3.SelStart = kah - 1
           Form1.RichTextBox3.SelLength = Sha
           Komak = Komak + (kah - 1) + Sha
          End If
          If Sha <> Empty Then
          Form1.Show
          Command2.Enabled = True
          Command4.Enabled = True
          Command1.Caption = "Next Find"
          Else
          Sb = MsgBox("Not Found", vbOKOnly, "Not Found")
          End If
          On Error Resume Next
End Sub

Private Sub Command2_Click()
 Form1.RichTextBox3.SelText = ""
 Command2.Enabled = True
 Command4.Enabled = True
End Sub

Private Sub Command3_Click()
 Me.Hide
 Form1.Show
 If Form1.RichTextBox3.Visible = True Then Form1.RichTextBox3.SetFocus
 Unload Me
End Sub

Private Sub Command4_Click()
 inq = InputBox("Please enter your word", "Enter your word...")
 Form1.RichTextBox3.SelText = inq
 Text1.Text = inq
 Command1_Click
End Sub

Private Sub Command5_Click()
Text1.Text = Empty
Command1.Caption = "Find"
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Form1.Width) / 2
Me.Top = (Screen.Height - Form1.Height) / 2
If Text1.Text = "" Then Command5.Enabled = False
Command1.Caption = "Find"
Call FormOnTop(Me.hWnd, True)
Text1.SetFocus
End Sub

Private Sub RichTextBox1_Change()
If Text1.Text = "" Then Command5.Enabled = False
If Text1.Text <> "" Then Command5.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub Text1_Change()
If Text1.Text <> Empty Then Command5.Enabled = True
End Sub
