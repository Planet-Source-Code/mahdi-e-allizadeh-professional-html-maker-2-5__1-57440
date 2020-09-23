VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New HTML"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin M2AHTMLMaker.chameleonButton Command1 
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   3600
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "Form5.frx":058A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   -240
      TabIndex        =   11
      Top             =   3240
      Width           =   5535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   910
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   4455
      Begin VB.TextBox Text1 
         Height          =   345
         Left            =   1680
         TabIndex        =   6
         Top             =   200
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   1680
         TabIndex        =   5
         Top             =   660
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Meta Description :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Meta Keywords :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Text            =   "0"
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "0"
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "0"
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Don't Underline Links"
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
      TabIndex        =   0
      Top             =   2760
      Width           =   2295
   End
   Begin M2AHTMLMaker.chameleonButton Command2 
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   3600
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "Form5.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin M2AHTMLMaker.chameleonButton Command12 
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Page'sTitle "
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
      MICON           =   "Form5.frx":05C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin M2AHTMLMaker.chameleonButton Command10 
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Color"
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
      MICON           =   "Form5.frx":05DE
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
      Left            =   2880
      TabIndex        =   16
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Custom Scrollbars"
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
      MICON           =   "Form5.frx":05FA
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
      Left            =   4200
      TabIndex        =   17
      Top             =   900
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   ""
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
      MICON           =   "Form5.frx":0616
      PICN            =   "Form5.frx":0632
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
      Caption         =   "Back Ground Image :"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Inp, Sa
Sa = InStr(1, Form1.RichTextBox3.Text, """description""", vbTextCompare)
Form1.RichTextBox3.SelStart = Sa + 22
Form1.RichTextBox3.SelLength = Text4.Text
Text4.Text = Len(Text1.Text)
Form1.RichTextBox3.SelText = Text1.Text

Sa = InStr(1, Form1.RichTextBox3.Text, """keywords""", vbTextCompare)
Form1.RichTextBox3.SelStart = Sa + 19
Form1.RichTextBox3.SelLength = Text5.Text
Text5.Text = Len(Text1.Text)
Form1.RichTextBox3.SelText = Text2.Text

If Check1.Value = 1 Then
Sa = InStr(1, Form1.RichTextBox3.Text, "</Head>", vbTextCompare)
Form1.RichTextBox3.SelStart = Sa - 1
Form1.RichTextBox3.SelText = "<style type="
Form1.RichTextBox3.SelText = """"
Form1.RichTextBox3.SelText = "text/css"
Form1.RichTextBox3.SelText = """"
Form1.RichTextBox3.SelText = "> "
Form1.RichTextBox3.SelText = " <!-- "
Form1.RichTextBox3.SelText = "A:link {text-decoration: none;} "
Form1.RichTextBox3.SelText = "A:visited {text-decoration: none;} "
Form1.RichTextBox3.SelText = "--> "
Form1.RichTextBox3.SelText = "</style>"
End If

Form1.RichTextBox3.Visible = False
Form1.Image1.Visible = False
Form1.RichTextBox3.Visible = True

Text1.Text = ""
Text2.Text = ""
Text3.Text = "0"
Text4.Text = "0"
Text5.Text = "0"
Check1.Value = 0

If Text6.Text <> Empty Then
 Sa = InStr(1, Form1.RichTextBox3.Text, "BGCOLOR=", vbTextCompare)
 Form1.RichTextBox3.SelStart = Sa - 2
 Form1.RichTextBox3.SelLength = 17
 Form1.RichTextBox3.SelText = Empty
 DoEvents
 Sb = InStr(1, Form1.RichTextBox3.Text, "VLINK=", vbTextCompare)
 Form1.RichTextBox3.SelStart = Sb + 13
 Form1.RichTextBox3.SelText = " background="""
 Form1.RichTextBox3.SelText = Text6.Text
 Form1.RichTextBox3.SelText = """"
Else
 Sa = InStr(1, Form1.RichTextBox3.Text, "VLINK=", vbTextCompare)
 Form1.RichTextBox3.SelStart = Sa + 16
End If

Var1 = 0
If Form1.RichTextBox3.Visible = True Then Form1.RichTextBox3.SetFocus
Text6.Text = Empty
Me.Hide
Form1.Show
End Sub

Private Sub Command10_Click()
Form4.Show
End Sub

Private Sub Command12_Click()
Dim Inp, Sa, Sa2
Sa = InStr(1, Form1.RichTextBox3.Text, "TITLE>", vbTextCompare)
Form1.RichTextBox3.SelStart = Sa + 5
Form1.RichTextBox3.SelLength = Text3.Text
Call FormOnTop(Me.hWnd, False)
Inp = InputBox("What is your Page's Name ?", "What is your Page's Name?")
Call FormOnTop(Me.hWnd, True)
Text3.Text = Len(Inp)
Form1.RichTextBox3.SelText = Inp
Form1.RichTextBox3.Visible = False
Form1.Image1.Visible = False
Form1.RichTextBox3.Visible = True
End Sub

Private Sub Command2_Click()
Me.Hide
Var1 = 0
Form1.Show
If Form1.RichTextBox3.Visible = True Then Form1.RichTextBox3.SetFocus
End Sub

Private Sub Command3_Click()
On Error GoTo ErrHandle
With CommonDialog1
  .DialogTitle = "Insert Back Ground Image"
  .Filter = "Image Files (*.bmp,*.gif,*.jpg,*.jpeg,*.png)|*.bmp;*.gif;*.jpg;*.jpeg;*.png"
  .ShowOpen
  If Len(.FileName) = 0 Then
    Exit Sub
  End If
  Text6.Text = .FileName
End With
ErrHandle:
 Exit Sub
End Sub

Private Sub Command4_Click()
Form17.Show
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Form1.Width) / 2
Me.Top = (Screen.Height - Form1.Height) / 2
Call FormOnTop(Me.hWnd, True)
End Sub
