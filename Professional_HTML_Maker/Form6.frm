VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form6 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Link"
   ClientHeight    =   1335
   ClientLeft      =   3555
   ClientTop       =   3405
   ClientWidth     =   5400
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin M2AHTMLMaker.chameleonButton Command1 
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   135
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
      MICON           =   "Form6.frx":058A
      PICN            =   "Form6.frx":05A6
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
      Left            =   -120
      TabIndex        =   1
      Top             =   600
      Width           =   5535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4560
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin M2AHTMLMaker.chameleonButton Command2 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      MICON           =   "Form6.frx":0928
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
      Left            =   2760
      TabIndex        =   2
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      MICON           =   "Form6.frx":0944
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error Resume Next
With CommonDialog1
  .DialogTitle = "Insert Back Ground Image"
  .Filter = "HTML , HTM Files (*.html,htm)|*.HTML;*.HTM"
  .ShowOpen
  If Len(.FileName) = 0 Then
    Exit Sub
  End If
  Text1.Text = .FileName
End With
End Sub

Private Sub Command2_Click()
bb = Form1.RichTextBox3.SelStart
ss = Form1.RichTextBox3.SelLength
     Form1.RichTextBox3.SelStart = Form1.RichTextBox3.SelStart
    Form1.RichTextBox3.SelText = "<A HREF="
    Form1.RichTextBox3.SelText = """"
    Form1.RichTextBox3.SelText = Text1.Text
    Form1.RichTextBox3.SelText = """"
    Form1.RichTextBox3.SelText = ">"
      Form1.RichTextBox3.SelStart = Val(Form1.RichTextBox3.SelStart) + Val(ss)
     Form1.RichTextBox3.SelText = "</A>"
  Me.Hide
  Form1.Show
  If Form1.RichTextBox3.Visible = True Then Form1.RichTextBox3.SetFocus
     Form1.RichTextBox3.SelStart = bb + (11 + Len(Text1.Text))
     Form1.RichTextBox3.SelLength = ss
End Sub

Private Sub Command3_Click()
Me.Hide
Form1.Show
If Form1.RichTextBox3.Visible = True Then Form1.RichTextBox3.SetFocus
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Form1.Width) / 2
Me.Top = (Screen.Height - Form1.Height) / 2
Call FormOnTop(Me.hWnd, True)
End Sub
