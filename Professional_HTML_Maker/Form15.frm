VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form15 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Date And Time"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4875
   Icon            =   "Form15.frx":0000
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   -120
      TabIndex        =   6
      Top             =   1440
      Width           =   5535
   End
   Begin VB.CheckBox Check2 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   240
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   960
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Text            =   "The Time is : "
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Text            =   "Today Is : "
      Top             =   240
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   52756481
      CurrentDate     =   38098
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   52756482
      CurrentDate     =   38098
   End
   Begin M2AHTMLMaker.chameleonButton Command2 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "Form15.frx":058A
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
      Left            =   1200
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "Form15.frx":05A6
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
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = Checked Then
 DTPicker2.Enabled = True
 Text2.Enabled = True
Else
 DTPicker2.Enabled = False
 Text2.Enabled = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = Checked Then
 DTPicker1.Enabled = True
 Text1.Enabled = True
Else
 DTPicker1.Enabled = False
 Text1.Enabled = False
End If
End Sub

Private Sub Command1_Click()
If Form1.RichTextBox3.Visible = True Then
 If Check2.Value = Checked Then Form1.RichTextBox3.SelText = Text1.Text & DTPicker1.Value
 If Check1.Value = Checked Then Form1.RichTextBox3.SelText = " " & Text2.Text & DTPicker2.Value
End If
Unload Me
Form1.Show
End Sub

Private Sub Command2_Click()
Me.Hide
Form1.Show
End Sub


Private Sub Form_Load()
Me.Left = (Screen.Width - Form1.Width) / 2
Me.Top = (Screen.Height - Form1.Height) / 2
DTPicker1.Value = Date
DTPicker2.Value = Time
Call FormOnTop(Me.hWnd, True)
End Sub
