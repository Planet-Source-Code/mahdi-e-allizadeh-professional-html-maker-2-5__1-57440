VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin M2AHTMLMaker.chameleonButton Command5 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
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
      MICON           =   "Form4.frx":058A
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
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   5535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   195
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin M2AHTMLMaker.chameleonButton Command1 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Back Ground Color"
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
      MICON           =   "Form4.frx":05A6
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
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Link Color"
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
      MICON           =   "Form4.frx":05C2
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
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Text Color"
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
      MICON           =   "Form4.frx":05DE
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
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Visited Link Color"
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
      MICON           =   "Form4.frx":05FA
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
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function SellColor()
Select Case CommonDialog1.Color
             Case 0
             Form1.RichTextBox3.SelText = "000000"
             Case 64
             Form1.RichTextBox3.SelText = "400000"
             Case 128
             Form1.RichTextBox3.SelText = "800000"
             Case 4210816
             Form1.RichTextBox3.SelText = "804040"
             Case 255
             Form1.RichTextBox3.SelText = "FF0000"
             Case 8421631
             Form1.RichTextBox3.SelText = "FF8080"
             Case 32896
             Form1.RichTextBox3.SelText = "808000"
             Case 16512
             Form1.RichTextBox3.SelText = "804000"
             Case 33023
             Form1.RichTextBox3.SelText = "FF8000"
             Case 4227327
             Form1.RichTextBox3.SelText = "FF8040"
             Case 65535
             Form1.RichTextBox3.SelText = "FFFF00"
             Case 8454143
             Form1.RichTextBox3.SelText = "FFFF80"
             Case 4227200
             Form1.RichTextBox3.SelText = "808040"
             Case 16384
             Form1.RichTextBox3.SelText = "004000"
             Case 32768
             Form1.RichTextBox3.SelText = "008000"
             Case 65280
             Form1.RichTextBox3.SelText = "00FF00"
             Case 65408
             Form1.RichTextBox3.SelText = "80FF00"
             Case 8454016
             Form1.RichTextBox3.SelText = "80FF80"
             Case 8421504
             Form1.RichTextBox3.SelText = "808080"
             Case 4210688
             Form1.RichTextBox3.SelText = "004040"
             Case 4227072
             Form1.RichTextBox3.SelText = "008040"
             Case 8421376
             Form1.RichTextBox3.SelText = "008080"
             Case 4259584
             Form1.RichTextBox3.SelText = "00FF40"
             Case 8453888
             Form1.RichTextBox3.SelText = "00FF80"
             Case 8421440
             Form1.RichTextBox3.SelText = "408080"
             Case 8388608
             Form1.RichTextBox3.SelText = "000080"
             Case 16711680
             Form1.RichTextBox3.SelText = "0000FF"
             Case 8404992
             Form1.RichTextBox3.SelText = "004080"
             Case 16776960
             Form1.RichTextBox3.SelText = "00FFFF"
             Case 16777088
             Form1.RichTextBox3.SelText = "80FFFF"
             Case 12632256
             Form1.RichTextBox3.SelText = "C0C0C0"
             Case 4194304
             Form1.RichTextBox3.SelText = "000040"
             Case 10485760
             Form1.RichTextBox3.SelText = "0000A0"
             Case 16744576
             Form1.RichTextBox3.SelText = "8080FF"
             Case 12615680
             Form1.RichTextBox3.SelText = "0080C0"
             Case 16744448
             Form1.RichTextBox3.SelText = "0080FF"
             Case 4194368
             Form1.RichTextBox3.SelText = "400040"
             Case 4194368
             Form1.RichTextBox3.SelText = "400040"
             Case 8388736
             Form1.RichTextBox3.SelText = "800080"
             Case 4194432
             Form1.RichTextBox3.SelText = "800040"
             Case 12615808
             Form1.RichTextBox3.SelText = "8080C0"
             Case 12615935
             Form1.RichTextBox3.SelText = "FF80C0"
             Case 16777215
             Form1.RichTextBox3.SelText = "FFFFFF"
             Case 8388672
             Form1.RichTextBox3.SelText = "400080"
             Case 16711808
             Form1.RichTextBox3.SelText = "8000FF"
             Case 8388863
             Form1.RichTextBox3.SelText = "FF0080"
             Case 16711935
             Form1.RichTextBox3.SelText = "FF00FF"
             Case 16744703
             Form1.RichTextBox3.SelText = "FF80FF"
            End Select
End Function

Private Sub Command1_Click()
On Error GoTo ErrHandler
Dim Sa, Sa2
With CommonDialog1
            .DialogTitle = "Select a color"
            .Flags = cdlCCPreventFullOpen
            .ShowColor
            End With
            Text1.Text = "BGCOLOR="""
            Sa = InStr(1, Form1.RichTextBox3.Text, Text1.Text, vbTextCompare)
            Form1.RichTextBox3.SelStart = Sa + 8
            Form1.RichTextBox3.SelLength = 6
            Form1.RichTextBox3.SelText = ""
            Form1.RichTextBox3.SelStart = Sa + 8
Call SellColor
Form1.RichTextBox3.Visible = False
Form1.Image1.Visible = False
Form1.RichTextBox3.Visible = True
ErrHandler:
 Exit Sub
End Sub

Private Sub Command2_Click()
On Error GoTo ErrHandler
Dim Sa, Sa2
With CommonDialog1
            .DialogTitle = "Select a color"
            .Flags = cdlCCPreventFullOpen
            .ShowColor
            End With
            Text1.Text = "TEXT="""
            Sa = InStr(1, Form1.RichTextBox3.Text, Text1.Text, vbTextCompare)
            Form1.RichTextBox3.SelStart = Sa + 5
            Form1.RichTextBox3.SelLength = 6
            Form1.RichTextBox3.SelText = ""
            Form1.RichTextBox3.SelStart = Sa + 5
Call SellColor
Form1.RichTextBox3.Visible = False
Form1.Image1.Visible = False
Form1.RichTextBox3.Visible = True
ErrHandler:
 Exit Sub
End Sub

Private Sub Command3_Click()
On Error GoTo ErrHandler
Dim Sa, Sa2
With CommonDialog1
            .DialogTitle = "Select a color"
            .Flags = cdlCCPreventFullOpen
            .ShowColor
            End With
            Text1.Text = "Link="""
            Sa = InStr(1, Form1.RichTextBox3.Text, Text1.Text, vbTextCompare)
            Form1.RichTextBox3.SelStart = Sa + 5
            Form1.RichTextBox3.SelLength = 6
            Form1.RichTextBox3.SelText = ""
            Form1.RichTextBox3.SelStart = Sa + 5
Call SellColor
Form1.RichTextBox3.Visible = False
Form1.Image1.Visible = False
Form1.RichTextBox3.Visible = True
ErrHandler:
 Exit Sub
End Sub

Private Sub Command4_Click()
On Error GoTo ErrHandler
Dim Sa, Sa2
With CommonDialog1
            .DialogTitle = "Select a color"
            .Flags = cdlCCPreventFullOpen
            .ShowColor
            End With
            Text1.Text = "VLINK="""
            Sa = InStr(1, Form1.RichTextBox3.Text, Text1.Text, vbTextCompare)
            Form1.RichTextBox3.SelStart = Sa + 6
            Form1.RichTextBox3.SelLength = 6
            Form1.RichTextBox3.SelText = ""
            Form1.RichTextBox3.SelStart = Sa + 6
Call SellColor
Form1.RichTextBox3.Visible = False
Form1.Image1.Visible = False
Form1.RichTextBox3.Visible = True
ErrHandler:
 Exit Sub
End Sub

Private Sub Command5_Click()
Me.Hide
If Var1 = 1 Then
 Form5.Show
Else
 Form1.Show
If Form1.RichTextBox3.Visible = True Then Form1.RichTextBox3.SetFocus
End If
Var1 = 0
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Form1.Width) / 2
Me.Top = (Screen.Height - Form1.Height) / 2
Call FormOnTop(Me.hWnd, True)
End Sub
