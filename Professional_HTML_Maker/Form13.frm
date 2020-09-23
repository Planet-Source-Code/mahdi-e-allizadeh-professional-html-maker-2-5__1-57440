VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form13 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Line"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5580
   Icon            =   "Form13.frx":0000
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox3 
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form13.frx":058A
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2400
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5295
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Value           =   100
         OrigLeft        =   1800
         OrigTop         =   360
         OrigRight       =   2055
         OrigBottom      =   735
         Increment       =   10
         Max             =   100
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   240
         TabIndex        =   4
         Text            =   "Size 1"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   3360
         TabIndex        =   3
         Text            =   "Left"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "100"
         Top             =   360
         Width           =   495
      End
      Begin M2AHTMLMaker.chameleonButton Command1 
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
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
         MICON           =   "Form13.frx":061A
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
         Caption         =   "Alignment :"
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
         Left            =   2240
         TabIndex        =   5
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Linear (%) :"
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
         Left            =   200
         TabIndex        =   2
         Top             =   390
         Width           =   1080
      End
   End
   Begin M2AHTMLMaker.chameleonButton Command3 
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   2880
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
      MICON           =   "Form13.frx":0636
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
      Left            =   1440
      TabIndex        =   9
      Top             =   2880
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
      MICON           =   "Form13.frx":0652
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000006&
      X1              =   720
      X2              =   4920
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Counter As Byte


Private Sub Combo2_Click()
St = Combo2.Text
If St = "Size 1" Then Line1.BorderWidth = "1"
If St = "Size 2" Then Line1.BorderWidth = "2"
If St = "Size 3" Then Line1.BorderWidth = "3"
If St = "Size 4" Then Line1.BorderWidth = "4"
If St = "Size 5" Then Line1.BorderWidth = "5"
If St = "Size 6" Then Line1.BorderWidth = "6"
If St = "Size 7" Then Line1.BorderWidth = "7"
End Sub

Private Sub Command1_Click()
On Error GoTo ErrHandler
With CommonDialog1
    .DialogTitle = "Select a color"
    .Flags = cdlCCPreventFullOpen
    .ShowColor
End With
Select Case CommonDialog1.Color
             Case 0
             RichTextBox3.Text = "000000"
             Case 64
             RichTextBox3.Text = "400000"
             Case 128
             RichTextBox3.Text = "800000"
             Case 4210816
             RichTextBox3.Text = "804040"
             Case 255
             RichTextBox3.Text = "FF0000"
             Case 8421631
             RichTextBox3.Text = "FF8080"
             Case 32896
             RichTextBox3.Text = "808000"
             Case 16512
             RichTextBox3.Text = "804000"
             Case 33023
             RichTextBox3.Text = "FF8000"
             Case 4227327
             RichTextBox3.Text = "FF8040"
             Case 65535
             RichTextBox3.Text = "FFFF00"
             Case 8454143
             RichTextBox3.Text = "FFFF80"
             Case 4227200
             RichTextBox3.Text = "808040"
             Case 16384
             RichTextBox3.Text = "004000"
             Case 32768
             RichTextBox3.Text = "008000"
             Case 65280
             RichTextBox3.Text = "00FF00"
             Case 65408
             RichTextBox3.Text = "80FF00"
             Case 8454016
             RichTextBox3.Text = "80FF80"
             Case 8421504
             RichTextBox3.Text = "808080"
             Case 4210688
             RichTextBox3.Text = "004040"
             Case 4227072
             RichTextBox3.Text = "008040"
             Case 8421376
             RichTextBox3.Text = "008080"
             Case 4259584
             RichTextBox3.Text = "00FF40"
             Case 8453888
             RichTextBox3.Text = "00FF80"
             Case 8421440
             RichTextBox3.Text = "408080"
             Case 8388608
             RichTextBox3.Text = "000080"
             Case 16711680
             RichTextBox3.Text = "0000FF"
             Case 8404992
             RichTextBox3.Text = "004080"
             Case 16776960
             RichTextBox3.Text = "00FFFF"
             Case 16777088
             RichTextBox3.Text = "80FFFF"
             Case 12632256
             RichTextBox3.Text = "C0C0C0"
             Case 4194304
             RichTextBox3.Text = "000040"
             Case 10485760
             RichTextBox3.Text = "0000A0"
             Case 16744576
             RichTextBox3.Text = "8080FF"
             Case 12615680
             RichTextBox3.Text = "0080C0"
             Case 16744448
             RichTextBox3.Text = "0080FF"
             Case 4194368
             RichTextBox3.Text = "400040"
             Case 4194368
             RichTextBox3.Text = "400040"
             Case 8388736
             RichTextBox3.Text = "800080"
             Case 4194432
             RichTextBox3.Text = "800040"
             Case 12615808
             RichTextBox3.Text = "8080C0"
             Case 12615935
             RichTextBox3.Text = "FF80C0"
             Case 16777215
             RichTextBox3.Text = "FFFFFF"
             Case 8388672
             RichTextBox3.Text = "400080"
             Case 16711808
             RichTextBox3.Text = "8000FF"
             Case 8388863
             RichTextBox3.Text = "FF0080"
             Case 16711935
             RichTextBox3.Text = "FF00FF"
             Case 16744703
             RichTextBox3.Text = "FF80FF"
            End Select
 Line1.BorderColor = CommonDialog1.Color
 Line2.BorderColor = CommonDialog1.Color
 Line3.BorderColor = CommonDialog1.Color
 Line4.BorderColor = CommonDialog1.Color
 Line5.BorderColor = CommonDialog1.Color
 Line6.BorderColor = CommonDialog1.Color
 Line7.BorderColor = CommonDialog1.Color
ErrHandler:
 Exit Sub
End Sub

Private Sub Command2_Click()
Form1.RichTextBox3.SelText = "<HR align="
Form1.RichTextBox3.SelText = Combo1.Text
Form1.RichTextBox3.SelText = " width="""
Form1.RichTextBox3.SelText = Text1.Text & "%"""
Form1.RichTextBox3.SelText = " color="
If RichTextBox3.Text <> Empty Then
 Form1.RichTextBox3.SelText = RichTextBox3.Text
Else
 Form1.RichTextBox3.SelText = "000000"
End If
Form1.RichTextBox3.SelText = " Size="
Combo2.SelStart = 5
Combo2.SelLength = Len(Combo2.Text)
Form1.RichTextBox3.SelText = Combo2.SelText & ">"
Me.Hide
Form1.Show
If Form1.RichTextBox3.Visible = True Then Form1.RichTextBox3.SetFocus
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
Combo1.AddItem "Left"
Combo1.AddItem "Center"
Combo1.AddItem "Right"
Combo2.AddItem "Size 1"
Combo2.AddItem "Size 2"
Combo2.AddItem "Size 3"
Combo2.AddItem "Size 4"
Combo2.AddItem "Size 5"
Combo2.AddItem "Size 6"
Combo2.AddItem "Size 7"
Combo2.AddItem "Size 8"
Combo2.AddItem "Size 9"
Combo2.AddItem "Size 10"
Combo2.AddItem "Size 11"
Combo2.AddItem "Size 12"
Combo2.AddItem "Size 13"
Combo2.AddItem "Size 14"
Combo2.AddItem "Size 15"
Combo2.AddItem "Size 16"
Combo2.AddItem "Size 17"
Combo2.AddItem "Size 18"
Combo2.AddItem "Size 19"
Combo2.AddItem "Size 20"
End Sub

Private Sub Text1_Change()
On Error Resume Next
Select Case Text1.Text
  Case 0
   Line1.X2 = 720
  Case 10
   Line1.X2 = 1200
  Case 20
   Line1.X2 = 1650
  Case 30
   Line1.X2 = 2100
  Case 40
   Line1.X2 = 2450
  Case 50
   Line1.X2 = 3000
  Case 60
   Line1.X2 = 3400
  Case 70
   Line1.X2 = 3800
  Case 80
   Line1.X2 = 4200
  Case 90
   Line1.X2 = 4500
  Case 100
   Line1.X2 = 4920
 End Select
End Sub

Private Sub VScroll1_Change()
Text1.Text = VScroll1.Value
End Sub

Private Sub UpDown1_Change()
Text1.Text = UpDown1.Value
End Sub
