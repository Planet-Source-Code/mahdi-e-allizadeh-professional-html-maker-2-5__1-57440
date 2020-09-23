VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form Form10 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Table"
   ClientHeight    =   3645
   ClientLeft      =   3315
   ClientTop       =   2955
   ClientWidth     =   5520
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin M2AHTMLMaker.chameleonButton Command2 
      Height          =   375
      Left            =   2760
      TabIndex        =   25
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Background Color"
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
      MICON           =   "Form10.frx":058A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   327681
      Value           =   2
      OrigLeft        =   1680
      OrigTop         =   240
      OrigRight       =   1935
      OrigBottom      =   615
      Max             =   100
      Enabled         =   -1  'True
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   0
      TabIndex        =   16
      Top             =   2880
      Width           =   5535
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   0
      TabIndex        =   15
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   0
      TabIndex        =   14
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
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
      Left            =   4200
      TabIndex        =   13
      Text            =   "100"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
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
      Left            =   4200
      TabIndex        =   12
      Text            =   "100"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
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
      Left            =   1440
      TabIndex        =   9
      Text            =   "2"
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
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
      Left            =   3240
      TabIndex        =   5
      Text            =   "2"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
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
      Left            =   3240
      TabIndex        =   4
      Text            =   "2"
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
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
      Left            =   960
      TabIndex        =   2
      Text            =   "2"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Left            =   960
      TabIndex        =   0
      Text            =   "2"
      Top             =   240
      Width           =   495
   End
   Begin ComCtl2.UpDown UpDown2 
      Height          =   375
      Left            =   1440
      TabIndex        =   18
      Top             =   840
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   327681
      Value           =   2
      OrigLeft        =   1680
      OrigTop         =   240
      OrigRight       =   1935
      OrigBottom      =   615
      Max             =   100
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown UpDown3 
      Height          =   375
      Left            =   3720
      TabIndex        =   19
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   327681
      Value           =   2
      OrigLeft        =   1680
      OrigTop         =   240
      OrigRight       =   1935
      OrigBottom      =   615
      Max             =   100
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown UpDown4 
      Height          =   375
      Left            =   3720
      TabIndex        =   20
      Top             =   840
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   327681
      Value           =   2
      OrigLeft        =   1680
      OrigTop         =   240
      OrigRight       =   1935
      OrigBottom      =   615
      Max             =   100
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown UpDown5 
      Height          =   375
      Left            =   1920
      TabIndex        =   21
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   327681
      Value           =   2
      OrigLeft        =   1680
      OrigTop         =   240
      OrigRight       =   1935
      OrigBottom      =   615
      Max             =   100
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown UpDown6 
      Height          =   375
      Left            =   4920
      TabIndex        =   22
      Top             =   1320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   327681
      Value           =   100
      OrigLeft        =   1680
      OrigTop         =   240
      OrigRight       =   1935
      OrigBottom      =   615
      Max             =   100
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown UpDown7 
      Height          =   375
      Left            =   4920
      TabIndex        =   23
      Top             =   1800
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   327681
      Value           =   100
      OrigLeft        =   1680
      OrigTop         =   240
      OrigRight       =   1935
      OrigBottom      =   615
      Max             =   100
      Enabled         =   -1  'True
   End
   Begin M2AHTMLMaker.chameleonButton Command1 
      Height          =   375
      Left            =   1200
      TabIndex        =   26
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Border Color"
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
      MICON           =   "Form10.frx":05A6
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
      Left            =   2760
      TabIndex        =   27
      Top             =   3120
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
      MICON           =   "Form10.frx":05C2
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
      Left            =   1440
      TabIndex        =   24
      Top             =   3120
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
      MICON           =   "Form10.frx":05DE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      Caption         =   "Table height :"
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
      Left            =   2880
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Table width :"
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
      Left            =   2880
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Border Width :"
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
      Left            =   40
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Cell Padding :"
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
      Left            =   1920
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Cell Spacing :"
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
      Left            =   1920
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Columns :"
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
      Left            =   40
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Rows : "
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
      Left            =   40
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo ErrHandler
Dim Sa, Sa2
With CommonDialog1
            .DialogTitle = "Select a color"
            .Flags = cdlCCPreventFullOpen
            .ShowColor
            End With
            Select Case CommonDialog1.Color
             Case 0
             Text8.Text = "000000"
             Case 64
             Text8.Text = "400000"
             Case 128
             Text8.Text = "800000"
             Case 4210816
             Text8.Text = "804040"
             Case 255
             Text8.Text = "FF0000"
             Case 8421631
             Text8.Text = "FF8080"
             Case 32896
             Text8.Text = "808000"
             Case 16512
             Text8.Text = "804000"
             Case 33023
             Text8.Text = "FF8000"
             Case 4227327
             Text8.Text = "FF8040"
             Case 65535
             Text8.Text = "FFFF00"
             Case 8454143
             Text8.Text = "FFFF80"
             Case 4227200
             Text8.Text = "808040"
             Case 16384
             Text8.Text = "004000"
             Case 32768
             Text8.Text = "008000"
             Case 65280
             Text8.Text = "00FF00"
             Case 65408
             Text8.Text = "80FF00"
             Case 8454016
             Text8.Text = "80FF80"
             Case 8421504
             Text8.Text = "808080"
             Case 4210688
             Text8.Text = "004040"
             Case 4227072
             Text8.Text = "008040"
             Case 8421376
             Text8.Text = "008080"
             Case 4259584
             Text8.Text = "00FF40"
             Case 8453888
             Text8.Text = "00FF80"
             Case 8421440
             Text8.Text = "408080"
             Case 8388608
             Text8.Text = "000080"
             Case 16711680
             Text8.Text = "0000FF"
             Case 8404992
             Text8.Text = "004080"
             Case 16776960
             Text8.Text = "00FFFF"
             Case 16777088
             Text8.Text = "80FFFF"
             Case 12632256
             Text8.Text = "C0C0C0"
             Case 4194304
             Text8.Text = "000040"
             Case 10485760
             Text8.Text = "0000A0"
             Case 16744576
             Text8.Text = "8080FF"
             Case 12615680
             Text8.Text = "0080C0"
             Case 16744448
             Text8.Text = "0080FF"
             Case 4194368
             Text8.Text = "400040"
             Case 4194368
             Text8.Text = "400040"
             Case 8388736
             Text8.Text = "800080"
             Case 4194432
             Text8.Text = "800040"
             Case 12615808
             Text8.Text = "8080C0"
             Case 12615935
             Text8.Text = "FF80C0"
             Case 16777215
             Text8.Text = "FFFFFF"
             Case 8388672
             Text8.Text = "400080"
             Case 16711808
             Text8.Text = "8000FF"
             Case 8388863
             Text8.Text = "FF0080"
             Case 16711935
             Text8.Text = "FF00FF"
             Case 16744703
             Text8.Text = "FF80FF"
            End Select
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
            Select Case CommonDialog1.Color
             Case 0
             Text9.Text = "000000"
             Case 64
             Text9.Text = "400000"
             Case 128
             Text9.Text = "800000"
             Case 4210816
             Text9.Text = "804040"
             Case 255
             Text9.Text = "FF0000"
             Case 8421631
             Text9.Text = "FF8080"
             Case 32896
             Text9.Text = "808000"
             Case 16512
             Text9.Text = "804000"
             Case 33023
             Text9.Text = "FF8000"
             Case 4227327
             Text9.Text = "FF8040"
             Case 65535
             Text9.Text = "FFFF00"
             Case 8454143
             Text9.Text = "FFFF80"
             Case 4227200
             Text9.Text = "808040"
             Case 16384
             Text9.Text = "004000"
             Case 32768
             Text9.Text = "008000"
             Case 65280
             Text9.Text = "00FF00"
             Case 65408
             Text9.Text = "80FF00"
             Case 8454016
             Text9.Text = "80FF80"
             Case 8421504
             Text9.Text = "808080"
             Case 4210688
             Text9.Text = "004040"
             Case 4227072
             Text9.Text = "008040"
             Case 8421376
             Text9.Text = "008080"
             Case 4259584
             Text9.Text = "00FF40"
             Case 8453888
             Text9.Text = "00FF80"
             Case 8421440
             Text9.Text = "408080"
             Case 8388608
             Text9.Text = "000080"
             Case 16711680
             Text9.Text = "0000FF"
             Case 8404992
             Text9.Text = "004080"
             Case 16776960
             Text9.Text = "00FFFF"
             Case 16777088
             Text9.Text = "80FFFF"
             Case 12632256
             Text9.Text = "C0C0C0"
             Case 4194304
             Text9.Text = "000040"
             Case 10485760
             Text9.Text = "0000A0"
             Case 16744576
             Text9.Text = "8080FF"
             Case 12615680
             Text9.Text = "0080C0"
             Case 16744448
             Text9.Text = "0080FF"
             Case 4194368
             Text9.Text = "400040"
             Case 4194368
             Text9.Text = "400040"
             Case 8388736
             Text9.Text = "800080"
             Case 4194432
             Text9.Text = "800040"
             Case 12615808
             Text9.Text = "8080C0"
             Case 12615935
             Text9.Text = "FF80C0"
             Case 16777215
             Text9.Text = "FFFFFF"
             Case 8388672
             Text9.Text = "400080"
             Case 16711808
             Text9.Text = "8000FF"
             Case 8388863
             Text9.Text = "FF0080"
             Case 16711935
             Text9.Text = "FF00FF"
             Case 16744703
             Text9.Text = "FF80FF"
            End Select
Form1.RichTextBox3.Visible = False
Form1.Image1.Visible = False
Form1.RichTextBox3.Visible = True
ErrHandler:
 Exit Sub
End Sub

Private Sub Command3_Click()
With Form1.RichTextBox3
  .SelText = "<table width="
  .SelText = """"
  .SelText = Text6.Text + "%"
  .SelText = """ "
  .SelText = "height="
  .SelText = """"
  .SelText = Text7.Text + "%"
  .SelText = """ "
  .SelText = "bordercolor="
  .SelText = """"
  .SelText = Text8.Text
  .SelText = """ "
  .SelText = "bgcolor="
  .SelText = """"
  .SelText = Text9.Text
  .SelText = """ "
  .SelText = "cellspacing="
  .SelText = """"
  .SelText = Text3.Text
  .SelText = """ "
  .SelText = "cellpadding="
  .SelText = """"
  .SelText = Text4.Text
  .SelText = """ "
  .SelText = "border="
  .SelText = Text5.Text
  .SelText = """"
  .SelText = ">"
  .SelText = "     "
For i = 1 To Text1.Text
   .SelText = "<tr>"
   .SelText = "    "
  For j = 1 To Text2.Text
     .SelText = "<td>  &nbsp;</td>"
     .SelText = "    "
  Next j
    .SelText = "    "
   .SelText = "</tr>"
Next i
  .SelText = "        "
  .SelText = "</table>"
Me.Hide
Form1.Show
If .Visible = True Then .SetFocus
End With
End Sub

Private Sub Command4_Click()
Me.Hide
Form1.Show
If Form1.RichTextBox3.Visible = True Then Form1.RichTextBox3.SetFocus
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Form1.Width) / 2
Me.Top = (Screen.Height - Form1.Height) / 2
Call FormOnTop(Me.hWnd, True)
End Sub

Private Sub UpDown1_Change()
Text1.Text = UpDown1.Value
End Sub

Private Sub UpDown2_Change()
Text2.Text = UpDown2.Value
End Sub

Private Sub UpDown3_Change()
Text3.Text = UpDown3.Value
End Sub

Private Sub UpDown4_Change()
Text4.Text = UpDown4.Value
End Sub

Private Sub UpDown5_Change()
Text1.Text = UpDown5.Value
End Sub

Private Sub UpDown6_Change()
Text6.Text = UpDown6.Value
End Sub

Private Sub UpDown7_Change()
Text7.Text = UpDown7.Value
End Sub
