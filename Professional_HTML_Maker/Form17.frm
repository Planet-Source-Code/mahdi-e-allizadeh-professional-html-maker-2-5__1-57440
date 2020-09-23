VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form17 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Scrollbars"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4725
   Icon            =   "Form17.frx":0000
   LinkTopic       =   "Form17"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin M2AHTMLMaker.chameleonButton Command4 
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   2400
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
      MICON           =   "Form17.frx":058A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6480
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   5655
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Custom Scrollbras Color (only in ie 5.5 or later)"
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
      Top             =   120
      Width           =   4575
   End
   Begin M2AHTMLMaker.chameleonButton Command5 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2400
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
      MICON           =   "Form17.frx":05A6
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
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Arrow Color"
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
      MICON           =   "Form17.frx":05C2
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
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Face Color"
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
      MICON           =   "Form17.frx":05DE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin M2AHTMLMaker.chameleonButton Command6 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "3D light Color"
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
      MICON           =   "Form17.frx":05FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin M2AHTMLMaker.chameleonButton Command8 
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Shadow Color"
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
      MICON           =   "Form17.frx":0616
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin M2AHTMLMaker.chameleonButton Command7 
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Highlight Color"
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
      MICON           =   "Form17.frx":0632
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin M2AHTMLMaker.chameleonButton Command9 
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Dark Shadow Color"
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
      MICON           =   "Form17.frx":064E
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
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Track Color"
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
      MICON           =   "Form17.frx":066A
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
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ForSelStart As Variant
Public ForOkClick As Byte

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

Private Sub Check1_Click()
If Check1.Value = Checked Then
 Command1.Enabled = True
 Command2.Enabled = True
 Command3.Enabled = True
 Command6.Enabled = True
 Command7.Enabled = True
 Command8.Enabled = True
 Command9.Enabled = True
 Form1.RichTextBox3.SelStart = 214
 Form1.RichTextBox3.SelText = Chr$(10)
 Form1.RichTextBox3.SelText = "<style>"
 Form1.RichTextBox3.SelText = Chr$(10)
 For i = 1 To 3
  Form1.RichTextBox3.SelText = Chr$(13)
 Next i
 Form1.RichTextBox3.SelText = " body{"
Else
 Command1.Enabled = False
 Command2.Enabled = False
 Command3.Enabled = False
 Command6.Enabled = False
 Command7.Enabled = False
 Command8.Enabled = False
 Command9.Enabled = False
 Form1.RichTextBox3.SelStart = 214
 Form1.RichTextBox3.SelLength = 19
 Form1.RichTextBox3.SelText = Empty
End If
End Sub

Private Sub Command1_Click()
On Error GoTo ErrHandler
Dim Sa, Sa2
ForOkClick = 1
With CommonDialog1
            .DialogTitle = "Select a color"
            .Flags = cdlCCPreventFullOpen
            .ShowColor
            End With
Form1.RichTextBox3.Visible = False
Form1.Image1.Visible = False
Form1.RichTextBox3.Visible = True
Form1.RichTextBox3.SelText = Chr$(10)
For i = 1 To 5
 Form1.RichTextBox3.SelText = Chr$(13)
Next i
Form1.RichTextBox3.SelText = "scrollbar-face-color: #"
Call SellColor
Form1.RichTextBox3.SelText = ";"
ForSelStart = Form1.RichTextBox3.SelStart
ErrHandler:
 Exit Sub
End Sub

Private Sub Command2_Click()
On Error GoTo ErrHandler
Dim Sa, Sa2
ForOkClick = 1
With CommonDialog1
            .DialogTitle = "Select a color"
            .Flags = cdlCCPreventFullOpen
            .ShowColor
            End With
Form1.RichTextBox3.Visible = False
Form1.Image1.Visible = False
Form1.RichTextBox3.Visible = True
Form1.RichTextBox3.SelText = Chr$(10)
For i = 1 To 5
 Form1.RichTextBox3.SelText = Chr$(13)
Next i
Form1.RichTextBox3.SelText = "scrollbar-track-color: #"
Call SellColor
Form1.RichTextBox3.SelText = ";"
ForSelStart = Form1.RichTextBox3.SelStart
ErrHandler:
 Exit Sub
End Sub

Private Sub Command3_Click()
On Error GoTo ErrHandler
Dim Sa, Sa2
ForOkClick = 1
With CommonDialog1
            .DialogTitle = "Select a color"
            .Flags = cdlCCPreventFullOpen
            .ShowColor
            End With
Form1.RichTextBox3.Visible = False
Form1.Image1.Visible = False
Form1.RichTextBox3.Visible = True
Form1.RichTextBox3.SelText = Chr$(10)
For i = 1 To 5
 Form1.RichTextBox3.SelText = Chr$(13)
Next i
Form1.RichTextBox3.SelText = "scrollbar-arrow-color: #"
Call SellColor
Form1.RichTextBox3.SelText = ";"
ForSelStart = Form1.RichTextBox3.SelStart
ErrHandler:
 Exit Sub
End Sub

Private Sub Command4_Click()
If ForOkClick = 0 Then GoTo ForOK
If Check1.Value = Checked Then
 Form1.RichTextBox3.SelStart = ForSelStart
 Form1.RichTextBox3.SelText = Chr$(10)
 For i = 1 To 5
  Form1.RichTextBox3.SelText = Chr$(13)
 Next i
 Form1.RichTextBox3.SelText = "}"
 Form1.RichTextBox3.SelText = Chr$(10)
 Form1.RichTextBox3.SelText = "</style>"
 
 Form1.RichTextBox3.SelStart = 214
 Form1.RichTextBox3.SelText = Chr$(10)
 Form1.RichTextBox3.SelText = "<style>"
 Form1.RichTextBox3.SelText = Chr$(10)
 For i = 1 To 3
  Form1.RichTextBox3.SelText = Chr$(13)
 Next i
 Form1.RichTextBox3.SelText = " body{"
End If
ForOK:
Check1.Value = Unchecked
 Me.Hide
 Form5.Show
End Sub

Private Sub Command5_Click()
Check1.Value = Unchecked
Me.Hide
Form5.Show
End Sub

Private Sub Command6_Click()
On Error GoTo ErrHandler
Dim Sa, Sa2
ForOkClick = 1
With CommonDialog1
            .DialogTitle = "Select a color"
            .Flags = cdlCCPreventFullOpen
            .ShowColor
            End With
Form1.RichTextBox3.Visible = False
Form1.Image1.Visible = False
Form1.RichTextBox3.Visible = True
Form1.RichTextBox3.SelText = Chr$(10)
For i = 1 To 5
 Form1.RichTextBox3.SelText = Chr$(13)
Next i
Form1.RichTextBox3.SelText = "scrollbar-3dlight-color: #"
Call SellColor
Form1.RichTextBox3.SelText = ";"
ForSelStart = Form1.RichTextBox3.SelStart
ErrHandler:
 Exit Sub
End Sub

Private Sub Command7_Click()
On Error GoTo ErrHandler
Dim Sa, Sa2
ForOkClick = 1
With CommonDialog1
            .DialogTitle = "Select a color"
            .Flags = cdlCCPreventFullOpen
            .ShowColor
            End With
Form1.RichTextBox3.Visible = False
Form1.Image1.Visible = False
Form1.RichTextBox3.Visible = True
Form1.RichTextBox3.SelText = Chr$(10)
For i = 1 To 5
 Form1.RichTextBox3.SelText = Chr$(13)
Next i
Form1.RichTextBox3.SelText = "scrollbar-highlight-color: #"
Call SellColor
Form1.RichTextBox3.SelText = ";"
ForSelStart = Form1.RichTextBox3.SelStart
ErrHandler:
 Exit Sub
End Sub

Private Sub Command8_Click()
On Error GoTo ErrHandler
Dim Sa, Sa2
ForOkClick = 1
With CommonDialog1
            .DialogTitle = "Select a color"
            .Flags = cdlCCPreventFullOpen
            .ShowColor
            End With
Form1.RichTextBox3.Visible = False
Form1.Image1.Visible = False
Form1.RichTextBox3.Visible = True
Form1.RichTextBox3.SelText = Chr$(10)
For i = 1 To 5
 Form1.RichTextBox3.SelText = Chr$(13)
Next i
Form1.RichTextBox3.SelText = "scrollbar-shadow-color: #"
Call SellColor
Form1.RichTextBox3.SelText = ";"
ForSelStart = Form1.RichTextBox3.SelStart
ErrHandler:
 Exit Sub
End Sub

Private Sub Command9_Click()
On Error GoTo ErrHandler
Dim Sa, Sa2
ForOkClick = 1
With CommonDialog1
            .DialogTitle = "Select a color"
            .Flags = cdlCCPreventFullOpen
            .ShowColor
            End With
Form1.RichTextBox3.Visible = False
Form1.Image1.Visible = False
Form1.RichTextBox3.Visible = True
Form1.RichTextBox3.SelText = Chr$(10)
For i = 1 To 5
 Form1.RichTextBox3.SelText = Chr$(13)
Next i
Form1.RichTextBox3.SelText = "scrollbar-darkshadow-color: #"
Call SellColor
Form1.RichTextBox3.SelText = ";"
ForSelStart = Form1.RichTextBox3.SelStart
ErrHandler:
 Exit Sub
End Sub

Private Sub Form_Load()
Call FormOnTop(Me.hWnd, True)
End Sub
