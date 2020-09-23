VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form7 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Picture"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin M2AHTMLMaker.chameleonButton Command1 
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   4635
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
      MICON           =   "Form7.frx":058A
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
      Left            =   7560
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1920
      Top             =   4560
   End
   Begin MSComCtl2.FlatScrollBar VS 
      Height          =   4095
      Left            =   7665
      TabIndex        =   8
      Top             =   120
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   7223
      _Version        =   393216
      Appearance      =   2
      LargeChange     =   50
      Orientation     =   1179648
      SmallChange     =   30
   End
   Begin MSComCtl2.FlatScrollBar HScroll1 
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   4200
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   2
      Arrows          =   65536
      LargeChange     =   50
      Orientation     =   1179649
      SmallChange     =   30
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4095
      Left            =   1920
      ScaleHeight     =   4035
      ScaleWidth      =   5670
      TabIndex        =   5
      Top             =   120
      Width           =   5730
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   7800
         Left            =   0
         ScaleHeight     =   516
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   803
         TabIndex        =   6
         Top             =   0
         Width           =   12105
      End
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin M2AHTMLMaker.chameleonButton Command2 
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   4635
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
      MICON           =   "Form7.frx":05A6
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
      Left            =   4320
      TabIndex        =   11
      Top             =   4635
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Save as other format"
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
      MICON           =   "Form7.frx":05C2
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
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 ss = Form1.RichTextBox3.SelLength
     Form1.RichTextBox3.SelStart = Form1.RichTextBox3.SelStart
    Form1.RichTextBox3.SelText = "<IMG SRC="
    Form1.RichTextBox3.SelText = """"
    If Len(Dir1.Path) > 3 Then Form1.RichTextBox3.SelText = Dir1.Path + "\" + File1.FileName
    If Len(Dir1.Path) = 3 Then Form1.RichTextBox3.SelText = Dir1.Path + File1.FileName
    Form1.RichTextBox3.SelText = """"
    Form1.RichTextBox3.SelText = " "
    Form1.RichTextBox3.SelText = "WIDTH="
    Form1.RichTextBox3.SelText = """"
    Form1.RichTextBox3.SelText = Text1.Text
    Form1.RichTextBox3.SelText = """"
    Form1.RichTextBox3.SelText = " "
    Form1.RichTextBox3.SelText = "HEIGHT="
    Form1.RichTextBox3.SelText = """"
    Form1.RichTextBox3.SelText = Text2.Text
    Form1.RichTextBox3.SelText = """"
    Form1.RichTextBox3.SelText = " "
    Form1.RichTextBox3.SelText = "BORDER=0>"
  Form1.RichTextBox3.Visible = False
  Form1.RichTextBox3.Visible = True
  Me.Hide
  Form1.Show
  Form1.RichTextBox3.SetFocus
End Sub

Private Sub Command2_Click()
Me.Hide
Form1.Show
If Form1.RichTextBox3.Visible = True Then Form1.RichTextBox3.SetFocus
End Sub

Private Sub Command3_Click()
On Error Resume Next
ForDir = Dir1.Path
With CommonDialog1
  If Len(Dir1.Path) = 3 Then .FileName = Dir1.Path + File1.FileName
  If Len(Dir1.Path) > 3 Then .FileName = Dir1.Path + "\" + File1.FileName
  .DialogTitle = "Save as...."
  .Filter = "Bit Map File (*.BMP)|*.BMP|Gif File (*.Gif)|*.Gif|Jpg File (*.Jpg)|*.Jpg|Tif File (*.Tif)|*.tif|Icon File (*.ico)|*.ico"
  .ShowSave
  If Len(.FileName) = 0 Then
    Exit Sub
  End If
  SavePicture Picture2.Picture, .FileName
End With
Dir1.Path = "D:\"
Dir1.Path = "C:\"
Dir1.Path = ForDir
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Form1.Width) / 2
Me.Top = (Screen.Height - Form1.Height) / 2
File1.Pattern = "*.Gif;*.Jpg;*.Bmp"
VS.Max = (Picture1.Height - Picture2.Height) + 20
HScroll1.Max = (Picture1.Width - Picture2.Width) + 20
Drive1.Drive = Form1.Drive1.Drive
Dir1.Path = Form1.Dir1.Path
File1.Path = Form1.File1.Path
Timer1.Enabled = True
Call FormOnTop(Me.hWnd, True)
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
Private Sub File1_Click()
On Error Resume Next
Dim a, a2, ab, ab2 As Integer
Form1.Image1.Visible = False
If Len(Dir1.Path) > 3 Then Picture2.Picture = LoadPicture(Dir1.Path + "\" + File1.FileName)
If Len(Dir1.Path) = 3 Then Picture2.Picture = LoadPicture(Dir1.Path + File1.FileName)
a = Picture2.Picture.Width
ab = (a / 26)
If ab > 10 And ab <= 500 Then ab = ab - 1
If ab > 500 Then ab = ab - 8
a2 = Picture2.Picture.Height
ab2 = (a2 / 26)
If ab2 > 10 And ab2 <= 500 Then ab2 = ab2 - 1
If ab2 > 500 Then ab2 = ab2 - 8
ab = Int(ab)
ab2 = Int(ab2)
Text1.Text = ab
Text2.Text = ab2

HScroll1.Value = 0
VS.Value = 0

If Picture2.Width < Picture1.Width And Picture2.Height < Picture1.Height Then
 Picture2.Left = (Picture1.Width - Picture2.Width) / 2
 Picture2.Top = (Picture1.Height - Picture2.Height) / 2
Else
 If Picture2.Width > Picture1.Width Or Picture2.Height > Picture1.Height Then
  Picture2.Left = 0
  Picture2.Top = 0
 End If
End If

End Sub
Private Sub File1_dblClick()
On Error Resume Next
Dim a, a2, ab, ab2 As Integer
Form1.Image1.Visible = False
If Len(Dir1.Path) > 3 Then Picture2.Picture = LoadPicture(Dir1.Path + "\" + File1.FileName)
If Len(Dir1.Path) = 3 Then Picture2.Picture = LoadPicture(Dir1.Path + File1.FileName)
a = Picture2.Picture.Width
ab = (a / 26)
If ab > 10 And ab <= 500 Then ab = ab - 1
If ab > 500 Then ab = ab - 8
a2 = Picture2.Picture.Height
ab2 = (a2 / 26)
If ab2 > 10 And ab2 <= 500 Then ab2 = ab2 - 1
If ab2 > 500 Then ab2 = ab2 - 8
ab = Int(ab)
ab2 = Int(ab2)
Text1.Text = ab
Text2.Text = ab2
End Sub
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
Unload Form7
End Sub

Private Sub HScroll1_Change()
Picture2.Left = HScroll1.Value
End Sub

Private Sub Timer1_Timer()
If Picture2.Width >= Picture1.Width Then
 HScroll1.Enabled = True
Else
 HScroll1.Enabled = False
End If

If Picture2.Height >= Picture1.Height Then
 VS.Enabled = True
Else
 VS.Enabled = False
End If
End Sub

Private Sub VS_Change()
Picture2.Top = VS.Value
End Sub
