VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Begin VB.Form Form16 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Flash"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7875
   Icon            =   "Form16.frx":0000
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   4170
      Left            =   1680
      TabIndex        =   7
      Top             =   75
      Width           =   6135
      _cx             =   4205125
      _cy             =   4201659
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      Stacking        =   "below"
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   60
      Pattern         =   "*.swf"
      TabIndex        =   2
      Top             =   2760
      Width           =   1530
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   1530
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   50
      Width           =   1530
   End
   Begin M2AHTMLMaker.chameleonButton Command2 
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   4440
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
      MICON           =   "Form16.frx":058A
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
      Left            =   4200
      TabIndex        =   9
      Top             =   4440
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
      MICON           =   "Form16.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "Width"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   4710
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Height"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   4365
      Width           =   495
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefLng A-Z
Private Const GWL_STYLE = (-16)
Private Const ES_NUMBER = &H2000&
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd&, ByVal nIndex&) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd&, ByVal nIndex&, ByVal dwNewLong&) As Long

Private Property Let OnlyNumber(ByVal m_hWnd As Long, ByVal rValue As Boolean)
    If (rValue = True) Then
        Call SetWindowLong(m_hWnd, GWL_STYLE, GetWindowLong(m_hWnd, GWL_STYLE) Or ES_NUMBER)
    Else
        Call SetWindowLong(m_hWnd, GWL_STYLE, GetWindowLong(m_hWnd, GWL_STYLE) And Not ES_NUMBER)
    End If
End Property

Private Sub Command1_Click()
If Text1.Text = Empty Or Text2.Text = Empty Then
 Text1.Text = 200
 Text2.Text = 400
End If
With Form1.RichTextBox3
 .SelText = "<OBJECT classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://active.macromedia.com/flash2/cabs/swflash.cab#version=4,0,0,0"" ID=Untitled WIDTH="
 .SelText = Text2.Text
 .SelText = " HEIGHT="
 .SelText = Text1.Text & ">"
 .SelText = Chr$(10)
 .SelText = "<PARAM NAME=movie VALUE="
 If Len(Dir1.Path) = 3 Then .SelText = """" & Dir1.Path + File1.FileName
 If Len(Dir1.Path) > 3 Then .SelText = """" & Dir1.Path + "\" + File1.FileName
 .SelText = """" & ">"
 .SelText = Chr$(10)
 .SelText = "<PARAM NAME=quality VALUE=high>"
 .SelText = Chr$(10)
 .SelText = "<PARAM NAME=loop VALUE=false>"
 .SelText = Chr$(10)
 .SelText = "</OBJECT>"
 .SelText = Chr$(10)
End With

Text1.Text = Empty
Text2.Text = Empty
ShockwaveFlash1.Stop
Me.Hide
Form1.Show
End Sub

Private Sub Command2_Click()
Text1.Text = Empty
Text2.Text = Empty
Me.Hide
Form1.Show
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_dblClick()
ShockwaveFlash1.Movie = ""
DoEvents
If Len(Dir1.Path) = 3 Then ShockwaveFlash1.Movie = Dir1.Path + File1.FileName
If Len(Dir1.Path) > 3 Then ShockwaveFlash1.Movie = Dir1.Path + "\" + File1.FileName
End Sub

Private Sub Form_Load()
OnlyNumber(Text1.hWnd) = True
OnlyNumber(Text2.hWnd) = True
Call FormOnTop(Me.hWnd, True)
End Sub
