VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6315
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   Moveable        =   0   'False
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   5850
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin M2AHTMLMaker.chameleonButton Command1 
      Height          =   375
      Left            =   2920
      TabIndex        =   0
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
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
      MICON           =   "Form9.frx":780F2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   3000
      Top             =   2280
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
Form1.Show
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
intBend = Abs(Val(40))     ' Abs() turns the number to positive
SetWindowRgn Me.hWnd, CreateRoundRectRgn(0, 0, _
                (Me.Width / 15), (Me.Height / 15), _
                intBend, intBend), _
                True
Form11.Hide
Unload Form11
End Sub

Private Sub Timer1_Timer()
Tim = Tim + 1
Form1.Show
Me.Show
If Tim = 2 Then
 Timer1.Enabled = False
 Me.Hide
End If
End Sub
