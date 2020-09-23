VERSION 5.00
Begin VB.Form Form14 
   BorderStyle     =   0  'None
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   2520
   LinkTopic       =   "Form14"
   ScaleHeight     =   3240
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Flash 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   120
      Picture         =   "Form14.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   43
      Top             =   1080
      Width           =   240
   End
   Begin VB.PictureBox ImageM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   260
      Index           =   11
      Left            =   2160
      Picture         =   "Form14.frx":0372
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   42
      Top             =   360
      Width           =   240
   End
   Begin VB.PictureBox Radio 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   2160
      Picture         =   "Form14.frx":07A4
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   41
      Top             =   720
      Width           =   200
   End
   Begin VB.PictureBox Chk 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   2160
      Picture         =   "Form14.frx":09EE
      ScaleHeight     =   195
      ScaleWidth      =   210
      TabIndex        =   40
      Top             =   960
      Width           =   210
   End
   Begin VB.PictureBox Reset 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   2160
      Picture         =   "Form14.frx":0C98
      ScaleHeight     =   195
      ScaleWidth      =   240
      TabIndex        =   39
      Top             =   1200
      Width           =   240
   End
   Begin VB.PictureBox Submit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   2160
      Picture         =   "Form14.frx":0F4A
      ScaleHeight     =   195
      ScaleWidth      =   240
      TabIndex        =   38
      Top             =   1440
      Width           =   240
   End
   Begin VB.PictureBox Button 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   2160
      Picture         =   "Form14.frx":11FC
      ScaleHeight     =   195
      ScaleWidth      =   240
      TabIndex        =   37
      Top             =   1680
      Width           =   240
   End
   Begin VB.PictureBox Textarea 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   2160
      Picture         =   "Form14.frx":14AE
      ScaleHeight     =   195
      ScaleWidth      =   240
      TabIndex        =   36
      Top             =   1920
      Width           =   240
   End
   Begin VB.PictureBox Filebrowser 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   2160
      Picture         =   "Form14.frx":18A8
      ScaleHeight     =   195
      ScaleWidth      =   240
      TabIndex        =   35
      Top             =   2160
      Width           =   240
   End
   Begin VB.PictureBox HiddenTxt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   2160
      Picture         =   "Form14.frx":1C32
      ScaleHeight     =   195
      ScaleWidth      =   240
      TabIndex        =   34
      Top             =   2400
      Width           =   240
   End
   Begin VB.PictureBox Pass 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   2160
      Picture         =   "Form14.frx":1FBC
      ScaleHeight     =   195
      ScaleWidth      =   240
      TabIndex        =   33
      Top             =   2640
      Width           =   240
   End
   Begin VB.PictureBox TextBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   2160
      Picture         =   "Form14.frx":2346
      ScaleHeight     =   195
      ScaleWidth      =   240
      TabIndex        =   32
      Top             =   2880
      Width           =   240
   End
   Begin VB.PictureBox Table1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      Picture         =   "Form14.frx":26D0
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   31
      Top             =   1440
      Width           =   230
   End
   Begin VB.PictureBox Ruller 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   110
      Index           =   0
      Left            =   120
      Picture         =   "Form14.frx":2ACA
      ScaleHeight     =   105
      ScaleWidth      =   210
      TabIndex        =   30
      Top             =   1800
      Width           =   210
   End
   Begin VB.PictureBox Sound 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      Picture         =   "Form14.frx":2C40
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   29
      Top             =   2160
      Width           =   240
   End
   Begin VB.PictureBox Help 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      Picture         =   "Form14.frx":303A
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   28
      Top             =   2520
      Width           =   240
   End
   Begin VB.PictureBox About 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      Picture         =   "Form14.frx":33AC
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   27
      Top             =   2880
      Width           =   255
   End
   Begin VB.PictureBox FontColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   480
      Picture         =   "Form14.frx":3762
      ScaleHeight     =   270
      ScaleWidth      =   240
      TabIndex        =   26
      Top             =   2520
      Width           =   240
   End
   Begin VB.PictureBox Font1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   480
      Picture         =   "Form14.frx":3BDC
      ScaleHeight     =   270
      ScaleWidth      =   240
      TabIndex        =   25
      Top             =   2880
      Width           =   240
   End
   Begin VB.PictureBox Bre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   480
      Picture         =   "Form14.frx":3F7E
      ScaleHeight     =   210
      ScaleWidth      =   165
      TabIndex        =   24
      Top             =   1080
      Width           =   165
   End
   Begin VB.PictureBox Para 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   0
      Left            =   480
      Picture         =   "Form14.frx":41B8
      ScaleHeight     =   165
      ScaleWidth      =   135
      TabIndex        =   23
      Top             =   1440
      Width           =   135
   End
   Begin VB.PictureBox Head 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   0
      Left            =   480
      Picture         =   "Form14.frx":432E
      ScaleHeight     =   165
      ScaleWidth      =   240
      TabIndex        =   22
      Top             =   1800
      Width           =   240
   End
   Begin VB.PictureBox FontSize 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   480
      Picture         =   "Form14.frx":45AC
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   21
      Top             =   2160
      Width           =   240
   End
   Begin VB.PictureBox Elink 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   960
      Picture         =   "Form14.frx":49EA
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   20
      Top             =   1440
      Width           =   230
   End
   Begin VB.PictureBox Link 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   960
      Picture         =   "Form14.frx":4DE4
      ScaleHeight     =   270
      ScaleWidth      =   225
      TabIndex        =   19
      Top             =   1800
      Width           =   230
   End
   Begin VB.PictureBox Left1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   960
      Picture         =   "Form14.frx":51CE
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   18
      Top             =   2160
      Width           =   240
   End
   Begin VB.PictureBox Center1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   960
      Picture         =   "Form14.frx":54B0
      ScaleHeight     =   195
      ScaleWidth      =   240
      TabIndex        =   17
      Top             =   2520
      Width           =   240
   End
   Begin VB.PictureBox Right1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   960
      Picture         =   "Form14.frx":5762
      ScaleHeight     =   195
      ScaleWidth      =   240
      TabIndex        =   16
      Top             =   2880
      Width           =   240
   End
   Begin VB.PictureBox Bold 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   0
      Left            =   960
      Picture         =   "Form14.frx":5A14
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   15
      Top             =   360
      Width           =   165
   End
   Begin VB.PictureBox Underline 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   960
      Picture         =   "Form14.frx":5BE2
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   14
      Top             =   600
      Width           =   165
   End
   Begin VB.PictureBox Italic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   960
      Picture         =   "Form14.frx":5DF8
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   13
      Top             =   840
      Width           =   180
   End
   Begin VB.PictureBox Image1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   960
      Picture         =   "Form14.frx":5FEA
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   240
   End
   Begin VB.PictureBox Find 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   1440
      Picture         =   "Form14.frx":6374
      ScaleHeight     =   240
      ScaleWidth      =   225
      TabIndex        =   11
      Top             =   1440
      Width           =   225
   End
   Begin VB.PictureBox Paste 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1440
      Picture         =   "Form14.frx":66F6
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   10
      Top             =   1800
      Width           =   225
   End
   Begin VB.PictureBox Delete 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   1440
      Picture         =   "Form14.frx":6A68
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   9
      Top             =   2160
      Width           =   225
   End
   Begin VB.PictureBox Cut 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1440
      Picture         =   "Form14.frx":6D1A
      ScaleHeight     =   255
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   2520
      Width           =   200
   End
   Begin VB.PictureBox Copy 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   1440
      Picture         =   "Form14.frx":7004
      ScaleHeight     =   270
      ScaleWidth      =   225
      TabIndex        =   7
      Top             =   2880
      Width           =   225
   End
   Begin VB.PictureBox Open 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1800
      Picture         =   "Form14.frx":73EE
      ScaleHeight     =   255
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   2880
      Width           =   225
   End
   Begin VB.PictureBox Close 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   1800
      Picture         =   "Form14.frx":7760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   720
      Width           =   225
   End
   Begin VB.PictureBox Print 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   1800
      Picture         =   "Form14.frx":7A72
      ScaleHeight     =   270
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   1080
      Width           =   225
   End
   Begin VB.PictureBox New 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   230
      Index           =   0
      Left            =   1800
      Picture         =   "Form14.frx":7E14
      ScaleHeight     =   225
      ScaleWidth      =   210
      TabIndex        =   3
      Top             =   1440
      Width           =   210
   End
   Begin VB.PictureBox SaveAs 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   1800
      Picture         =   "Form14.frx":8116
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   1800
      Width           =   225
   End
   Begin VB.PictureBox Close 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   1800
      Picture         =   "Form14.frx":8428
      ScaleHeight     =   240
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   2160
      Width           =   225
   End
   Begin VB.PictureBox Save 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   1800
      Picture         =   "Form14.frx":87AA
      ScaleHeight     =   270
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   2520
      Width           =   225
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
