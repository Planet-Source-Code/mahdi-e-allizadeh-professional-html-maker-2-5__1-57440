VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C3DF5D2F-40CD-4CDD-B283-AA3D32054C81}#1.0#0"; "AutoResize.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "M2A HTML Maker  -  http://www.IranM2A.Tk"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MousePointer    =   1  'Arrow
   ScaleHeight     =   7890
   ScaleWidth      =   11880
   Begin M2AHTMLMaker.chameleonButton Command1 
      Height          =   375
      Left            =   2400
      TabIndex        =   14
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "New HTML Page"
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
      MICON           =   "Form1.frx":058A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.AutoResize Resize 
      Left            =   5520
      Tag             =   "NO"
      Top             =   4560
      _ExtentX        =   714
      _ExtentY        =   714
      AspectRatioValue=   0
   End
   Begin RichTextLib.RichTextBox RichTextBox3 
      Height          =   6615
      Left            =   2460
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   11668
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      MousePointer    =   3
      TextRTF         =   $"Form1.frx":05A6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6885
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   12144
      _Version        =   393216
      MousePointer    =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Browser"
      TabPicture(0)   =   "Form1.frx":07E7
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape22"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Drive1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Dir1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "File1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Java Script"
      TabPicture(1)   =   "Form1.frx":0803
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape34"
      Tab(1).Control(1)=   "Command24"
      Tab(1).Control(2)=   "Command16"
      Tab(1).Control(3)=   "Command15"
      Tab(1).ControlCount=   4
      Begin VB.FileListBox File1 
         Height          =   3015
         Left            =   120
         TabIndex        =   10
         Top             =   3760
         Width           =   2055
      End
      Begin VB.DirListBox Dir1 
         Height          =   3015
         Left            =   100
         TabIndex        =   9
         Top             =   750
         Width           =   2055
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   115
         TabIndex        =   8
         Top             =   440
         Width           =   2070
      End
      Begin M2AHTMLMaker.chameleonButton Command15 
         Height          =   375
         Left            =   -74880
         TabIndex        =   27
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "Scrolling Banner"
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
         MICON           =   "Form1.frx":081F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin M2AHTMLMaker.chameleonButton Command16 
         Height          =   375
         Left            =   -74880
         TabIndex        =   28
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "Last Data Modified"
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
         MICON           =   "Form1.frx":083B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin M2AHTMLMaker.chameleonButton Command24 
         Height          =   375
         Left            =   -74880
         TabIndex        =   29
         Top             =   1800
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "Get More Java script"
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
         MICON           =   "Form1.frx":0857
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Shape Shape34 
         BorderColor     =   &H00FFC0C0&
         FillColor       =   &H00FFC0C0&
         FillStyle       =   3  'Vertical Line
         Height          =   1920
         Left            =   -74925
         Top             =   435
         Width           =   2145
      End
      Begin VB.Shape Shape22 
         BorderColor     =   &H00FF8080&
         FillColor       =   &H00FF8080&
         FillStyle       =   3  'Vertical Line
         Height          =   6430
         Left            =   90
         Top             =   390
         Width           =   2115
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   120
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   3
      FontName        =   "MS Sans Serif"
      FontSize        =   10
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   480
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0873
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0985
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A97
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0BA9
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CBB
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0DCD
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0EDF
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0FF1
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1103
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1215
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1327
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1439
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":154B
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":165D
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":177D
            Key             =   "Undo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      DisabledImageList=   "imlToolbarIcons"
      HotImageList    =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageKey        =   "Align Left"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageKey        =   "Center"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageKey        =   "Align Right"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageKey        =   "Find"
         EndProperty
      EndProperty
      MousePointer    =   1
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   6855
         Left            =   0
         TabIndex        =   3
         Top             =   1320
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   12091
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"Form1.frx":1DF9
      End
      Begin RichTextLib.RichTextBox RichTextBox2 
         Height          =   6855
         Left            =   120
         TabIndex        =   2
         Top             =   2520
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   12091
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"Form1.frx":1E89
      End
      Begin VB.Shape Shape23 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   7440
         Top             =   -120
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Status Bar"
      Top             =   7575
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15293
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "05-02-2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "13.19"
         EndProperty
      EndProperty
      MousePointer    =   1
   End
   Begin VB.TextBox Text7 
      Height          =   6495
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Form1.frx":1F19
      Top             =   1200
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   495
      Left            =   5400
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2520
      Top             =   5040
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5400
      TabIndex        =   13
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2880
      Top             =   6840
   End
   Begin M2AHTMLMaker.chameleonButton Command4 
      Height          =   1455
      Left            =   11160
      TabIndex        =   15
      Top             =   6000
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   2566
      BTYPE           =   14
      TX              =   "Clear  Sel  Text"
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
      MICON           =   "Form1.frx":2069
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
      Left            =   4520
      TabIndex        =   16
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Tag High Lighting"
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
      MICON           =   "Form1.frx":2085
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
      Left            =   6800
      TabIndex        =   17
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Test HTML Page"
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
      MICON           =   "Form1.frx":20A1
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
      Left            =   9080
      TabIndex        =   18
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Refresh"
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
      MICON           =   "Form1.frx":20BD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin M2AHTMLMaker.chameleonButton Command12 
      Height          =   495
      Left            =   11160
      TabIndex        =   19
      Top             =   480
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Page's Title "
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
      MICON           =   "Form1.frx":20D9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin M2AHTMLMaker.chameleonButton Command5 
      Height          =   495
      Left            =   11160
      TabIndex        =   20
      Top             =   1080
      Width           =   615
      _ExtentX        =   1085
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
      MICON           =   "Form1.frx":20F5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin M2AHTMLMaker.chameleonButton Command10 
      Height          =   495
      Left            =   11160
      TabIndex        =   21
      Top             =   1680
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
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
      MICON           =   "Form1.frx":2111
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
      Left            =   11160
      TabIndex        =   22
      Top             =   2520
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "<BR>"
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
      MICON           =   "Form1.frx":212D
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
      Left            =   11160
      TabIndex        =   23
      Top             =   3120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "<P>"
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
      MICON           =   "Form1.frx":2149
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin M2AHTMLMaker.chameleonButton Command13 
      Height          =   495
      Left            =   11160
      TabIndex        =   24
      Top             =   3960
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Insert Link"
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
      MICON           =   "Form1.frx":2165
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin M2AHTMLMaker.chameleonButton Command14 
      Height          =   495
      Left            =   11160
      TabIndex        =   25
      Top             =   4560
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Insert Picture"
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
      MICON           =   "Form1.frx":2181
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin M2AHTMLMaker.chameleonButton Command11 
      Height          =   495
      Left            =   11160
      TabIndex        =   26
      Top             =   5160
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Email Link"
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
      MICON           =   "Form1.frx":219D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape32 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   11190
      Top             =   5780
      Width           =   585
   End
   Begin VB.Shape Shape31 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   11190
      Top             =   3730
      Width           =   585
   End
   Begin VB.Shape Shape30 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   155
      Left            =   11190
      Top             =   2290
      Width           =   580
   End
   Begin VB.Shape Shape27 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   11160
      Top             =   6000
      Width           =   615
   End
   Begin VB.Shape Shape26 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   1432
      Left            =   11235
      Top             =   6090
      Width           =   607
   End
   Begin VB.Shape Shape25 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   11160
      Top             =   5160
      Width           =   615
   End
   Begin VB.Shape Shape24 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   11220
      Top             =   5240
      Width           =   600
   End
   Begin VB.Shape Shape21 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   11160
      Top             =   4560
      Width           =   615
   End
   Begin VB.Shape Shape20 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   11220
      Top             =   4640
      Width           =   600
   End
   Begin VB.Shape Shape19 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   11160
      Top             =   3960
      Width           =   615
   End
   Begin VB.Shape Shape18 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   11220
      Top             =   4040
      Width           =   600
   End
   Begin VB.Shape Shape17 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   11160
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape16 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   11220
      Top             =   3200
      Width           =   600
   End
   Begin VB.Shape Shape15 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   11160
      Top             =   2520
      Width           =   615
   End
   Begin VB.Shape Shape14 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   11220
      Top             =   2600
      Width           =   600
   End
   Begin VB.Shape Shape13 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   11160
      Top             =   1680
      Width           =   615
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   11220
      Top             =   1760
      Width           =   600
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   11160
      Top             =   480
      Width           =   615
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   11210
      Top             =   560
      Width           =   600
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   11160
      Top             =   1080
      Width           =   615
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   475
      Left            =   11230
      Top             =   1160
      Width           =   585
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   9080
      Top             =   480
      Width           =   1935
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   9180
      Top             =   570
      Width           =   1885
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6800
      Top             =   480
      Width           =   2055
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   6900
      Top             =   570
      Width           =   1995
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   380
      Left            =   4520
      Top             =   480
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   4620
      Top             =   570
      Width           =   1995
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   3
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2400
      Top             =   480
      Width           =   1935
   End
   Begin VB.Shape Shape28 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   2505
      Top             =   570
      Width           =   1880
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   6495
      Left            =   2460
      Picture         =   "Form1.frx":21B9
      Stretch         =   -1  'True
      Top             =   960
      Width           =   8580
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuFileKh0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveas 
         Caption         =   "Save as"
      End
      Begin VB.Menu mnuFileKh1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Print"
      End
      Begin VB.Menu Khatsh1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRecentFiles 
         Caption         =   "Recent Files"
         Begin VB.Menu mnuFileRecentFiles1 
            Caption         =   "Untitled 1"
         End
         Begin VB.Menu mnuFileRecentFiles2 
            Caption         =   "Untitled 2"
         End
         Begin VB.Menu mnuFileRecentFiles3 
            Caption         =   "Untitled 3"
         End
         Begin VB.Menu mnuFileRecentFiles4 
            Caption         =   "Untitled 4"
         End
         Begin VB.Menu mnuFileRecentFiles5 
            Caption         =   "Untitled 5"
         End
         Begin VB.Menu mnuFileRecentFiles6 
            Caption         =   "Untitled 6"
         End
         Begin VB.Menu mnuFileRecentFiles7 
            Caption         =   "Untitled 7"
         End
      End
      Begin VB.Menu mnuFileKh2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuEditkhg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu khasd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditSelAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuKh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewWebBrowser 
         Caption         =   "Test HTML Page"
      End
      Begin VB.Menu mnuViewMap 
         Caption         =   "Map Of The Text Color"
      End
   End
   Begin VB.Menu mnuTags 
      Caption         =   "&Insert Tags"
      Begin VB.Menu mnuTagsInsertLink 
         Caption         =   "Insert Link"
      End
      Begin VB.Menu mnuTagsInsertPicture 
         Caption         =   "Insert Picture"
      End
      Begin VB.Menu mnuTagsEmail 
         Caption         =   "Insert Email Link"
      End
      Begin VB.Menu mnuTagsSound 
         Caption         =   "Insert Sound"
      End
      Begin VB.Menu mnuTagsFlash 
         Caption         =   "Insert Flash"
      End
      Begin VB.Menu mnuTagsTable 
         Caption         =   "Insert Table"
      End
      Begin VB.Menu mnuTagsLine 
         Caption         =   "Insert Line"
      End
      Begin VB.Menu mnuTagsDateAndTime 
         Caption         =   "Insert Date And Time"
      End
      Begin VB.Menu mnuKh0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTagsFonts 
         Caption         =   "Font"
         Begin VB.Menu mnuFontS 
            Caption         =   "Font Size"
            Begin VB.Menu mnuFF1 
               Caption         =   "Font Size 1"
            End
            Begin VB.Menu mnuFF2 
               Caption         =   "Font Size 2"
            End
            Begin VB.Menu mnuFF3 
               Caption         =   "Font Size 3"
            End
            Begin VB.Menu mnuFF4 
               Caption         =   "Font Size 4"
            End
            Begin VB.Menu mnuFF5 
               Caption         =   "Font Size 5"
            End
            Begin VB.Menu mnuFF6 
               Caption         =   "Font Size 6"
            End
            Begin VB.Menu mnuFF7 
               Caption         =   "Font Size 7"
            End
         End
         Begin VB.Menu mnuFontC 
            Caption         =   "Font Color"
         End
      End
      Begin VB.Menu mnuTagsHSize 
         Caption         =   "Header Size"
         Begin VB.Menu mnuHD1 
            Caption         =   "Header Size 1"
         End
         Begin VB.Menu mnuHD2 
            Caption         =   "Header Size 2"
         End
         Begin VB.Menu mnuHD3 
            Caption         =   "Header Size 3"
         End
         Begin VB.Menu mnuHD4 
            Caption         =   "Header Size 4"
         End
         Begin VB.Menu mnuHD5 
            Caption         =   "Header Size 5"
         End
         Begin VB.Menu mnuHD6 
            Caption         =   "Header Size 6"
         End
      End
      Begin VB.Menu mnuKh1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTagsLeft 
         Caption         =   "Left"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuTagsCenter 
         Caption         =   "Center"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuTagsRight 
         Caption         =   "Right"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuKh2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTagsParagraph 
         Caption         =   "<P> Parageraph"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuTagsBreak 
         Caption         =   "<BR> Break"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuKh3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTagsBold 
         Caption         =   "Bold"
      End
      Begin VB.Menu mnuTagsItalic 
         Caption         =   "Italic"
      End
      Begin VB.Menu mnuTagsUnder 
         Caption         =   "Under Line"
      End
      Begin VB.Menu khat12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTagsChkBox 
         Caption         =   "Insert Check Box"
      End
      Begin VB.Menu mnuTagsRadioButton 
         Caption         =   "Insert Radio Button"
      End
      Begin VB.Menu mnuTagsImage 
         Caption         =   "Insert Image Box"
      End
      Begin VB.Menu mnuTagskh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTagsTxtBox 
         Caption         =   "Insert Text Box"
      End
      Begin VB.Menu mnuTagsPassTxtBox 
         Caption         =   "Insert Password Text Box"
      End
      Begin VB.Menu mnuTagsHiddenTextBox 
         Caption         =   "Insert Hidden Text Box"
      End
      Begin VB.Menu mnuTagsFileBrowserTxtBox 
         Caption         =   "Insert File Browser Text Box"
      End
      Begin VB.Menu mnuTagsTxtArea 
         Caption         =   "Insert Text Area"
      End
      Begin VB.Menu mnuTagskhat100 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTagsButton 
         Caption         =   "Insert Button"
      End
      Begin VB.Menu mnuTagsSubmitButton 
         Caption         =   "Insert Submit Button"
      End
      Begin VB.Menu mnuTagsResetButton 
         Caption         =   "Insert Reset Button"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTest 
         Caption         =   "Test HTML Page"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuTagh 
         Caption         =   "Tag High Lighting"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsOp 
         Caption         =   "Settings Of M2A HTML Maker"
      End
      Begin VB.Menu mnuOptionskhat 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsFont 
         Caption         =   "Set The Font"
      End
      Begin VB.Menu mnuOptionsDefaultFont 
         Caption         =   "Set Default Font"
      End
      Begin VB.Menu mnuOptionskhat2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsSave 
         Caption         =   "Save All Settings"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpH 
         Caption         =   "HTML Help - "
         Shortcut        =   {F1}
      End
      Begin VB.Menu Khat 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpwww 
         Caption         =   "www.IranM2A.tk"
      End
      Begin VB.Menu mnuHelpA 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnupopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnupopupR 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnupopupkh 
         Caption         =   "-"
      End
      Begin VB.Menu mnupopupCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnupopupCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnupopupPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnupopupDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuPopKh 
         Caption         =   "-"
      End
      Begin VB.Menu mnupopupSelAll 
         Caption         =   "Select All"
      End
   End
   Begin VB.Menu mnupopup2 
      Caption         =   "Popup2"
      Visible         =   0   'False
      Begin VB.Menu mnupopup2Del 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnupopup3 
      Caption         =   "Popup3"
      Visible         =   0   'False
      Begin VB.Menu mnupopup3Mkdir 
         Caption         =   "Make Directory"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CmdCounter As Byte 'This varriable is for to prevent repeat alittle command whit an if
Public CmdCounter1 As Byte 'See Command1_MouseDown
Public CmdCounter2 As Byte
Public CmdCounter3 As Byte
Public CmdCounter4 As Byte
Public TypeOfSave As Byte
Public For_File_Name As Byte
Public cnt As Byte
Public ForOneClickBR, ForOneClickP As Byte
Public SaveCaption As Variant
Public He, Wi As Variant

Private Sub LoadNewDoc()
'This sub is for create a new html page
On Error Resume Next
    DoEvents
    RichTextBox3.FileName = Empty
    If RichTextBox3.Text = "" Then RichTextBox3.Text = Text7.Text
    Image1.Visible = False
    RichTextBox3.Visible = True
    Command1.Caption = "Close HTML Page"
    mnuFileClose.Enabled = True
    Form5.Show
    RichTextBox3.SelStart = 270
    TypeOfSave = 0
    Var1 = 1
    Txt2 = 0
    For_File_Name = 0
End Sub

Public Function SelColor()
'This function is a function that can change the selected color in HTML code
With Form1.RichTextBox3
Select Case dlgCommonDialog.Color
             Case 0
             .SelText = "000000"
             Case 64
              .SelText = "400000"
             Case 128
              .SelText = "800000"
             Case 4210816
              .SelText = "804040"
             Case 255
              .SelText = "FF0000"
             Case 8421631
              .SelText = "FF8080"
             Case 32896
              .SelText = "808000"
             Case 16512
              .SelText = "804000"
             Case 33023
              .SelText = "FF8000"
             Case 4227327
              .SelText = "FF8040"
             Case 65535
              .SelText = "FFFF00"
             Case 8454143
              .SelText = "FFFF80"
             Case 4227200
              .SelText = "808040"
             Case 16384
              .SelText = "004000"
             Case 32768
              .SelText = "008000"
             Case 65280
              .SelText = "00FF00"
             Case 65408
              .SelText = "80FF00"
             Case 8454016
              .SelText = "80FF80"
             Case 8421504
              .SelText = "808080"
             Case 4210688
              .SelText = "004040"
             Case 4227072
              .SelText = "008040"
             Case 8421376
              .SelText = "008080"
             Case 4259584
              .SelText = "00FF40"
             Case 8453888
              .SelText = "00FF80"
             Case 8421440
              .SelText = "408080"
             Case 8388608
              .SelText = "000080"
             Case 16711680
              .SelText = "0000FF"
             Case 8404992
              .SelText = "004080"
             Case 16776960
              .SelText = "00FFFF"
             Case 16777088
              .SelText = "80FFFF"
             Case 12632256
              .SelText = "C0C0C0"
             Case 4194304
              .SelText = "000040"
             Case 10485760
              .SelText = "0000A0"
             Case 16744576
              .SelText = "8080FF"
             Case 12615680
              .SelText = "0080C0"
             Case 16744448
              .SelText = "0080FF"
             Case 4194368
              .SelText = "400040"
             Case 4194368
              .SelText = "400040"
             Case 8388736
              .SelText = "800080"
             Case 4194432
              .SelText = "800040"
             Case 12615808
              .SelText = "8080C0"
             Case 12615935
              .SelText = "FF80C0"
             Case 16777215
              .SelText = "FFFFFF"
             Case 8388672
              .SelText = "400080"
             Case 16711808
              .SelText = "8000FF"
             Case 8388863
              .SelText = "FF0080"
             Case 16711935
              .SelText = "FF00FF"
             Case 16744703
              .SelText = "FF80FF"
            End Select
End With
End Function

Public Function ChangeBkColor()
'This function is a function that can change button's back ground color
DoEvents
CmdCounter = 0

For Each ctl In Controls
 If TypeOf ctl Is Shape Then
     ctl.BorderColor = 16761024
 End If
Next
End Function

Private Sub Recent()
'It is for recent menu in file menu
 SaveCaption = mnuFileRecentFiles1.Caption
 mnuFileRecentFiles1.Caption = RichTextBox3.FileName
If mnuFileRecentFiles1.Caption = SaveCaption Then
 Exit Sub
Else
 mnuFileRecentFiles7.Caption = mnuFileRecentFiles6.Caption
 mnuFileRecentFiles6.Caption = mnuFileRecentFiles5.Caption
 mnuFileRecentFiles5.Caption = mnuFileRecentFiles4.Caption
 mnuFileRecentFiles4.Caption = mnuFileRecentFiles3.Caption
 mnuFileRecentFiles3.Caption = mnuFileRecentFiles2.Caption
 mnuFileRecentFiles2.Caption = SaveCaption
End If
End Sub

Private Sub ForFont()
RichTextBox3.Font.Size = dlgCommonDialog.FontSize
RichTextBox3.Font.Name = dlgCommonDialog.FontName
RichTextBox3.Font.Bold = dlgCommonDialog.FontBold
RichTextBox3.Font.Italic = dlgCommonDialog.FontItalic
RichTextBox3.Font.Strikethrough = dlgCommonDialog.FontStrikethru
RichTextBox3.Font.Underline = dlgCommonDialog.FontUnderline
End Sub

Private Sub ForLeft()
  If RichTextBox3.Visible = True Then
       bb = RichTextBox3.SelStart
       ss = RichTextBox3.SelLength
      RichTextBox3.SelStart = RichTextBox3.SelStart
     RichTextBox3.SelText = "<div align="
     RichTextBox3.SelText = """"
     RichTextBox3.SelText = "Left"
     RichTextBox3.SelText = """"
     RichTextBox3.SelText = ">"
       RichTextBox3.SelStart = Val(RichTextBox3.SelStart) + Val(ss)
      RichTextBox3.SelText = "</div>"
       Command7_Click
      RichTextBox3.SelStart = bb + 18
      RichTextBox3.SelLength = ss
   End If
End Sub

Private Sub ForCenter()
If RichTextBox3.Visible = True Then
    bb = RichTextBox3.SelStart
    ss = RichTextBox3.SelLength
      RichTextBox3.SelStart = RichTextBox3.SelStart
     RichTextBox3.SelText = "<Center>"
       RichTextBox3.SelStart = Val(RichTextBox3.SelStart) + Val(ss)
      RichTextBox3.SelText = "</Center>"
  Command7_Click
   RichTextBox3.SelStart = bb + 8
   RichTextBox3.SelLength = ss
End If
End Sub

Private Sub ForRight()
  If RichTextBox3.Visible = True Then
      bb = RichTextBox3.SelStart
      ss = RichTextBox3.SelLength
      RichTextBox3.SelStart = RichTextBox3.SelStart
     RichTextBox3.SelText = "<div align="
     RichTextBox3.SelText = """"
     RichTextBox3.SelText = "Right"
     RichTextBox3.SelText = """"
     RichTextBox3.SelText = ">"
       RichTextBox3.SelStart = Val(RichTextBox3.SelStart) + Val(ss)
      RichTextBox3.SelText = "</div>"
     Command7_Click
        RichTextBox3.SelStart = bb + 19
   RichTextBox3.SelLength = ss
  End If
End Sub

Private Sub Refresh_The_Dir()
On Error Resume Next
SaveDir = Dir1.Path
Dir1.Path = "c:\"
Dir1.Path = "d:\"
Dir1.Path = SaveDir
End Sub

Public Function ForFontSize(MenuName As String)
On Error Resume Next
'It's for set the html code font
With RichTextBox3
bb = .SelStart
ss = .SelLength
      .SelStart = .SelStart
     .SelText = "<Font Size="
     .SelText = """"
     Select Case MenuName
      Case "mnuFF1"
       .SelText = "1"
      Case "mnuFF2"
       .SelText = "2"
      Case "mnuFF3"
       .SelText = "3"
      Case "mnuFF4"
       .SelText = "4"
      Case "mnuFF5"
       .SelText = "5"
      Case "mnuFF6"
       .SelText = "6"
      Case "mnuFF7"
       .SelText = "7"
     End Select
     .SelText = """"
     .SelText = ">"
       .SelStart = Val(.SelStart) + Val(ss)
      .SelText = "</Font>"
  Command7_Click
 .SelStart = bb + 15
 .SelLength = ss
End With
End Function

Public Function ForHeadingFont(MenuName2 As String)
On Error Resume Next
With RichTextBox3
bb = .SelStart
ss = .SelLength
      .SelStart = .SelStart
      Select Case MenuName2
       Case "mnuHD1"
        .SelText = "<H1>"
        .SelStart = Val(.SelStart) + Val(ss)
        .SelText = "</H1>"
       Case "mnuHD2"
        .SelText = "<H2>"
        .SelStart = Val(.SelStart) + Val(ss)
        .SelText = "</H2>"
       Case "mnuHD3"
        .SelText = "<H3>"
        .SelStart = Val(.SelStart) + Val(ss)
        .SelText = "</H3>"
       Case "mnuHD4"
        .SelText = "<H4>"
        .SelStart = Val(.SelStart) + Val(ss)
        .SelText = "</H4>"
       Case "mnuHD5"
        .SelText = "<H5>"
        .SelStart = Val(.SelStart) + Val(ss)
        .SelText = "</H5>"
       Case "mnuHD6"
        .SelText = "<H6>"
        .SelStart = Val(.SelStart) + Val(ss)
        .SelText = "</H6>"
      End Select
  Command7_Click
 .SelStart = bb + 4
 .SelLength = ss
End With
End Function

Public Function ForRecents(ForRecentName As String)
On Error Resume Next
If Command1.Caption = "Close HTML Page" Then
 If Txt2 = 1 Then
  Beep
  myq = MsgBox("Do you want to save this page ?", vbYesNoCancel + vbQuestion, "Save...")
 End If
 If myq = 6 Then
  mnuFileSave_Click
  mnuFileClose.Enabled = False
  RichTextBox3.Text = ""
  RichTextBox3.Visible = False
  Image1.Visible = True
  Command1.Caption = "New HTML Page"
  Txt1 = "0"
 End If
 If myq = 7 Then
  mnuFileClose.Enabled = False
  RichTextBox3.Text = ""
  RichTextBox3.Visible = False
  Image1.Visible = True
  Command1.Caption = "New HTML Page"
  Txt1 = "0"
 End If
 If myq = 2 Then Exit Function
End If
Image1.Visible = False
RichTextBox3.Visible = True
RichTextBox3.Text = ""
Select Case ForRecentName
 Case "mnuFileRecentFiles1"
  RichTextBox3.FileName = mnuFileRecentFiles1.Caption
 Case "mnuFileRecentFiles2"
  RichTextBox3.FileName = mnuFileRecentFiles2.Caption
 Case "mnuFileRecentFiles3"
  RichTextBox3.FileName = mnuFileRecentFiles3.Caption
 Case "mnuFileRecentFiles4"
  RichTextBox3.FileName = mnuFileRecentFiles4.Caption
 Case "mnuFileRecentFiles5"
  RichTextBox3.FileName = mnuFileRecentFiles5.Caption
 Case "mnuFileRecentFiles6"
  RichTextBox3.FileName = mnuFileRecentFiles6.Caption
 Case "mnuFileRecentFiles7"
  RichTextBox3.FileName = mnuFileRecentFiles7.Caption
End Select
Txt2 = 0
If Command1.Caption = "New HTML Page" Then Command1.Caption = "Close HTML Page"
End Function

Private Sub Command1_Click()
On Error Resume Next
If Command1.Caption = "Close HTML Page" Then
 Form5.Hide
 If Txt2 = 1 Then
  Beep
  myq = MsgBox("Do you want to save this page ?", vbYesNoCancel + vbQuestion, "Save...")
 End If
 If myq = 6 Then  'If myq = Yes then
  mnuFileSave_Click
  mnuFileClose.Enabled = False
  RichTextBox3.Text = ""
  RichTextBox3.Visible = False
  Image1.Visible = True
  Command1.Caption = "New HTML Page"
  Txt1 = "0"
 End If
 
 If myq = 7 Then 'If myq = No then
  mnuFileClose.Enabled = False
  RichTextBox3.Text = ""
  RichTextBox3.Visible = False
  Image1.Visible = True
  Command1.Caption = "New HTML Page"
  Txt1 = "0"
 Call Recent
 
 End If
 If myq = 2 Then 'If myq = Cancel then
  Exit Sub
 End If
If Txt2 = 0 Then
  mnuFileClose.Enabled = False
  RichTextBox3.Text = ""
  RichTextBox3.Visible = False
  Image1.Visible = True
  Command1.Caption = "New HTML Page"
  Txt1 = "0"
End If
End If
If Txt1 = "0" Then
 Txt1 = "1"
 GoTo 100
End If
If CmdCounter4 = 1 Then
 CmdCounter4 = 0
 GoTo 100
End If
If Command1.Caption = "New HTML Page" Then
  NameSave = "Untitled"
  LoadNewDoc
End If
100
Txt2 = 0
CmdCounter3 = 1
Call Refresh_The_Dir
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CmdCounter1 = 0 Then If Button = 1 Then Command1_Click 'In the first time you click on Command1(New HTML Page) this click not do any thing
'Then for destroy this problem i use above code and under code
CmdCounter1 = 1
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Command1.Caption = "Close HTML Page" Then
 Command1.ToolTipText = "Close HTML Page"
 sbStatusBar.SimpleText = "Close HTML Page"
Else
 Command1.ToolTipText = "New HTML Page"
 sbStatusBar.SimpleText = "New HTML Page"
End If
If CmdCounter = 0 Then
 If Rang = 1 Then Call ChangeBkColor
 Shape7.BorderColor = 8454143
 '16744576
 Rang = 1
End If
CmdCounter = 1
End Sub
Private Sub Command10_Click()
If Form1.RichTextBox3.Visible = True Then Form4.Show
End Sub

Private Sub Command10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "Background Color - Link Color And ..."
If CmdCounter = 0 Then
 If Rang = 1 Then Call ChangeBkColor
 Shape13.BorderColor = 8454143
 Rang = 1
End If
If Button = 1 Then Command10_Click
CmdCounter = 1
End Sub

Private Sub Command11_Click()
Dim Email
If Form1.RichTextBox3.Visible = True Then
Email = InputBox("E_mail Address to Link to:", "Insert Email Link")
With RichTextBox3
 bb = .SelStart
 ss = .SelLength
      .SelStart = .SelStart
     .SelText = "<A HREF="
     .SelText = """"
     .SelText = "mailto:"
     .SelText = Email
     .SelText = """"
     .SelText = ">"
       .SelStart = Val(.SelStart) + Val(ss)
      .SelText = "</A>"
     Command7_Click
      .SelStart = bb + (18 + Len(Email))
      .SelLength = ss
End With
End If
End Sub

Private Sub Command11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "Insert Link In Your Selected Text"
If CmdCounter = 0 Then
 If Rang = 1 Then Call ChangeBkColor
 Shape25.BorderColor = 8454143
 Rang = 1
End If
If Button = 1 Then Command11_Click
CmdCounter = 1
End Sub

Private Sub Command12_Click()
If Form1.RichTextBox3.Visible = True Then
Dim Inp, Sa, Sa2
Sa = InStr(1, RichTextBox3.Text, "TITLE>", vbTextCompare)
RichTextBox3.SelStart = Sa + 5
RichTextBox3.SelLength = Form5.Text3.Text
Inp = InputBox("What is your Page's Name ?", "What is your Page's Name?")
Form5.Text3.Text = Len(Inp)
RichTextBox3.SelText = Inp
RichTextBox3.Visible = False
Image1.Visible = False
RichTextBox3.Visible = True
Command1.Caption = "Close HTML Page"
RichTextBox3.SelStart = 270
End If
End Sub

Private Sub Command12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "Name Of The Page"
If CmdCounter = 0 Then
 If Rang = 1 Then Call ChangeBkColor
 Shape11.BorderColor = 8454143
 Rang = 1
End If
If Button = 1 Then Command12_Click
CmdCounter = 1
End Sub

Private Sub Command13_Click()
If Form1.RichTextBox3.Visible = True Then
Form6.Show
End If
End Sub

Private Sub Command13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "Insert Link In Your Selected Text"
If CmdCounter = 0 Then
 If Rang = 1 Then Call ChangeBkColor
 Shape19.BorderColor = 8454143
 Rang = 1
End If
If Button = 1 Then Command13_Click
CmdCounter = 1
End Sub

Private Sub Command14_Click()
If Form1.RichTextBox3.Visible = True Then
Form7.Show
End If
End Sub

Private Sub Command14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "Insert Picture"
If CmdCounter = 0 Then
 If Rang = 1 Then Call ChangeBkColor
 Shape21.BorderColor = 8454143
 Rang = 1
End If
If Button = 1 Then Command14_Click
CmdCounter = 1
End Sub

Private Sub Command15_Click()
If Form1.RichTextBox3.Visible = True Then
Image1.Visible = False
Command1.Caption = "Close HTML Page"
With RichTextBox3
 .SelText = "<script language="
 .SelText = """"
 .SelText = "JavaScript"
 .SelText = """"
 .SelText = ">                             "
 .SelText = "var id,pause=0,position=0,revol=9;  function banner() {  var i,k;   var msg="
 .SelText = """"
 .SelText = "                         Your Text Here                             "
 .SelText = """"
 .SelText = ";  var speed=10;  document.thisform.thisbanner.value=msg.substring(position,position+50);  if(position++==msg.length)   {      if (revol-- < 2) return;      position=0;  }  id=setTimeout("
 .SelText = """"
 .SelText = "banner()"
 .SelText = """"
 .SelText = ",1000/speed);}</script></head><body bgcolor="
 .SelText = """"
 .SelText = "ffffff"
 .SelText = """"
 .SelText = "onload="
 .SelText = """"
 .SelText = "banner()"
 .SelText = """"
 .SelText = "><form name="
 .SelText = """"
 .SelText = "thisform"
 .SelText = """"
 .SelText = "><input type="
 .SelText = """"
 .SelText = "text"
 .SelText = """"
 .SelText = " name="
 .SelText = """"
 .SelText = "thisbanner"
 .SelText = """"
 .SelText = " size="
 .SelText = """"
 .SelText = "40"
 .SelText = """"
 .SelText = "></FORM>"
End With
Command7_Click
End If
End Sub

Private Sub Command15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "It Is A Java Script"
If CmdCounter = 0 Then
 If Rang = 1 Then Call ChangeBkColor
 Command15.BackColor = &HFFFFFF
 Command15.FontSize = 10
 Rang = 1
End If
CmdCounter = 1
End Sub

Private Sub Command16_Click()
If Form1.RichTextBox3.Visible = True Then
Image1.Visible = False
Command1.Caption = "Close HTML Page"
With RichTextBox3
 .SelText = "<SCRIPT LANGUAGE="
 .SelText = """"
 .SelText = "JavaScript"
 .SelText = """"
 .SelText = ">  var dateMod = "
 .SelText = """"""
 .SelText = "  ;dateMod = document.lastModified  ;document.write("
 .SelText = """"
 .SelText = "Last Updated:  "
 .SelText = """"
 .SelText = ");  document.write(dateMod);  document.write(); </SCRIPT>"
End With
Command7_Click
End If
End Sub

Private Sub Command16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "It Is A Java Script"
If CmdCounter = 0 Then
 If Rang = 1 Then Call ChangeBkColor
 Command16.BackColor = &HFFFFFF
 Command16.FontSize = 10
 Rang = 1
End If
CmdCounter = 1
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Command2_Click
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "Only Go To Next Line"
If CmdCounter = 0 Then
 If Rang = 1 Then Call ChangeBkColor
 Shape17.BorderColor = 8454143
 Rang = 1
End If
CmdCounter = 1
End Sub

Private Sub Command24_Click()
On Error Resume Next
MsgBox "Get More Java Scipt in Program's Directory (In JS.zip (In instalation version)) and you can Get More Java Scipt in http://Javascript.internet.com  And get Funy Picture and Funy Photo in http://www.Andyart.com", vbOKOnly, "Get more Java Script"
Shell "Explorer http://Javascript.internet.com"
End Sub

Private Sub Command24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "Get More Java Script"
If CmdCounter = 0 Then
 If Rang = 1 Then Call ChangeBkColor
 Command24.BackColor = &HFFFFFF
 Command24.FontSize = 10
 Rang = 1
End If
CmdCounter = 1
End Sub
Private Sub Command25_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "See The Browser"
End Sub
Private Sub Command26_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "Use The Java Script In your HTML Page"
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Command3_Click
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "Go To Next Lines "
If CmdCounter = 0 Then
 If Rang = 1 Then Call ChangeBkColor
 Shape15.BorderColor = 8454143
 Rang = 1
End If
CmdCounter = 1
End Sub

Private Sub Command4_Click()
If Form1.RichTextBox3.Visible = True Then RichTextBox3.SelText = " "
End Sub

Private Sub Command2_Click()
If ForOneClickP = 0 Then
 If Form1.RichTextBox3.Visible = True Then
 If RichTextBox3.SelLength <= 0 Then
  RichTextBox3.SelText = "<P>"
  RichTextBox3.SelText = Chr$(10)
 End If
 If RichTextBox3.SelLength > 0 Then
  bb = RichTextBox3.SelStart
  ss = RichTextBox3.SelLength
  RichTextBox3.SelStart = RichTextBox3.SelStart
  RichTextBox3.SelText = "<P>"
  RichTextBox3.SelStart = RichTextBox3.SelStart + ss
  RichTextBox3.SelText = "</P>"
  RichTextBox3.SelStart = bb + 3
  RichTextBox3.SelLength = ss
 End If
 Command7_Click
 End If
 ForOneClickP = 1
Else
 ForOneClickP = 0
End If
End Sub

Private Sub Command3_Click()
If ForOneClickBR = 0 Then
 If Form1.RichTextBox3.Visible = True Then
 If RichTextBox3.SelLength <= 0 Then
  RichTextBox3.SelText = "<BR>"
  RichTextBox3.SelText = Chr$(10)
 End If
 If RichTextBox3.SelLength > 0 Then
  bb = RichTextBox3.SelStart
  ss = RichTextBox3.SelLength
  RichTextBox3.SelStart = RichTextBox3.SelStart
  RichTextBox3.SelText = "<BR>"
  RichTextBox3.SelStart = RichTextBox3.SelStart + ss
  RichTextBox3.SelText = "</BR>"
  RichTextBox3.SelStart = bb + 4
  RichTextBox3.SelLength = ss
 End If
 Command7_Click
 End If
 ForOneClickBR = 1
Else
 ForOneClickBR = 0
End If
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "Clear Selected Text"
If CmdCounter = 0 Then
 If Rang = 1 Then Call ChangeBkColor
 Shape27.BorderColor = 8454143
 Rang = 1
End If
If Button = 1 Then Command4_Click
CmdCounter = 1
End Sub
Private Sub Command5_Click()
On Error GoTo ErrHandler
If Form1.RichTextBox3.Visible = True Then
 Image1.Visible = False
 Command1.Caption = "Close HTML Page"
 With dlgCommonDialog
            .DialogTitle = "Select a color"
            .Flags = cdlCCPreventFullOpen
            .ShowColor
End With
With RichTextBox3
           bb = .SelStart
           ss = .SelLength
     .SelStart = .SelStart
     .SelText = "<FONT COLOR="
     .SelText = """"
  Call SelColor
     .SelText = """"
     .SelText = ">"
       .SelStart = Val(.SelStart) + Val(ss)
      .SelText = "</Font>"
      .Visible = True
 Image1.Visible = False
 Command1.Caption = "Close HTML Page"
     Command7_Click
      .SelStart = bb + 21
      .SelLength = ss
End With
End If
ErrHandler:
 Exit Sub
End Sub


Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "Text Color"
If CmdCounter = 0 Then
 If Rang = 1 Then Call ChangeBkColor
 Shape9.BorderColor = 8454143
 Rang = 1
End If
If Button = 1 Then Command5_Click
CmdCounter = 1
End Sub

Private Sub Command6_Click()
If Form1.RichTextBox3.Visible = True Then mnuViewWebBrowser_Click
Sleep 500
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "Test The HTML Page In M2A HTML Viewer"
If CmdCounter = 0 Then
 If Rang = 1 Then Call ChangeBkColor
 Shape3.BorderColor = 8454143
 Rang = 1
End If
If Button = 1 Then Command6_Click
CmdCounter = 1
End Sub

Private Sub Command7_Click()
 If RichTextBox3.Visible = True Then
  Image1.Visible = False
  RichTextBox3.Visible = False
  RichTextBox3.Visible = True
 End If
 If RichTextBox3.Visible = False Then
  RichTextBox3.Visible = False
  Image1.Visible = False
  Image1.Visible = True
 End If
If Form1.RichTextBox3.Visible = True Then Form1.RichTextBox3.SetFocus
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "Refresh"
If CmdCounter = 0 Then
 If Rang = 1 Then Call ChangeBkColor
 Shape5.BorderColor = 8454143
 Rang = 1
End If
If Button = 1 Then Command7_Click
CmdCounter = 1
End Sub
Private Sub Command9_Click()
On Error GoTo er
Dim MyA, MyA1, MyB, MyB1, Adad
Dim Var1 As Byte
If Form1.RichTextBox3.Visible = True Then
If Txt2 = 0 Then
 Var1 = 0
Else
 Var1 = 1
End If
With RichTextBox3
Me.MousePointer = vbHourglass
Adad = 1
Do While Len(.Text) > Adad
 .SelStart = Adad
MyA = InStr(.SelStart, .Text, "<", vbTextCompare)
 .SelStart = MyA
If .SelStart = 0 Then Exit Do
MyB = InStr(.SelStart, .Text, ">", vbTextCompare)
MyB = MyB - 1
 .SelLength = MyB - MyA
 .SelColor = 16711680
Adad = MyB
Loop

Adad = 1
Do While Len(.Text) > Adad
 .SelStart = Adad
MyA = InStr(.SelStart, .Text, "<A", vbTextCompare)
 .SelStart = MyA
If .SelStart = 0 Then Exit Do
MyB = InStr(.SelStart, .Text, ">", vbTextCompare)
MyB = MyB - 1
 .SelLength = MyB - MyA
 .SelColor = 8421631
Adad = MyB
Loop

Adad = 1
Do While Len(.Text) > Adad
 .SelStart = Adad
MyA = InStr(.SelStart, .Text, "<A title", vbTextCompare)
 .SelStart = MyA
If .SelStart = 0 Then Exit Do
MyB = InStr(.SelStart, .Text, ">", vbTextCompare)
MyB = MyB - 1
 .SelLength = MyB - MyA
 .SelColor = 8421631
Adad = MyB
Loop

Adad = 1
Do While Len(.Text) > Adad
 .SelStart = Adad
MyA = InStr(.SelStart, .Text, "<A href", vbTextCompare)
 .SelStart = MyA
If .SelStart = 0 Then Exit Do
MyB = InStr(.SelStart, .Text, ">", vbTextCompare)
MyB = MyB - 1
 .SelLength = MyB - MyA
 .SelColor = 8421631
Adad = MyB
Loop

Adad = 1
Do While Len(.Text) > Adad
 .SelStart = Adad
MyA = InStr(.SelStart, .Text, "<embed", vbTextCompare)
 .SelStart = MyA
If .SelStart = 0 Then Exit Do
MyB = InStr(.SelStart, .Text, ">", vbTextCompare)
MyB = MyB - 1
 .SelLength = MyB - MyA
 .SelColor = 16711935
Adad = MyB
Loop

Adad = 1
Do While Len(.Text) > Adad
 .SelStart = Adad
MyA = InStr(.SelStart, .Text, "</embed", vbTextCompare)
 .SelStart = MyA
If .SelStart = 0 Then Exit Do
MyB = InStr(.SelStart, .Text, ">", vbTextCompare)
MyB = MyB - 1
 .SelLength = MyB - MyA
 .SelColor = 16711935
Adad = MyB
Loop

Adad = 1
Do While Len(.Text) > Adad
 .SelStart = Adad
MyA = InStr(.SelStart, .Text, "<A on", vbTextCompare)
 .SelStart = MyA
If .SelStart = 0 Then Exit Do
MyB = InStr(.SelStart, .Text, ">", vbTextCompare)
MyB = MyB - 1
 .SelLength = MyB - MyA
 .SelColor = 8421631
Adad = MyB
Loop

Adad = 1
Do While Len(.Text) > Adad
 .SelStart = Adad
MyA = InStr(.SelStart, .Text, "</A", vbTextCompare)
 .SelStart = MyA
If .SelStart = 0 Then Exit Do
MyB = InStr(.SelStart, .Text, ">", vbTextCompare)
MyB = MyB - 1
 .SelLength = MyB - MyA
 .SelColor = 8421631
Adad = MyB
Loop

Adad = 1
Do While Len(.Text) > Adad
 .SelStart = Adad
MyA = InStr(.SelStart, .Text, "<img", vbTextCompare)
 .SelStart = MyA
If .SelStart = 0 Then Exit Do
MyB = InStr(.SelStart, .Text, ">", vbTextCompare)
MyB = MyB - 1
 .SelLength = MyB - MyA
 .SelColor = 32768
Adad = MyB
Loop

Adad = 1
Do While Len(.Text) > Adad
 .SelStart = Adad
MyA = InStr(.SelStart, .Text, "<script", vbTextCompare)
 .SelStart = MyA - 1
If .SelStart = 0 Then Exit Do
MyB = InStr(.SelStart, .Text, "</script>", vbTextCompare)
MyB = MyB + 9
 .SelLength = MyB - MyA
 .SelColor = 255
Adad = MyB
Loop
er:

 .SelStart = 0
 .SelLength = Len(.Text)
 .SelBold = Not .SelBold
Me.MousePointer = vbArrow
Command7_Click
 .SelStart = 1
End With
If Var1 = 0 Then Txt2 = 0
End If
End Sub

Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "Tag High Lighting"
If CmdCounter = 0 Then
 If Rang = 1 Then Call ChangeBkColor
 Shape1.BorderColor = 8454143
 Rang = 1
End If
If Button = 1 Then Command9_Click
CmdCounter = 1
End Sub

Private Sub Dir1_Change()
On Error Resume Next
File1.Path = Dir1.Path
ChDir Dir1.Path
NameDir = CurDir
Drive1.Drive = Left(Dir1.Path, 2)
End Sub
Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "See The Directory And Choose It"
If Rang = 1 Then Call ChangeBkColor
Rang = 0
End Sub

Private Sub Dir1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then Me.PopupMenu mnupopup3
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
ChDrive Drive1.Drive
End Sub
Private Sub Drive2_Change()
On Error Resume Next
Dir2.Path = Drive2.Drive
End Sub

Private Sub File1_Click()
On Error Resume Next
NameF = File1.FileName
End Sub

Private Sub File1_dblClick()
On Error Resume Next
NameF = File1.FileName
If RichTextBox3.Visible = True Then
If Txt2 = 1 Then
  Beep
  myq = MsgBox("Do you want to save this page ?", vbYesNoCancel + vbQuestion, "Save...")
End If
 If myq = 6 Then
  mnuFileSave_Click
  mnuFileClose.Enabled = False
  RichTextBox3.FileName = Empty
  RichTextBox3.Visible = False
  Image1.Visible = True
  Txt1 = "0"
 End If
 If myq = 7 Then
  mnuFileClose.Enabled = False
  RichTextBox3.Text = ""
  RichTextBox3.Visible = False
  Image1.Visible = True
  Txt1 = "0"
 End If
 If myq = 2 Then Exit Sub
For_File_Name = 1
Image1.Visible = False
RichTextBox3.Visible = True
RichTextBox3.FileName = Dir1.Path + "\" + NameF
If Len(Dir1.Path) = 3 Then RichTextBox3.FileName = Dir1.Path + NameF
 
Call Recent

Command7_Click
Command1.Caption = "Close HTML Page"
NameSave = File1.FileName
TypeOfSave = 1
CmdCounter4 = 1
Timer1.Enabled = True
Else
For_File_Name = 1
 Image1.Visible = False
 RichTextBox3.Visible = True
 RichTextBox3.Text = ""
 RichTextBox3.FileName = Dir1.Path + "\" + File1.FileName
 If Len(Dir1.Path) = 3 Then RichTextBox3.FileName = Dir1.Path + File1.FileName
 
 Call Recent
 
 Command7_Click
 Command1.Caption = "Close HTML Page"
 NameSave = File1.FileName
 TypeOfSave = 1
 Timer1.Enabled = True
 CmdCounter4 = 1
End If
If Form1.RichTextBox3.Visible = True Then Form1.RichTextBox3.SetFocus
Txt1 = "0"
CmdCounter4 = 1
End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "See The File And Choose It"
If Rang = 1 Then Call ChangeBkColor
Rang = 0
End Sub

Private Sub File1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If File1.ListIndex <> -1 Then
 If Button = vbRightButton Then Me.PopupMenu mnupopup2
End If
End Sub

Private Sub Form_Activate()
 Call ChangeBkColor
 Command7_Click
End Sub

Private Sub Form_Load()
On Error Resume Next

lLeft = GetSetting(App.Title, "Settings", "MainLeft")
If lLeft = Empty Then
 Me.Left = (Screen.Width - Form1.Width) / 2
 Me.Top = (Screen.Height - Form1.Height) / 2
Else
 Me.Left = GetSetting(App.Title, "Settings", "MainLeft")
 Me.Top = GetSetting(App.Title, "Settings", "MainTop")
End If

Wi = GetSetting(App.Title, "Settings", "MainWidth")
He = GetSetting(App.Title, "Settings", "MainHeight")
Timer2.Enabled = True

s = GetSetting(App.Title, "Settings", "File1")
If s = Empty Then SaveSetting App.Title, "Settings", "File1", dlgCommonDialog.FileName

mnuFileRecentFiles1.Caption = GetSetting(App.Title, "Settings", "RecentFile1")
mnuFileRecentFiles2.Caption = GetSetting(App.Title, "Settings", "RecentFile2")
mnuFileRecentFiles3.Caption = GetSetting(App.Title, "Settings", "RecentFile3")
mnuFileRecentFiles4.Caption = GetSetting(App.Title, "Settings", "RecentFile4")
mnuFileRecentFiles5.Caption = GetSetting(App.Title, "Settings", "RecentFile5")
mnuFileRecentFiles6.Caption = GetSetting(App.Title, "Settings", "RecentFile6")
mnuFileRecentFiles7.Caption = GetSetting(App.Title, "Settings", "RecentFile7")

Form11.Check1.Value = GetSetting(App.Title, "Settings", "StatuseBar")
Form11.Check2.Value = GetSetting(App.Title, "Settings", "ToolBar")

RichTextBox3.Font.Name = GetSetting(App.Title, "Settings", "FontName")
RichTextBox3.Font.Size = GetSetting(App.Title, "Settings", "FontSize")
RichTextBox3.Font.Bold = GetSetting(App.Title, "Settings", "FontBold")
RichTextBox3.Font.Italic = GetSetting(App.Title, "Settings", "FontItalic")
RichTextBox3.Font.Underline = GetSetting(App.Title, "Settings", "FontUnderline")
RichTextBox3.Font.Strikethrough = GetSetting(App.Title, "Settings", "FontStrikethru")

dlgCommonDialog.FontName = GetSetting(App.Title, "Settings", "FontName")
dlgCommonDialog.FontBold = GetSetting(App.Title, "Settings", "FontBold")
dlgCommonDialog.FontItalic = GetSetting(App.Title, "Settings", "FontItalic")
dlgCommonDialog.FontSize = GetSetting(App.Title, "Settings", "FontSize")
dlgCommonDialog.FontUnderline = GetSetting(App.Title, "Settings", "FontUnderline")
dlgCommonDialog.FontStrikethru = GetSetting(App.Title, "Settings", "FontStrikethru")

ForInternetBrowser = GetSetting(App.Title, "Settings", "ForInternetBrowser")
If ForInternetBrowser = 0 Then YourFavorite = 0
If ForInternetBrowser = 1 Then YourFavorite = 1

ForHidden = GetSetting(App.Title, "Settings", "ForHidden")
If ForHidden = "No" Then File1.Hidden = False
If ForHidden = "Yes" Then File1.Hidden = True

Me.BackColor = GetSetting(App.Title, "Settings", "Color")

Adres = GetSetting(App.Title, "Settings", "ForStartAddress")

If Adres = Empty Then Adres = "C:\"
Form11.Text1.Text = Adres
Unload Form11
Text12.Text = Adres
Text12.SelStart = 0
Text12.SelLength = 2
Text12.Text = Text12.SelText
Drive1.Drive = Text12.Text
Dir1.Path = Adres
File1.Path = Dir1.Path
File1.Pattern = "*.HTML;*.HTM"
If Clipboard.GetText = "" Then
 mnuEditPaste.Enabled = False
 mnupopupPaste.Enabled = False
Else
 mnuEditPaste.Enabled = True
 mnupopupPaste.Enabled = True
End If

 Call MenuAddBitmap(Form1.hWnd, 0, 0, Form14.New(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 0, 1, Form14.Open(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 0, 2, Form14.Close(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 0, 4, Form14.Save(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 0, 5, Form14.SaveAs(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 0, 7, Form14.Print(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 0, 9, Form14.New(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 0, 11, Form14.Close(1).Picture)
 
 Call MenuAddBitmap(Form1.hWnd, 1, 2, Form14.Cut(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 1, 3, Form14.Copy(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 1, 4, Form14.Paste(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 1, 5, Form14.Delete(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 1, 7, Form14.Find(0).Picture)
 
 Call MenuAddBitmap(Form1.hWnd, 3, 0, Form14.Link(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 1, Form14.Image1(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 2, Form14.Elink(1).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 3, Form14.Sound(1).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 4, Form14.Flash.Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 5, Form14.Table1(1).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 6, Form14.Ruller(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 9, Form14.Font1(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 10, Form14.Head(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 12, Form14.Left1(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 13, Form14.Center1(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 14, Form14.Right1(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 16, Form14.Para(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 17, Form14.Bre(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 19, Form14.Bold(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 20, Form14.Italic(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 21, Form14.Underline(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 23, Form14.Chk(9).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 24, Form14.Radio(10).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 25, Form14.ImageM(11).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 27, Form14.TextBox(1).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 28, Form14.Pass(2).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 29, Form14.HiddenTxt(3).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 30, Form14.Filebrowser(4).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 31, Form14.Textarea(5).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 33, Form14.Button(6).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 34, Form14.Submit(7).Picture)
 Call MenuAddBitmap(Form1.hWnd, 3, 35, Form14.Reset(8).Picture)
 
 Call MenuAddBitmap(Form1.hWnd, 6, 0, Form14.Help(0).Picture)
 Call MenuAddBitmap(Form1.hWnd, 6, 3, Form14.About(0).Picture)
 
 Timer1.Enabled = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbArrow
sbStatusBar.SimpleText = "M2A HTML Maker  -  http://www.IranM2A.Tk"
If Rang = 1 Then Call ChangeBkColor
Rang = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If RichTextBox3.Visible = True Then
 If Txt2 = 1 Then
  Beep
  myq = MsgBox("Do you want to save this page ?", vbYesNoCancel + vbQuestion, "Save...")
 End If
 If myq = 6 Then
  mnuFileSave_Click
  mnuFileClose.Enabled = False
  RichTextBox3.Text = ""
  RichTextBox3.Visible = False
  Image1.Visible = True
  Command1.Caption = "New HTML Page"
  Txt1 = "0"
 End If
 If myq = 7 Then
  mnuFileClose.Enabled = False
  RichTextBox3.Text = ""
  RichTextBox3.Visible = False
  Image1.Visible = True
  Command1.Caption = "New HTML Page"
  Txt1 = "0"
  
 Call Recent
 
 End If
 If myq = 2 Then
  Cancel = True
  Exit Sub
 End If
End If
If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
End If
        SaveSetting App.Title, "Settings", "RecentFile1", mnuFileRecentFiles1.Caption
        SaveSetting App.Title, "Settings", "RecentFile2", mnuFileRecentFiles2.Caption
        SaveSetting App.Title, "Settings", "RecentFile3", mnuFileRecentFiles3.Caption
        SaveSetting App.Title, "Settings", "RecentFile4", mnuFileRecentFiles4.Caption
        SaveSetting App.Title, "Settings", "RecentFile5", mnuFileRecentFiles5.Caption
        SaveSetting App.Title, "Settings", "RecentFile6", mnuFileRecentFiles6.Caption
        SaveSetting App.Title, "Settings", "RecentFile7", mnuFileRecentFiles7.Caption
 For intctr = (Forms.Count - 1) To 0 Step -1
  Unload Forms(intctr)
 Next intctr
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "M2A HTML Maker  -  http://www.IranM2A.Tk"
If Rang = 1 Then Call ChangeBkColor
Rang = 0
End Sub

Private Sub mnuEditDelete_Click()
If Form1.RichTextBox3.Visible = True Then RichTextBox3.SelText = " "
End Sub
 
Private Sub mnuEditFind_Click()
If Form1.RichTextBox3.Visible = True Then Form3.Show
End Sub

Private Sub mnuEditRefresh_Click()
Command7_Click
End Sub

Private Sub mnuEditSelAll_Click()
If Form1.RichTextBox3.Visible = True Then
 Mystr = Len(RichTextBox3.Text)
 RichTextBox3.SelStart = 0
 RichTextBox3.SelLength = Mystr
End If
End Sub

Private Sub mnuFF1_Click()
Call ForFontSize("mnuFF1")
End Sub

Private Sub mnuFF2_Click()
Call ForFontSize("mnuFF2")
End Sub

Private Sub mnuFF3_Click()
Call ForFontSize("mnuFF3")
End Sub

Private Sub mnuFF4_Click()
Call ForFontSize("mnuFF4")
End Sub

Private Sub mnuFF5_Click()
Call ForFontSize("mnuFF5")
End Sub

Private Sub mnuFF6_Click()
Call ForFontSize("mnuFF6")
End Sub

Private Sub mnuFF7_Click()
Call ForFontSize("mnuFF7")
End Sub

Private Sub mnuFileClose_Click()
Command1_Click
End Sub

Private Sub mnuFileRecentFiles1_Click()
Call ForRecents("mnuFileRecentFiles1")
End Sub

Private Sub mnuFileRecentFiles2_Click()
Call ForRecents("mnuFileRecentFiles2")
End Sub

Private Sub mnuFileRecentFiles3_Click()
Call ForRecents("mnuFileRecentFiles3")
End Sub

Private Sub mnuFileRecentFiles4_Click()
Call ForRecents("mnuFileRecentFiles4")
End Sub

Private Sub mnuFileRecentFiles5_Click()
Call ForRecents("mnuFileRecentFiles5")
End Sub

Private Sub mnuFileRecentFiles6_Click()
Call ForRecents("mnuFileRecentFiles6")
End Sub

Private Sub mnuFileRecentFiles7_Click()
Call ForRecents("mnuFileRecentFiles7")
End Sub

Private Sub mnuFileSaveas_Click()
On Error Resume Next
If Form1.RichTextBox3.Visible = True Then
    Dim sFile As String
        With dlgCommonDialog
            .DialogTitle = "Save As..."
            .FileName = Dir1.Path + "\" + "Untitled"
            If Len(Dir1.Path) = 3 Then .FileName = "Untitled"
            .Filter = "HTML File (*.html)|*.html|HTM File (*.htm)|*.htm"
            .ShowSave
            If Err.Number = 32755 Then GoTo ErrHandler
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
            RichTextBox3.FileName = .FileName
        End With
        If FileLen(sFile) = Empty Then
         F = FreeFile
         Open sFile For Output As #F
         Print #F, RichTextBox3.Text
         Close #F
       End If
End If
Txt2 = 0
If RichTextBox3.FileName = Empty Then RichTextBox3.FileName = sFile
TypeOfSave = 1
Call Recent
Call Refresh_The_Dir
ErrHandler:
 Exit Sub
End Sub

Private Sub mnuFontC_Click()
On Error Resume Next
Command5_Click
End Sub

Private Sub mnuHD1_Click()
Call ForHeadingFont("mnuHD1")
End Sub

Private Sub mnuHD2_Click()
Call ForHeadingFont("mnuHD2")
End Sub

Private Sub mnuHD3_Click()
Call ForHeadingFont("mnuHD3")
End Sub

Private Sub mnuHD4_Click()
Call ForHeadingFont("mnuHD4")
End Sub

Private Sub mnuHD5_Click()
Call ForHeadingFont("mnuHD5")
End Sub

Private Sub mnuHD6_Click()
Call ForHeadingFont("mnuHD6")
End Sub

Private Sub mnuHelpA_Click()
On Error Resume Next
Form9.Timer1.Enabled = False
Form9.MousePointer = 1
Form9.Command1.Visible = True
Form9.Show vbModal
End Sub

Private Sub mnuHelpH_Click()
On Error Resume Next
MsgBox "For see this help ,you must Download M2A HTML Maker's Help From http://www.IranM2A.Tk"
End Sub

Private Sub mnuEdit_Click()
On Error Resume Next
If Clipboard.GetText = "" Then
 mnuEditPaste.Enabled = False
Else
 mnuEditPaste.Enabled = True
End If
If RichTextBox3.Visible = False Then
  mnuEditRefresh.Enabled = False
  mnuEditCut.Enabled = False
  mnuEditCopy.Enabled = False
  mnuEditPaste.Enabled = False
  mnuEditDelete.Enabled = False
  mnuEditSelAll.Enabled = False
  mnuEditFind.Enabled = False
 Else
  mnuEditRefresh.Enabled = True
  mnuEditCut.Enabled = True
  mnuEditCopy.Enabled = True
  mnuEditPaste.Enabled = True
  mnuEditDelete.Enabled = True
  mnuEditSelAll.Enabled = True
  mnuEditFind.Enabled = True
End If
End Sub

Private Sub mnuHelpwww_Click()
On Error Resume Next
Shell "Explorer http://www.iranm2a.tk"
End Sub

Private Sub mnuOptionsDefaultFont_Click()
On Error Resume Next
dlgCommonDialog.FontSize = 10
dlgCommonDialog.FontBold = False
dlgCommonDialog.FontItalic = False
dlgCommonDialog.FontUnderline = False
dlgCommonDialog.FontStrikethru = False
dlgCommonDialog.FontName = "MS Sans Serif"
RichTextBox3.Font.Name = "MS Sans Serif"
Call ForFont
mnuOptionsSave_Click
Command7_Click
End Sub

Private Sub mnuOptionsFont_Click()
On Error GoTo ErrHandler
dlgCommonDialog.ShowFont
Call ForFont
Command7_Click
ErrHandler:
 Exit Sub
End Sub

Private Sub mnuOptionsSave_Click()
On Error Resume Next
       SaveSetting App.Title, "Settings", "FontSize", RichTextBox3.Font.FontSize
       SaveSetting App.Title, "Settings", "FontName", RichTextBox3.Font.FontName
       SaveSetting App.Title, "Settings", "FontBold", RichTextBox3.Font.FontBold
       SaveSetting App.Title, "Settings", "FontStrikethru", RichTextBox3.Font.FontStrikethru
       SaveSetting App.Title, "Settings", "FontItalic", RichTextBox3.Font.FontItalic
       SaveSetting App.Title, "Settings", "FontUnderline", RichTextBox3.Font.FontUnderline
       SaveSetting App.Title, "Settings", "Color", Form11.Picture1.BackColor
       Form11.Hide
End Sub

Private Sub mnupopup2Del_Click()
On Error Resume Next
myq = MsgBox("Are you sure you want to remove this file", vbYesNo, "Delete...")
If myq = 6 Then
SetAttr NameF, vbNormal
Kill NameF
Refresh_The_Dir
File1.Visible = False
File1.Visible = True
End If
End Sub
Private Sub mnuOptionsOp_Click()
Form11.Drive1.Drive = Drive1.Drive
Form11.Dir1.Path = Dir1.Path
Form11.Show
End Sub

Private Sub mnupopup3Mkdir_Click()
On Error Resume Next
Myinput = InputBox("Enter directory name for create it ", "Directory name")
If Len(NameDir) > 3 Then
MkDir NameDir + "\" + Myinput
Else
MkDir NameDir + Myinput
End If
Refresh_The_Dir
Dir1.Visible = False
Dir1.Visible = True
End Sub

Private Sub mnupopupSelAll_Click()
mnuEditSelAll_Click
End Sub
Private Sub mnupopupCopy_Click()
mnuEditCopy_Click
End Sub

Private Sub mnupopupCut_Click()
mnuEditCut_Click
End Sub

Private Sub mnupopupDelete_Click()
RichTextBox3.SelText = " "
End Sub

Private Sub mnupopupPaste_Click()
mnuEditPaste_Click
End Sub

Private Sub mnupopupR_Click()
Command7_Click
End Sub

Private Sub mnuTagh_Click()
Command9_Click
End Sub

Private Sub mnuTags_Click()
If RichTextBox3.Visible = False Then
 mnuTagsInsertPicture.Enabled = False
 mnuTagsInsertLink.Enabled = False
 mnuTagsSound.Enabled = False
 mnuTagsTable.Enabled = False
 mnuTagsLine.Enabled = False
 mnuTagsEmail.Enabled = False
 mnuTagsFonts.Enabled = False
 mnuTagsHSize.Enabled = False
 mnuTagsLeft.Enabled = False
 mnuTagsCenter.Enabled = False
 mnuTagsRight.Enabled = False
 mnuTagsBreak.Enabled = False
 mnuTagsParagraph.Enabled = False
 mnuTagsBold.Enabled = False
 mnuTagsItalic.Enabled = False
 mnuTagsUnder.Enabled = False
 mnuTagsChkBox.Enabled = False
 mnuTagsRadioButton.Enabled = False
 mnuTagsImage.Enabled = False
 mnuTagsTxtBox.Enabled = False
 mnuTagsPassTxtBox.Enabled = False
 mnuTagsHiddenTextBox.Enabled = False
 mnuTagsFileBrowserTxtBox.Enabled = False
 mnuTagsTxtArea.Enabled = False
 mnuTagsButton.Enabled = False
 mnuTagsSubmitButton.Enabled = False
 mnuTagsResetButton.Enabled = False
 mnuTagsDateAndTime.Enabled = False
 mnuTagsFlash.Enabled = False
Else
 mnuTagsInsertPicture.Enabled = True
 mnuTagsInsertLink.Enabled = True
 mnuTagsSound.Enabled = True
 mnuTagsTable.Enabled = True
 mnuTagsLine.Enabled = True
 mnuTagsEmail.Enabled = True
 mnuTagsFonts.Enabled = True
 mnuTagsHSize.Enabled = True
 mnuTagsLeft.Enabled = True
 mnuTagsCenter.Enabled = True
 mnuTagsRight.Enabled = True
 mnuTagsBreak.Enabled = True
 mnuTagsParagraph.Enabled = True
 mnuTagsBold.Enabled = True
 mnuTagsItalic.Enabled = True
 mnuTagsUnder.Enabled = True
 mnuTagsChkBox.Enabled = True
 mnuTagsRadioButton.Enabled = True
 mnuTagsImage.Enabled = True
 mnuTagsTxtBox.Enabled = True
 mnuTagsPassTxtBox.Enabled = True
 mnuTagsHiddenTextBox.Enabled = True
 mnuTagsFileBrowserTxtBox.Enabled = True
 mnuTagsTxtArea.Enabled = True
 mnuTagsButton.Enabled = True
 mnuTagsSubmitButton.Enabled = True
 mnuTagsResetButton.Enabled = True
 mnuTagsDateAndTime.Enabled = True
 mnuTagsFlash.Enabled = True
End If
End Sub

Private Sub mnuTagsBold_Click()
With RichTextBox3
ss = .SelLength
      .SelStart = .SelStart
     .SelText = "<B>"
       .SelStart = Val(.SelStart) + Val(ss)
      .SelText = "</B>"
     Command7_Click
End With
End Sub

Private Sub mnuTagsBreak_Click()
Command3_Click
End Sub

Private Sub mnuTagsButton_Click()
RichTextBox3.SelText = "<input type=""button"" value="""">"
Command7_Click
End Sub

Private Sub mnuTagsCenter_Click()
Call ForCenter
End Sub

Private Sub mnuTagsChkBox_Click()
RichTextBox3.SelText = "<input type=""checkbox"" name="""" value="""">"
Command7_Click
End Sub

Private Sub mnuTagsDateAndTime_Click()
If RichTextBox3.Visible = True Then Form15.Show
End Sub

Private Sub mnuTagsEmail_Click()
Command11_Click
End Sub

Private Sub mnuTagsFileBrowserTxtBox_Click()
RichTextBox3.SelText = "<input type=""file"" name="""" value="""">"
Command7_Click
End Sub

Private Sub mnuTagsFlash_Click()
Form16.Show
End Sub

Private Sub mnuTagsHiddenTextBox_Click()
RichTextBox3.SelText = "<input type=""hidden"" name="""" value="""">"
Command7_Click
End Sub

Private Sub mnuTagsImage_Click()
RichTextBox3.SelText = "<input type=""image"" name="""" value="""">"
Command7_Click
End Sub

Private Sub mnuTagsInsertLink_Click()
Command13_Click
End Sub

Private Sub mnuTagsInsertPicture_Click()
Command14_Click
End Sub

Private Sub mnuTagsItalic_Click()
With RichTextBox3
ss = .SelLength
      .SelStart = .SelStart
     .SelText = "<I>"
       .SelStart = Val(.SelStart) + Val(ss)
      .SelText = "</I>"
     Command7_Click
End With
End Sub

Private Sub mnuTagsLeft_Click()
Call ForLeft
End Sub

Private Sub mnuTagsLine_Click()
If RichTextBox3.Visible = True Then Form13.Show
End Sub

Private Sub mnuTagsParagraph_Click()
Command2_Click
End Sub

Private Sub mnuTagsPassTxtBox_Click()
RichTextBox3.SelText = "<input type=""password"" name="""" value="""">"
Command7_Click
End Sub

Private Sub mnuTagsRadioButton_Click()
RichTextBox3.SelText = "<input type=""radio"" name="""" value="""">"
Command7_Click
End Sub

Private Sub mnuTagsResetButton_Click()
RichTextBox3.SelText = "<input type=""reset"">"
Command7_Click
End Sub

Private Sub mnuTagsRight_Click()
Call ForRight
End Sub

Private Sub mnuTagsSound_Click()
If RichTextBox3.Visible = True Then Form8.Show
End Sub

Private Sub mnuTagsSubmitButton_Click()
RichTextBox3.SelText = "<input type=""submit"">"
Command7_Click
End Sub

Private Sub mnuTagsTable_Click()
If RichTextBox3.Visible = True Then Form10.Show
End Sub

Private Sub mnuTagsTxtArea_Click()
RichTextBox3.SelText = "<textarea name="""" rows="""" cols=""""></textarea>"
Command7_Click
End Sub

Private Sub mnuTagsTxtBox_Click()
RichTextBox3.SelText = "<input type=""text"" name="""" value="""">"
Command7_Click
End Sub

Private Sub mnuTagsUnder_Click()
With RichTextBox3
ss = .SelLength
      .SelStart = .SelStart
     .SelText = "<U>"
       .SelStart = Val(.SelStart) + Val(ss)
      .SelText = "</U>"
     Command7_Click
End With
End Sub

Private Sub mnuTest_Click()
Command6_Click
End Sub

Private Sub mnuTools_Click()
If RichTextBox3.Visible = False Then
 mnuTest.Enabled = False
 mnuTagh.Enabled = False
Else
 mnuTest.Enabled = True
 mnuTagh.Enabled = True
End If
End Sub

Private Sub mnuView_Click()
If RichTextBox3.Visible = False Then
 mnuViewWebBrowser.Enabled = False
Else
 mnuViewWebBrowser.Enabled = True
End If
End Sub

Private Sub mnuViewMap_Click()
On Error Resume Next
Form12.Show
End Sub

Private Sub mnuViewWebBrowser_Click()
On Error Resume Next
If RichTextBox3.Visible = True Then
    Dim MyPath
     MyPath = Dir1.Path
     Text2.Text = MyPath
     Text2.SelStart = Len(Text2.Text)
    If Txt2 = 1 Then
      If Len(Text2.Text) > 3 Then Text2.SelText = "\html1.html"
      If Len(Text2.Text) = 3 Then Text2.SelText = "html1.html"
    ElseIf Txt2 = 0 And RichTextBox3.FileName <> Empty Then
      Text2.Text = RichTextBox3.FileName
      GoTo Start1
    ElseIf Txt2 = 0 And RichTextBox3.FileName = Empty Then
      If Len(Text2.Text) > 3 Then Text2.SelText = "\html1.html"
      If Len(Text2.Text) = 3 Then Text2.SelText = "html1.html"
    End If
     F = FreeFile
     SetAttr Text2.Text, vbNormal
     Open Text2.Text For Output As #F
     Print #F, RichTextBox3.Text
     Close #F
Start1:
If Form11.Option2.Value = True Then
 YourFavorite = 1
ElseIf Form11.Option1.Value = True Then
 YourFavorite = 0
End If
    If YourFavorite = 1 Then Shell "Explorer " & Text2.Text 'See Module1
    If YourFavorite = 0 Then
     If Txt3 <> 1 Then '
      Form2.cboAddress.Text = Text2.Text
      Form2.Show
     Else
      Form2.cboAddress.Text = Text2.Text
      Form2.Timer1.Enabled = True
      Form2.Show
     End If
    End If
    If Txt2 = 0 And RichTextBox3.FileName <> Empty Then GoTo Start2
     SetAttr Text2.Text, vbHidden
Start2:
    Txt3 = 1
End If
Unload Form11
End Sub
Private Sub RichTextBox3_Change()
On Error Resume Next
Txt2 = 1
End Sub

Private Sub RichTextBox3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbIbeam
If RichTextBox3.FileName = Empty Then
 sbStatusBar.SimpleText = "HTML Code"
ElseIf For_File_Name = 1 Then
 sbStatusBar.SimpleText = RichTextBox3.FileName
ElseIf For_File_Name = 0 Then
 sbStatusBar.SimpleText = "HTML Code"
End If
If Rang = 1 Then Call ChangeBkColor
Rang = 0
End Sub

Private Sub RichTextBox3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then Me.PopupMenu mnupopup
If Clipboard.GetText = "" Then
 mnupopupPaste.Enabled = False
Else
 mnupopupPaste.Enabled = True
End If
End Sub

Private Sub sbStatusBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "M2A HTML Maker  -  http://www.IranM2A.Tk"
If Rang = 1 Then Call ChangeBkColor
Rang = 0
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbStatusBar.SimpleText = "You Can See The Browser And Use It And You Can Insert Java Script In Your HTML File"
If Rang = 1 Then Call ChangeBkColor
Rang = 0
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    Select Case Button.Key
        Case "New"
        If Command1.Caption = "New HTML Page" Then
         If CmdCounter3 = 0 Then
          Call Command1_Click
          Call Command1_Click
          CmdCounter3 = 1
          CmdCounter4 = 1
         Else
          Call Command1_Click
         End If
        End If
        Case "Open"
            mnuFileOpen_Click
           If RichTextBox3.Visible = True Then RichTextBox3.SetFocus
        Case "Save"
           If RichTextBox3.Visible = True Then
            mnuFileSave_Click
            RichTextBox3.SetFocus
           End If
        Case "Print"
           If RichTextBox3.Visible = True Then
            mnuFilePrint_Click
             RichTextBox3.SetFocus
           End If
        Case "Cut"
           If RichTextBox3.Visible = True Then
            mnuEditCut_Click
             RichTextBox3.SetFocus
           End If
        Case "Copy"
           If RichTextBox3.Visible = True Then
            mnuEditCopy_Click
             RichTextBox3.SetFocus
           End If
        Case "Paste"
            If RichTextBox3.Visible = True Then
            If Clipboard.GetText = "" Then Button.Enabled = False
            mnuEditPaste_Click
             RichTextBox3.SetFocus
            End If
        Case "Bold"
        If RichTextBox3.Visible = True Then
            bb = RichTextBox3.SelStart
            ss = RichTextBox3.SelLength
      RichTextBox3.SelStart = RichTextBox3.SelStart
     RichTextBox3.SelText = "<B>"
       RichTextBox3.SelStart = Val(RichTextBox3.SelStart) + Val(ss)
      RichTextBox3.SelText = "</B>"
     Command7_Click
      RichTextBox3.SelStart = bb + 3
      RichTextBox3.SelLength = ss
     End If
        Case "Italic"
          If RichTextBox3.Visible = True Then
          bb = RichTextBox3.SelStart
          ss = RichTextBox3.SelLength
      RichTextBox3.SelStart = RichTextBox3.SelStart
     RichTextBox3.SelText = "<I>"
       RichTextBox3.SelStart = Val(RichTextBox3.SelStart) + Val(ss)
      RichTextBox3.SelText = "</I>"
     Command7_Click
      RichTextBox3.SelStart = bb + 3
      RichTextBox3.SelLength = ss
     End If
        Case "Underline"
          If RichTextBox3.Visible = True Then
          bb = RichTextBox3.SelStart
          ss = RichTextBox3.SelLength
      RichTextBox3.SelStart = RichTextBox3.SelStart
     RichTextBox3.SelText = "<U>"
       RichTextBox3.SelStart = Val(RichTextBox3.SelStart) + Val(ss)
      RichTextBox3.SelText = "</U>"
     Command7_Click
      RichTextBox3.SelStart = bb + 3
      RichTextBox3.SelLength = ss
     End If
       Case "Align Left"
          Call ForLeft
       Case "Center"
         Call ForCenter
       Case "Align Right"
        Call ForRight
       Case "Find"
       If RichTextBox3.Visible = True Then Form3.Show
    End Select
End Sub

Private Sub mnuWindowNewWindow_Click()
    LoadNewDoc
End Sub
Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
    Command7_Click
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
    Command7_Click
End Sub



Private Sub mnuEditPaste_Click()
    If Clipboard.GetText = "" Then Me.Enabled = False
    On Error Resume Next
    RichTextBox3.SelRTF = Clipboard.GetText
Command7_Click
End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText RichTextBox3.SelRTF
End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText RichTextBox3.SelRTF
    RichTextBox3.SelText = vbNullString
Command7_Click
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnuFilePrint_Click()
On Error GoTo ErrHandler
 If Form1.RichTextBox3.Visible = True Then
    With dlgCommonDialog
        .DialogTitle = "Print"
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        .Flags = .Flags + cdlPDAllPages
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            RichTextBox3.SelPrint .hdc
        End If
    End With
 End If
ErrHandler:
 Exit Sub
End Sub
Private Sub mnuFileSave_Click()
On Error Resume Next
If Form1.RichTextBox3.Visible = True Then
  If TypeOfSave = 1 Then
   If RichTextBox3.FileName = Empty Then
    If Len(Dir1.Path) > 3 Then Text2.Text = Dir1.Path + "\" + File1.FileName
    If Len(Dir1.Path) = 3 Then Text2.Text = Dir1.Path + File1.FileName
    RichTextBox3.FileName = Text2.Text
   Else
    Text2.Text = RichTextBox3.FileName
   End If
    F = FreeFile
    Open Text2.Text For Output As #F
    Print #F, RichTextBox3.Text
    Close #F
   End If
   If TypeOfSave = 0 Then
    TypeOfSave = 1
    Dim sFile As String
        With dlgCommonDialog
            .DialogTitle = "Save"
            If Len(Dir1.Path) = 3 Then .FileName = Dir1.Path + NameSave
            If Len(Dir1.Path) > 3 Then .FileName = Dir1.Path + "\" + NameSave
            .Filter = "HTML File (*.html)|*.html|HTM File (*.htm)|*.htm"
            .ShowSave
             If Err.Number = 32755 Then GoTo ErrHandler
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
            RichTextBox3.FileName = .FileName
        End With
         F = FreeFile
         Open sFile For Output As #F
         Print #F, RichTextBox3.Text
         Close #F
    End If
End If
Txt2 = 0
If RichTextBox3.FileName = Empty Then RichTextBox3.FileName = sFile
Call Recent
Call Refresh_The_Dir
Exit Sub
ErrHandler:
 TypeOfSave = 0
 Exit Sub
End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo ErrHandler
If RichTextBox3.Visible = True Then
If Txt2 = 1 Then
  Beep
  myq = MsgBox("Do you want to save this page ?", vbYesNoCancel + vbQuestion, "Save...")
 End If
 If myq = 6 Then
  mnuFileSave_Click
  mnuFileClose.Enabled = False
  RichTextBox3.Text = ""
  RichTextBox3.Visible = False
  Image1.Visible = True
  Command1.Caption = "New HTML Page"
  Txt1 = "0"
 End If
 If myq = 7 Then
  mnuFileClose.Enabled = False
  RichTextBox3.Text = ""
  RichTextBox3.Visible = False
  Image1.Visible = True
  Command1.Caption = "New HTML Page"
  Txt1 = "0"
  CmdCounter4 = 1
 End If
 If myq = 2 Then Exit Sub
 For_File_Name = 1
    Dim sFile As String
    With dlgCommonDialog
        .DialogTitle = "Open"
        .FileName = GetSetting(App.Title, "Settings", "File1")
        .FileName = Empty
        .Filter = "HTML,HTM files (*.HTML,*.HTM)|*.html;*.htm|HTML File (*.html)|*.html|HTM File (*.htm)|*.htm"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    SaveSetting App.Title, "Settings", "File1", dlgCommonDialog.FileName
    NameSave = dlgCommonDialog.FileName
    RichTextBox3.Visible = True
    RichTextBox3.LoadFile sFile
    
   Call Recent

    Command7_Click
    Command1.Caption = "Close HTML Page"
    If Form1.RichTextBox3.Visible = True Then Form1.RichTextBox3.SetFocus
    Dir1.Path = CurDir
    Txt1 = "0"
    CmdCounter4 = 1
    Timer1.Enabled = True
Else
For_File_Name = 1
    RichTextBox3.Visible = True
    RichTextBox3.Text = ""
    With dlgCommonDialog
        .DialogTitle = "Open"
        .FileName = GetSetting(App.Title, "Settings", "File1")
        .FileName = Empty
        .Filter = "HTML , HTM Files (*.HTML,*.HTM)|*.html;*.htm|HTML File (*.html)|*.html|HTM File (*.htm)|*.htm"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    SaveSetting App.Title, "Settings", "File1", dlgCommonDialog.FileName
    NameSave = dlgCommonDialog.FileName
    RichTextBox3.LoadFile sFile
    
   Call Recent

    Command7_Click
    Command1.Caption = "Close HTML Page"
    If Form1.RichTextBox3.Visible = True Then Form1.RichTextBox3.SetFocus
    Dir1.Path = CurDir
    Timer1.Enabled = True
    Txt1 = "0"
    CmdCounter4 = 1
End If
Call Refresh_The_Dir
ErrHandler:
 If RichTextBox3.Text = Empty Then
  Image1.Visible = True
  RichTextBox3.Visible = False
  Timer1.Enabled = True
 End If
 Exit Sub
End Sub

Private Sub mnuFileNew_Click()
On Error Resume Next
  If Command1.Caption = "New HTML Page" Then
        If CmdCounter3 = 0 Then
          Call Command1_Click
          Call Command1_Click
          CmdCounter3 = 1
          CmdCounter4 = 1
        Else
          Call Command1_Click
        End If
  End If
End Sub
Private Sub mnuFile_Click()
On Error Resume Next

If RichTextBox3.Visible = False Then
 mnuFileClose.Enabled = False
 mnuFileSave.Enabled = False
 mnuFileSaveas.Enabled = False
 mnuFilePrint.Enabled = False
Else
 mnuFileClose.Enabled = True
 mnuFileSave.Enabled = True
 mnuFileSaveas.Enabled = True
 mnuFilePrint.Enabled = True
End If

If mnuFileRecentFiles1.Caption = Empty Then
 mnuFileRecentFiles1.Caption = mnuFileRecentFiles2.Caption
 mnuFileRecentFiles2.Caption = mnuFileRecentFiles3.Caption
 mnuFileRecentFiles3.Caption = mnuFileRecentFiles4.Caption
 mnuFileRecentFiles4.Caption = mnuFileRecentFiles5.Caption
 mnuFileRecentFiles5.Caption = mnuFileRecentFiles6.Caption
 mnuFileRecentFiles6.Caption = mnuFileRecentFiles7.Caption
End If

If mnuFileRecentFiles2.Caption = Empty Then
 mnuFileRecentFiles2.Caption = mnuFileRecentFiles3.Caption
 mnuFileRecentFiles3.Caption = mnuFileRecentFiles3.Caption
 mnuFileRecentFiles4.Caption = mnuFileRecentFiles5.Caption
 mnuFileRecentFiles5.Caption = mnuFileRecentFiles6.Caption
 mnuFileRecentFiles6.Caption = mnuFileRecentFiles7.Caption
End If

If mnuFileRecentFiles3.Caption = Empty Then
 mnuFileRecentFiles3.Caption = mnuFileRecentFiles4.Caption
 mnuFileRecentFiles4.Caption = mnuFileRecentFiles5.Caption
 mnuFileRecentFiles5.Caption = mnuFileRecentFiles6.Caption
 mnuFileRecentFiles6.Caption = mnuFileRecentFiles7.Caption
End If

If mnuFileRecentFiles4.Caption = Empty Then
 mnuFileRecentFiles4.Caption = mnuFileRecentFiles5.Caption
 mnuFileRecentFiles5.Caption = mnuFileRecentFiles6.Caption
 mnuFileRecentFiles6.Caption = mnuFileRecentFiles7.Caption
End If

If mnuFileRecentFiles5.Caption = Empty Then
 mnuFileRecentFiles5.Caption = mnuFileRecentFiles6.Caption
 mnuFileRecentFiles6.Caption = mnuFileRecentFiles7.Caption
End If

If mnuFileRecentFiles6.Caption = Empty Then
 mnuFileRecentFiles6.Caption = mnuFileRecentFiles7.Caption
End If

End Sub

Private Sub tbToolBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
sbStatusBar.SimpleText = "M2A HTML Maker  -  http://www.IranM2A.Tk"
If Rang = 1 Then Call ChangeBkColor
Rang = 0
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
 Txt2 = 0
 Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If Wi = Empty And He = Empty Then
 Me.Width = 12000
 Me.Height = 8700
 Timer2.Enabled = False
 Exit Sub
End If
 Me.Width = Wi
 Me.Height = He
 Timer2.Enabled = False
End Sub
