VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{C3DF5D2F-40CD-4CDD-B283-AA3D32054C81}#1.0#0"; "AutoResize.ocx"
Begin VB.Form Form2 
   Caption         =   "M2A Web Browser"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8160
   ScaleWidth      =   11880
   Begin Project1.AutoResize Resize 
      Left            =   4560
      Tag             =   "NO"
      Top             =   3600
      _ExtentX        =   714
      _ExtentY        =   714
      AspectRatioValue=   0
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2760
      Top             =   5280
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   2640
      Top             =   2640
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   5400
      Top             =   3840
   End
   Begin VB.PictureBox Picture1 
      Height          =   7335
      Left            =   0
      ScaleHeight     =   7275
      ScaleWidth      =   11835
      TabIndex        =   4
      Top             =   960
      Width           =   11895
      Begin SHDocVwCtl.WebBrowser brwWebBrowser 
         Height          =   7095
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   11835
         ExtentX         =   20876
         ExtentY         =   12515
         ViewMode        =   1
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   -1  'True
         NoClientEdge    =   -1  'True
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Go"
            Object.ToolTipText     =   "Go"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.ComboBox cboAddress 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   10935
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4440
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":058A
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0956
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0CE9
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1097
            Key             =   "Go"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "G  O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Address :"
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
      TabIndex        =   2
      Top             =   650
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public He2, Wi2 As Variant
Private Sub Command1_Click()
    brwWebBrowser.Navigate cboAddress.Text
End Sub
Private Sub Form_Load()
On Error Resume Next
Me.Left = (Screen.Width - Form1.Width) / 2
Me.Top = (Screen.Height - Form1.Height) / 2
Me.Left = GetSetting(App.Title, "Settings", "MainLeft2")
Me.Top = GetSetting(App.Title, "Settings", "MainTop2")
Wi2 = GetSetting(App.Title, "Settings", "MainWidth2")
He2 = GetSetting(App.Title, "Settings", "MainHeight2")
Timer3.Enabled = True
brwWebBrowser.Offline = True
End Sub

Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
End Sub

Private Sub brwWebBrowser_NavigateComplete(ByVal URL As String)
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
         bFound = True
         Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
     cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
End Sub

Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    brwWebBrowser.Navigate cboAddress.Text
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft2", Me.Left
        SaveSetting App.Title, "Settings", "MainTop2", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth2", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight2", Me.Height
End If
Unload Me
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Back"
            brwWebBrowser.GoBack
        Case "Stop"
            brwWebBrowser.Stop
        Case "Refresh"
            Command1_Click
        Case "Go"
            Command1_Click
    End Select
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
If Me.WindowState = vbMinimized Then Me.WindowState = vbMaximized
Command1_Click
Timer1.Enabled = False
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
Me.Width = Wi2
Me.Height = He2
Timer3.Enabled = False
End Sub
