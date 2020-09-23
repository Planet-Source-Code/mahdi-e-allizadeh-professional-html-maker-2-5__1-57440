Attribute VB_Name = "Module1"
Option Explicit
Public Const MF_BYPOSITION = &H400&
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function CreateRoundRectRgn Lib "gdi32" _
    (ByVal RectX1 As Long, ByVal RectY1 As Long, ByVal RectX2 As Long, _
    ByVal RectY2 As Long, ByVal EllipseWidth As Long, _
    ByVal EllipseHeight As Long) As Long
Public Txt1 As Byte
Public Txt2 As Byte 'It is for show msgbox for save or don't show it . if Txt2 = 0 then don't show msgbox and if Txt2 = 1 then show msgbox
Public Txt3 As Byte 'It is for to prevent more than 1 Page open M2A web browser
Public YourFavorite As Byte 'This varriable is for that if you select M2A Web Browser (In Options) then when you click Test see your
'HTML page in M2A Web Browser and if you select MS Internet Explorer and then click Test then you can see your HTML Page in MS Internet Explorer
Public Adres As Variant
Public NameF As Variant
Public NameDir As Variant
Public NameSave As Variant
Public Rang As Variant
Public Tim As Variant
Public Komak As Variant
Public Var1 As Byte

Sub MenuAddBitmap( _
                   FrmHwnd&, _
                   MainMnuIndex%, _
                   SubMnuIndex%, _
                   sixteen_by_sixteen_picBox As StdPicture _
                                    )
Dim hMenu As Long, hSubMenu As Long
'get the handle of the menu
hMenu = GetMenu(FrmHwnd&)
'get the first submenu
hSubMenu = GetSubMenu(hMenu, MainMnuIndex%)
'set the menu bitmap
SetMenuItemBitmaps hSubMenu, SubMnuIndex%, MF_BYPOSITION, sixteen_by_sixteen_picBox, sixteen_by_sixteen_picBox '  End Sub

End Sub
