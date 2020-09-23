Attribute VB_Name = "basLightBox"
' /*
'  * Módulo:                          basLightBox
'  * Tipo de módulo:                  Módulo
'  *
'  *           Copyright© 1996-2010 CyberZone Software
'  *                                Santiago Sidera
'  */

' indica que requiere declaración explícita de variables
Option Explicit

' indica que las comparaciones que se hagan, serán en modo binario, y no en modo texto, "a<>A"
Option Compare Binary

' base de vectores y matrices: 0
Option Base 0

Public Enum FoxEffectFlags
            FOX_USE_MASK = &H1
            FOX_ANTI_ALIAS = &H2
            FOX_CHROME_LINEAR = &H4
            FOX_SRC_INVERT = &H100
            FOX_DST_INVERT = &H200
            FOX_MASK_INVERT = &H400
            FOX_SRC_GREYSCALE = &H1000
            FOX_DST_GREYSCALE = &H2000
            FOX_FLIP_X = &H40000
            FOX_FLIP_Y = &H80000
            FOX_TURN_LEFT = &H10000
            FOX_TURN_RIGHT = FOX_FLIP_X Or FOX_FLIP_Y
            FOX_TURN_90DEG = FOX_TURN_LEFT
            FOX_TURN_180DEG = FOX_TURN_RIGHT
            FOX_TURN_270DEG = FOX_FLIP_X Or FOX_FLIP_Y Or FOX_TURN_LEFT
End Enum

Public Declare Function FoxBrightness Lib "lightbox.dll" (ByVal hDC As Long, ByVal handle As Long, ByVal hSrcDC As Long, ByVal srchandle As Long, ByVal brightness As Long, Optional ByVal TransColor As Long, Optional ByVal Flags As FoxEffectFlags) As Long

Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function RegisterWindowMessage Lib "user32.dll" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private Const WS_EX_APPWINDOW     As Long = &H40000
Private Const GWL_EXSTYLE         As Long = (-20)

Public Type RECT
            Left As Long
            Top As Long
            Right As Long
            Bottom As Long
End Type

Public Sub CapturarVentana(ByRef frmForm As Form, ByRef picTemp As PictureBox, ByVal strArchivo As String)
Dim lngRet As Long
Dim lnghWnd As Long
Dim rtWindowRect As RECT
Dim lngHeight As Long
Dim lngWidth As Long

Set picTemp.Picture = Nothing
picTemp.Cls
picTemp.Visible = False
picTemp.AutoRedraw = True
picTemp.AutoSize = False

lnghWnd = frmForm.hWnd
lngRet = GetWindowRect(lnghWnd, rtWindowRect)

lngWidth = rtWindowRect.Right - rtWindowRect.Left
lngHeight = rtWindowRect.Bottom - rtWindowRect.Top

picTemp.Width = (lngWidth + 4) * Screen.TwipsPerPixelX
picTemp.Height = (lngHeight + 4) * Screen.TwipsPerPixelY

lngRet = BitBlt(picTemp.hDC, 0, 0, lngWidth, lngHeight, GetWindowDC(lnghWnd), 0, 0, vbSrcCopy)

On Error Resume Next
Kill strArchivo
SavePicture picTemp.Image, strArchivo
End Sub

Public Sub MostrarFormularioModal(ByRef frmFormPadre As Form, ByRef frmForm As Form, Optional ByVal lngGradoBrillo As Long = -140, Optional ByVal lngIncrementoBrillo As Long = 10, Optional ByVal lngWindowState As Long = vbNormal)
Dim lngBrillo As Long

CapturarVentana frmFormPadre, frmFormPadre.picTemp, strAppPath & "captemp.bmp"

Load frmLightBox
frmLightBox.Visible = False

frmLightBox.Top = frmFormPadre.Top
frmLightBox.Left = frmFormPadre.Left
frmLightBox.Width = frmFormPadre.Width
frmLightBox.Height = frmFormPadre.Height
frmLightBox.Picture = LoadPicture(strAppPath & "captemp.bmp")
Kill strAppPath & "captemp.bmp"
frmLightBox.Visible = True

frmFormPadre.Visible = False
If (frmFormPadre Is frmMain) Then AgregarATaskBar frmFormPadre.hWnd

lngBrillo = 0

If lngGradoBrillo < 0 Then
   While (lngBrillo > lngGradoBrillo)
         lngBrillo = lngBrillo - lngIncrementoBrillo
         If (frmLightBox.Width >= (800 * Screen.TwipsPerPixelX)) Or (frmLightBox.Height >= (600 * Screen.TwipsPerPixelY)) Then lngBrillo = lngGradoBrillo
         FoxBrightness frmLightBox.hDC, frmLightBox.Image.handle, frmLightBox.hDC, frmLightBox.Picture.handle, lngBrillo, RGB(0, 0, 0), FOX_USE_MASK
         frmLightBox.Refresh
         DoEvents
   Wend
ElseIf lngGradoBrillo > 0 Then
   While (lngBrillo < lngGradoBrillo)
         lngBrillo = lngBrillo + lngIncrementoBrillo
         If (frmLightBox.Width >= (800 * Screen.TwipsPerPixelX)) Or (frmLightBox.Height >= (600 * Screen.TwipsPerPixelY)) Then lngBrillo = lngGradoBrillo
         FoxBrightness frmLightBox.hDC, frmLightBox.Image.handle, frmLightBox.hDC, frmLightBox.Picture.handle, lngBrillo, RGB(0, 0, 0), FOX_USE_MASK
         frmLightBox.Refresh
         DoEvents
   Wend
End If

Load frmForm
frmForm.Show vbModal

If lngGradoBrillo < 0 Then
   While (lngBrillo <> 0)
         lngBrillo = lngBrillo + lngIncrementoBrillo
         If (frmLightBox.Width >= (800 * Screen.TwipsPerPixelX)) Or (frmLightBox.Height >= (600 * Screen.TwipsPerPixelY)) Then lngBrillo = 0
         FoxBrightness frmLightBox.hDC, frmLightBox.Image.handle, frmLightBox.hDC, frmLightBox.Picture.handle, lngBrillo, RGB(0, 0, 0), FOX_USE_MASK
         frmLightBox.Refresh
         DoEvents
   Wend
ElseIf lngGradoBrillo > 0 Then
   While (lngBrillo <> 0)
         lngBrillo = lngBrillo - lngIncrementoBrillo
         If (frmLightBox.Width >= (800 * Screen.TwipsPerPixelX)) Or (frmLightBox.Height >= (600 * Screen.TwipsPerPixelY)) Then lngBrillo = 0
         FoxBrightness frmLightBox.hDC, frmLightBox.Image.handle, frmLightBox.hDC, frmLightBox.Picture.handle, lngBrillo, RGB(0, 0, 0), FOX_USE_MASK
         frmLightBox.Refresh
         DoEvents
   Wend
End If

frmFormPadre.Visible = True

frmLightBox.Visible = False
Unload frmLightBox
End Sub

Public Sub AgregarATaskBar(ByVal lnghWnd As Long)
Dim lStyle As Long
Dim lResult As Long
Dim lHook As Long
Dim lTrayhWnd As Long
Dim lTBhWnd As Long

lStyle = GetWindowLong(lnghWnd, GWL_EXSTYLE)
lResult = SetWindowLong(lnghWnd, GWL_EXSTYLE, lStyle Or WS_EX_APPWINDOW)
lHook = RegisterWindowMessage("SHELLHOOK")
lTrayhWnd = FindWindowEx(0, 0, "Shell_TrayWnd", vbNullString)

If lTrayhWnd Then
   lTBhWnd = FindWindowEx(lTrayhWnd, 0, "RebarWindow32", vbNullString)
   If lTBhWnd Then
      lTBhWnd = FindWindowEx(lTBhWnd, 0, "MSTaskSwWClass", vbNullString)
      If lTBhWnd Then
         lResult = PostMessage(lTBhWnd, lHook, 1, ByVal lnghWnd)
      End If
   End If
End If
End Sub
