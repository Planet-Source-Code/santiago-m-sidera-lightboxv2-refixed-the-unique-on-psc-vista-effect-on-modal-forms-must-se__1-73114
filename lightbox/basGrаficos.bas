Attribute VB_Name = "basGráficos"
' /*
'  * Módulo:                          basGráficos
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

Public Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Public Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long
Public Declare Function LoadCursorFromFile Lib "user32.dll" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Public Declare Function SetSystemCursor Lib "user32.dll" (ByVal hCur As Long, ByVal id As Long) As Long
Public Declare Function GetCursor Lib "user32.dll" () As Long
Public Declare Function CopyIcon Lib "user32.dll" (ByVal hCur As Long) As Long

Public Declare Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Declare Function EnumDisplaySettings Lib "user32.dll" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32.dll" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long

Public Declare Function GetMenu Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemInfo Lib "user32.dll" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, lpstructMenuItem As structMenuItem) As Long
Public Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcstructMenuItem As structMenuItem) As Long
Public Declare Function DrawMenuBar Lib "user32.dll" (ByVal hWnd As Long) As Long

Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

Public Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndParent As Long) As Long
Public Declare Function EnableWindow Lib "user32.dll" (ByVal lnghWnd As Long, ByVal bEnabled As Boolean) As Boolean

Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Const HALFTONE                     As Long = 4

' SetWindowPos()
Public Const HWND_TOP                     As Long = 0
Public Const HWND_DESKTOP                 As Long = 0
Public Const HWND_BROADCAST               As Long = &HFFFF&
Public Const HWND_BOTTOM                  As Long = 1
Public Const HWND_TOPMOST                 As Long = (-1)
Public Const HWND_NOTOPMOST               As Long = (-2)
Public Const SWP_NOACTIVATE               As Long = &H10
Public Const SWP_SHOWWINDOW               As Long = &H40
Public Const SWP_FRAMECHANGED             As Long = &H20
Public Const SWP_DRAWFRAME                As Long = SWP_FRAMECHANGED
Public Const SWP_HIDEWINDOW               As Long = &H80
Public Const SWP_NOCOPYBITS               As Long = &H100
Public Const SWP_NOMOVE                   As Long = &H2
Public Const SWP_NOOWNERZORDER            As Long = &H200
Public Const SWP_NOREDRAW                 As Long = &H8
Public Const SWP_NOREPOSITION             As Long = SWP_NOOWNERZORDER
Public Const SWP_NOSIZE                   As Long = &H1
Public Const SWP_NOZORDER                 As Long = &H4

' SetWindowLong()
Public Const WS_OVERLAPPED               As Long = &H0&
Public Const WS_BORDER                   As Long = &H800000
Public Const WS_CAPTION                  As Long = &HC00000
Public Const WS_CHILD                    As Long = &H40000000
Public Const WS_CLIPCHILDREN             As Long = &H2000000
Public Const WS_CLIPSIBLINGS             As Long = &H4000000
Public Const WS_DISABLED                 As Long = &H8000000
Public Const WS_DLGFRAME                 As Long = &H400000
Public Const WS_EX_ACCEPTFILES           As Long = &H10&
Public Const WS_EX_DLGMODALFRAME         As Long = &H1&
Public Const WS_EX_NOPARENTNOTIFY        As Long = &H4&
Public Const WS_EX_TOPMOST               As Long = &H8&
Public Const WS_EX_TRANSPARENT           As Long = &H20&
Public Const WS_GROUP                    As Long = &H20000
Public Const WS_HSCROLL                  As Long = &H100000
Public Const WS_MAXIMIZE                 As Long = &H1000000
Public Const WS_MINIMIZE                 As Long = &H20000000
Public Const WS_MAXIMIZEBOX              As Long = &H10000
Public Const WS_MINIMIZEBOX              As Long = &H20000
Public Const WS_POPUP                    As Long = &H80000000
Public Const WS_SYSMENU                  As Long = &H80000
Public Const WS_TABSTOP                  As Long = &H10000
Public Const WS_THICKFRAME               As Long = &H40000
Public Const WS_VSCROLL                  As Long = &H200000
Public Const WS_VISIBLE                  As Long = &H10000000
Public Const WS_ICONIC                   As Long = WS_MINIMIZE
Public Const WS_SIZEBOX                  As Long = WS_THICKFRAME
Public Const WS_TILED                    As Long = WS_OVERLAPPED
Public Const WS_OVERLAPPEDWINDOW         As Long = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX Or WS_POPUP)
Public Const WS_TILEDWINDOW              As Long = WS_OVERLAPPEDWINDOW
Public Const WS_CHILDWINDOW              As Long = (WS_CHILD)
Public Const WS_POPUPWINDOW              As Long = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)

Public Const GWL_HINSTANCE               As Long = (-6)
Public Const GWL_STYLE                   As Long = (-16)
Public Const GWL_EXSTYLE                 As Long = (-20)
Public Const GWL_WNDPROC                 As Long = (-4)
Public Const GWL_ID                      As Long = (-12)
Public Const GWL_USERDATA                As Long = (-21)

' SendMessage()
Public Const WM_USER                     As Long = &H400
Public Const CCM_FIRST                   As Long = &H2000&
Public Const CCM_SETBKCOLOR              As Long = (CCM_FIRST + 1)
Public Const PBM_SETBKCOLOR              As Long = CCM_SETBKCOLOR
Public Const PBM_SETBARCOLOR             As Long = (WM_USER + 9)

Public Const strFuentePredeterminada          As String = "calibri"
Public Const lngTamañoFuentePredeterminado    As Long = 10

Public Const MF_STRING                   As Long = &H0&
Public Const MF_HELP                     As Long = &H4000&
Public Const MFS_DEFAULT                 As Long = &H1000&
Public Const MIIM_ID                     As Long = &H2
Public Const MIIM_SUBMENU                As Long = &H4
Public Const MIIM_TYPE                   As Long = &H10
Public Const MIIM_DATA                   As Long = &H20

Public Type structMenuItem
            cbSize                       As Long
            fMask                        As Long
            fType                        As Long
            fState                       As Long
            wid                          As Long
            hSubMenu                     As Long
            hbmpChecked                  As Long
            hbmpUnchecked                As Long
            dwItemData                   As Long
            dwTypeData                   As String
            cch                          As Long
End Type

Public Const OCR_NORMAL                  As Long = 32512

Public Enum estado
            [estado-desocupado] = 0
            [estado-ocupado] = 1
End Enum

Public lngCursorViejo                    As Long
Public lngCursorViejoAntesDeIniciar      As Long
Public blnFormsBloqueados                As Boolean

Public Sub Centrar(ByVal frm As Form, Optional ByRef objRef As Object)
If (objRef Is Nothing) Then
   frm.Top = (Screen.Height / 2) - (frm.Height / 2)
   frm.Left = (Screen.Width / 2) - (frm.Width / 2)
Else
   frm.Top = objRef.Top + ((objRef.Height / 2) - (frm.Height / 2))
   frm.Left = objRef.Left + ((objRef.Width / 2) - (frm.Width / 2))
End If
End Sub

Public Sub AlFrente(ByVal frmFormulario As Form, Optional ByVal blnEstado As Boolean = True)
If blnEstado Then
   SetWindowPos frmFormulario.hWnd, HWND_TOPMOST, frmFormulario.Left / Screen.TwipsPerPixelX, frmFormulario.Top / Screen.TwipsPerPixelY, frmFormulario.Width / Screen.TwipsPerPixelX, frmFormulario.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
Else
   SetWindowPos frmFormulario.hWnd, HWND_NOTOPMOST, frmFormulario.Left / Screen.TwipsPerPixelX, frmFormulario.Top / Screen.TwipsPerPixelY, frmFormulario.Width / Screen.TwipsPerPixelX, frmFormulario.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End If
End Sub

Public Sub SetProgressBarColor(ByRef objObjeto As Object, ByVal lngColor As Long)
SendMessage objObjeto.hWnd, PBM_SETBARCOLOR, 0&, ByVal lngColor
End Sub

Public Sub SetProgressBarBackColor(ByRef objObjeto As Object, ByVal lngColor As Long)
SendMessage objObjeto.hWnd, PBM_SETBKCOLOR, 0&, ByVal lngColor
End Sub

Public Function PrepararLabel(ByRef lblLabel As Label, Optional ByVal strCaption As String = vbNullString, Optional ByVal lngHeight As Long = -1, Optional ByVal lngWidth As Long = -1, Optional ByVal blnAutoSize As Boolean = False, Optional ByVal lngLeft As Long = -1, Optional ByVal lngTop As Long = -1, Optional ByVal strFuente As String = vbNullString, Optional ByVal lngTamañoFuente As Long = -1, Optional ByVal blnNegrita As Boolean = False, Optional ByVal blnCursiva As Boolean = False, Optional ByVal lngColorFuente As ColorConstants = vbBlack, Optional ByVal lngBackStyle As Long = 0) As Boolean
If (strCaption <> vbNullString) Then lblLabel.Caption = strCaption
lblLabel.AutoSize = blnAutoSize
If (Not blnAutoSize) Then
   If (lngHeight <> -1) Then lblLabel.Height = lngHeight
   If (lngWidth <> -1) Then lblLabel.Width = lngWidth
End If
If (lngLeft <> -1) Then lblLabel.Left = lngLeft
If (lngTop <> -1) Then lblLabel.Top = lngTop

If (strFuente = vbNullString) Then lblLabel.Font = strFuentePredeterminada
If (lngTamañoFuente = -1) Then lblLabel.Font.Size = lngTamañoFuentePredeterminado

lblLabel.ForeColor = lngColorFuente
lblLabel.Font.Bold = blnNegrita
lblLabel.Font.Italic = blnCursiva
lblLabel.BackStyle = lngBackStyle
End Function

Public Sub Posicionar(ByRef lstLista As Object, ByVal strCad As String)
Dim i As Long
Dim lngPos As Long

lngPos = -1
For i = 0 To (lstLista.ListCount - 1)
    If (Trim$(LCase$(lstLista.List(i))) = Trim$(LCase$(strCad))) Then
       lngPos = i
       Exit For
    End If
Next i

lstLista.ListIndex = lngPos
End Sub

Public Function Foco(ByRef ctlControl As Object) As Boolean
Foco = False

On Error GoTo Errores 'Resume Next

ctlControl.SetFocus
ctlControl.SelStart = Len(ctlControl.Text)
'If (TypeOf ctlControl Is MaskEdBox) Then ctlControl.SelStart = 0
ctlControl.SetFocus

Foco = True
Exit Function

Errores:
End Function

Public Function MoveMenuLeft(FORMhwnd As Long) As Boolean
Dim mnuItemInfo As structMenuItem
Dim hMenu As Long
Dim BuffStr As String * 255
Dim iMenuCount As Integer

MoveMenuLeft = False
hMenu = GetMenu(FORMhwnd)
BuffStr = Space$(80)

With mnuItemInfo
    .cbSize = Len(mnuItemInfo)
    .dwTypeData = BuffStr & Chr$(0)
    .fType = MF_STRING
    .cch = Len(mnuItemInfo.dwTypeData)
    .fState = MFS_DEFAULT
    .fMask = MIIM_ID Or MIIM_DATA Or MIIM_TYPE Or MIIM_SUBMENU
End With

For iMenuCount = 100 To 1 Step -1
    If GetMenuItemInfo(hMenu, iMenuCount, True, mnuItemInfo) <> 0 Then Exit For
Next iMenuCount

If iMenuCount > 0 Then
   mnuItemInfo.fType = mnuItemInfo.fType Or MF_HELP
   SetMenuItemInfo hMenu, iMenuCount, True, mnuItemInfo
   DrawMenuBar (FORMhwnd)
   MoveMenuLeft = True
End If
End Function

Public Sub AnimarCursor(ByVal strArchivo As String)
EstablecerCursor strArchivo
End Sub

Public Sub ObtenerCursor()
lngCursorViejo = CopyIcon(GetCursor())
End Sub

Public Sub ObtenerCursorAntesDeIniciar()
lngCursorViejoAntesDeIniciar = CopyIcon(GetCursor())
End Sub

Public Sub RestablecerCursorAntesDeIniciar()
SetSystemCursor lngCursorViejoAntesDeIniciar, OCR_NORMAL
End Sub

Public Sub EstablecerCursor(ByVal strArchivo As String)
Dim lngCursorNuevo As Long

ObtenerCursor
lngCursorNuevo = LoadCursorFromFile(strArchivo)
SetSystemCursor lngCursorNuevo, OCR_NORMAL
End Sub

Public Sub RestablecerCursor()
SetSystemCursor lngCursorViejo, OCR_NORMAL
End Sub

Public Function SetearHijoMDI(ByRef frmChildForm As Form, ByRef frmMDI As MDIForm) As Boolean
Dim lngRet As Long
Dim lngStyle As Long

SetearHijoMDI = False

lngRet = SetParent(frmChildForm.hWnd, frmMDI.hWnd)
lngStyle = GetWindowLong(frmChildForm.hWnd, GWL_STYLE)
lngStyle = SetWindowLong(frmChildForm.hWnd, GWL_STYLE, lngStyle Or WS_CHILD Or WS_OVERLAPPEDWINDOW)
frmChildForm.Hide
frmChildForm.Show , frmMDI

SetearHijoMDI = (lngRet <> 0)
End Function

Public Sub BloqForms(Optional ByVal strFormExcepción As String = vbNullString, Optional ByVal strFormMain As String = vbNullString)
Dim frmForm As Form

If (blnFormsBloqueados) Then Exit Sub
For Each frmForm In Forms
    If (Trim$(LCase$(strFormExcepción)) <> vbNullString) Then
       If (Trim$(LCase$(frmForm.Name)) <> Trim$(LCase$(strFormExcepción))) Then
          If (Trim$(LCase$(frmForm.Name)) <> Trim$(LCase$(strFormMain))) Then
             frmForm.Enabled = False
             'EnableWindow frmForm.hWnd, False
          End If
       End If
    Else
       If (Trim$(LCase$(frmForm.Name)) <> Trim$(LCase$(strFormMain))) Then
          frmForm.Enabled = False
          'EnableWindow frmForm.hWnd, False
       End If
    End If
Next frmForm

blnFormsBloqueados = True
End Sub

Public Sub UnBloqForms(Optional ByVal strFormExcepción As String = vbNullString, Optional ByVal strFormMain As String = vbNullString)
Dim frmForm As Form

If (Not blnFormsBloqueados) Then Exit Sub
For Each frmForm In Forms
    If (Trim$(LCase$(strFormExcepción)) <> vbNullString) Then
       If (Trim$(LCase$(frmForm.Name)) <> Trim$(LCase$(strFormExcepción))) Then
          If (Trim$(LCase$(frmForm.Name)) <> Trim$(LCase$(strFormMain))) Then
             frmForm.Enabled = True
             'EnableWindow frmForm.hWnd, True
          End If
       End If
    Else
       If (Trim$(LCase$(frmForm.Name)) <> Trim$(LCase$(strFormMain))) Then
          frmForm.Enabled = True
          'EnableWindow frmForm.hWnd, True
       End If
    End If
Next frmForm

blnFormsBloqueados = False
End Sub
