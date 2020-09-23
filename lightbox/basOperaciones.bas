Attribute VB_Name = "basOperaciones"
' /*
'  * Módulo:                          basOperaciones
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

' **********************************************************************************************
' declaraciones de funciones y subrutinas API (application programming interface)
' **********************************************************************************************

Public Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long

Public Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Declare Function CreateDirectory Lib "kernel32.dll" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITYSTRUCT) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Public Declare Function CopyFileA Lib "kernel32.dll" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function CopyFileEx Lib "kernel32.dll" Alias "CopyFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal lpProgressRoutine As Long, ByVal lpData As Long, ByRef pbCancel As Long, ByVal dwCopyFlags As Long) As Long

Public Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32.dll" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesAvailableToCaller As LargeInt, TotalNumberOfBytes As LargeInt, TotalNumberOfFreeBytes As LargeInt) As Long

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

' **********************************************************************************************
' constantes y tipos de datos de las funciones y subrutinas API
' **********************************************************************************************

Public Const MAX_PATH                     As Long = 260
Public Const INVALID_HANDLE_VALUE         As Long = -1
Public Const FILE_ATTRIBUTE_DIRECTORY     As Long = &H10

Public Const FO_DELETE                    As Long = &H3
Public Const FO_COPY                      As Long = &H2
Public Const FO_RENAME                    As Long = &H4
Public Const FO_MOVE                      As Long = &H1
Public Const FOF_SIMPLEPROGRESS           As Long = &H100
Public Const FOF_ALLOWUNDO                As Long = &H40
Public Const FOF_NOCONFIRMATION           As Long = &H10
Public Const FOF_SILENT                   As Long = &H4

' CopyFileEx()
Public Const PROGRESS_CONTINUE                 As Long = &H0
Public Const PROGRESS_CANCEL                   As Long = &H1
Public Const PROGRESS_STOP                     As Long = &H2
Public Const PROGRESS_QUIET                    As Long = &H3

Public Const CALLBACK_CHUNK_FINISHED           As Long = &H0
Public Const CALLBACK_STREAM_SWITCH            As Long = &H1

Public Const COPY_FILE_REPLACE_IF_EXISTS       As Long = &H0
Public Const COPY_FILE_FAIL_IF_EXISTS          As Long = &H1
Public Const COPY_FILE_RESTARTABLE             As Long = &H2
Public Const COPY_FILE_OPEN_SOURCE_FOR_WRITE   As Long = &H4

Public Const lngAllAttr                        As Long = vbArchive Or vbDirectory Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem

Public Type FILETIME
            dwLowDateTime                 As Long
            dwHighDateTime                As Long
End Type

Public Type WIN32_FIND_DATA
            dwFileAttributes              As Long
            ftCreationTime                As FILETIME
            ftLastAccessTime              As FILETIME
            ftLastWriteTime               As FILETIME
            nFileSizeHigh                 As Long
            nFileSizeLow                  As Long
            dwReserved0                   As Long
            dwReserved1                   As Long
            cFileName                     As String * MAX_PATH
            cAlternate                    As String * 14
End Type

Public Type SHFILEOPSTRUCT
            hWnd                          As Long
            wFunc                         As Long
            pFrom                         As String
            pTo                           As String
            fFlags                        As Integer
            fAnyOperationsAborted         As Long
            hNameMappings                 As Long
            lpszProgressTitle             As Long
End Type

Public Type SECURITYSTRUCT
            nLength                       As Long
            lpSecurityDescriptor          As Long
            bInheritHandle                As Boolean
End Type

Public Type ULong ' Unsigned Long
            Byte1                         As Byte
            Byte2                         As Byte
            Byte3                         As Byte
            Byte4                         As Byte
End Type

Public Type LargeInt ' Large Integer
            LoDWord                       As ULong
            HiDWord                       As ULong
            LoDWord2                      As ULong
            HiDWord2                      As ULong
End Type

Public Enum e_Sistema
            [systemdir] = 0
            [tempdir] = 1
            [windir] = 2
End Enum

Public pgbProgresoCopiaCallback           As Object
Public blnCancelaCopia                    As Boolean

Public Function Existe(ByVal strArgumento As String) As Boolean
Dim blnExisteCarpeta As Boolean
Dim blnExisteArchivo As Boolean

blnExisteCarpeta = ExisteCarpeta(strArgumento)
blnExisteArchivo = ExisteArchivo(strArgumento)

Select Case True
       Case blnExisteCarpeta And blnExisteArchivo
            Existe = (blnExisteCarpeta And blnExisteArchivo)
       Case blnExisteCarpeta
            Existe = blnExisteCarpeta
       Case blnExisteArchivo
            Existe = blnExisteArchivo
       Case Else
            Existe = False
End Select
End Function

Public Function ExisteArchivo(ByVal strArchivo As String) As Boolean
Dim WFD As WIN32_FIND_DATA
Dim lngHandle As Long

ExisteArchivo = False
On Error GoTo Errores

lngHandle = FindFirstFile(strArchivo, WFD)
ExisteArchivo = (lngHandle <> INVALID_HANDLE_VALUE)
FindClose lngHandle
Exit Function

Errores:
FindClose lngHandle
End Function

Public Function ExisteCarpeta(ByVal strPath As String) As Boolean
Dim lngHandle As Long
Dim WFD As WIN32_FIND_DATA

ExisteCarpeta = False
On Error GoTo Errores

strPath = AddRemoveSlash(strPath, True)
lngHandle = FindFirstFile(strPath, WFD)
ExisteCarpeta = (lngHandle <> INVALID_HANDLE_VALUE) And (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY)
FindClose lngHandle
Exit Function

Errores:
FindClose lngHandle
End Function

Public Function AddRemoveSlash(ByVal strPath As String, Optional ByVal blnRemover As Boolean = False) As String
Dim strSlash As String
Dim strPathIntermedio As String
Dim strUltCar As String
Dim blnPathUNC As String

strSlash = "\"
AddRemoveSlash = strPath
blnPathUNC = False
On Error GoTo Errores

strPathIntermedio = Trim$(strPath)

blnPathUNC = EsRutaUNC(strPath)
If blnPathUNC Then Mid$(strPathIntermedio, 1, 2) = "  "
strPathIntermedio = Trim$(strPathIntermedio)

Do
   strPathIntermedio = Replace$(strPathIntermedio, (strSlash & strSlash), strSlash)
   strUltCar = Right$(strPathIntermedio, 1)
   If strUltCar = "\" Then strPathIntermedio = Left$(strPathIntermedio, (Len(strPathIntermedio) - 1))
Loop Until (strUltCar <> "\")

strPathIntermedio = Replace$(strPathIntermedio, (strSlash & strSlash), strSlash)
If (Not blnRemover) Then strPathIntermedio = strPathIntermedio & "\"

If blnPathUNC Then
   strPathIntermedio = "\\" & strPathIntermedio
End If

AddRemoveSlash = strPathIntermedio
Exit Function

Errores:
End Function

Public Function ObtenerLíneaComando(Optional ByVal lngMaxArgs As Long = 10) As String()
Dim strCar As String
Dim strLCMD As String
Dim lngLenLCMD As String
Dim I As Long
Dim blnArgIn As Boolean
Dim lngNumArgs As Long
Dim strArgArray() As String

On Error GoTo Errores
ReDim strArgArray(lngMaxArgs)
lngNumArgs = 0
blnArgIn = False

strLCMD = Command()
lngLenLCMD = Len(strLCMD)

For I = 1 To lngLenLCMD
    strCar = Mid(strLCMD, I, 1)
    If ((strCar <> " ") And (strCar <> vbTab)) Then
       If Not blnArgIn Then
          If lngNumArgs >= lngMaxArgs Then Exit For
          lngNumArgs = lngNumArgs + 1
          blnArgIn = True
       End If
       strArgArray(lngNumArgs) = strArgArray(lngNumArgs) & LCase$(strCar)
    Else
       blnArgIn = False
    End If
Next I

ReDim Preserve strArgArray(lngNumArgs)
ObtenerLíneaComando = strArgArray()
Exit Function

Errores:
MsgBox "Error: " & LCase$(Err.Description) & vbNewLine & "Origen: " & "basOperaciones.ObtenerLíneaComando()", vbOKOnly Or vbCritical, "Error"
End Function

Public Function ObtenerDirectorio(Optional ByVal eDir As e_Sistema = [tempdir]) As String
Dim strTemp As String

ObtenerDirectorio = "#ERR#"

On Error GoTo Errores
strTemp = Space$(255)

Select Case eDir
       Case [tempdir]
            GetTempPath Len(strTemp), strTemp
       Case [windir]
            GetWindowsDirectory strTemp, Len(strTemp)
       Case [systemdir]
            GetSystemDirectory strTemp, Len(strTemp)
End Select

strTemp = AddRemoveSlash(Left$(Trim$(strTemp), Len(Trim$(strTemp)) - 1))
ObtenerDirectorio = strTemp
Exit Function

Errores:
End Function

Public Function GetFileName(ByVal strFilePath As String) As String
On Error Resume Next
GetFileName = Right$(strFilePath, Len(strFilePath) - InStrRev(strFilePath, "\"))
End Function

Public Function GetFileDir(ByVal strFilePath As String) As String
On Error Resume Next
GetFileDir = Left$(strFilePath, Len(strFilePath) - Len(GetFileName$(strFilePath)) - 1)
End Function

Public Function Borrar(ByRef strParam As String) As Boolean 'Long
Dim lngRet As Long
Dim typOperation As SHFILEOPSTRUCT

Borrar = False
On Error GoTo Errores

typOperation.wFunc = FO_DELETE
typOperation.pFrom = AddRemoveSlash(strParam, True)
typOperation.fFlags = FOF_SILENT + FOF_NOCONFIRMATION
lngRet = SHFileOperation(typOperation)
Borrar = (lngRet = 0)
Exit Function

Errores:
MsgBox Err.Number
End Function

Public Function CrearDirectorio(ByVal strPath As String) As Boolean
Dim intContador As Integer
Dim strTempDir As String
Dim strRuta As String
Dim strSep() As String
Dim lngRet As Long
Dim lpSA As SECURITYSTRUCT
Dim strTemp As String

CrearDirectorio = False
On Error GoTo Errores

strRuta = AddRemoveSlash(strPath, True)
strSep() = Split(strRuta, "\")

For intContador = 0 To UBound(strSep())
    strTempDir = strTempDir & AddRemoveSlash(strSep(intContador))
    strTemp = AddRemoveSlash(strTempDir, True)
    If Len(strTemp) > 2 Then
       If (Not Existe(strTemp)) Then
          lpSA.nLength = Len(lpSA)
          lngRet = CreateDirectory(strTemp, lpSA)
          If lngRet = 0 Then GoTo Errores
       End If
    End If
Next intContador

CrearDirectorio = Existe(strPath)
Exit Function

Errores:
End Function

Public Function EsRutaUNC(ByVal strPath As String) As Boolean
Dim strPathIntermedio As String

EsRutaUNC = False
strPathIntermedio = Trim$(strPath)
If Left$(strPathIntermedio, 2) = "\\" Then EsRutaUNC = True
End Function

Public Function QuitarRutaUNC(ByVal strPath As String) As String
Dim strPathIntermedio As String

strPathIntermedio = Trim$(strPath)
QuitarRutaUNC = strPathIntermedio
If (Not EsRutaUNC(strPath)) Then Exit Function

strPathIntermedio = Trim$(strPathIntermedio)
Mid$(strPathIntermedio, 1, 2) = "  "
strPathIntermedio = Trim$(strPathIntermedio)

QuitarRutaUNC = strPathIntermedio
End Function

Public Function CopiarArchivo(ByRef objObjetoProgreso As Object, ByVal strOrigen As String, ByVal strDestino As String, Optional ByVal blnReemplazarSiExiste As Boolean = True) As Boolean
Dim lngRet As Long

blnCancelaCopia = False
On Error GoTo Errores

' FALTA HACER:
' Chequear si es WinXP o Win98
' Si es WinXP, utilizar CopyFileEx
' Si es Win98, utilizar CopyFileA
Vincular objObjetoProgreso, pgbProgresoCopiaCallback
lngRet = CopyFileEx(strOrigen, strDestino, AddressOf CopyFileCallback, 0&, CLng(blnCancelaCopia), IIf(blnReemplazarSiExiste, COPY_FILE_REPLACE_IF_EXISTS, COPY_FILE_FAIL_IF_EXISTS))
RomperVínculo pgbProgresoCopiaCallback
CopiarArchivo = (lngRet <> 0)
Exit Function

Errores:
End Function

Private Function CopyFileCallback(ByVal TotalFileSize As Currency, ByVal TotalBytesTransferred As Currency, ByVal StreamSize As Currency, ByVal StreamBytesTransferred As Currency, ByVal dwStreamNumber As Long, ByVal dwCallbackReason As Long, ByVal hSourceFile As Long, ByVal hDestinationFile As Long, ByRef lpData As Long) As Long
Dim lngCnt As Long
Dim lngRet As Long

On Error GoTo Errores

If blnCancelaCopia Then
   lngRet = PROGRESS_CANCEL
   CopyFileCallback = lngRet
   Exit Function
End If

'If blnSegundoPlanoTemporal Then Esperar lngTEC
'Esperar 0.0001

Select Case dwCallbackReason
       Case CALLBACK_STREAM_SWITCH         'cambiaron los stream o un nuevo archivo se está copiando
            lngRet = PROGRESS_CONTINUE
       
       Case CALLBACK_CHUNK_FINISHED        'un pedazo de los datos se está copiando
            lngCnt = (TotalBytesTransferred * 100 / TotalFileSize)
            pgbProgresoCopiaCallback.Value = lngCnt

            lngRet = PROGRESS_CONTINUE
End Select

CopyFileCallback = lngRet
Exit Function

Errores:
End Function

Public Sub Vincular(ByRef Objeto1 As Object, ByRef Objeto2 As Object)
On Error Resume Next
Set Objeto2 = Objeto1
End Sub

Public Sub RomperVínculo(ByRef Objeto As Object)
On Error Resume Next
Set Objeto = Nothing
End Sub

Public Function CopiarArchivoAPI(ByVal strOrigen As String, ByVal strDestino As String) As Long
Dim SHFileOp As SHFILEOPSTRUCT
Dim lngRet As Long

strOrigen = AddRemoveSlash(strOrigen, True)
strDestino = AddRemoveSlash(strDestino, True)

strOrigen = strOrigen & Chr$(0) & Chr$(0)
strDestino = strDestino & Chr$(0) & Chr$(0)

SHFileOp.wFunc = FO_COPY
SHFileOp.pFrom = strOrigen
SHFileOp.pTo = strDestino
SHFileOp.fFlags = FOF_NOCONFIRMATION Or FOF_SILENT ' Or FOF_NOCONFIRMMKDIR 'Or FOF_NOCONFIRMATION Or FOF_SILENT

'Esperar 0.0001

lngRet = SHFileOperation(SHFileOp)
CopiarArchivoAPI = lngRet
End Function

Public Function ObtenerTipoUnidad(ByVal strPath As String) As String
Dim strUnidad As String
Dim lngRet As Long
Dim strTipo As String

ObtenerTipoUnidad = vbNullString
On Error GoTo Errores

strTipo = vbNullString

If EsRutaUNC(strPath) Then
   strTipo = "Unidad remota"
   ObtenerTipoUnidad = strTipo
   Exit Function
End If

strUnidad = LetraUnidad(strPath) & ":\"
Select Case GetDriveType(strUnidad)
       Case 2
            strTipo = "Unidad extraíble"
       Case 3
            strTipo = "Unidad de disco fijo"
       Case 4
            strTipo = "Unidad remota"
       Case 5
            strTipo = "Unidad de CD-ROM"
       Case 6
            strTipo = "Unidad de RAMDrive"
       Case Else
            strTipo = "Unidad desconocida"
End Select

ObtenerTipoUnidad = strTipo
Exit Function

Errores:
End Function

Public Function LetraUnidad(ByVal strRuta As String) As String
Dim strLetra As String
Dim strUnidad As String
Dim strSep As String

LetraUnidad = vbNullString
On Error Resume Next

If (Len(strRuta) < 2) Then Exit Function

If EsRutaUNC(strRuta) Then
   LetraUnidad = vbNullString
   Exit Function
End If

strUnidad = Left$(strRuta, 2)
strLetra = Left$(strRuta, 1)
strSep = Right$(strUnidad, 1)
If strSep = ":" Then
   Select Case LCase$(strLetra)
          Case "a" To "z"
               LetraUnidad = strLetra
   End Select
End If
End Function

Public Function ConvBytesMB(ByVal curBytes As Currency) As Currency
Dim curMB As Currency

On Error Resume Next
curMB = CSng(((curBytes) / 1024) / 1024)
ConvBytesMB = curMB
End Function

Public Function FormatMB(ByVal curBytes As Currency) As String
On Error Resume Next
FormatMB = Format$(ConvBytesMB(curBytes), "###0.00")
End Function

Public Function FormatNumero(ByVal curNum As Currency, Optional ByVal blnDecimales As Boolean = False) As String
On Error Resume Next

If blnDecimales Then
   FormatNumero = Format$(curNum, "###0.00")
Else
   FormatNumero = Format$(curNum, "##00")
End If
End Function

Public Function ObtenerEspacioLibre(ByVal strPath As String) As Double
Dim lintTamaño As LargeInt
Dim lintEspacioLibre As LargeInt
Dim lintEspacioDisponible As LargeInt
Dim lngRet As Long

ObtenerEspacioLibre = 0
On Error GoTo Errores

lngRet = GetDiskFreeSpaceEx(strPath, lintEspacioDisponible, lintTamaño, lintEspacioLibre)

If lngRet <> 0 Then ObtenerEspacioLibre = CULong(lintEspacioDisponible.HiDWord.Byte1, lintEspacioDisponible.HiDWord.Byte2, lintEspacioDisponible.HiDWord.Byte3, lintEspacioDisponible.HiDWord.Byte4) * 2 ^ 32 + CULong(lintEspacioDisponible.LoDWord.Byte1, lintEspacioDisponible.LoDWord.Byte2, lintEspacioDisponible.LoDWord.Byte3, lintEspacioDisponible.LoDWord.Byte4)
Exit Function

Errores:
End Function

Public Function ObtenerTamañoUnidad(ByVal strPath As String) As Double
Dim lintTamaño As LargeInt
Dim lintEspacioLibre As LargeInt
Dim lintEspacioDisponible As LargeInt
Dim lngRet As Long

ObtenerTamañoUnidad = 0
On Error GoTo Errores

lngRet = GetDiskFreeSpaceEx(strPath, lintEspacioDisponible, lintTamaño, lintEspacioLibre)

If lngRet <> 0 Then ObtenerTamañoUnidad = CULong(lintTamaño.HiDWord.Byte1, lintTamaño.HiDWord.Byte2, lintTamaño.HiDWord.Byte3, lintTamaño.HiDWord.Byte4) * 2 ^ 32 + CULong(lintTamaño.LoDWord.Byte1, lintTamaño.LoDWord.Byte2, lintTamaño.LoDWord.Byte3, lintTamaño.LoDWord.Byte4)
Exit Function

Errores:
End Function

Public Function CULong(Byte1 As Byte, Byte2 As Byte, Byte3 As Byte, Byte4 As Byte) As Double
CULong = Byte4 * 2 ^ 24 + Byte3 * 2 ^ 16 + Byte2 * 2 ^ 8 + Byte1
End Function
