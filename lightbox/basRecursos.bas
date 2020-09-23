Attribute VB_Name = "basRecursos"
' /*
'  * Módulo:                          basRecursos
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

Public Sub ExtraerRecurso(ByVal strNombreRecurso As String, ByVal strExtensiónArchivo As String, Optional ByVal strRutaOpcional As String = vbNullString, Optional ByVal blnRegistrar As Boolean = False)
Dim intFF As Integer
Dim bytBuff() As Byte
Dim strArchivo As String

If (strRutaOpcional <> vbNullString) Then
   strArchivo = AddRemoveSlash(strRutaOpcional) & strNombreRecurso & "." & strExtensiónArchivo
Else
   strArchivo = AddRemoveSlash(App.Path) & strNombreRecurso & "." & strExtensiónArchivo
End If

If (Not Existe(strArchivo)) Then
   bytBuff() = LoadResData(strNombreRecurso, "CUSTOM")
   intFF = FreeFile()
   
   If (Not Existe(GetFileDir(strArchivo))) Then CrearDirectorio GetFileDir(strArchivo)
   
   'EscribirLOG "*** " & "extrayendo recurso " & strArchivo
   Open strArchivo For Binary As #intFF
        Put #intFF, , bytBuff()
   Close #intFF
   
   'If blnRegistrar Then
   '   'EscribirLOG "*** " & "registrando dll " & strArchivo
   '   'If (Not RegistrarDLL(strArchivo)) Then
   '   '   'EscribirLOG "*** " & "error al extraer recurso dll y registrarlo en DLLRegisterServer."
   '   '   MsgBox "Error al extraer recurso DLL y registrarlo en DLLRegisterServer.", vbOKOnly Or vbCritical, "Error"
   '   'End If
   'End If
End If
End Sub
