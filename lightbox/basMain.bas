Attribute VB_Name = "basMain"
' /*
'  * Módulo:                          basMain
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

Public strAppPath As String

Public Sub Main()
strAppPath = AddRemoveSlash(App.Path)
ExtraerRecurso "lightbox", "dll", strAppPath

Load frmMain
Centrar frmMain

frmMain.Caption = "Formulario principal"
frmMain.cmdFormSecundario.Caption = "Mostrar form secundario"
frmMain.cmdFormSecundario.Default = True
frmMain.cmdSalir.Caption = "Salir"
frmMain.cmdSalir.Cancel = True

frmMain.lblGradoBrillo.Caption = "Grado del brillo del formulario cuando esté en segundo plano"
frmMain.sldGradoBrillo.TickFrequency = 10
frmMain.sldGradoBrillo.TickStyle = sldBoth
frmMain.sldGradoBrillo.Min = -255
frmMain.sldGradoBrillo.Max = 255
frmMain.sldGradoBrillo.Value = -150

frmMain.lblIncrementoBrillo.Caption = "Unidades de incremento de brillo"
frmMain.sldIncrementoBrillo.TickFrequency = 1
frmMain.sldIncrementoBrillo.TickStyle = sldBoth
frmMain.sldIncrementoBrillo.Min = 0
frmMain.sldIncrementoBrillo.Max = 18
frmMain.sldIncrementoBrillo.Value = 10

frmMain.Show
End Sub

Public Sub Salir()
End
End Sub
