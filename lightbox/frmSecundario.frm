VERSION 5.00
Begin VB.Form frmSecundario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmSecundario"
   ClientHeight    =   3750
   ClientLeft      =   14250
   ClientTop       =   3030
   ClientWidth     =   4605
   Icon            =   "frmSecundario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4605
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "cmdCerrar"
      Default         =   -1  'True
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSecundario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /*
'  * Módulo:                          frmSecundario
'  * Tipo de módulo:                  Formulario
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

Private Sub cmdCerrar_Click()
Unload Me
End Sub
