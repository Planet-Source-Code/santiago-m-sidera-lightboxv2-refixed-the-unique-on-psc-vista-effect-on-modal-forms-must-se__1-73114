VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "frmMain"
   ClientHeight    =   6885
   ClientLeft      =   14265
   ClientTop       =   1500
   ClientWidth     =   5985
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   5985
   Begin MSComctlLib.Slider sldGradoBrillo 
      Height          =   615
      Left            =   960
      TabIndex        =   4
      Top             =   2280
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1085
      _Version        =   393216
      Min             =   -255
      Max             =   255
      TickStyle       =   2
      TickFrequency   =   10
      TextPosition    =   1
   End
   Begin VB.CommandButton cmdFormSecundario 
      Caption         =   "cmdFormSecundario"
      Default         =   -1  'True
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "cmdSalir"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   6240
      Width           =   1815
   End
   Begin VB.PictureBox picTemp 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.Slider sldIncrementoBrillo 
      Height          =   615
      Left            =   960
      TabIndex        =   5
      Top             =   3600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1085
      _Version        =   393216
      Max             =   15
      TickStyle       =   2
      TextPosition    =   1
   End
   Begin VB.Label lblIncrementoBrillo 
      BackStyle       =   0  'Transparent
      Caption         =   "lblIncrementoBrillo"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   3120
      Width           =   4335
   End
   Begin VB.Label lblGradoBrillo 
      BackStyle       =   0  'Transparent
      Caption         =   "lblGradoBrillo"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   1800
      Width           =   4335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /*
'  * Módulo:                          frmMain
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

Private Sub cmdFormSecundario_Click()
Load frmSecundario
Centrar frmSecundario, frmMain

frmSecundario.Caption = "Formulario secundario modal"
frmSecundario.cmdCerrar.Caption = "Cerrar"
frmSecundario.cmdCerrar.Cancel = True
frmSecundario.cmdCerrar.Default = True

MostrarFormularioModal frmMain, frmSecundario, sldGradoBrillo.Value, sldIncrementoBrillo.Value, frmMain.WindowState
End Sub

Private Sub cmdSalir_Click()
Salir
End Sub
