VERSION 5.00
Begin VB.Form frmCirculosFord 
   Caption         =   "Círculos de Ford"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   13995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCirculosFord 
      Caption         =   "Círculos Ford"
      Height          =   495
      Left            =   12600
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmCirculosFord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim miAncho As Long
Dim miAlto As Long
Dim miInicio As Long
Dim miFin As Long
Dim miCentroX As Long
Dim miCentroY As Long
Dim miRadio As Long

Private Sub cmdCirculosFord_Click()
' Inicializacion de variables
  miAncho = 7000
  miAlto = 7000
  miInicio = 3500
  miFin = miInicio + miAncho
  ' Línea de base
  Line (miInicio - 3500, miAlto)-(miFin + 3500, miAlto)
  ' Círculo 0/1
  miRadio = miAncho / 2
  miCentroY = miAlto - miRadio
  miCentroX = miInicio
  Circle (miCentroX, miCentroY), miRadio
  ' Círculo 1/1
  miRadio = miAncho / 2
  miCentroY = miAlto - miRadio
  miCentroX = miFin
  Circle (miCentroX, miCentroY), miRadio
  ' Círculo 1/2
  miRadio = miAncho / 8
  miCentroY = miAlto - miRadio
  miCentroX = (miInicio + miFin) / 2
  Circle (miCentroX, miCentroY), miRadio
End Sub
