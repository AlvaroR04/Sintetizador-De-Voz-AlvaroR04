VERSION 5.00
Begin VB.Form Sintetizador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sintetizador"
   ClientHeight    =   12780
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   14610
   Icon            =   "Sintetizador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12780
   ScaleWidth      =   14610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Texto 
      Appearance      =   0  'Flat
      Height          =   12255
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   14655
   End
   Begin VB.CommandButton Read 
      Caption         =   "Leer"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   12360
      Width           =   14655
   End
   Begin VB.Menu About 
      Caption         =   "Sobre sintetizador"
   End
End
Attribute VB_Name = "Sintetizador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim voz

Private Sub About_Click()
    MsgBox "AlvaroR04's Synthesizer versión: " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & _
    "El programa está en una fase beta, y este puede congelarse mientras se realiza la lectura en voz alta hasta que termine. Se recomienda discreción."
End Sub

Private Sub Form_Load()
    Set voz = CreateObject("SAPI.SpVoice")
    voz.Rate = 0
    Texto.Width = Me.Width
End Sub

Private Sub Read_Click()
    voz.speak Texto.Text
End Sub
