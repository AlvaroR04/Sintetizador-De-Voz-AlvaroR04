VERSION 5.00
Begin VB.Form Sintetizador 
   Caption         =   "Sintetizador"
   ClientHeight    =   6840
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   14610
   Icon            =   "Sintetizador.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   14610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Texto 
      Appearance      =   0  'Flat
      Height          =   6855
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   14655
   End
   Begin VB.Menu Leer 
      Caption         =   "Leer"
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
    Texto.Height = Me.Height
End Sub

Private Sub Form_Resize()
    Texto.Width = Me.Width
    Texto.Height = Me.Height
End Sub

Private Sub Leer_Click()
    voz.speak Texto.Text
End Sub
