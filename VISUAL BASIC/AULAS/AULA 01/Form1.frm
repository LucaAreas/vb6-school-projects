VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   1245
      Left            =   1950
      TabIndex        =   2
      Top             =   1830
      Width           =   2625
   End
   Begin VB.ComboBox cboNumero 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   3390
      List            =   "Form1.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   210
      Width           =   2745
   End
   Begin VB.Label lblNumero 
      Caption         =   "Número"
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   210
      Width           =   2745
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExibir_Click()

    Select Case cboNumero.Text
        Case "10"
            Call MsgBox("Palmeiras tem mundial?")
        Case "20"
            Call MsgBox("Viva o Vasco da Gama!")
        Case "30"
            Call MsgBox("Bom dia seus vagabundos!")
        Case "40"
            Call MsgBox("Receba!")
        Case "50"
            Call MsgBox("Ooohhhh ma gawd!")
    End Select
    
End Sub
