VERSION 5.00
Begin VB.Form frmjuncao 
   BackColor       =   &H00000000&
   Caption         =   "Form3"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtinicio 
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ListBox lsttabuada 
      Height          =   2010
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtfim 
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdexibir 
      Caption         =   "Exibir"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Tabuada 
      Caption         =   "Tabuada"
      Height          =   975
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Digite a tabuada"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "frmjuncao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexibir_Click()
Dim x As Integer
For y = txtinicio.Text To txtfim.Text
    For x = 1 To 10
      lsttabuada.AddItem y & "x" & x & "=" & y * x
    Next
    lsttabuada.AddItem "-------------"
Next
End Sub
