VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000FFFF&
   Caption         =   "FORMULARIO INICIAL"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "ABRIR FORMULARIO2"
      Height          =   945
      Left            =   1020
      TabIndex        =   0
      Top             =   750
      Width           =   2475
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call Form2.Show
    Call Me.Hide

End Sub

