VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0000C000&
   Caption         =   "FORMULARIO 2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "VOLTAR"
      Height          =   705
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   3195
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call Form1.Show
    Call Me.Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form1.Show

End Sub
