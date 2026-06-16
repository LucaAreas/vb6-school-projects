VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2490
      TabIndex        =   1
      Top             =   2940
      Width           =   2385
   End
   Begin VB.ComboBox cboNumero 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      ItemData        =   "Form1.frx":0000
      Left            =   3660
      List            =   "Form1.frx":0010
      TabIndex        =   0
      Top             =   660
      Width           =   2385
   End
   Begin VB.Label label 
      Caption         =   "Número"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1650
      TabIndex        =   2
      Top             =   720
      Width           =   1875
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdExibir_Click()
    Select Case cboNumero.Text
     Case "10"
        Call MsgBox("lol", vbCritical, "oba")
             Case "20"
        Call MsgBox("loko")
             Case "30"
        Call MsgBox("SIUUUUUUUUU", vbQuestion, "oi ")
        
             Case "40"
        Call MsgBox("PALMEIRASSSS")
        End Select
End Sub

