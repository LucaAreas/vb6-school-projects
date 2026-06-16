VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Limpar"
      Height          =   555
      Left            =   840
      TabIndex        =   4
      Top             =   3780
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   555
      Left            =   840
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   555
      Left            =   840
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1950
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   555
      Left            =   840
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1050
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   555
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   270
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Call LimparTela
    

End Sub

Private Sub Form_Load()
    
    Call LimparTela
    
End Sub

Private Sub LimparTela()

    Text1.Text = Empty
    Text2.Text = Empty
    Text3.Text = Empty
    Text4.Text = Empty

End Sub
