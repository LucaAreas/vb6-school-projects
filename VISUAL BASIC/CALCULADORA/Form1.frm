VERSION 5.00
Begin VB.Form Calculadora 
   BackColor       =   &H8000000D&
   Caption         =   "Calculadora"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdResultado 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   4560
      TabIndex        =   16
      Top             =   7200
      Width           =   885
   End
   Begin VB.CommandButton cmdDivisao 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   4500
      TabIndex        =   15
      Top             =   6030
      Width           =   885
   End
   Begin VB.CommandButton cmdVezes 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   4500
      TabIndex        =   14
      Top             =   4740
      Width           =   885
   End
   Begin VB.CommandButton cmdMenos 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   4500
      TabIndex        =   13
      Top             =   3540
      Width           =   885
   End
   Begin VB.CommandButton cmdMais 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   4500
      TabIndex        =   12
      Top             =   2280
      Width           =   885
   End
   Begin VB.CommandButton cmdDecimal 
      Caption         =   ","
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   690
      TabIndex        =   11
      Top             =   5940
      Width           =   885
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3150
      TabIndex        =   10
      Top             =   6000
      Width           =   885
   End
   Begin VB.CommandButton cmdNum0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1860
      TabIndex        =   9
      Top             =   5970
      Width           =   885
   End
   Begin VB.CommandButton cmdNum3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3150
      TabIndex        =   8
      Top             =   4710
      Width           =   885
   End
   Begin VB.CommandButton cmdNum2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1860
      TabIndex        =   7
      Top             =   4770
      Width           =   885
   End
   Begin VB.CommandButton cmdNum1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   690
      TabIndex        =   6
      Top             =   4740
      Width           =   885
   End
   Begin VB.CommandButton cmdNum6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3150
      TabIndex        =   5
      Top             =   3570
      Width           =   885
   End
   Begin VB.CommandButton cmdNum8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1860
      TabIndex        =   4
      Top             =   2250
      Width           =   885
   End
   Begin VB.CommandButton cmdNum9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3150
      TabIndex        =   3
      Top             =   2250
      Width           =   885
   End
   Begin VB.CommandButton cmdNum4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   720
      TabIndex        =   2
      Top             =   3540
      Width           =   885
   End
   Begin VB.CommandButton cmdNum5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1860
      TabIndex        =   1
      Top             =   3600
      Width           =   885
   End
   Begin VB.CommandButton cmdNum7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   690
      TabIndex        =   0
      Top             =   2250
      Width           =   885
   End
   Begin VB.Label lblResultado 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   630
      TabIndex        =   17
      Top             =   600
      Width           =   4755
   End
End
Attribute VB_Name = "Calculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim num1, num2 As Double
Dim operacao As String


Private Sub cmdDecimal_Click()

    If InStr(1, lblResultado.Caption, ",") = 0 Then
        lblResultado.Caption = lblResultado.Caption & ","
    End If

End Sub

Private Sub cmdDivisao_Click()
num1 = Val(lblResultado.Caption)
        operacao = "/"
        lblResultado.Caption = ""
End Sub

Private Sub cmdLimpar_Click()

    lblResultado.Caption = ""

End Sub

Private Sub cmdMais_Click()

    num1 = Val(lblResultado.Caption)
        operacao = "+"
        lblResultado.Caption = ""

End Sub

Private Sub cmdMenos_Click()
num1 = Val(lblResultado.Caption)
        operacao = "-"
        lblResultado.Caption = ""
End Sub

Private Sub cmdNum7_Click()
lblResultado.Caption = lblResultado.Caption & 7
End Sub

Private Sub cmdNum0_Click()
lblResultado.Caption = lblResultado.Caption & 0


End Sub

Private Sub cmdNum5_Click()
lblResultado.Caption = lblResultado.Caption & 5
End Sub

Private Sub cmdNum4_Click()
lblResultado.Caption = lblResultado.Caption & 4
End Sub

Private Sub cmdNum9_Click()
lblResultado.Caption = lblResultado.Caption & 9
End Sub

Private Sub cmdNum8_Click()
lblResultado.Caption = lblResultado.Caption & 8
End Sub

Private Sub cmdNum6_Click()
lblResultado.Caption = lblResultado.Caption & 6
End Sub

Private Sub cmdNum1_Click()
lblResultado.Caption = lblResultado.Caption & 1
End Sub

Private Sub cmdNum2_Click()
lblResultado.Caption = lblResultado.Caption & 2
End Sub

Private Sub cmdNum3_Click()
lblResultado.Caption = lblResultado.Caption & 3
End Sub

Private Sub cmdResultado_Click()

    num1 = CDbl(lblResultado.Caption)
    num2 = CDbl(lblResultado.Caption)
        Select Case operacao
        Case "+": lblResultado.Caption = num1 + num2
          Case "-": lblResultado.Caption = num1 - num2
            Case "/": lblResultado.Caption = num1 / num2
              Case "X": lblResultado.Caption = num1 * num2
        End Select
End Sub

Private Sub cmdVezes_Click()


    num1 = Val(lblResultado.Caption)
        operacao = "X"
        lblResultado.Caption = ""

End Sub
