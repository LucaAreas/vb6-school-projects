VERSION 5.00
Begin VB.Form cmdcalcular 
   Caption         =   "Calcular de Salário"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Limpar"
      Height          =   345
      Left            =   3390
      TabIndex        =   4
      Top             =   2910
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   405
      Left            =   3120
      TabIndex        =   2
      Top             =   930
      Width           =   1245
   End
   Begin VB.TextBox cmdcalcular 
      Height          =   345
      Left            =   3180
      TabIndex        =   1
      Top             =   210
      Width           =   4005
   End
   Begin VB.Label Label2 
      Caption         =   "Salário Láquido"
      Height          =   315
      Left            =   1545
      TabIndex        =   3
      Top             =   1965
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Salário Bruto"
      Height          =   345
      Left            =   1478
      TabIndex        =   0
      Top             =   240
      Width           =   1185
   End
End
Attribute VB_Name = "cmdcalcular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcalcular_Click()
 If IsNumeric(txtSalarioBruto) = True Then
    vSalarioBruto = CCur(txtSalarioBruto.Text)
    If vSalarioBruto < 1700 Then
        vSalarioLiquido = vSalarioBruto * 0.93
    Else
        Call MsgBox("Digite Somente Números")
    End If
End Sub

Private Sub Text1_Change()

End Sub

