VERSION 5.00
Begin VB.Form txtvSalarioBruto 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalcularSalarioLiquido 
      Caption         =   "CALCULAR SALARIO LÍQUIDO"
      Height          =   645
      Left            =   3600
      TabIndex        =   3
      Top             =   3330
      Width           =   4635
   End
   Begin VB.TextBox txtvSalarioBruto 
      Height          =   765
      Left            =   5940
      TabIndex        =   1
      Top             =   1110
      Width           =   2085
   End
   Begin VB.Label IblvSalarioLiquido 
      BorderStyle     =   1  'Fixed Single
      Height          =   675
      Left            =   5910
      TabIndex        =   4
      Top             =   2220
      Width           =   2145
   End
   Begin VB.Label Label2 
      Caption         =   "SALÁRIO LIQUIDO"
      Height          =   675
      Left            =   3570
      TabIndex        =   2
      Top             =   2220
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "SALARIO BRUTO"
      Height          =   765
      Left            =   3570
      TabIndex        =   0
      Top             =   1110
      Width           =   2085
   End
End
Attribute VB_Name = "txtvSalarioBruto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcalcular_Click()

End Sub

Private Sub cmdCalcularSalarioLiquido_Click()
    
    Dim vSalarioBruto       As Double
    Dim vSalarioLiquido     As Double
    
        If IsNumeric(txtvSalarioBruto.Text) Then
            vSalarioBruto = CDbl(Trim(txtvSalarioBruto.Text))
            If vSalarioBruto > 0 Then
                If vSalarioBruto < 1700 Then
                    vSalarioLiquido = vSalarioBruto * 0.93
                Else
                    If vSalarioBruto <= 2400 Then
                        vSalarioLiquido = vSalarioBruto * 0.92
                    Else
                    
                        If vSalarioBruto <= 3200 Then
                            vSalarioLiquido = vSalarioBruto * 0.91
                        Else
                            vSalarioLiquido = vSalarioBruto * 0.89
                        End If
                    End If
                End If
                IblvSalarioLiquido.Caption = Format(vSalarioLiquido, "R$ #.00")
            Else
                Call MsgBox("Digite um valor maior que zero")
                Call txtSalarioBruto.SetFocus
            End If
        Else
            Call MsgBox("Digite somente um valor numérico.")
            Call txtvSalarioBruto.SetFocus
        End If
        
End Sub

Private Sub lblValorSalarioLiquido_Click()

End Sub

Private Sub vSalarioLiquido_Click()

End Sub

