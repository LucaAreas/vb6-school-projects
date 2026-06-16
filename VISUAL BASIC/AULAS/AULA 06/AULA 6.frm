VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHora 
      Height          =   585
      Left            =   5670
      TabIndex        =   2
      Top             =   1230
      Width           =   1575
   End
   Begin VB.CommandButton cmdcalcular 
      Caption         =   "EXIBIR HORA"
      Height          =   705
      Left            =   3780
      TabIndex        =   0
      Top             =   2100
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "HORARIO"
      Height          =   315
      Left            =   2070
      TabIndex        =   1
      Top             =   1350
      Width           =   2505
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdcalcular_Click()

    If IsNumeric(txtHora.Text) = True Then
        If CInt(txtHora.Text) >= 0 Then
        If CInt(txtHora.Text) <= 23 Then
        If CInt(txtHora.Text) < 12 Then
        Call MsgBox("Bom dia!")
        Else
            If CInt(txtHora.Text) < 18 Then
            
            Call MsgBox("Boa tarde")
            Else
            Call MsgBox("Boa noite")
                End If
             End If
       Else
        Call MsgBox("Não é permitido valor acima de 23.")
        txtHora.SetFocus
        
        End If
    Else
        Call MsgBox("não é possivel valor abaixo de 0")
        txtHora.SetFocus
        
    End If
    Else
        Call MsgBox("Digite somente horas numéricas de 0 a 23")
    End If
    
End Sub

