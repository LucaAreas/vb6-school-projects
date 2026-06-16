VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalcularvText1 
      Caption         =   "Calcular"
      Height          =   1155
      Left            =   1380
      TabIndex        =   3
      Top             =   4170
      Width           =   5595
   End
   Begin VB.TextBox txtvNumero2 
      Alignment       =   2  'Center
      Height          =   1365
      Left            =   3090
      TabIndex        =   2
      Top             =   1380
      Width           =   3465
   End
   Begin VB.TextBox txtvNumero1 
      Alignment       =   2  'Center
      Height          =   1065
      Left            =   3090
      TabIndex        =   1
      Top             =   150
      Width           =   3465
   End
   Begin VB.Label Label5 
      Caption         =   "RESULTADO"
      Height          =   915
      Left            =   1410
      TabIndex        =   6
      Top             =   3000
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "N2"
      Height          =   1335
      Left            =   1410
      TabIndex        =   5
      Top             =   1350
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "N1"
      Height          =   1035
      Left            =   1410
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblValorResultado 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   765
      Left            =   3120
      TabIndex        =   0
      Top             =   3090
      Width           =   3435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalcular_Click()

End Sub
Private Sub cmdCalcularvText1_Click()

Dim vNumero1 As Double
Dim vNumero2 As Double

If IsNumeric(Trim(txtvNumero1)) Then
    If IsNumeric(Trim(txtvNumero2)) Then
        vNumero1 = CDbl(Trim(txtvNumero1))
        vNumero2 = CDbl(Trim(txtvNumero2))
        lblValorResultado.Caption = vNumero1 + vNumero2
    Else
        Call MsgBox("Digite um valor numérico no número 2.")
        Call txtvNumero1.SetFocus
    End If
Else
    Call MsgBox("Digite um valor númerico no número 1.")
    Call txtvNumero2.SetFocus
End If

End Sub

Private Sub iblValorResultado_Click()

End Sub

Private Sub txtText1_Change()




End Sub

Private Sub txtText2_Change()

End Sub

