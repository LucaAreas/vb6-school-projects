VERSION 5.00
Begin VB.Form frmVendas 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13410
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   13410
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbProduto1 
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   2040
      Width           =   1995
   End
   Begin VB.TextBox txtQtd5 
      Height          =   495
      Left            =   6960
      TabIndex        =   22
      Top             =   5730
      Width           =   1995
   End
   Begin VB.ComboBox cmbProduto5 
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   5730
      Width           =   1995
   End
   Begin VB.TextBox txtQtd4 
      Height          =   495
      Left            =   6960
      TabIndex        =   18
      Top             =   4680
      Width           =   1995
   End
   Begin VB.ComboBox cmbProduto4 
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   4680
      Width           =   1995
   End
   Begin VB.TextBox txtQtd3 
      Height          =   495
      Left            =   6960
      TabIndex        =   14
      Top             =   3810
      Width           =   1995
   End
   Begin VB.ComboBox cmbProduto3 
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3810
      Width           =   1995
   End
   Begin VB.TextBox txtQtd2 
      Height          =   495
      Left            =   6960
      TabIndex        =   10
      Top             =   2850
      Width           =   1995
   End
   Begin VB.ComboBox cmbProduto2 
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2850
      Width           =   1995
   End
   Begin VB.TextBox txtQtd1 
      Height          =   495
      Left            =   6960
      TabIndex        =   5
      Top             =   2010
      Width           =   1995
   End
   Begin VB.Label lblValortotal5 
      Height          =   495
      Left            =   9420
      TabIndex        =   24
      Top             =   5730
      Width           =   1995
   End
   Begin VB.Label lblValorunitario5 
      Height          =   495
      Left            =   4470
      TabIndex        =   23
      Top             =   5730
      Width           =   1995
   End
   Begin VB.Label lblValortotal4 
      Height          =   495
      Left            =   9420
      TabIndex        =   20
      Top             =   4680
      Width           =   1995
   End
   Begin VB.Label lblValorunitario4 
      Height          =   495
      Left            =   4470
      TabIndex        =   19
      Top             =   4680
      Width           =   1995
   End
   Begin VB.Label lblValortotal3 
      Height          =   495
      Left            =   9420
      TabIndex        =   16
      Top             =   3810
      Width           =   1995
   End
   Begin VB.Label lblValorunitario3 
      Height          =   495
      Left            =   4470
      TabIndex        =   15
      Top             =   3810
      Width           =   1995
   End
   Begin VB.Label lblValortotal2 
      Height          =   495
      Left            =   9420
      TabIndex        =   12
      Top             =   2850
      Width           =   1995
   End
   Begin VB.Label lblValorunitario2 
      Height          =   495
      Left            =   4470
      TabIndex        =   11
      Top             =   2850
      Width           =   1995
   End
   Begin VB.Label lblSoma 
      Caption         =   "0"
      Height          =   495
      Left            =   9420
      TabIndex        =   8
      Top             =   6960
      Width           =   1995
   End
   Begin VB.Label lblValortotal1 
      Height          =   495
      Left            =   9420
      TabIndex        =   7
      Top             =   2010
      Width           =   1995
   End
   Begin VB.Label lblValorunitario1 
      Height          =   495
      Left            =   4470
      TabIndex        =   6
      Top             =   2010
      Width           =   1995
   End
   Begin VB.Label Label8 
      Caption         =   "Soma"
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   6960
      Width           =   1995
   End
   Begin VB.Label Label6 
      Caption         =   "Valor total"
      Height          =   495
      Left            =   9420
      TabIndex        =   3
      Top             =   990
      Width           =   1995
   End
   Begin VB.Label lblQuantidade 
      Caption         =   "Qtd"
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   990
      Width           =   1995
   End
   Begin VB.Label lbl 
      Caption         =   "Valor unitário"
      Height          =   495
      Left            =   4470
      TabIndex        =   1
      Top             =   990
      Width           =   1995
   End
   Begin VB.Label lblProduto 
      Caption         =   "Produto"
      Height          =   495
      Left            =   1860
      TabIndex        =   0
      Top             =   990
      Width           =   1995
   End
End
Attribute VB_Name = "frmVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbProduto1_GotFocus()
    
    cmbProduto1.BackColor = &HFFFF&
    
End Sub

Private Sub cmbProduto1_LostFocus()
    
    cmbProduto1.BackColor = &H80000005
    
End Sub

Private Sub cmbProduto2_GotFocus()
    
    cmbProduto2.BackColor = &HFFFF&
    
End Sub

Private Sub cmbProduto2_LostFocus()
    
    cmbProduto2.BackColor = &H80000005
    
End Sub

Private Sub cmbProduto3_GotFocus()
    
    cmbProduto3.BackColor = &HFFFF&
    
End Sub

Private Sub cmbProduto3_LostFocus()
    
    cmbProduto3.BackColor = &H80000005
    
End Sub

Private Sub cmbProduto4_GotFocus()
    
    cmbProduto4.BackColor = &HFFFF&
    
End Sub

Private Sub cmbProduto4_LostFocus()
    
    cmbProduto4.BackColor = &H80000005
    
End Sub

Private Sub cmbProduto5_GotFocus()
    
    cmbProduto5.BackColor = &HFFFF&
    
End Sub

Private Sub cmbProduto5_LostFocus()
    
    cmbProduto5.BackColor = &H80000005
    
End Sub

Private Sub cmbProduto1_Click()
    
    If cmbProduto1.Text = "Banana" Then
        lblValorunitario1.Caption = "1,20"
    ElseIf cmbProduto1.Text = "Uva" Then
        lblValorunitario1.Caption = "3,00"
    ElseIf cmbProduto1.Text = "Abacate" Then
        lblValorunitario1.Caption = "2,50"
    ElseIf cmbProduto1.Text = "Melão" Then
        lblValorunitario1.Caption = "5,00"
    ElseIf cmbProduto1.Text = "Caqui" Then
        lblValorunitario1.Caption = "1,10"
    End If

    If txtQtd1 <> Empty Then
        If IsNumeric(txtQtd1.Text) Then
            lblValortotal1 = txtQtd1.Text * CDbl(lblValorunitario1)
            lblSoma.Caption = "R$" & CCur(lblValortotal1.Caption) + lblSoma.Caption
        Else
            lblValortotal1.Caption = ""
            Call MsgBox("Digite um número!")
        End If
    Else
        lblValortotal1 = ""
    End If
    
End Sub

Private Sub txtQtd1_Change()
    
    If txtQtd1 <> Empty Then
        If IsNumeric(txtQtd1.Text) Then
            lblValortotal1 = txtQtd1.Text * CDbl(lblValorunitario1)
            lblSoma.Caption = "R$" & CCur(lblValortotal1.Caption) + lblSoma.Caption
        Else
            lblValortotal1.Caption = ""
            Call MsgBox("Digite um número!")
        End If
    Else
        lblValortotal1 = ""
    End If
    
End Sub

Private Sub cmbProduto2_Click()
    
    If cmbProduto2.Text = "Banana" Then
        lblValorunitario2.Caption = "1,20"
    ElseIf cmbProduto2.Text = "Uva" Then
        lblValorunitario2.Caption = "3,00"
    ElseIf cmbProduto2.Text = "Abacate" Then
        lblValorunitario2.Caption = "2,50"
    ElseIf cmbProduto2.Text = "Melão" Then
        lblValorunitario2.Caption = "5,00"
    ElseIf cmbProduto2.Text = "Caqui" Then
        lblValorunitario2.Caption = "1,10"
    End If
    
    If txtQtd2 <> Empty Then
        If IsNumeric(txtQtd2.Text) Then
            lblValortotal2 = txtQtd2.Text * CDbl(lblValorunitario2)
            lblSoma.Caption = "R$" & CCur(lblValortotal2.Caption) + lblSoma.Caption
        Else
            lblValortotal2.Caption = ""
            Call MsgBox("Digite um número!")
        End If
    Else
        lblValortotal2 = ""
    End If
    
End Sub

Private Sub txtQtd2_Change()
    
    If txtQtd2 <> Empty Then
        If IsNumeric(txtQtd2.Text) Then
            lblValortotal2 = txtQtd2.Text * CDbl(lblValorunitario2)
            lblSoma.Caption = "R$" & CCur(lblValortotal2.Caption) + lblSoma.Caption
        Else
            lblValortotal2.Caption = ""
            Call MsgBox("Digite um número!")
        End If
    Else
        lblValortotal2 = ""
    End If
    
End Sub

Private Sub cmbProduto3_Click()
    
    If cmbProduto3.Text = "Banana" Then
        lblValorunitario3.Caption = "1,20"
    ElseIf cmbProduto3.Text = "Uva" Then
        lblValorunitario3.Caption = "3,00"
    ElseIf cmbProduto3.Text = "Abacate" Then
        lblValorunitario3.Caption = "2,50"
    ElseIf cmbProduto3.Text = "Melão" Then
        lblValorunitario3.Caption = "5,00"
    ElseIf cmbProduto3.Text = "Caqui" Then
        lblValorunitario3.Caption = "1,10"
    End If
    
    If txtQtd3 <> Empty Then
        If IsNumeric(txtQtd3.Text) Then
            lblValortotal3 = txtQtd3.Text * CDbl(lblValorunitario3)
            lblSoma.Caption = "R$" & CCur(lblValortotal3.Caption) + lblSoma.Caption
        Else
            lblValortotal3.Caption = ""
            Call MsgBox("Digite um número!")
        End If
    Else
        lblValortotal3 = ""
    End If
    
End Sub

Private Sub txtQtd3_Change()
    
    If txtQtd3 <> Empty Then
        If IsNumeric(txtQtd3.Text) Then
            lblValortotal3 = txtQtd3.Text * CDbl(lblValorunitario3)
            lblSoma.Caption = "R$" & CCur(lblValortotal3.Caption) + lblSoma.Caption
        Else
            lblValortotal3.Caption = ""
            Call MsgBox("Digite um número!")
        End If
    Else
        lblValortotal3 = ""
    End If
      
End Sub

Private Sub cmbProduto4_Click()
    
    If cmbProduto4.Text = "Banana" Then
        lblValorunitario4.Caption = "1,20"
    ElseIf cmbProduto4.Text = "Uva" Then
        lblValorunitario4.Caption = "3,00"
    ElseIf cmbProduto4.Text = "Abacate" Then
        lblValorunitario4.Caption = "2,50"
    ElseIf cmbProduto4.Text = "Melão" Then
        lblValorunitario4.Caption = "5,00"
    ElseIf cmbProduto4.Text = "Caqui" Then
        lblValorunitario4.Caption = "1,10"
    End If
    
    If txtQtd4 <> Empty Then
        If IsNumeric(txtQtd4.Text) Then
            lblValortotal4 = txtQtd4.Text * CDbl(lblValorunitario4)
            lblSoma.Caption = "R$" & CCur(lblValortotal4.Caption) + lblSoma.Caption
        Else
            lblValortotal4.Caption = ""
            Call MsgBox("Digite um número!")
        End If
    Else
        lblValortotal4 = ""
    End If

End Sub
Private Sub txtQtd4_change()
    
    If txtQtd4 <> Empty Then
        If IsNumeric(txtQtd4.Text) Then
            lblValortotal4 = txtQtd4.Text * CDbl(lblValorunitario4)
            lblSoma.Caption = "R$" & CCur(lblValortotal4.Caption) + lblSoma.Caption
        Else
            lblValortotal4.Caption = ""
            Call MsgBox("Digite um número!")
        End If
    Else
        lblValortotal4 = ""
    End If
     
End Sub

Private Sub cmbProduto5_Click()
    
    If cmbProduto5.Text = "Banana" Then
        lblValorunitario5.Caption = "1,20"
    ElseIf cmbProduto5.Text = "Uva" Then
        lblValorunitario5.Caption = "3,00"
    ElseIf cmbProduto5.Text = "Abacate" Then
        lblValorunitario5.Caption = "2,50"
    ElseIf cmbProduto5.Text = "Melão" Then
        lblValorunitario5.Caption = "5,00"
    ElseIf cmbProduto5.Text = "Caqui" Then
        lblValorunitario5.Caption = "1,10"
    End If
    
    If txtQtd5 <> Empty Then
        If IsNumeric(txtQtd5.Text) Then
            lblValortotal5 = txtQtd5.Text * CDbl(lblValorunitario5)
            lblSoma.Caption = "R$" & CCur(lblValortotal5.Caption) + lblSoma.Caption
        Else
            lblValortotal5.Caption = ""
            Call MsgBox("Digite um número!")
        End If
    Else
        lblValortotal5 = ""
    End If

End Sub
Private Sub txtQtd5_change()
    
    If txtQtd5 <> Empty Then
        If IsNumeric(txtQtd5.Text) Then
            lblValortotal5 = txtQtd5.Text * CDbl(lblValorunitario5)
            lblSoma.Caption = "R$" & CCur(lblValortotal5.Caption) + lblSoma.Caption
        Else
            lblValortotal5.Caption = ""
            Call MsgBox("Digite um número!")
        End If
    Else
        lblValortotal5 = ""
    End If
      
End Sub

Private Sub Form_Load()
    
    'Carrega os podutos na combo Produtos 2
    cmbProduto2.AddItem "Banana"
    cmbProduto2.AddItem "Uva"
    cmbProduto2.AddItem "Abacate"
    cmbProduto2.AddItem "Melão"
    cmbProduto2.AddItem "Caqui"
    
    'Carrega os podutos na combo Produtos 3
    cmbProduto3.AddItem "Banana"
    cmbProduto3.AddItem "Uva"
    cmbProduto3.AddItem "Abacate"
    cmbProduto3.AddItem "Melão"
    cmbProduto3.AddItem "Caqui"
    
    'Carrega os podutos na combo Produtos 4
    cmbProduto4.AddItem "Banana"
    cmbProduto4.AddItem "Uva"
    cmbProduto4.AddItem "Abacate"
    cmbProduto4.AddItem "Melão"
    cmbProduto4.AddItem "Caqui"
    
    'Carrega os podutos na combo Produtos 5
    cmbProduto5.AddItem "Banana"
    cmbProduto5.AddItem "Uva"
    cmbProduto5.AddItem "Abacate"
    cmbProduto5.AddItem "Melão"
    cmbProduto5.AddItem "Caqui"
    
End Sub

Private Sub txtQtd1_GotFocus()
    
    txtQtd1.BackColor = &HFFFF&
    
End Sub

Private Sub txtQtd1_LostFocus()
    
    txtQtd1.BackColor = &H80000005
    
End Sub

Private Sub txtQtd2_GotFocus()

    txtQtd2.BackColor = &HFFFF&
    
End Sub

Private Sub txtQtd2_LostFocus()
    
    txtQtd2.BackColor = &H80000005
    
End Sub

Private Sub txtQtd3_GotFocus()

    txtQtd3.BackColor = &HFFFF&
    
End Sub

Private Sub txtQtd3_LostFocus()
    
    txtQtd3.BackColor = &H80000005
    
End Sub

Private Sub txtQtd4_GotFocus()

    txtQtd4.BackColor = &HFFFF&
    
End Sub


Private Sub txtQtd4_LostFocus()
    
    txtQtd4.BackColor = &H80000005
    
End Sub

Private Sub txtQtd5_GotFocus()

    txtQtd5.BackColor = &HFFFF&
    
End Sub


Private Sub txtQtd5_LostFocus()
    
    txtQtd5.BackColor = &H80000005
    
End Sub

