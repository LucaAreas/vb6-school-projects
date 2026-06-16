VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000C0&
   Caption         =   "Tic-Tac-Toe"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      Caption         =   "Frame1"
      Height          =   8925
      Left            =   600
      TabIndex        =   4
      Top             =   270
      Width           =   9945
      Begin VB.Label lbl9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   120
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   7020
         TabIndex        =   13
         Top             =   6270
         Width           =   2715
      End
      Begin VB.Label lbl8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   120
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2475
         Left            =   3690
         TabIndex        =   12
         Top             =   6300
         Width           =   2865
      End
      Begin VB.Label lbl7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   120
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2475
         Left            =   420
         TabIndex        =   11
         Top             =   6270
         Width           =   2955
      End
      Begin VB.Label lbl6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   120
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   6990
         TabIndex        =   10
         Top             =   3300
         Width           =   2655
      End
      Begin VB.Label lbl5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   120
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2715
         Left            =   3660
         TabIndex        =   9
         Top             =   3300
         Width           =   2955
      End
      Begin VB.Label lbl4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   120
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2715
         Left            =   480
         TabIndex        =   8
         Top             =   3330
         Width           =   2955
      End
      Begin VB.Label lbl3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   120
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2865
         Left            =   6960
         TabIndex        =   7
         Top             =   210
         Width           =   2685
      End
      Begin VB.Label lbl2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   120
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2865
         Left            =   3690
         TabIndex        =   6
         Top             =   210
         Width           =   2925
      End
      Begin VB.Label lbl1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   120
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2835
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   600
      TabIndex        =   0
      Top             =   9270
      Width           =   3435
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000000C0&
      Caption         =   "Frame2"
      Height          =   2085
      Left            =   10890
      TabIndex        =   1
      Top             =   180
      Width           =   4515
      Begin VB.OptionButton joga2 
         BackColor       =   &H000000C0&
         Caption         =   "Jogador 2 (O)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   450
         TabIndex        =   21
         Top             =   1170
         Width           =   3795
      End
      Begin VB.OptionButton joga1 
         BackColor       =   &H000000C0&
         Caption         =   "Jogador 1 (X)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   420
         TabIndex        =   2
         Top             =   630
         Width           =   3825
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000C0&
         Caption         =   "Primeiro a jogar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   210
         Width           =   3765
      End
   End
   Begin VB.Label resultadoEmpate 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   12240
      TabIndex        =   20
      Top             =   7230
      Width           =   1365
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Empate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   11850
      TabIndex        =   19
      Top             =   6330
      Width           =   2145
   End
   Begin VB.Label resultadoO 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   13530
      TabIndex        =   18
      Top             =   5490
      Width           =   1365
   End
   Begin VB.Label resultadoX 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   10710
      TabIndex        =   17
      Top             =   5550
      Width           =   1365
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   13560
      TabIndex        =   16
      Top             =   4620
      Width           =   1365
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   10650
      TabIndex        =   15
      Top             =   4590
      Width           =   1365
   End
   Begin VB.Label lblPlacar 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Placar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   11280
      TabIndex        =   14
      Top             =   3450
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vControl As String

Private Sub cmdLimpar_Click()

    joga1.Enabled = True
    joga2.Enabled = True
    joga1.Value = False
    joga2.Value = False

End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

    If joga1.Value = False And joga2.Value = False Then
        lbl1.Enabled = False
        lbl2.Enabled = False
        lbl3.Enabled = False
        lbl4.Enabled = False
        lbl5.Enabled = False
        lbl6.Enabled = False
        lbl7.Enabled = False
        lbl8.Enabled = False
        lbl9.Enabled = False
    End If
End Sub

Private Sub joga1_Click()

    
    If joga1.Value = True Then
        joga1.Enabled = False
        joga2.Value = False
        joga2.Enabled = False
        joga1.Value = True
    End If
    
    vControl = "X"
    
End Sub



Private Sub joga2_Click()

    
    If joga2.Value = True Then
        joga2.Enabled = False
        joga1.Value = False
        joga1.Enabled = False
    End If
    
    vControl = "O"
    
End Sub



Private Sub lbl9_Click()

    
    If vControl = "X" Then
        lbl9 = "X"
        vControl = "O"
    Else
        lbl9 = "O"
        vControl = "X"
    End If

End Sub

Private Sub lbl1_Click()
    
    If vControl = "X" Then
        lbl1 = "X"
        vControl = "O"
    Else
        lbl1 = "O"
        vControl = "X"
    End If
        
End Sub



Private Sub lbl2_Click()

    
    If vControl = "X" Then
        lbl2 = "X"
        vControl = "O"
    Else
        lbl2 = "O"
        vControl = "X"
    End If

End Sub

Private Sub lbl3_Click()

    
    If vControl = "X" Then
        lbl3 = "X"
        vControl = "O"
    Else
        lbl3 = "O"
        vControl = "X"
    End If

End Sub

Private Sub lbl4_Click()

    
     If vControl = "X" Then
        lbl4 = "X"
        vControl = "O"
    Else
        lbl4 = "O"
        vControl = "X"
    End If

End Sub

Private Sub lbl5_Click()

    
    If vControl = "X" Then
        lbl5 = "X"
        vControl = "O"
    Else
        lbl5 = "O"
        vControl = "X"
    End If
End Sub

Private Sub lbl6_Click()

    
    If vControl = "X" Then
        lbl6 = "X"
        vControl = "O"
    Else
        lbl6 = "O"
        vControl = "X"
    End If

End Sub

Private Sub lbl7_Click()

    
    If vControl = "X" Then
        lbl7 = "X"
        vControl = "O"
    Else
        lbl7 = "O"
        vControl = "X"
    End If

End Sub

Private Sub lbl8_Click()

    If vControl = "X" Then
        lbl8 = "X"
        vControl = "O"
    Else
        lbl8 = "O"
        vControl = "X"
    End If

End Sub
