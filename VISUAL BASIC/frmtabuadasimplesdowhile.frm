VERSION 5.00
Begin VB.Form frmtabuadasimplesdowhile 
   BackColor       =   &H00000040&
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ListBox LstTabuada 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3810
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   4695
   End
   Begin VB.TextBox txtTabuada 
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   2280
      Width           =   4695
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   2880
      Width           =   4695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Tabuada Simples Do while"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   240
      Width           =   7335
   End
   Begin VB.Label Label1 
      Caption         =   "Tabuada"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Digite a tabuada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      TabIndex        =   3
      Top             =   1320
      Width           =   4695
   End
End
Attribute VB_Name = "frmtabuadasimplesdowhile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexibir_Click()
Dim x As Integer
x = 1
lsttabuada.Clear
    
    Do While x <= 10
        lsttabuada.AddItem txtTabuada.Text & "X" & x & "=" & txtTabuada.Text * x
        x = x + 1
    Loop
    
End Sub
