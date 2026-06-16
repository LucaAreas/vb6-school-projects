VERSION 5.00
Begin VB.Form frmfaixatabuadadowhile 
   BackColor       =   &H00808000&
   Caption         =   "Form4"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexibir 
      Caption         =   "Exibir"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox txtfim 
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   2880
      Width           =   1695
   End
   Begin VB.ListBox lsttabuada 
      Height          =   2010
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtinicio 
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Digite a tabuada"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Tabuada 
      Caption         =   "Tabuada"
      Height          =   975
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "frmfaixatabuadadowhile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
