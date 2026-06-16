VERSION 5.00
Begin VB.Form frmtabuadasimplesfornext 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexibir 
      Caption         =   "Exibir"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   3600
      Width           =   1695
   End
   Begin VB.ListBox lsttabuada 
      Height          =   2010
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtinicio 
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Digite a tabuada"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Tabuada 
      Caption         =   "Tabuada"
      Height          =   975
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "frmtabuadasimplesfornext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
