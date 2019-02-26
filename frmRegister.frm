VERSION 5.00
Begin VB.Form frmRegister 
   BackColor       =   &H008080FF&
   Caption         =   "Register Player"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRegister 
      Caption         =   "&Yay Register Me!"
      Height          =   1635
      Left            =   3120
      TabIndex        =   2
      Top             =   3240
      Width           =   3255
   End
   Begin VB.TextBox Player2Name 
      Height          =   1035
      Left            =   5340
      TabIndex        =   1
      Top             =   960
      Width           =   2115
   End
   Begin VB.TextBox Player1Name 
      Height          =   975
      Left            =   1140
      TabIndex        =   0
      Top             =   780
      Width           =   1995
   End
   Begin VB.Label Label2 
      Caption         =   "ENTER PLAYER 2 "
      Height          =   555
      Left            =   5520
      TabIndex        =   4
      Top             =   240
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "ENTER PLAYER 1"
      Height          =   495
      Left            =   1260
      TabIndex        =   3
      Top             =   240
      Width           =   1875
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRegister_Click()
frmRegister.Visible = False
frmPlay.Visible = True
End Sub


