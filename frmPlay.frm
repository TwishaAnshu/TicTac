VERSION 5.00
Begin VB.Form frmPlay 
   BackColor       =   &H00C0C0FF&
   Caption         =   "LET'S PLAY!"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   12108
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   12108
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdAI 
      Caption         =   "AI"
      Height          =   915
      Left            =   10020
      TabIndex        =   17
      Top             =   5400
      Width           =   1515
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "QUIT"
      Height          =   1035
      Left            =   10080
      TabIndex        =   16
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "RESET"
      Height          =   855
      Left            =   9840
      TabIndex        =   15
      Top             =   2700
      Width           =   1215
   End
   Begin VB.CommandButton cmdNewGame 
      Caption         =   "NEWGAME"
      Height          =   555
      Left            =   9780
      TabIndex        =   14
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Line Line70 
      Visible         =   0   'False
      X1              =   1140
      X2              =   8520
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line LINE8 
      Visible         =   0   'False
      X1              =   1560
      X2              =   8580
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line LINE6 
      Visible         =   0   'False
      X1              =   1740
      X2              =   8460
      Y1              =   2340
      Y2              =   2400
   End
   Begin VB.Line LINE5 
      Visible         =   0   'False
      X1              =   7440
      X2              =   7500
      Y1              =   1800
      Y2              =   6480
   End
   Begin VB.Line LINE4 
      Visible         =   0   'False
      X1              =   2520
      X2              =   2520
      Y1              =   1620
      Y2              =   6240
   End
   Begin VB.Line Line3 
      Visible         =   0   'False
      X1              =   4740
      X2              =   4740
      Y1              =   1740
      Y2              =   6240
   End
   Begin VB.Line LINE2 
      Visible         =   0   'False
      X1              =   8580
      X2              =   1500
      Y1              =   1740
      Y2              =   5940
   End
   Begin VB.Line LINE1 
      Visible         =   0   'False
      X1              =   1800
      X2              =   8040
      Y1              =   1740
      Y2              =   6000
   End
   Begin VB.Label lblNumTies 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   360
      TabIndex        =   20
      Top             =   5760
      Width           =   915
   End
   Begin VB.Label lblNumOWins 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   300
      TabIndex        =   19
      Top             =   5040
      Width           =   1035
   End
   Begin VB.Label lblNumXWins 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   360
      TabIndex        =   18
      Top             =   4500
      Width           =   1095
   End
   Begin VB.Label lblTie 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TIE"
      Height          =   555
      Left            =   9480
      TabIndex        =   13
      Top             =   540
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label lblOWins 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O Wins!"
      Height          =   615
      Left            =   6120
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblXWins 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X WINS!"
      Height          =   555
      Left            =   3060
      TabIndex        =   11
      Top             =   420
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Label lblTurn 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   420
      TabIndex        =   10
      Top             =   1380
      Width           =   795
   End
   Begin VB.Label lblPlayerName 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   480
      TabIndex        =   9
      Top             =   540
      Width           =   1575
   End
   Begin VB.Label lbl9 
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   6540
      TabIndex        =   8
      Top             =   5220
      Width           =   1695
   End
   Begin VB.Label lbl8 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   4200
      TabIndex        =   7
      Top             =   5220
      Width           =   1275
   End
   Begin VB.Label lbl7 
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   1740
      TabIndex        =   6
      Top             =   5280
      Width           =   1035
   End
   Begin VB.Label lbl6 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   6720
      TabIndex        =   5
      Top             =   3420
      Width           =   1515
   End
   Begin VB.Label lbl5 
      BorderStyle     =   1  'Fixed Single
      Height          =   675
      Left            =   4200
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lbl3 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   6840
      TabIndex        =   3
      Top             =   2040
      Width           =   1515
   End
   Begin VB.Label lbl2 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   4260
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lbl4 
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   1980
      TabIndex        =   1
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lbl1 
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   1800
      TabIndex        =   0
      Top             =   1860
      Width           =   1215
   End
   Begin VB.Line Line 
      Index           =   1
      X1              =   6300
      X2              =   6300
      Y1              =   1680
      Y2              =   6240
   End
   Begin VB.Line Line 
      Index           =   3
      X1              =   1620
      X2              =   8820
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line 
      Index           =   2
      X1              =   1560
      X2              =   8640
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Line Line 
      Index           =   0
      X1              =   3660
      X2              =   3660
      Y1              =   1680
      Y2              =   6240
   End
End
Attribute VB_Name = "frmPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As String
Dim O As String
Dim wins As Double


Sub WhoWins()
If lbl1.Caption = "X" And lbl2.Caption = "X" And lbl3.Caption = "X" Then
LINE6.Visible = True
lblXWins.Visible = True


End If
If lbl8.Caption = "O" And lbl7.Caption = "O" And lbl9.Caption = "O" Then
LINE8.Visible = True
lblOWins.Visible = True

End If

If lbl8.Caption = "X" And lbl7.Caption = "X" And lbl9.Caption = "X" Then
LINE8.Visible = True
End If


If lbl1.Caption = "O" And lbl2.Caption = "O" And lbl3.Caption = "O" Then
LINE6.Visible = True
lblOWins.Visible = True
End If

If lbl4.Caption = "X" And lbl5.Caption = "X" And lbl6.Caption = "X" Then
Line70.Visible = True
lblXWins.Visible = True
End If

If lbl4.Caption = "O" And lbl5.Caption = "O" And lbl6.Caption = "O" Then
Line70.Visible = True
lblOWins.Visible = True
End If

If lbl7.Caption = "X" And lbl8.Caption = "X" And lbl9.Caption = "X" Then
LINE8.Visible = True
lblXWins.Visible = True
End If


If lbl7.Caption = "O" And lbl8.Caption = "O" And lbl9.Caption = "0" Then
LINE8.Visible = True
lblOWins.Visible = True
End If

If lbl1.Caption = "X" And lbl4.Caption = "X" And lbl7.Caption = "X" Then
LINE4.Visible = True
lblXWins.Visible = True
End If

If lbl1.Caption = "O" And lbl4.Caption = "O" And lbl7.Caption = "O" Then
LINE4.Visible = True
lblXWins.Visible = True
End If

If lbl2.Caption = "X" And lbl5.Caption = "X" And lbl8.Caption = "X" Then
Line3.Visible = True
lblXWins.Visible = True
End If

If lbl2.Caption = "O" And lbl5.Caption = "O" And lbl8.Caption = "O" Then
Line3.Visible = True
lblOWins.Visible = True
End If



If lbl2.Caption = "X" And lbl5.Caption = "X" And lbl8.Caption = "X" Then
LINE5.Visible = True
lblXWins.Visible = True
End If

If lbl3.Caption = "O" And lbl6.Caption = "O" And lbl9.Caption = "O" Then
LINE5.Visible = True
lblOWins.Visible = True
End If



If lbl1.Caption = "X" And lbl5.Caption = "X" And lbl9.Caption = "X" Then
LINE1.Visible = True
lblXWins.Visible = True
End If

If lbl1.Caption = "O" And lbl5.Caption = "O" And lbl9.Caption = "O" Then
LINE1.Visible = True
lblOWins.Visible = True
End If


If lbl3.Caption = "X" And lbl5.Caption = "X" And lbl7.Caption = "X" Then
LINE2.Visible = True
lblXWins.Visible = True
End If

If lbl3.Caption = "O" And lbl5.Caption = "O" And lbl7.Caption = "O" Then
LINE2.Visible = True
lblOWins.Visible = True
End If

If lbl1.Enabled = False And lbl2.Enabled = False And lbl3.Enabled = False And lbl4.Enabled = False And lbl5.Enabled = False And lbl6.Enabled = False And lbl7.Enabled = False And lbl8.Enabled = False And lbl9.Enabled = False And lblXWins.Visible = False And lblOWins.Visible = False Then
lblTie.Visible = True
End If


' double win

If lblXWins.Visible = True Then
lblNumXWins.Caption = Val(lblNumXWins) + 1

End If

If lblOWins.Visible = True Then
lblNumOWins.Caption = Val(lblNumXWins) + 1
End If

If lblTie.Visible = True Then
lblNumTies.Caption = Val(lblNumTies) + 1
End If

' DOUBLE WIN

If LINE1.Visible = True And LINE2.Visible = True Or LINE1.Visible = True And LINE2.Visible = True Or LINE1.Visible = True And LINE2.Visible = True Then
lblNumOWins.Caption = Val(lblNumXWins) + 2
End If

If LINE1.Visible = True And Line3.Visible = True Then
lblNumOWins.Caption = Val(lblNumXWins) + 2
End If

If LINE1.Visible = True And LINE5.Visible = True Then
lblNumOWins.Caption = Val(lblNumXWins) + 2
End If

If LINE1.Visible = True And LINE6.Visible = True Then
lblNumOWins.Caption = Val(lblNumXWins) + 2
End If

If LINE1.Visible = True And Line70.Visible = True Then
lblNumOWins.Caption = Val(lblNumXWins) + 2
End If


If LINE1.Visible = True And LINE4.Visible = True Then
lblNumOWins.Caption = Val(lblNumXWins) + 2
End If
If LINE2.Visible = True And Line3.Visible = True Or LINE2.Visible = True And LINE4.Visible = True Or LINE2.Visible = True And LINE5.Visible = True Or LINE2.Visible = True And LINE2.Visible = True Or LINE2.Visible = True And LINE4.Visible = True Or LINE2.Visible = True And LINE5.Visible = True Or LINE2.Visible = True And LINE6.Visible = True Or LINE2.Visible = True And Line70.Visible = True Or LINE1.Visible = True And LINE8.Visible = True Then
lblNumOWins.Caption = Val(lblNumXWins) + 2
End If

If LINE1.Visible = True And LINE2.Visible = True Then
lblNumOWins.Caption = Val(lblNumXWins) + 2
End If

If Line3.Visible = True And LINE4.Visible = True Or Line3.Visible = True And Line3.Visible = True Or Line3.Visible = True And LINE4.Visible = True Or Line3.Visible = True And LINE5.Visible = True Or Line3.Visible = True And LINE6.Visible = True Or Line3.Visible = True And Line70.Visible = True Or Line3.Visible = True And LINE8.Visible = True Then
lblNumOWins.Caption = Val(lblNumXWins) + 2
End If

If LINE4.Visible = True And LINE5.Visible = True Or LINE4.Visible = True And LINE6.Visible = True Or LINE4.Visible = True And Line70.Visible = True Or LINE4.Visible = True And LINE8.Visible = True Then
lblNumOWins.Caption = Val(lblNumXWins) + 2
End If

If LINE5.Visible = True And LINE6.Visible = True Or LINE5.Visible = True And Line70.Visible = True Or LINE5.Visible = True And LINE8.Visible = True Then
lblNumOWins.Caption = Val(lblNumXWins) + 2
End If

If LINE6.Visible = True And Line70.Visible = True Or LINE6.Visible = True And LINE8.Visible = True Or LINE6.Visible = True And LINE8.Visible = True Then
lblNumOWins.Caption = Val(lblNumXWins) + 2
End If

If Line70.Visible = True And LINE8.Visible = True Then
lblNumOWins.Caption = Val(lblNumXWins) + 2
End If




 
 ' for x
 
 
If LINE1.Visible = True And LINE2.Visible = True Or LINE1.Visible = True And Line3.Visible = True Or LINE1.Visible = True And LINE4.Visible = True Or LINE1.Visible = True And LINE5.Visible = True Or LINE1.Visible = True And LINE6.Visible = True Or LINE1.Visible = True And Line70.Visible = True Or LINE1.Visible = True And LINE8.Visible = True Or LINE1.Visible = True And LINE2.Visible = True Or LINE1.Visible = True And LINE2.Visible = True Then
lblNumXWins.Caption = Val(lblNumXWins) + 2
End If

If LINE2.Visible = True And Line3.Visible = True Or LINE2.Visible = True And LINE4.Visible = True Or LINE2.Visible = True And LINE5.Visible = True Or LINE2.Visible = True And LINE2.Visible = True Or LINE2.Visible = True And LINE4.Visible = True Or LINE2.Visible = True And LINE5.Visible = True Or LINE2.Visible = True And LINE6.Visible = True Or LINE2.Visible = True And Line70.Visible = True Or LINE1.Visible = True And LINE8.Visible = True Then
lblNumXWins.Caption = Val(lblNumXWins) + 2
End If

If Line3.Visible = True And LINE4.Visible = True Or Line3.Visible = True And Line3.Visible = True Or Line3.Visible = True And LINE4.Visible = True Or Line3.Visible = True And LINE5.Visible = True Or Line3.Visible = True And LINE6.Visible = True Or Line3.Visible = True And Line70.Visible = True Or Line3.Visible = True And LINE8.Visible = True Then
lblNumXWins.Caption = Val(lblNumXWins) + 2
End If

If LINE4.Visible = True And LINE5.Visible = True Or LINE4.Visible = True And LINE6.Visible = True Or LINE4.Visible = True And Line70.Visible = True Or LINE4.Visible = True And LINE8.Visible = True Then
lblNumXWins.Caption = Val(lblNumXWins) + 2
End If

If LINE5.Visible = True And LINE6.Visible = True Or LINE5.Visible = True And Line70.Visible = True Or LINE5.Visible = True And LINE8.Visible = True Then
lblNumXWins.Caption = Val(lblNumXWins) + 2
End If

If LINE6.Visible = True And Line70.Visible = True Or LINE6.Visible = True And LINE8.Visible = True Or LINE6.Visible = True And LINE8.Visible = True Then
lblNumXWins.Caption = Val(lblNumXWins) + 2
End If

If Line70.Visible = True And LINE8.Visible = True Then
lblNumXWins.Caption = Val(lblNumXWins) + 2
End If

If LINE8.Visible = True And LINE1.Visible = True Then
lblNumXWins.Caption = Val(lblNumXWins) + 2
End If




End Sub




Private Sub cmdAI_Click()
lblTurn.Caption = "X"
If lbl1 = "X" Then
lbl2 = "O"
End If

If lbl2 = "X" Then
lbl5 = "O"
End If

If lbl3 = "X" Then
lbl6 = "O"
End If

If lbl4 = "X" Then
lbl7 = "O"
End If

If lbl5 = "X" Then
lbl9 = "O"
End If
If lbl5 = "X" Then
lbl6 = "O"
End If

If lbl7 = "X" Then
lbl6 = "O"
End If

If lbl8 = "X" Then
lbl6 = "O"
End If

If lbl9 = "X" Then
lbl6 = "O"
End If






End Sub

Private Sub cmdNewGame_Click()
lbl1.Caption = ""
lbl2.Caption = ""
lbl3.Caption = ""
lbl4.Caption = ""
lbl5.Caption = ""
lbl6.Caption = ""
lbl7.Caption = ""
lbl8.Caption = ""
lbl9.Caption = ""
lblXWins.Caption = ""
lblOWins.Caption = ""
 LINE1.Visible = False
 LINE2.Visible = False
 Line3.Visible = False
 LINE4.Visible = False
 LINE5.Visible = False
 LINE6.Visible = False
 Line70.Visible = False
 LINE8.Visible = False
 
 lbl1.Enabled = True
 lbl2.Enabled = True
lbl3.Enabled = True
lbl4.Enabled = True
lbl5.Enabled = True
lbl6.Enabled = True
lbl7.Enabled = True
lbl8.Enabled = True
lbl9.Enabled = True



 
 
 lblNumXWins.Caption = ""
 lblNumOWins.Caption = ""
 
 lblNumTies = ""




frmPlay.Hide
frmRegister.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReset_Click()
lbl1.Caption = ""
lbl2.Caption = ""
lbl3.Caption = ""
lbl4.Caption = ""
lbl5.Caption = ""
lbl6.Caption = ""
lbl7.Caption = ""
lbl8.Caption = ""
lbl9.Caption = ""
lblTurn.Caption = "X"
lblPlayerName.Caption = frmRegister.Player1Name.Text
LINE1.Visible = False
LINE2.Visible = False
Line3.Visible = False
LINE4.Visible = False
LINE5.Visible = False
LINE6.Visible = False
Line70.Visible = False
LINE8.Visible = False
lbl1.Enabled = True
lbl2.Enabled = True
lbl3.Enabled = True
lbl4.Enabled = True
lbl5.Enabled = True
lbl6.Enabled = True
lbl7.Enabled = True
lbl8.Enabled = True
lbl9.Enabled = True
lblXWins.Visible = False
lblOWins.Visible = False
lblTie.Visible = False

End Sub

Private Sub Form_Load()
If lblTurn.Caption = "" Then
lblTurn.Caption = "X"
lblPlayerName.Caption = frmRegister.Player1Name.Text  ' displays player 1 name

Call WhoWins




'If Val(wins) = 1 Then
'lblNumXWins = Val(wins) + 1
 'End If

End If

End Sub

Private Sub lbl1_Click()
If lbl1 = "" Then
    If lblTurn.Caption = "X" Then
    
        lbl1.Caption = "X"
        lbl1.Enabled = False
        lblTurn.Caption = "O"
        lblPlayerName.Caption = frmRegister.Player2Name.Text
    Else
        lbl1.Caption = "O"
        lbl1.Enabled = False
        lblTurn.Caption = "X"
        lblPlayerName.Caption = frmRegister.Player1Name.Text
     End If
End If
Call WhoWins
End Sub

Private Sub lbl2_Click()
  If lblTurn.Caption = "X" Then
        lbl2.Caption = "X"
        lbl2.Enabled = False
        lblTurn.Caption = "O"
        lblPlayerName.Caption = frmRegister.Player2Name.Text
  Else
        lbl2.Caption = "O"
          lbl2.Enabled = False
        lblTurn.Caption = "X"
          lblPlayerName.Caption = frmRegister.Player1Name.Text
    End If
Call WhoWins

End Sub

Private Sub lbl3_Click()
  If lblTurn.Caption = "X" Then
        lbl3.Caption = "X"
        lbl3.Enabled = False
        lblTurn.Caption = "O"
          lblPlayerName.Caption = frmRegister.Player2Name.Text
  Else
        lbl3.Caption = "O"
        lblTurn.Caption = "X"
          lbl3.Enabled = False
            lblPlayerName.Caption = frmRegister.Player1Name.Text
    End If
Call WhoWins

End Sub

Private Sub lbl4_Click()
  If lblTurn.Caption = "X" Then
        lbl4.Caption = "X"
        lbl4.Enabled = False
        lblTurn.Caption = "O"
          lblPlayerName.Caption = frmRegister.Player2Name.Text
  Else
        lbl4.Caption = "O"
        lblTurn.Caption = "X"
          lbl4.Enabled = False
            lblPlayerName.Caption = frmRegister.Player1Name.Text
    End If

Call WhoWins
End Sub

Private Sub lbl5_Click()
  If lblTurn.Caption = "X" Then
        lbl5.Caption = "X"
        lbl5.Enabled = False
        lblTurn.Caption = "O"
          lblPlayerName.Caption = frmRegister.Player2Name.Text
  Else
        lbl5.Caption = "O"
        lblTurn.Caption = "X"
          lbl5.Enabled = False
            lblPlayerName.Caption = frmRegister.Player1Name.Text
    End If

Call WhoWins
End Sub

Private Sub lbl6_Click()
  If lblTurn.Caption = "X" Then
        lbl6.Caption = "X"
        lbl6.Enabled = False
        lblTurn.Caption = "O"
          lblPlayerName.Caption = frmRegister.Player2Name.Text
  Else
        lbl6.Caption = "O"
        lblTurn.Caption = "X"
          lbl6.Enabled = False
            lblPlayerName.Caption = frmRegister.Player1Name.Text
    End If

Call WhoWins
End Sub

Private Sub lbl7_Click()
  If lblTurn.Caption = "X" Then
        lbl7.Caption = "X"
        lbl7.Enabled = False
        lblTurn.Caption = "O"
          lblPlayerName.Caption = frmRegister.Player2Name.Text
  Else
        lbl7.Caption = "O"
        lbl7.Enabled = False
        lblTurn.Caption = "X"
          lbl7.Enabled = False
            lblPlayerName.Caption = frmRegister.Player1Name.Text
    End If

Call WhoWins
End Sub

Private Sub lbl8_Click()
  If lblTurn.Caption = "X" Then
        lbl8.Caption = "X"
        lbl8.Enabled = False
        lblTurn.Caption = "O"
          lblPlayerName.Caption = frmRegister.Player2Name.Text
  Else
        lbl8.Caption = "O"
        lblTurn.Caption = "X"
          lbl8.Enabled = False
            lblPlayerName.Caption = frmRegister.Player1Name.Text
    End If
Call WhoWins
End Sub

Private Sub lbl9_Click()
  If lblTurn.Caption = "X" Then
        lbl9.Caption = "X"
        lbl9.Enabled = False
        lblTurn.Caption = "O"
          lblPlayerName.Caption = frmRegister.Player2Name.Text
  Else
        lbl9.Caption = "O"
        lblTurn.Caption = "X"
          lbl9.Enabled = False
            lblPlayerName.Caption = frmRegister.Player1Name.Text
    End If
Call WhoWins

End Sub



