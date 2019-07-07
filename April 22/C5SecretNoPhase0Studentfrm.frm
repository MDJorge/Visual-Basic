VERSION 5.00
Begin VB.Form C5SecretNoPhase0frm 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Secret Number Game"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnReplay 
      BackColor       =   &H00C0C000&
      Caption         =   "Replay"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8280
      Width           =   2055
   End
   Begin VB.CommandButton btnDone 
      BackColor       =   &H00C0C000&
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8160
      Width           =   2175
   End
   Begin VB.CommandButton btnCheckGuess 
      BackColor       =   &H00C0C000&
      Caption         =   "Check Guess"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8280
      Width           =   2055
   End
   Begin VB.TextBox txtGuess 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   5040
      TabIndex        =   1
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblWins 
      Caption         =   "Label1"
      Height          =   1335
      Left            =   11400
      TabIndex        =   8
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label lblPlayed 
      Caption         =   "Label1"
      Height          =   1215
      Left            =   11160
      TabIndex        =   7
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label lblMsg 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   4680
      Width           =   8655
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00FFFF80&
      Caption         =   "My Guess"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   2
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Guess a number between 1 and 50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   8415
   End
End
Attribute VB_Name = "C5SecretNoPhase0frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AUTHOR:    Jorge Monzon Diaz

'DATE:      April 22, 2014
'DESC:      Ch.5 Guessing game (secret no game first pgm)

Option Explicit

Dim intSecretNumber, ctrClicks, intWins, intPlayed As Integer

Private Sub btnCheckGuess_Click()
    If IsNumeric(txtGuess) Then
        If txtGuess >= 1 And txtGuess <= 50 Then
            ctrClicks = ctrClicks + 1
            If ctrClicks < 5 Then
                If txtGuess < intSecretNumber Then
                    lblMsg.BackColor = vbRed
                    lblMsg = "Too low"
                ElseIf txtGuess > intSecretNumber Then
                    lblMsg.BackColor = vbRed
                    lblMsg = "Too high"
                Else
                    intWins = intWins + 1
                    lblMsg.BackColor = vbGreen
                    lblMsg = "You win!!!"
                    lblWins = intWins
                End If
            Else
                lblMsg = "Too many guesses. Number was: " & intSecretNumber
                lblMsg.BackColor = vbRed
                btnReplay.Enabled = True
                btnCheckGuess.Enabled = False
            End If
        Else
            lblMsg = "Enter a number 1-50"
        End If
    Else
        lblMsg = "Input must be numeric"
    End If
End Sub

Private Sub btnDone_Click()
    Unload Me
End Sub

Private Sub btnReplay_Click()
    ctrClicks = 0
    intPlayed = intPlayed + 1
    lblPlayed = intPlayed
    Form_Load
End Sub

Private Sub Form_Load()
  C5SecretNoPhase0frm.WindowState = 2          ' maximize form
  txtGuess.TabIndex = 1                           ' set cursor
  btnCheckGuess.Default = True                    ' set default ENTER key to calc button
  lblMsg = " "                                    ' clear fields on form
  txtGuess = ""
  lblPlayed = intPlayed
  lblWins = intWins
  
  ' set random number
  Randomize
  intSecretNumber = Int(50 * Rnd + 1)
  
  btnReplay.Enabled = False
  btnCheckGuess.Enabled = True
End Sub

Private Sub txtGuess_Change()
    lblMsg = ""
End Sub
