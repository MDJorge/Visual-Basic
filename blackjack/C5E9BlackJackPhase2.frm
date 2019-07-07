VERSION 5.00
Begin VB.Form frmC5E9BlackJackPhase2 
   BackColor       =   &H00008000&
   Caption         =   "Game 21 BlackJack"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H00404040&
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   12960
      Width           =   1695
   End
   Begin VB.CommandButton cmdPlayAgain 
      BackColor       =   &H00808080&
      Caption         =   "Play Again"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   12960
      Width           =   1695
   End
   Begin VB.CommandButton cmdDrawCard 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Draw Card"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   12960
      Width           =   1695
   End
   Begin VB.CommandButton cmdCheckScore 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Check Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   12960
      Width           =   1695
   End
   Begin VB.Label lblCompcard5 
      BackColor       =   &H008080FF&
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   15720
      TabIndex        =   23
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label lblCompCard4 
      BackColor       =   &H008080FF&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   12960
      TabIndex        =   22
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label lblPlayCard5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   6240
      TabIndex        =   21
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label lblPlayCard4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   3480
      TabIndex        =   20
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label lblInstruc2 
      BackColor       =   &H0000C000&
      Caption         =   "set in form load"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   19
      Top             =   600
      Width           =   17775
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H0000FF00&
      Caption         =   "TO PLAY BLACKJACK (Game of 21) GAME:"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   12615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "The Computer drew:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   17
      Top             =   6720
      Width           =   6015
   End
   Begin VB.Label lblCompScore 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14880
      TabIndex        =   16
      Top             =   11040
      Width           =   1695
   End
   Begin VB.Label lblCompMessage 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Computer Score is:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9840
      TabIndex        =   15
      Top             =   11040
      Width           =   4455
   End
   Begin VB.Label lblPlayScore 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Left            =   5280
      TabIndex        =   14
      Top             =   11040
      Width           =   1575
   End
   Begin VB.Label lblPlayerMessage 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Players Score is:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   13
      Top             =   11040
      Width           =   4335
   End
   Begin VB.Label lblCompCard3 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   10200
      TabIndex        =   12
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label lblCompCard2 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   14280
      TabIndex        =   9
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label lblCompCard1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   11160
      TabIndex        =   8
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label lblPlayCard3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   720
      TabIndex        =   7
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label lblPlayCard1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   1920
      TabIndex        =   6
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblPlayCard2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   5040
      TabIndex        =   5
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "The Computer is dealt:"
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
      Left            =   9840
      TabIndex        =   4
      Top             =   2040
      Width           =   7695
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "The Player drew:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   6720
      Width           =   6015
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "The Player is dealt:"
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
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   7695
   End
End
Attribute VB_Name = "frmC5E9BlackJackPhase2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AUTHOR: TEACHER
'DATE:   Feb 2010
'DESC:   C5 Ex 9 Game of 21 - Phase 2 ANSWER (player & computr can draw up to 3 cards)

' define all numeric variables
Dim playCard1, playCard2, playCard3, playCard4, playCard5 As Integer
Dim compCard1, compCard2, compCard3, compCard4, compCard5 As Integer
Dim playScore, compScore As Integer
Dim counter As Integer

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub cmdCheckScore_Click()           ' sub runs when CHECK SCORES button clicked

' Reveal computer cards on form
    lblCompCard1 = compCard1                    ' display computer card on form
    lblCompCard2 = compCard2                    ' display computer card on form
        
    ' LOGIC to determine if computer should draw more cards or not
    '  (NOTE: not ElseIFs becasue ALL 3 Ifs need to be run and tested)
    '
    If (compScore <= playScore) And (playScore <= 21) Then      ' player winning-draw card neither busted
        compCard3 = Int(10 * Rnd()) + 1                         ' then get another card
        lblCompCard3 = compCard3                                ' assign value to card on form
        compScore = compScore + compCard3                       ' add card to comp score
        If (compScore <= playScore) And (playScore <= 21) Then  ' if comp.still not winning
            compCard4 = Int(10 * Rnd()) + 1                     ' then get another card
            lblCompCard4 = compCard4                            ' assign value to card on form
            compScore = compScore + compCard4                   ' add card to comp score
            If (compScore <= playScore) And (playScore <= 21) Then  ' if comp.still not winning
                compCard5 = Int(10 * Rnd()) + 1                 ' then get another card
                lblCompcard5 = compCard5                        ' assign value to card on form
                compScore = compScore + compCard5               ' add card to comp score
             End If
         End If
     End If
    lblCompScore = compScore                                    ' display comp.score on form
    
    ' LOGIC to determine WHO won
    '
    If compScore > 21 And playScore > 21 Then
        MsgBox " Game is a draw!!"
    ElseIf compScore > playScore And compScore <= 21 Then
        MsgBox " computer Wins!!"
    ElseIf playScore > compScore And playScore <= 21 Then
        MsgBox " Player Wins!!"
    ElseIf playScore = compScore And playScore <= 21 And compScore <= 21 Then
        MsgBox " DRAW- no winners!!"
    ElseIf playScore > 21 Then
        MsgBox "Computer wins!!"
    Else
        MsgBox "Player wins!!"
    End If
    
    ' disable the draw and checkscores button on form
    cmdDrawCard.Enabled = False
    cmdCheckScore.Enabled = False
End Sub

Private Sub cmdDrawCard_Click()         ' SUB runs when DRAW button clicked
    ' max of 3 extra cards-keep counter to count only 3 given
    '
    If counter < 3 Then                             ' not more than 3 cards have been drawn
        counter = counter + 1                       ' increment counter by 1
        If counter = 1 Then                         ' FIRST card drawn
            playCard3 = Int((10 - 1 + 1) * Rnd + 1)       ' get 3rd card
            playScore = playScore + playCard3             ' update score with 3rd card
            lblPlayCard3 = playCard3                      ' move 3rd player card to form
         ElseIf counter = 2 Then                     ' SECOND card drawn
            playCard4 = Int((10 - 1 + 1) * Rnd + 1)       ' get 4th card
            playScore = playScore + playCard4             ' update score with 4th card
            lblPlayCard4 = playCard4                      ' move 4th player card to form
         ElseIf counter = 3 Then                     ' THIRD card drawn
            playCard5 = Int((10 - 1 + 1) * Rnd + 1)       ' get 5th card
            playScore = playScore + playCard5             ' update score with 5th card
            lblPlayCard5 = playCard5                      ' move 5th player card to form
         End If
    Else: MsgBox " You already drew 3 extra cards-only 3 extra cards allowed! Do CheckScores"
    End If
    lblPlayScore = playScore                      ' move updated player score to form
End Sub

Private Sub cmdPlayAgain_Click()
    ' set game up to play again-refresh scores, cards, and counters
    Call Form_Load
      
End Sub

Private Sub Form_Load()
    ' clear all form objects
    lblPlayCard1 = " "
    lblPlayCard2 = " "
    lblPlayCard3 = " "
    lblPlayCard4 = " "
    lblPlayCard5 = " "
    lblCompCard1 = " "
    lblCompCard2 = " "
    lblCompCard3 = " "
    lblCompCard4 = " "
    lblCompcard5 = " "
    lblPlayScore = " "
    lblCompScore = " "

    ' initialize random numbers
    Randomize
    
    ' get player cards (random no 1-10) for 2 cards, calc score and move to form
    playCard1 = Int((10 - 1 + 1) * Rnd + 1)
    playCard2 = Int((10 - 1 + 1) * Rnd + 1)
    playScore = playCard1 + playCard2
    lblPlayCard1 = playCard1
    lblPlayCard2 = playCard2
    lblPlayScore = playScore
    
    ' get computer cards (random no 1-10) for 3 cards and calc score
    compCard1 = Int((10 - 1 + 1) * Rnd + 1)
    compCard2 = Int((10 - 1 + 1) * Rnd + 1)
    compScore = compCard1 + compCard2
        
    ' HIDE computer cards and Score on form  (display with asterik)
    lblCompCard1 = "*"
    lblCompCard2 = "*"
    lblCompScore = "*"
    
    ' set counter fro player drawing cards (to count how many draws are done)
    counter = 0
    
    ' Enable all buttons
    cmdDrawCard.Enabled = True
    cmdCheckScore.Enabled = True
       
    ' set instructions on form
    lblInstruc2 = "Game starts with 2 cards being dealt.  Player can DRAW up to 3 additional cards. Computer cards and score are hidden until Player is done drawing.  Goal is to reach score of 21, without going over 21.  Hit CHECK SCORES to see winner."
    
End Sub

