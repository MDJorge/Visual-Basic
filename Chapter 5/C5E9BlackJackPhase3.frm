VERSION 5.00
Begin VB.Form frmC5E9BlackJack 
   BackColor       =   &H00004000&
   Caption         =   "Game 21 BlackJack"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnDone 
      BackColor       =   &H0000FFFF&
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
      Height          =   975
      Left            =   8280
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9720
      Width           =   1695
   End
   Begin VB.CommandButton btnPlayAgain 
      BackColor       =   &H0000FFFF&
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
      Height          =   975
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9720
      Width           =   1575
   End
   Begin VB.CommandButton btnDrawCard 
      BackColor       =   &H0000FFFF&
      Caption         =   "Draw Card"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9720
      Width           =   1575
   End
   Begin VB.CommandButton btnCheckScore 
      BackColor       =   &H0000FFFF&
      Caption         =   "Check Score"
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
      Height          =   975
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9720
      Width           =   1575
   End
   Begin VB.Label lblCompWins 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11640
      TabIndex        =   27
      Top             =   8880
      Width           =   1455
   End
   Begin VB.Label lblPlayWins 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   26
      Top             =   8880
      Width           =   1455
   End
   Begin VB.Label lblComputerWins 
      BackColor       =   &H000000FF&
      Caption         =   "# of Computer wins:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   25
      Top             =   8880
      Width           =   3255
   End
   Begin VB.Label lblPlayerWins 
      BackColor       =   &H000000FF&
      Caption         =   "# of Player wins:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   24
      Top             =   8880
      Width           =   3735
   End
   Begin VB.Label lblCompCard5 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   11880
      TabIndex        =   23
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label lblCompCard4 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   9840
      TabIndex        =   22
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label lblPlayCard5 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   4440
      TabIndex        =   21
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label lblPlayCard4 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2400
      TabIndex        =   20
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label lblInstruc2 
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   19
      Top             =   600
      Width           =   13815
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
      Width           =   13695
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "The Computer drew:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   17
      Top             =   5040
      Width           =   5295
   End
   Begin VB.Label lblCompScore 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11640
      TabIndex        =   16
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label lblCompMessage 
      BackColor       =   &H000000FF&
      Caption         =   "Computer Score is:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   15
      Top             =   8160
      Width           =   3255
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   14
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label lblPlayerMessage 
      BackColor       =   &H000000FF&
      Caption         =   "Players Score is:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   8160
      Width           =   3735
   End
   Begin VB.Label lblCompCard3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   7800
      TabIndex        =   12
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label lblCompCard2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   9960
      TabIndex        =   9
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblCompCard1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   7800
      TabIndex        =   8
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblPlayCard3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   360
      TabIndex        =   7
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label lblPlayCard1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   480
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblPlayCard2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2520
      TabIndex        =   5
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C000C0&
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
      Height          =   495
      Left            =   7680
      TabIndex        =   4
      Top             =   2160
      Width           =   5535
   End
   Begin VB.Label lbl3 
      BackColor       =   &H000000FF&
      Caption         =   "The Player drew:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   3735
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00C000C0&
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
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   5055
   End
End
Attribute VB_Name = "frmC5E9BlackJack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AUTHOR: Jorge Monzon Diaz
'DATE:   May 29, 2014
'DESC:   C5 Ex 9 BlackJack PHASE III

Option Explicit

Dim intUserCard1, intUserCard2, intUserCard3, intUserCard4, intUserCard5 As Integer     ' user cards
Dim intCompCard1, intCompCard2, intCompCard3, inCompCard As Integer     ' computer cards
Dim intUserTotal, intCompTotal As Integer                                               ' computer and user card sums
Dim intPlayWins, intCompWins As Integer                                                 ' used to keep track of number of wins
Dim ctrHits, ctrComp As Integer                                                         ' counters used for checking how many cards should be drawn

Private Sub btnCheckScore_Click()           ' sub runs when CHECK SCORES button clicked

    ' Reveal computer cards on form
    lblCompCard1 = intCompCard1                    ' display computer card on form
    lblCompCard2 = intCompCard2                    ' display computer card on form
    
    ' LOGIC to determine if computer should draw more cards or not
    ' Loop only runs if the player is winning, the computer hasn't busted, and that only 3 cards have been drawn
    Do While intUserTotal < 21 And intCompTotal < 21 And intCompTotal < intUserTotal And ctrComp <= 3
        ' max of 3 extra cards available for computer use
        ' if logic combined with counter is used to check how many have already been drawn
        ctrComp = ctrComp + 1                           ' increment computer counter by one
        
        If ctrComp = 1 Then                             ' draw 3th card incase computer is losing
            intCompCard3 = Int(10 * Rnd() + 1)              ' get another card
            lblCompCard3 = intCompCard3                     ' display card on form
            intCompTotal = intCompTotal + intCompCard3      ' add card to computer total
        ElseIf ctrComp = 2 Then                         ' draw 4th card incase computer is still losing
            intCompCard4 = Int(10 * Rnd() + 1)              ' get another card
            lblCompCard4 = intCompCard4                     ' display card on form
            intCompTotal = intCompTotal + intCompCard4      ' add card to computer total
        Else                                            ' draw 5th card incase computer is still losing
            intCompCard5 = Int(10 * Rnd() + 1)              ' get another card
            lblCompCard5 = intUserCard5                     ' display card on form
            intCompTotal = intCompTotal + intCompCard5      ' add card to computer total
        End If
    Loop
    
    ' display computer total on form
    lblCompScore = intCompTotal
    
    ' LOGIC to determine WHO won
    If intCompTotal > 21 And intUserTotal > 21 Then
        MsgBox "Draw!"
    ElseIf intCompTotal > intUserTotal And intCompTotal <= 21 Then
        MsgBox "You lose! The computer wins!"
        intCompWins = intCompWins + 1                       ' update number of times computer has won
    ElseIf intUserTotal > intCompTotal And intUserTotal <= 21 Then
        MsgBox "You win!"
        intPlayWins = intPlayWins + 1                       ' update number of times user has won
    ElseIf intUserTotal = intCompTotal And intUserTotal And intCompTotal <= 21 Then
        MsgBox "Draw!"
    ElseIf intUserTotal > 21 Then
        MsgBox "You lose! The computer wins!"
        intCompWins = intCompWins + 1                       ' update number of times computer has won
    Else
        MsgBox "You win!"
        intPlayWins = intPlayWins + 1                       ' update number of times user has won
    End If
     
     
    ' show number of times the computer and user have won
    lblPlayWins = intPlayWins
    lblCompWins = intCompWins
    
    ' disable the draw and checkscores button on form, and enable the play again button
    btnDrawCard.Enabled = False
    btnCheckScore.Enabled = False
    btnPlayAgain.Enabled = True
End Sub

Private Sub btnDone_Click()
    Unload Me
End Sub

Private Sub btnDrawCard_Click()                 ' SUB runs when DRAW button clicked
    ' max of 3 extra cards available
    ' if logic combined with counter is used to check how many have already been drawn
    ctrHits = ctrHits + 1                           ' increment counter by 1
    If ctrHits = 1 Then                             ' FIRST card drawn
        intUserCard3 = Int(10 * Rnd() + 1)              ' get 3rd card
        lblPlayCard3 = intUserCard3                     ' update score with 3rd card
        intUserTotal = intUserTotal + intUserCard3      ' move 3rd player card to form
    ElseIf ctrHits = 2 Then                         ' SECOND card drawn
        intUserCard4 = Int(10 * Rnd() + 1)              ' get 4th card
        lblPlayCard4 = intUserCard4                     ' update score with 4th card
        intUserTotal = intUserTotal + intUserCard4      ' move 4th player card to form
    ElseIf ctrHits = 3 Then                         ' THIRD card drawn
        intUserCard5 = Int(10 * Rnd() + 1)              ' get 5th card
        lblPlayCard5 = intUserCard5                     ' move 5th player card to form
        intUserTotal = intUserTotal + intUserCard5      ' update score with 5th card
    Else                                            ' if more than 3 cards were already drawn produce error and disable button
        btnDrawCard.Enabled = False
        MsgBox "ERROR: Maximum number of cards (3) already drawn."
    End If
    
    ' move updated player score to form
    lblPlayScore = intUserTotal
End Sub

Private Sub btnPlayAgain_Click()
    ' set game up to play again-refresh scores, cards, and counters
    Call Form_Load
End Sub

Private Sub Form_Load()
    frmC5E9BlackJack.WindowState = 2
    btnDrawCard.Default = True
    btnDrawCard.TabIndex = 1
    btnCheckScore.TabIndex = 2
    btnPlayAgain.TabIndex = 3
    btnDone.TabIndex = 4
    
    ' enable/disable buttons
    btnDrawCard.Enabled = True
    btnCheckScore.Enabled = True
    btnPlayAgain.Enabled = False
    
    ' initialize random numbers
    Randomize
    
    ' reset counters
    ctrHits = 0
    ctrComp = 0
    
    ' clear out forms objects
    lblPlayCard3 = ""
    lblPlayCard4 = ""
    lblPlayCard5 = ""
    lblCompCard3 = ""
    lblCompCard4 = ""
    lblCompCard5 = ""
    
    ' show number of times the computer and user have won
    lblPlayWins = intPlayWins
    lblCompWins = intCompWins
    
    ' get player cards (random no 1-10) for 2 cards, calc score and move to form
    intUserCard1 = Int(10 * Rnd() + 1)
    intUserCard2 = Int(10 * Rnd() + 1)
    lblPlayCard1 = intUserCard1
    lblPlayCard2 = intUserCard2
    intUserTotal = intUserCard1 + intUserCard2
    lblPlayScore = intUserTotal
    
    ' get computer cards (random no 1-10) for 3 cards and calc score
    intCompCard1 = Int(10 * Rnd() + 1)
    intCompCard2 = Int(10 * Rnd() + 1)
    intCompCard3 = Int(10 * Rnd() + 1)
    intCompTotal = intCompCard1 + intCompCard2
    
    ' HIDE computer cards and Score on form  (display with asterik)
    lblCompCard1 = "*"
    lblCompCard2 = "*"
    lblCompScore = "*"
    
    ' set instructions on form
    lblInstruc2 = "Game starts with 2 cards being dealt.  Player can DRAW up to 3 additional cards. Computer cards and score are hidden until Player is done drawing.  Goal is to reach score of 21, without going over 21.  Hit CHECK SCORES to see winner."
End Sub

