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
      Top             =   9120
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
      Top             =   9120
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
      Top             =   9120
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
      Top             =   9120
      Width           =   1575
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
      Top             =   8040
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
      Top             =   8040
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
      Left            =   12120
      TabIndex        =   12
      Top             =   2880
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
'DATE:   April 13, 2014
'DESC:   C5 Ex 9 BlackJack

Option Explicit

Dim intUserCard1, intUserCard2, intUserCard3, intCompCard1, intCompCard2, intCompCard3, intUserTotal, intCompTotal, ctrHits As Integer

Private Sub btnCheckScore_Click()
    lblCompCard1 = intCompCard1
    lblCompCard2 = intCompCard2
    lblCompCard3 = intCompCard3
    intCompTotal = intCompCard1 + intCompCard2 + intCompCard3
    lblCompScore = intCompTotal
    If intCompTotal > 21 And intUserTotal > 21 Then
        MsgBox "Draw!"
    ElseIf intCompTotal > intUserTotal And intCompTotal <= 21 Then
        MsgBox "You lose! The computer wins!"
    ElseIf intUserTotal > intCompTotal And intUserTotal <= 21 Then
        MsgBox "You win!"
    ElseIf intUserTotal = intCompTotal And intUserTotal And intCompTotal <= 21 Then
        MsgBox "Draw!"
    ElseIf intUserTotal > 21 Then
        MsgBox "You lose! The computer wins!"
    Else
        MsgBox "You win!"
    End If
    btnDrawCard.Enabled = False
    btnCheckScore.Enabled = False
    btnPlayAgain.Enabled = True
End Sub

Private Sub btnDone_Click()
    Unload Me
End Sub

Private Sub btnDrawCard_Click()
    If ctrHits < 1 Then
        intUserCard3 = Int(10 * Rnd() + 1)
        lblPlayCard3 = intUserCard3
        intUserTotal = intUserCard1 + intUserCard2 + intUserCard3
        lblPlayScore = intUserTotal
    Else
        MsgBox "ERROR: Card already drawn."
    End If
    ctrHits = ctrHits + 1
End Sub

Private Sub btnPlayAgain_Click()
    btnDrawCard.Enabled = True
    btnCheckScore.Enabled = True
    Form_Load
End Sub

Private Sub Form_Load()
    frmC5E9BlackJack.WindowState = 2
    btnDrawCard.Default = True
    btnDrawCard.TabIndex = 1
    btnCheckScore.TabIndex = 2
    btnPlayAgain.TabIndex = 3
    btnDone.TabIndex = 4
    btnPlayAgain.Enabled = False
    Randomize
    ctrHits = 0
    'user cards
    intUserCard1 = Int(10 * Rnd() + 1)
    intUserCard2 = Int(10 * Rnd() + 1)
    lblPlayCard1 = intUserCard1
    lblPlayCard2 = intUserCard2
    lblPlayCard3 = ""
    intUserTotal = intUserCard1 + intUserCard2
    lblPlayScore = intUserTotal
    'computer cards
    intCompCard1 = Int(10 * Rnd() + 1)
    intCompCard2 = Int(10 * Rnd() + 1)
    intCompCard3 = Int(10 * Rnd() + 1)
    lblCompCard1 = "*"
    lblCompCard2 = "*"
    lblCompCard3 = "*"
    lblCompScore = ""
End Sub
