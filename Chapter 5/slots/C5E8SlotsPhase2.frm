VERSION 5.00
Begin VB.Form C5E8SlotsPhase1a 
   BackColor       =   &H00008000&
   Caption         =   "Slot Machine Game"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12375
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtWager 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      TabIndex        =   12
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton btnReplay 
      BackColor       =   &H000000FF&
      Caption         =   "REPLAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8640
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8040
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H000000FF&
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5760
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8040
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton cmdPull 
      BackColor       =   &H000000FF&
      Caption         =   "PULL"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2280
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8040
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Amount of Wager"
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
      Left            =   120
      TabIndex        =   11
      Top             =   6000
      Width           =   4335
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
      Height          =   1695
      Left            =   0
      TabIndex        =   9
      Top             =   600
      Width           =   10215
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H0000FF00&
      Caption         =   "TO PLAY SLOT MACHINE GAME:"
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
      TabIndex        =   8
      Top             =   0
      Width           =   10095
   End
   Begin VB.Label lblMsg 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   120
      TabIndex        =   7
      Top             =   6960
      Width           =   12255
   End
   Begin VB.Image Image1 
      Height          =   2505
      Left            =   7920
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   4530
   End
   Begin VB.Label lblTokens 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
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
      Left            =   4920
      TabIndex        =   6
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
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
      Height          =   1815
      Left            =   4920
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
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
      Height          =   1815
      Left            =   3000
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1080
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblAnswer 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tokens You Have:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   4335
   End
End
Attribute VB_Name = "C5E8SlotsPhase1a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AUTHOR:    TEACHER
'DATE:      5/13/11
'DESC:      C5 Ex 8 SLOT MACHINE Phase 1a (w/lblMsg not msgbox)

' GENERAL PROCEDURE SECTION
Option Explicit
Private tokens As Integer                 ' define global variable for intTokens
Dim slot1, slot2, slot3 As Integer        ' define 3 slot variables
Dim wager As Integer

Private Sub btnReplay_Click()
    Call Form_Load
End Sub

Private Sub Form_Load()
    C5E8SlotsPhase1a.WindowState = 2
    cmdPull.Default = True
       
    Randomize                       ' init randomize feature
    
    lbl1 = " "                      ' clear 3 random numbers on form
    lbl2 = " "
    lbl3 = " "
    
    tokens = 100                      ' set intTokens to be 100
    lblTokens = tokens                ' move intTokens to form field
   
    btnReplay.Enabled = False
    cmdPull.Enabled = True
End Sub

Private Sub cmdPull_Click()
    If tokens > 0 Then
         If txtWager > 0 And txtWager <= tokens Then
            wager = Val(txtWager)
            
            tokens = tokens - wager                ' subtract 1 from intTokens
             
            slot1 = Int(Rnd * 3 + 1)           ' get 3 random whole numbers between 1-3
            slot2 = Int(Rnd * 3 + 1)
            slot3 = Int(Rnd * 3 + 1)
             
            lbl1 = slot1                      ' move 3 slot numbers to form fields
            lbl2 = slot2
            lbl3 = slot3
             
            If (slot1 = 1) And (slot2 = 1) And (slot3 = 1) Then       ' if all numbers = 1
                tokens = tokens + (wager * 4)
                lblMsg = "You WIN. You won " & (wager * 4) & " Tokens!"
            ElseIf (slot1 = 2) And (slot2 = 2) And (slot3 = 2) Then   ' if all numbers = 2
                tokens = tokens + (wager * 8)
                lblMsg = "You WIN. You won " & (wager * 8) & " Tokens!"
            ElseIf (slot1 = 3) And (slot2 = 3) And (slot3 = 3) Then   ' if all numbers = 3
                tokens = tokens + (wager * 12)
                lblMsg = "You WIN. You won " & (wager * 12) & " Tokens!"
            Else                                                            ' else YOU LOSE
                lblMsg = "You LOSE. Try Again!"
            End If
             
            
            lblTokens = tokens               ' move intTokens to form
         Else
            lblMsg = "Wager is wrong. Must be greater than 0 and less than amount of tokens"
         End If
   Else
        lblMsg = "Out of tokens. Replay or quit."
        cmdPull.Enabled = False
        btnReplay.Enabled = True
    End If
End Sub


Private Sub cmdDone_Click()
    Unload Me
End Sub

