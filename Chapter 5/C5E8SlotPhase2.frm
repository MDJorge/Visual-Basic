VERSION 5.00
Begin VB.Form frmC5E8SlotComplete 
   BackColor       =   &H00008000&
   Caption         =   "Slot Machine Game"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10080
   FillColor       =   &H00404040&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00008000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   12990
   ScaleWidth      =   21480
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtWager 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   12600
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   480
      Width           =   3015
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
      Height          =   1095
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9360
      Width           =   3615
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
      Height          =   1095
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9240
      Width           =   3255
   End
   Begin VB.Label lblWager 
      Alignment       =   2  'Center
      Caption         =   "Wager"
      Height          =   735
      Left            =   12600
      TabIndex        =   9
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00008000&
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
      Left            =   1560
      TabIndex        =   7
      Top             =   4080
      Width           =   10215
   End
   Begin VB.Label lblTokens 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   54
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9840
      TabIndex        =   6
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label lblSlot3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   9360
      TabIndex        =   5
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblSlot2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   5040
      TabIndex        =   4
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblSlot1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblAnswer 
      Caption         =   "Tokens You Have:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   6240
      Width           =   8535
   End
End
Attribute VB_Name = "frmC5E8SlotComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AUTHOR:    Jorge Monzon Diaz
'DATE:      April 12, 2014
'DESC:      C5 Ex 8 SLOT MACHINE PHASE II

Option Explicit

Dim intTokens, intNo1, intNo2, intNo3, intWager, intWon As Integer
    
Private Sub Form_Load()
    frmC5E8SlotComplete.WindowState = 2
    cmdPull.Default = True
    Randomize
    intTokens = 100
    lblTokens = intTokens
    lblSlot1 = ""
    lblSlot2 = ""
    lblSlot3 = ""
    txtWager = ""
End Sub

Private Sub cmdPull_Click()
    intNo1 = Int(3 * Rnd + 1)
    intNo2 = Int(3 * Rnd + 1)
    intNo3 = Int(3 * Rnd + 1)
    lblSlot1 = intNo1
    lblSlot2 = intNo2
    lblSlot3 = intNo3
    If txtWager = "" Then
        intWager = 1
    ElseIf IsNumeric(txtWager) Then
        intWager = txtWager
    Else
        MsgBox "Enter a valid wager number"
    End If
    If intTokens > 0 Then
        If intNo1 = intNo2 And intNo2 = intNo3 And intNo3 = intNo1 Then
            If intNo1 = 1 Then
                intWon = intWager * 4
                intTokens = intTokens + intWon
                lblMessage = "You won " & intWon & " tokens!"
            ElseIf intNo1 = 2 Then
                intWon = intWager * 8
                intTokens = intTokens + intWon
                lblMessage = "You won " & intWon & " tokens!"
            Else
                intWon = intWager * 12
                intTokens = intTokens + intWon
                lblMessage = "You got " & intWon & " tokens!"
            End If
        Else
            intTokens = intTokens - intWager
            lblMessage = "You lose " & intWager & " tokens!"
        End If
        lblTokens = intTokens
    Else
        lblMessage = "You lose!"
    End If
    txtWager = ""
End Sub


Private Sub cmdDone_Click()
    Unload Me
End Sub

