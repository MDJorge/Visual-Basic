VERSION 5.00
Begin VB.Form frmC5E5Computer 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Computer TroubleShootingp"
   ClientHeight    =   3084
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3084
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnDone 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Done"
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8280
      Width           =   2055
   End
   Begin VB.CommandButton btnCalc 
      BackColor       =   &H00C0C0C0&
      Caption         =   "What To Do?"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox txtSpin 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7800
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtBeep 
      BackColor       =   &H00FF8080&
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
      Left            =   7800
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblAnswer 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   480
      TabIndex        =   6
      Top             =   5760
      Width           =   6975
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00FF8080&
      Caption         =   "Does the hard drive spin (Y/N)? "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   6495
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00FF8080&
      Caption         =   "Does the computer Beep on start up (Y/N)?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   6135
   End
End
Attribute VB_Name = "frmC5E5Computer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AUTHOR:    Jorge Monzon Diaz
'DATE:      April 8, 2014
'DESC:      C5 Ex 5 Computer Troubleshooting

Dim booBeep, booSpin As Boolean

Private Sub btnCalc_Click()
    If txtBeep = "Y" Or txtBeep = "y" Then
        booBeep = True
    Else
        booBeep = False
    End If
    If txtSpin = "Y" Or txtSpin = "y" Then
        booSpin = True
    Else
        booSpin = False
    End If
    If booSpin And booBeep Then
        lblAnswer = "Contact tech support."
    ElseIf booBeep And booSpin = False Then
        lblAnswer = "Check drive contacts."
    ElseIf booBeep = False And booSpin = False Then
        lblAnswer = "Bring computer to repair center."
    ElseIf booBeep = False And booSpin Then
        lblAnswer = "Check the speaker connections."
    Else
        lblAnswer = "Invalid Inputs"
    End If
End Sub

Private Sub btnDone_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    frmC5E5Computer.WindowState = 2
    btnCalc.Default = True
    lblAnswer = ""
    txtBeep = ""
    txtSpin = ""
End Sub

Private Sub txtBeep_Change()
    lblAnswer = ""
End Sub

Private Sub txtSpin_Change()
    lblAnswer = ""
End Sub
