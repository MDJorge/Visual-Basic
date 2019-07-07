VERSION 5.00
Begin VB.Form frmC5RandomNoPractice 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Random Number Practice"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnDone 
      BackColor       =   &H00808000&
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
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton btnCalc 
      BackColor       =   &H00808000&
      Caption         =   "Calc"
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
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label lblErrorMessage 
      BackColor       =   &H00FFFF80&
      Height          =   615
      Left            =   1200
      TabIndex        =   8
      Top             =   7920
      Width           =   8535
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H0080FFFF&
      Caption         =   "Press CALC button to get 5 random numbers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   8415
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00FFFF00&
      Caption         =   "lbl1"
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
      Left            =   1560
      TabIndex        =   4
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00FFFF00&
      Caption         =   "lbl2"
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
      Left            =   6480
      TabIndex        =   3
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00FFFF00&
      Caption         =   "lbl3"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label lbl4 
      BackColor       =   &H00FFFF00&
      Caption         =   "lbl4"
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
      Left            =   6480
      TabIndex        =   1
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label lbl5 
      BackColor       =   &H00FFFF00&
      Caption         =   "lbl5"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   4800
      Width           =   2895
   End
End
Attribute VB_Name = "frmC5RandomNoPractice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' AUTHOR:   NAME
' DATE:     April 2012
' DESC      C5 Random no practice

Private Sub btnCalc_Click()
    lbl1 = Int(Rnd * 11 + 50)
    lbl2 = Int(Rnd * 11 + 50)
    lbl3 = Int(Rnd * 11 + 50)
    lbl4 = Int(Rnd * 11 + 50)
    lbl5 = Int(Rnd * 11 + 50)
    lblErrorMessage = "You got 5 random numbers between 50-60"
End Sub

Private Sub btnDone_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Randomize       'use diff randon numbers each run

    lbl1 = ""
    lbl2 = ""
    lbl3 = ""
    lbl4 = ""
    lbl5 = ""
End Sub
