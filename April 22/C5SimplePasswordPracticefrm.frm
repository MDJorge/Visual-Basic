VERSION 5.00
Begin VB.Form C5SimplePasswordPracticefrm 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Secret Number Game"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   12990
   ScaleWidth      =   21480
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtPassword 
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
      Height          =   555
      IMEMode         =   3  'DISABLE
      Left            =   4680
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   3360
      Width           =   3375
   End
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H00C0C000&
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton cmdCheckID 
      BackColor       =   &H00C0C000&
      Caption         =   "Check Id and password"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   2295
   End
   Begin VB.TextBox txtID 
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
      Height          =   555
      Left            =   4680
      TabIndex        =   1
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00FFFF80&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   5
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00FFFF80&
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Enter your ID and Password below"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5295
   End
End
Attribute VB_Name = "C5SimplePasswordPracticefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AUTHOR:    Teacher
'DATE:      May 2012
'DESC:      Ch.5 Password Practice

Option Explicit

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   
    txtID = ""
    txtPassword = ""
    txtID.TabIndex = 1
    txtPassword.TabIndex = 2
End Sub

Private Sub txtGuess_Change()
    
End Sub

Private Sub txtSecretNumber_Change()
    
End Sub
