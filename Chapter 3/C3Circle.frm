VERSION 5.00
Begin VB.Form frmC3Circle 
   Caption         =   "Chapter 3 Circle"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1932
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1932
   End
   Begin VB.Image imgCircle 
      Height          =   2400
      Left            =   480
      Picture         =   "C3Circle.frx":0000
      Top             =   1920
      Width           =   2400
   End
   Begin VB.Label lblAnswer 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Width           =   7212
   End
   Begin VB.Label lblPreAnswerText 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   2412
   End
   Begin VB.Label lblUserPrompt 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   16.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   8652
   End
End
Attribute VB_Name = "frmC3Circle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AUTHOR: Jorge Monzon Diaz
'DATE:   March 13, 2014
'DESC:   Chapter 3 Circle



Private Sub cmdCalc_Click()
    Dim dblPi As Double
    dblPi = 4 * Atn(1) 'You need to calculate pi since it isn't included by default
    dblAns = dblPi * (10 * 10)
    lblAnswer.Caption = dblAns & " Units"
End Sub

Private Sub cmdDone_Click()
    frmC3Circle.Hide
End Sub

Private Sub Form_Load()
    frmC3Circle.WindowState = 2
    frmC3Circle.BackColor = RGB(9, 129, 102)
    cmdCalc.BackColor = RGB(101, 179, 84)
    cmdDone.BackColor = RGB(18, 215, 15)
    lblAnswer.FontBold = True
    lblAnswer.Caption = ""
    lblUserPrompt.Caption = "What is the area of a circle with a radius of 10 Units?"
    lblPreAnswerText.Caption = "The answer is:"
End Sub

