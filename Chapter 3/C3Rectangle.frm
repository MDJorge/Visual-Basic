VERSION 5.00
Begin VB.Form frmC3Rectangle 
   Caption         =   "Rectangle Area & Perimeter"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDone 
      Caption         =   "Finish"
      Height          =   852
      Left            =   6360
      TabIndex        =   6
      Top             =   8160
      Width           =   2172
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Filler"
      Height          =   1092
      Left            =   6240
      TabIndex        =   5
      Top             =   6360
      Width           =   2052
   End
   Begin VB.Label lblPerimeterAnswer 
      Caption         =   "P-Answer"
      Height          =   852
      Left            =   4200
      TabIndex        =   4
      Top             =   6720
      Width           =   1092
   End
   Begin VB.Label lblAreaAnswer 
      Caption         =   "A-Answer"
      Height          =   972
      Left            =   3960
      TabIndex        =   3
      Top             =   5040
      Width           =   1212
   End
   Begin VB.Label lblPerimeter 
      Caption         =   "PERIMETER"
      Height          =   1212
      Left            =   840
      TabIndex        =   2
      Top             =   6480
      Width           =   2412
   End
   Begin VB.Label lblArea 
      Caption         =   "AREA"
      Height          =   972
      Left            =   840
      TabIndex        =   1
      Top             =   4920
      Width           =   2292
   End
   Begin VB.Label lblPrompt 
      Caption         =   "FILLER TEXT"
      Height          =   2652
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   6732
   End
End
Attribute VB_Name = "frmC3Rectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jorge Monzon Diaz
'Date:   March 21, 2014
'Desc:   Chapter 3 Exercise 10 Rectangle Area & Perimeter

Option Explicit


Private Sub cmdCalculate_Click()
    Dim intLength, intWidth As Integer
    intLength = 5
    intWidth = 3
    lblAreaAnswer = CInt(intLength * intWidth)
    lblPerimeterAnswer = CInt((2 * intLength) + (2 * intWidth))
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    frmC3Rectangle.WindowState = 2
    lblPrompt = "What is the area and perimeter of a rectangle with a length of 5 and a width of 3?"
    lblArea = "Area ="
    lblAreaAnswer = " "
    lblPerimeter = "Perimeter ="
    lblPerimeterAnswer = " "
    cmdCalculate.Caption = "Calculate"
    cmdDone.Caption = "Done"
End Sub
