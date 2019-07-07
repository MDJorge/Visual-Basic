VERSION 5.00
Begin VB.Form frmPythagoreanTheorem 
   Caption         =   "Pythagorean Theorem"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtB 
      Height          =   732
      Left            =   5400
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1920
      Width           =   972
   End
   Begin VB.TextBox txtA 
      Height          =   732
      Left            =   5400
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   720
      Width           =   1212
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Command1"
      Height          =   2412
      Left            =   5880
      TabIndex        =   4
      Top             =   6120
      Width           =   1932
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Command1"
      Height          =   1572
      Left            =   2040
      TabIndex        =   3
      Top             =   5880
      Width           =   2532
   End
   Begin VB.Label lblAnswer 
      Caption         =   "Label1"
      Height          =   852
      Left            =   6720
      TabIndex        =   7
      Top             =   4080
      Width           =   1212
   End
   Begin VB.Label lblHypotenuse 
      Caption         =   "Label1"
      Height          =   732
      Left            =   1920
      TabIndex        =   2
      Top             =   4080
      Width           =   3972
   End
   Begin VB.Label lblEnterB 
      Caption         =   "Label1"
      Height          =   852
      Left            =   960
      TabIndex        =   1
      Top             =   1920
      Width           =   3372
   End
   Begin VB.Label lblEnterA 
      Caption         =   "Label1"
      Height          =   612
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   3492
   End
End
Attribute VB_Name = "frmPythagoreanTheorem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jorge Monzon Diaz
'Date:   March 28, 2014
'Desc:   Chapter 4 Exercise 7 Pythagorean Theorem

Option Explicit

Private Sub cmdCalc_Click()
    Dim dblA, dblB, dblRawSum As Double
    dblA = Val(txtA)
    dblB = Val(txtB)
    dblRawSum = dblA ^ 2 + dblB ^ 2
    lblAnswer = Math.Sqr(dblRawSum)
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    frmPythagoreanTheorem.WindowState = 2
    lblEnterA = "Enter the length of side A:"
    lblEnterB = "Enter the length of side B:"
    txtA = " "
    txtB = " "
    lblHypotenuse = "The length of the hypotenuse is:"
    lblAnswer = " "
    cmdDone.Caption = "Done"
    cmdCalc.Caption = "Calculate Hypotenuse"
End Sub

Private Sub txtA_Change()
    lblAnswer = " "
End Sub

Private Sub txtB_Change()
    lblAnswer = " "
End Sub
