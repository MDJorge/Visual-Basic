VERSION 5.00
Begin VB.Form frmDigitsOfANumber 
   Caption         =   "Digits of a Number"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDone 
      Caption         =   "Command1"
      Height          =   1092
      Left            =   4920
      TabIndex        =   7
      Top             =   8400
      Width           =   1812
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Command1"
      Height          =   732
      Left            =   5040
      TabIndex        =   6
      Top             =   6960
      Width           =   1332
   End
   Begin VB.TextBox txtUserInput 
      Height          =   732
      Left            =   5520
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   1932
   End
   Begin VB.Label lblSecondDigitAnswer 
      Caption         =   "Label1"
      Height          =   852
      Left            =   5160
      TabIndex        =   5
      Top             =   5160
      Width           =   972
   End
   Begin VB.Label lblFirstDigitAnswer 
      Caption         =   "Label1"
      Height          =   852
      Left            =   5160
      TabIndex        =   4
      Top             =   3000
      Width           =   1092
   End
   Begin VB.Label lblSecondDigit 
      Caption         =   "Label1"
      Height          =   1452
      Left            =   840
      TabIndex        =   3
      Top             =   4680
      Width           =   3012
   End
   Begin VB.Label lblFirstDigit 
      Caption         =   "Label1"
      Height          =   1092
      Left            =   840
      TabIndex        =   2
      Top             =   2760
      Width           =   3372
   End
   Begin VB.Label lblPrompt 
      Caption         =   "Label1"
      Height          =   1332
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   4092
   End
End
Attribute VB_Name = "frmDigitsOfANumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jorge Monzon Diaz
'Date:   March 28, 2014
'Desc:   Chapter 4 Exercise 9

Option Explicit


Private Sub cmdCalc_Click()
    Dim intUserInput As Integer
    intUserInput = Val(txtUserInput)
    lblFirstDigitAnswer = intUserInput \ 10
    lblSecondDigitAnswer = intUserInput - (lblFirstDigitAnswer * 10)
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    frmDigitsOfANumber.WindowState = 2
    lblPrompt = "Enter a two-digit number:"
    lblFirstDigit = "The first digit is:"
    lblSecondDigit = "The second digit is:"
    cmdCalc.Caption = "Calculate"
    cmdDone.Caption = "Done"
    txtUserInput = " "
    lblFirstDigitAnswer = " "
    lblSecondDigitAnswer = " "
End Sub

Private Sub txtUserInput_Change()
    lblFirstDigitAnswer = " "
    lblSecondDigitAnswer = " "
End Sub
