VERSION 5.00
Begin VB.Form frmC5E6Auto 
   Caption         =   "Form1"
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
      Height          =   612
      Left            =   6240
      TabIndex        =   4
      Top             =   8640
      Width           =   2532
   End
   Begin VB.CommandButton cmdEvaluate 
      Caption         =   "Command1"
      Height          =   1572
      Left            =   6120
      TabIndex        =   3
      Top             =   6000
      Width           =   3492
   End
   Begin VB.TextBox txtInput 
      Height          =   1092
      Left            =   6000
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1200
      Width           =   3012
   End
   Begin VB.Image imgCar 
      Height          =   1812
      Left            =   600
      Picture         =   "C5E6Auto.frx":0000
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   4908
   End
   Begin VB.Label lblResult 
      Caption         =   "Label1"
      Height          =   1332
      Left            =   1440
      TabIndex        =   2
      Top             =   3600
      Width           =   7092
   End
   Begin VB.Label lblPrompt 
      Caption         =   "Label1"
      Height          =   1452
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   4812
   End
End
Attribute VB_Name = "frmC5E6Auto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author:    Jorge Monzon Diaz
'Date:      April 13, 2014
'Desc:      Chapter 5 Exercise 6 Defective Cars

Option Explicit

Dim intInput As Integer

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub cmdEvaluate_Click()
    If IsNumeric(txtInput) Then
        intInput = txtInput
        If intInput = 119 Or intInput = 179 Or intInput >= 189 And intInput <= 195 Or intInput = 221 Or intInput = 780 Then
            lblResult = "Your car is defective. Please have it fixed."
            imgCar.Visible = True
        Else
            lblResult = "Your car is not defective"
            imgCar.Visible = False
        End If
    Else
        MsgBox "Enter a valid number!"
    End If
End Sub

Private Sub Form_Load()
    frmC5E6Auto.WindowState = 2
    lblPrompt = "Enter your car model number:"
    txtInput = ""
    lblResult = ""
    cmdEvaluate.Caption = "Evaluate"
    cmdDone.Caption = "Done"
    imgCar.Visible = False
    cmdEvaluate.Default = True
    txtInput.TabIndex = 1
End Sub

Private Sub txtInput_Change()
    lblResult = ""
    imgCar.Visible = False
End Sub
