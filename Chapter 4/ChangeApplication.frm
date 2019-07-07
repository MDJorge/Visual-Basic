VERSION 5.00
Begin VB.Form frmChangeApplication 
   Caption         =   "Change"
   ClientHeight    =   2325
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDone 
      Caption         =   "Command1"
      Height          =   852
      Left            =   7560
      TabIndex        =   7
      Top             =   8640
      Width           =   1212
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Command1"
      Height          =   972
      Left            =   4920
      TabIndex        =   6
      Top             =   8400
      Width           =   1692
   End
   Begin VB.TextBox txtUserInput 
      Height          =   852
      Left            =   6120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   1332
   End
   Begin VB.Image imgCoins 
      Height          =   2652
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   3372
   End
   Begin VB.Label lblPennyAmount 
      Caption         =   "Label1"
      Height          =   612
      Left            =   4560
      TabIndex        =   11
      Top             =   6720
      Width           =   1092
   End
   Begin VB.Label lblNickelAmount 
      Caption         =   "Label1"
      Height          =   852
      Left            =   4680
      TabIndex        =   10
      Top             =   5280
      Width           =   1212
   End
   Begin VB.Label lblDimeAmount 
      Caption         =   "Label1"
      Height          =   732
      Left            =   4800
      TabIndex        =   9
      Top             =   3720
      Width           =   1092
   End
   Begin VB.Label lblQuarterAmount 
      Caption         =   "Label1"
      Height          =   852
      Left            =   4800
      TabIndex        =   8
      Top             =   2280
      Width           =   1092
   End
   Begin VB.Label lblPennies 
      Caption         =   "Label1"
      Height          =   1092
      Left            =   1320
      TabIndex        =   5
      Top             =   6720
      Width           =   2172
   End
   Begin VB.Label lblNickels 
      Caption         =   "Label1"
      Height          =   972
      Left            =   1200
      TabIndex        =   4
      Top             =   5400
      Width           =   2172
   End
   Begin VB.Label lblDimes 
      Caption         =   "Label1"
      Height          =   972
      Left            =   1080
      TabIndex        =   3
      Top             =   3600
      Width           =   2172
   End
   Begin VB.Label lblQuarters 
      Caption         =   "Label1"
      Height          =   972
      Left            =   1200
      TabIndex        =   2
      Top             =   2160
      Width           =   2412
   End
   Begin VB.Label lblPrompt 
      Caption         =   "Label1"
      Height          =   612
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   4092
   End
End
Attribute VB_Name = "frmChangeApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jorge Monzon Diaz
'Date:   March 28, 2014
'Desc:   Chapter 4 Exercise 8 Change Application

Option Explicit

Private Sub cmdCalc_Click()
    imgCoins.Picture = LoadPicture("C:\Users\Jorge\Downloads\coins.jpg")
    Dim intCoins As Integer
    intCoins = Val(txtUserInput)
    lblQuarterAmount = intCoins \ 25
    lblDimeAmount = (intCoins - (lblQuarterAmount * 25)) \ 10
    lblNickelAmount = (intCoins - (lblQuarterAmount * 25) - (lblDimeAmount * 10)) \ 5
    lblPennyAmount = (intCoins - (lblQuarterAmount * 25) - (lblDimeAmount * 10) - (lblNickelAmount * 5)) \ 1
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    frmChangeApplication.WindowState = 2
    lblPrompt = "Enter the change in cents:"
    lblQuarters = "Quarters:"
    lblDimes = "Dimes:"
    lblNickels = "Nickels:"
    lblPennies = "Pennies:"
    cmdCalc.Caption = "Calculate"
    cmdDone.Caption = "Done"
    imgCoins.Picture = LoadPicture("")
    txtUserInput = " "
    lblQuarterAmount = " "
    lblDimeAmount = " "
    lblNickelAmount = " "
    lblPennyAmount = " "
End Sub

Private Sub txtUserInput_Change()
    lblQuarterAmount = " "
    lblDimeAmount = " "
    lblNickelAmount = " "
    lblPennyAmount = " "
End Sub
