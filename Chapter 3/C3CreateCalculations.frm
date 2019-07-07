VERSION 5.00
Begin VB.Form frmCreateCalculations 
   Caption         =   "Create Calculations"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSubtraction 
      Caption         =   "Subtract"
      Height          =   732
      Left            =   1200
      TabIndex        =   4
      Top             =   4440
      Width           =   1812
   End
   Begin VB.CommandButton cmdAddition 
      Caption         =   "Add"
      Height          =   732
      Left            =   1200
      TabIndex        =   3
      Top             =   6000
      Width           =   1812
   End
   Begin VB.CommandButton cmdMultiplication 
      Caption         =   "Multiply"
      Height          =   732
      Left            =   1200
      TabIndex        =   2
      Top             =   2880
      Width           =   1812
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Finish"
      Height          =   972
      Left            =   8160
      TabIndex        =   1
      Top             =   5880
      Width           =   1812
   End
   Begin VB.CommandButton cmdDivision 
      Caption         =   "Divide"
      Height          =   732
      Left            =   1200
      TabIndex        =   0
      Top             =   1320
      Width           =   1812
   End
   Begin VB.Label lblSum 
      Caption         =   "Label1"
      Height          =   612
      Left            =   3840
      TabIndex        =   8
      Top             =   6120
      Width           =   1452
   End
   Begin VB.Label lblDifference 
      Caption         =   "Label1"
      Height          =   612
      Left            =   3840
      TabIndex        =   7
      Top             =   4560
      Width           =   1452
   End
   Begin VB.Label lblProduct 
      Caption         =   "Label1"
      Height          =   612
      Left            =   3720
      TabIndex        =   6
      Top             =   3120
      Width           =   1452
   End
   Begin VB.Label lblQuotient 
      Caption         =   "Label1"
      Height          =   612
      Left            =   3600
      TabIndex        =   5
      Top             =   1560
      Width           =   1452
   End
End
Attribute VB_Name = "frmCreateCalculations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AUTHOR: Jorge Monzon Diaz
'Date:   March 21, 2014
'Desc:   Chapter 3 Exercise 8 Create Calculations

Option Explicit


Private Sub cmdAddition_Click()
    lblSum = CInt(6 + 5)
End Sub

Private Sub cmdDivision_Click()
    lblQuotient = CDbl(6 / 5)
End Sub

Private Sub cmdMultiplication_Click()
    lblProduct = CInt(6 * 5)
End Sub

Private Sub cmdSubtraction_Click()
    lblDifference = CInt(6 - 5)
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    frmCreateCalculations.WindowState = 2
    cmdDivision.Caption = "6 / 5 ="
    cmdMultiplication.Caption = "6 * 5 ="
    cmdAddition.Caption = "6 + 5 ="
    cmdSubtraction.Caption = "6 - 5 ="
    cmdDone.Caption = "Done"
    lblSum = " "
    lblQuotient = " "
    lblProduct = " "
    lblDifference = " "
End Sub
