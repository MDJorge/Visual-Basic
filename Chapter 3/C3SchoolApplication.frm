VERSION 5.00
Begin VB.Form frmSchoolApplication 
   Caption         =   "School"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDone 
      Caption         =   "FILLER"
      Height          =   1332
      Left            =   5400
      TabIndex        =   0
      Top             =   4080
      Width           =   1692
   End
   Begin VB.Label lblSchoolMascot 
      Caption         =   "Label1"
      Height          =   612
      Left            =   1320
      TabIndex        =   2
      Top             =   2160
      Width           =   2652
   End
   Begin VB.Label lblSchoolName 
      Caption         =   "Label1"
      Height          =   852
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   2292
   End
End
Attribute VB_Name = "frmSchoolApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AUTHOR: Jorge Monzon Diaz
'DATE:   March 21, 2014
'DESC:   Chapter 3 Exercise 2a School Application

Option Explicit

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    frmSchoolApplication.WindowState = 2
    lblSchoolName = "Nyack High School"
    lblSchoolMascot = "Nyack Indians"
    cmdDone.Caption = "Done"
End Sub
