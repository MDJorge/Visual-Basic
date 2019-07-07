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
      Left            =   8160
      TabIndex        =   0
      Top             =   5160
      Width           =   1692
   End
   Begin VB.Image imgIndian 
      Height          =   4428
      Left            =   2040
      Picture         =   "C3SchoolApplicationB.frx":0000
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   4560
   End
   Begin VB.Label lblSchoolMascot 
      Caption         =   "Label1"
      Height          =   612
      Left            =   3240
      TabIndex        =   2
      Top             =   6600
      Width           =   2652
   End
   Begin VB.Label lblSchoolName 
      Caption         =   "Label1"
      Height          =   852
      Left            =   3000
      TabIndex        =   1
      Top             =   600
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
'DESC:   Chapter 3 Exercise 2b School Application

Option Explicit

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    frmSchoolApplication.WindowState = 2
    lblSchoolName.Visible = False
    lblSchoolMascot.Visible = False
    cmdDone.Caption = "Done"
End Sub

Private Sub imgIndian_Click()
    lblSchoolName = "Nyack High School"
    lblSchoolMascot = "Nyack Indians"
    lblSchoolName.Visible = True
    lblSchoolMascot.Visible = True
End Sub
