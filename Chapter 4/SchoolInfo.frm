VERSION 5.00
Begin VB.Form frmSchoolInfo 
   Caption         =   "School Information"
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
      Height          =   1332
      Left            =   5760
      TabIndex        =   9
      Top             =   7800
      Width           =   2172
   End
   Begin VB.Frame fraSelect 
      Caption         =   "Frame1"
      Height          =   8652
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   4452
      Begin VB.OptionButton optYale 
         Caption         =   "Option8"
         Height          =   492
         Left            =   360
         TabIndex        =   8
         Top             =   7200
         Width           =   3252
      End
      Begin VB.OptionButton optUPenn 
         Caption         =   "Option7"
         Height          =   732
         Left            =   360
         TabIndex        =   7
         Top             =   6120
         Width           =   3372
      End
      Begin VB.OptionButton optPrinceton 
         Caption         =   "Option6"
         Height          =   852
         Left            =   360
         TabIndex        =   6
         Top             =   4920
         Width           =   3012
      End
      Begin VB.OptionButton optHarvard 
         Caption         =   "Option5"
         Height          =   732
         Left            =   360
         TabIndex        =   5
         Top             =   3840
         Width           =   3012
      End
      Begin VB.OptionButton optDartmouth 
         Caption         =   "Option4"
         Height          =   732
         Left            =   360
         TabIndex        =   4
         Top             =   2880
         Width           =   2652
      End
      Begin VB.OptionButton optCornell 
         Caption         =   "Option3"
         Height          =   612
         Left            =   240
         TabIndex        =   3
         Top             =   2040
         Width           =   3252
      End
      Begin VB.OptionButton optColumbia 
         Caption         =   "Option2"
         Height          =   492
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   3732
      End
      Begin VB.OptionButton optBrown 
         Caption         =   "Option1"
         Height          =   372
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   3252
      End
   End
   Begin VB.Label lblState 
      Caption         =   "Label1"
      Height          =   852
      Left            =   5760
      TabIndex        =   12
      Top             =   3840
      Width           =   2892
   End
   Begin VB.Label lblCity 
      Caption         =   "Label1"
      Height          =   612
      Left            =   5760
      TabIndex        =   11
      Top             =   2880
      Width           =   2772
   End
   Begin VB.Label lblLocation 
      Caption         =   "Label1"
      Height          =   372
      Left            =   5280
      TabIndex        =   10
      Top             =   2040
      Width           =   2292
   End
End
Attribute VB_Name = "frmSchoolInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jorge Monzon Diaz
'Date:   March 28, 2014
'Desc:   Chapter 4 Exercise 11 School Information

Option Explicit

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    frmSchoolInfo.WindowState = 2
    fraSelect = "Select a school"
    lblLocation = "Location:"
    lblCity = " "
    lblState = " "
    cmdDone.Caption = "Done"
    optBrown.Caption = "Brown University"
    optColumbia.Caption = "Columbia University"
    optCornell.Caption = "Cornell University"
    optDartmouth.Caption = "Dartmouth College"
    optHarvard.Caption = "Harvard University"
    optPrinceton.Caption = "Princeton University"
    optUPenn.Caption = "University of Pennsylvania"
    optYale.Caption = "Yale University"
End Sub

Private Sub optBrown_Click()
    lblCity = "Providence"
    lblState = "Rhode Island"
End Sub

Private Sub optColumbia_Click()
    lblCity = "New York"
    lblState = "New York"
End Sub

Private Sub optCornell_Click()
    lblCity = "Ithaca"
    lblState = "New York"
End Sub

Private Sub optDartmouth_Click()
    lblCity = "Hanover"
    lblState = "New Hampshire"
End Sub

Private Sub optHarvard_Click()
    lblCity = "Cambridge"
    lblState = "Massachusetts"
End Sub

Private Sub optPrinceton_Click()
    lblCity = "Princeton"
    lblState = "New Jersey"
End Sub

Private Sub optUPenn_Click()
    lblCity = "Philadelphia"
    lblState = "Pennsylvania"
End Sub

Private Sub optYale_Click()
    lblCity = "New Haven"
    lblState = "Connecticut"
End Sub
