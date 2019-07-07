VERSION 5.00
Begin VB.Form frmTimeConversion 
   Caption         =   "Time Conversion"
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
      Left            =   6480
      TabIndex        =   9
      Top             =   7560
      Width           =   1812
   End
   Begin VB.Frame fraSelect 
      Caption         =   "Frame1"
      Height          =   2652
      Left            =   600
      TabIndex        =   2
      Top             =   2400
      Width           =   7092
      Begin VB.OptionButton optHoursAndMinutes 
         Caption         =   "Option1"
         Height          =   1212
         Left            =   3960
         TabIndex        =   4
         Top             =   840
         Width           =   2652
      End
      Begin VB.OptionButton optSeconds 
         Caption         =   "Option1"
         Height          =   1092
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   2412
      End
   End
   Begin VB.TextBox txtUserInput 
      Height          =   1092
      Left            =   5520
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   1932
   End
   Begin VB.Label lblSeconds 
      Caption         =   "Label1"
      Height          =   1092
      Left            =   5760
      TabIndex        =   8
      Top             =   5640
      Width           =   2172
   End
   Begin VB.Label lblMinutes 
      Caption         =   "Label1"
      Height          =   972
      Left            =   7320
      TabIndex        =   7
      Top             =   5760
      Width           =   1572
   End
   Begin VB.Label lblHours 
      Caption         =   "Label1"
      Height          =   852
      Left            =   5280
      TabIndex        =   6
      Top             =   5760
      Width           =   1452
   End
   Begin VB.Label lblAnswerPrompt 
      Caption         =   "Label1"
      Height          =   1332
      Left            =   840
      TabIndex        =   5
      Top             =   5640
      Width           =   3972
   End
   Begin VB.Label lblPrompt 
      Caption         =   "Label1"
      Height          =   1092
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   4572
   End
End
Attribute VB_Name = "frmTimeConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jorge Monzon Diaz
'Date:   March 28, 2014
'Desc:   Chapter 4 Exercise 13 Time Conversion

Option Explicit

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    frmTimeConversion.WindowState = 2
    lblPrompt = "Enter the time in minutes"
    fraSelect = "Select time conversion"
    optSeconds.Caption = "Minutes to seconds"
    optHoursAndMinutes.Caption = "Minutes to hour:minute format"
    cmdDone.Caption = "Done"
    lblSeconds = " "
    lblMinutes = " "
    lblHours = " "
    lblAnswerPrompt = " "
    txtUserInput = " "
End Sub

Private Sub optHoursAndMinutes_Click()
    lblSeconds.Visible = False
    lblMinutes.Visible = True
    lblHours.Visible = True
    lblAnswerPrompt = "The time in hours:minute format is"
    Dim intUserInput As Integer
    intUserInput = Val(txtUserInput)
    lblHours = intUserInput \ 60
    lblMinutes = intUserInput - (lblHours * 60)
End Sub

Private Sub optSeconds_Click()
    lblSeconds.Visible = True
    lblMinutes.Visible = False
    lblHours.Visible = False
    lblAnswerPrompt = "The time in seconds is"
    Dim intUserInput As Integer
    intUserInput = Val(txtUserInput)
    lblSeconds = intUserInput * 60
End Sub
