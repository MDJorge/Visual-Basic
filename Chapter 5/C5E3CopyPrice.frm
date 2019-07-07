VERSION 5.00
Begin VB.Form frmC5E3CopyPrice 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Printing Prices"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H0000C000&
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H0000C000&
      Caption         =   "Calculate"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox txtNoCopies 
      BackColor       =   &H0080FF80&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   3
      Text            =   "txtNoCopies"
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label lblPricePerCopy 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      Caption         =   "lblPricePerCopy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   7
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label lblTotPrice 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      Caption         =   "lblTotPrice"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   6
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label lblAns 
      BackColor       =   &H0080FF80&
      Caption         =   "The Total Price is $"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   5415
   End
   Begin VB.Label lblPrice 
      BackColor       =   &H0080FF80&
      Caption         =   "Price per Copy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   5415
   End
   Begin VB.Label lblCopies 
      BackColor       =   &H0080FF80&
      Caption         =   "Enter the Number of Copies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5415
   End
End
Attribute VB_Name = "frmC5E3CopyPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author:   Jorge Monzon Diaz
'Date:     April 8, 2014
'Desc:     c5e3 Printing

Option Explicit

Dim intCopies As Integer

Private Sub cmdCalc_Click()
    If IsNumeric(txtNoCopies) Then
        intCopies = Val(txtNoCopies)
        If (intCopies >= 0) And (intCopies <= 499) Then
            lblPricePerCopy = 0.3
            lblTotPrice = intCopies * 0.3
        ElseIf (intCopies >= 500) And (intCopies <= 749) Then
            lblPricePerCopy = 0.28
            lblTotPrice = intCopies * 0.28
        ElseIf (intCopies >= 750) And (intCopies <= 999) Then
            lblPricePerCopy = 0.27
            lblTotPrice = intCopies * 0.27
        ElseIf (intCopies >= 1000) Then
            lblPricePerCopy = 0.25
            lblTotPrice = intCopies * 0.25
        Else
            MsgBox "Enter a number greater than 0"
        End If
    Else
        MsgBox "Enter a valid number"
    End If
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtNoCopies = ""
    lblPricePerCopy = ""
    lblTotPrice = ""
End Sub

Private Sub txtNoOfCopies_Change()
    lblPrice = Nothing
    lblPricePerCopy = Nothing
    lblAns = Nothing
    lblTotPrice = Nothing
End Sub

