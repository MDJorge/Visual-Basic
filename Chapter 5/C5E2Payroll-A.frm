VERSION 5.00
Begin VB.Form frmC5E2Payroll 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Payroll"
   ClientHeight    =   4152
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   5628
   LinkTopic       =   "Form1"
   ScaleHeight     =   12396
   ScaleWidth      =   22824
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame fraExempt 
      BackColor       =   &H0080C0FF&
      Caption         =   "Exempt from taxes?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   5295
      Begin VB.OptionButton optNotTaxExempt 
         BackColor       =   &H0080C0FF&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton taxExempt 
         BackColor       =   &H0080C0FF&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   840
         TabIndex        =   2
         Top             =   1260
         Width           =   1455
      End
   End
   Begin VB.TextBox txtRate 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton btnCalc 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Calc Pay"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton btnDone 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox txtHours 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblPay 
      BackColor       =   &H000080FF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      TabIndex        =   9
      Top             =   6600
      Width           =   6375
   End
   Begin VB.Label lblPayMsg 
      BackColor       =   &H0080C0FF&
      Caption         =   "Pay is:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   26.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   8
      Top             =   6480
      Width           =   2415
   End
   Begin VB.Label lblRate 
      BackColor       =   &H0080C0FF&
      Caption         =   "Enter the Hourly Pay Rate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   7
      Top             =   1440
      Width           =   7695
   End
   Begin VB.Label lblInput 
      BackColor       =   &H0080C0FF&
      Caption         =   "Enter the number of Hours Worked"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   7695
   End
End
Attribute VB_Name = "frmC5E2Payroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AUTHOR:    Jorge Monzon Diaz
'DATE:      April 8, 2014
'DESC:       Ch.5-Ex.2 Payroll

Option Explicit

Private Sub btnCalc_Click()
    Dim dblHours, dblRate As Integer, dblOvertime, dblTaxes As Double
    If (IsNumeric(txtHours) And IsNumeric(txtRate)) Then
        dblHours = txtHours
        dblRate = txtRate
        If txtHours < 40 Then
            lblPay = dblHours * dblRate
        Else
            dblOvertime = (dblHours - 40) * (dblRate / 2)
            lblPay = dblHours * dblRate + dblOvertime
        End If
    Else
        MsgBox "Hours worked and hourly pay rate fields must contain valid numbers"
    End If
        
End Sub

Private Sub btnDone_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    frmC5E2Payroll.WindowState = 2  ' max form
    txtHours.TabIndex = 1           ' set cursor on field
    txtRate.TabIndex = 2
    btnCalc.Default = True          ' set ENTER key to calc button
    
    txtHours = ""                   ' clear form fields
    txtRate = ""
    lblPay = ""
End Sub

Private Sub txtHours_Change()
    lblPay = ""
End Sub

Private Sub txtRate_Change()
    lblPay = ""
End Sub
