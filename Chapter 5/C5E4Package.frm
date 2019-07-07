VERSION 5.00
Begin VB.Form frmC5E4Package 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Package Check"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00C0C000&
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H00C0C000&
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7560
      Width           =   1455
   End
   Begin VB.TextBox txtHeight 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
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
      Left            =   2760
      TabIndex        =   3
      Text            =   "txtHeight"
      Top             =   5040
      Width           =   2295
   End
   Begin VB.TextBox txtWidth 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
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
      Left            =   2760
      TabIndex        =   2
      Text            =   "txtWidth"
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox txtLength 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
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
      Left            =   2760
      TabIndex        =   1
      Text            =   "txtLength"
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
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
      Left            =   2760
      TabIndex        =   0
      Text            =   "txtWeight"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      Caption         =   "width in centimeters:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   7
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblAnswer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "lblAnswer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   9
      Top             =   6240
      Width           =   8295
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      Caption         =   "height in centimeters:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   8
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      Caption         =   "length in centimeters:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      Caption         =   "weight in kilograms:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Enter the package's data:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "frmC5E4Package"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AUTHOR: Jorge Monzon Diaz
'DATE:   April 8, 2014
'DESC:   C5 Ex 4 Package Delivery
 
 Option Explicit
 
' define all numeric variables
Dim intWeight, intLength, intWidth, intHeight As Integer, dblCubed As Double


Private Sub cmdCalc_Click()
    If IsNumeric(txtWeight) And IsNumeric(txtLength) And IsNumeric(txtWidth) And IsNumeric(txtHeight) Then
        intWeight = txtWeight
        intLength = txtLength
        intWidth = txtWidth
        intHeight = txtHeight
        If intWeight < 27 Then
            dblCubed = intLength * intWidth * intHeight
            If dblCubed < 100000 Then
                lblAnswer = "Accepted. Your package meets requirements."
            Else
                lblAnswer = "Rejected. Too big."
            End If
        Else
            lblAnswer = "Rejected. Too heavy."
        End If
    Else
        MsgBox "Please enter valid numbers"
    End If
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    frmC5E4Package.WindowState = 2
    txtWeight = ""
    txtLength = ""
    txtWidth = ""
    txtHeight = ""
    lblAnswer = ""
    cmdCalc.Default = True
End Sub


Private Sub txtHeight_Change()
    lblAnswer = ""
End Sub

Private Sub txtLength_Change()
    lblAnswer = ""
End Sub

Private Sub txtWeight_Change()
    lblAnswer = ""
End Sub

Private Sub txtWidth_Change()
    lblAnswer = ""
End Sub
