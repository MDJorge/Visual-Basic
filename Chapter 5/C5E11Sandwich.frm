VERSION 5.00
Begin VB.Form frmC5E11Sandwich 
   BackColor       =   &H000000FF&
   Caption         =   "Sandwich Order"
   ClientHeight    =   3084
   ClientLeft      =   252
   ClientTop       =   456
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   12396
   ScaleWidth      =   22824
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10200
      TabIndex        =   13
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      TabIndex        =   12
      Top             =   8280
      Width           =   1695
   End
   Begin VB.Frame frmMethod 
      BackColor       =   &H000000FF&
      Caption         =   "Sandwich Size"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1440
      TabIndex        =   9
      Top             =   2280
      Width           =   6975
      Begin VB.OptionButton optLarge 
         BackColor       =   &H000000FF&
         Caption         =   "Large"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   11
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optSmall 
         BackColor       =   &H000000FF&
         Caption         =   "Small"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.CheckBox chkMayo 
      BackColor       =   &H000000FF&
      Caption         =   "Mayo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9840
      TabIndex        =   5
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CheckBox chkCheese 
      BackColor       =   &H000000FF&
      Caption         =   "Cheese"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7320
      TabIndex        =   4
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CheckBox chkOnion 
      BackColor       =   &H000000FF&
      Caption         =   "Onion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CheckBox chkMustard 
      BackColor       =   &H000000FF&
      Caption         =   "Mustard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   2
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CheckBox chkTomato 
      BackColor       =   &H000000FF&
      Caption         =   "Tomato"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CheckBox chkLettuce 
      BackColor       =   &H000000FF&
      Caption         =   "Lettuce"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H000000FF&
      Caption         =   "Total Price = $"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1440
      TabIndex        =   14
      Top             =   6840
      Width           =   3855
   End
   Begin VB.Label lblAnswer 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      TabIndex        =   8
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Label lblQuestion 
      BackColor       =   &H000000FF&
      Caption         =   "Select Sandwich Options"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1440
      TabIndex        =   7
      Top             =   960
      Width           =   6735
   End
   Begin VB.Label lblRegular 
      BackColor       =   &H000000FF&
      Caption         =   "Fixings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   6
      Top             =   4800
      Width           =   2655
   End
End
Attribute VB_Name = "frmC5E11Sandwich"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AUTHOR:    Jorge Monzon Diaz
'DATE:      April 12, 2014
'DESC:      Chapter 5 Exercise 11 Sandwich Order

Option Explicit

Const SMALLSANDWICH As Double = 2.5
Const LARGESANDWICH As Integer = 4
Const MUSTARD As Integer = 0
Const MAYO As Integer = 0
Const ONION As Double = 0.1
Const LETTUCE As Double = 0.1
Const TOMATO As Double = 0.25
Const CHEESE As Double = 0.5
Dim curTotal As Currency, intCounter As Integer

Private Sub chkCheese_Click()
    If chkCheese Then
        curTotal = curTotal + CHEESE
    Else
        curTotal = curTotal - CHEESE
    End If
    lblAnswer = curTotal
End Sub

Private Sub chkLettuce_Click()
    If chkLettuce Then
        curTotal = curTotal + LETTUCE
    Else
        curTotal = curTotal - LETTUCE
    End If
    lblAnswer = curTotal
End Sub

Private Sub chkMayo_Click()
    If chkMayo Then
        curTotal = curTotal + MAYO
    Else
        curTotal = curTotal - MAYO
    End If
    lblAnswer = curTotal
End Sub

Private Sub chkMustard_Click()
    If chkMustard Then
        curTotal = curTotal + MUSTARD
    Else
        curTotal = curTotal - MUSTARD
    End If
    lblAnswer = curTotal
End Sub

Private Sub chkOnion_Click()
    If chkOnion Then
        curTotal = curTotal + ONION
    Else
        curTotal = curTotal - ONION
    End If
    lblAnswer = curTotal
End Sub

Private Sub chkTomato_Click()
    If chkTomato Then
        curTotal = curTotal + TOMATO
    Else
        curTotal = curTotal - TOMATO
    End If
    lblAnswer = curTotal
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Me.PrintForm
End Sub

Private Sub Form_Load()
    frmC5E11Sandwich.WindowState = 2
    cmdPrint.Default = True
    cmdPrint.TabIndex = 1
    cmdExit.TabIndex = 2
    lblAnswer = curTotal
End Sub

Private Sub optLarge_Click()
    If intCounter = 0 Then
        curTotal = curTotal + LARGESANDWICH
    Else
        curTotal = curTotal - SMALLSANDWICH
        curTotal = curTotal + LARGESANDWICH
    End If
    intCounter = intCounter + 1
    lblAnswer = curTotal
End Sub

Private Sub optSmall_Click()
    If intCounter = 0 Then
        curTotal = curTotal + SMALLSANDWICH
    Else
        curTotal = curTotal - LARGESANDWICH
        curTotal = curTotal + SMALLSANDWICH
    End If
    intCounter = intCounter + 1
    lblAnswer = curTotal
End Sub
