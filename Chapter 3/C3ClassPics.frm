VERSION 5.00
Begin VB.Form frmC3ClassPics 
   Caption         =   "Chapter 3 Class Pictures"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   BeginProperty Font 
      Name            =   "Lucida Sans"
      Size            =   13.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImg3 
      Appearance      =   0  'Flat
      Caption         =   "Show Picture"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   3
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3960
      Width           =   1692
   End
   Begin VB.CommandButton cmdImg2 
      Appearance      =   0  'Flat
      Caption         =   "Show Picture"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   2
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3960
      Width           =   1692
   End
   Begin VB.CommandButton cmdImg1 
      Appearance      =   0  'Flat
      Caption         =   "Show Picture"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   1
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   1692
   End
   Begin VB.CommandButton cmdDone 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   16.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   2412
   End
   Begin VB.CommandButton cmdImg0 
      Caption         =   "Show Picture"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   0
      Left            =   480
      MaskColor       =   &H80000000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   1692
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Class Pictures"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   19.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   480
      TabIndex        =   4
      Top             =   360
      Width           =   8172
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "NAME4"
      Height          =   252
      Index           =   3
      Left            =   6960
      TabIndex        =   3
      Top             =   1200
      Width           =   1692
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "NAME3"
      Height          =   252
      Index           =   2
      Left            =   4800
      TabIndex        =   2
      Top             =   1200
      Width           =   1692
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "NAME2"
      Height          =   252
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   1200
      Width           =   1572
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "NAME1"
      Height          =   252
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1692
   End
   Begin VB.Image imgPerson 
      Height          =   2052
      Index           =   0
      Left            =   480
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1692
   End
   Begin VB.Image imgPerson 
      Height          =   2052
      Index           =   2
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1692
   End
   Begin VB.Image imgPerson 
      Height          =   2052
      Index           =   1
      Left            =   2640
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1692
   End
   Begin VB.Image imgPerson 
      Height          =   2052
      Index           =   3
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1692
   End
End
Attribute VB_Name = "frmC3ClassPics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AUTHOR: Jorge Monzon Diaz
'DATE: March 13, 2014
'DESC: Chapter 3 Class Pics



Private Sub bgColors()
    'Put all color properties in one place
    'RGB is used since it has a larger variety of colors
    frmC3ClassPics.BackColor = RGB(241, 241, 241)
    'Labels
    lblTitle.FontUnderline = True
    lblTitle.ForeColor = RGB(221, 77, 59)
    lblName(0).BackStyle = 0
    lblName(1).BackStyle = 0
    lblName(2).BackStyle = 0
    lblName(3).BackStyle = 0
    lblName(0).ForeColor = RGB(123, 168, 255)
    lblName(1).ForeColor = RGB(123, 168, 255)
    lblName(2).ForeColor = RGB(123, 168, 255)
    lblName(3).ForeColor = RGB(123, 168, 255)
    'Command Buttons
    cmdImg0(0).BackColor = RGB(87, 135, 225)
    cmdImg1(1).BackColor = RGB(93, 93, 181)
    cmdImg2(2).BackColor = RGB(91, 162, 253)
    cmdImg3(3).BackColor = RGB(64, 102, 175)
    cmdDone.BackColor = RGB(101, 179, 84)
End Sub

Private Sub cmdDone_Click()
    frmC3ClassPics.Hide
End Sub

Private Sub cmdImg0_Click(Index As Integer)
    imgPerson(0).Picture = LoadPicture("C:\Users\Jorge\Desktop\test.jpg")
    imgPerson(1).Picture = LoadPicture("")
    imgPerson(2).Picture = LoadPicture("")
    imgPerson(3).Picture = LoadPicture("")
End Sub

Private Sub cmdImg1_Click(Index As Integer)
    imgPerson(0).Picture = LoadPicture("")
    imgPerson(1).Picture = LoadPicture("C:\Users\Jorge\Desktop\test.jpg")
    imgPerson(2).Picture = LoadPicture("")
    imgPerson(3).Picture = LoadPicture("")
End Sub

Private Sub cmdImg2_Click(Index As Integer)
    imgPerson(0).Picture = LoadPicture("")
    imgPerson(1).Picture = LoadPicture("")
    imgPerson(2).Picture = LoadPicture("C:\Users\Jorge\Desktop\test.jpg")
    imgPerson(3).Picture = LoadPicture("")
End Sub

Private Sub cmdImg3_Click(Index As Integer)
    imgPerson(0).Picture = LoadPicture("")
    imgPerson(1).Picture = LoadPicture("")
    imgPerson(2).Picture = LoadPicture("")
    imgPerson(3).Picture = LoadPicture("C:\Users\Jorge\Desktop\test.jpg")
End Sub

Private Sub Form_Load()
    frmC3ClassPics.WindowState = 2 'Maximize window
    bgColors 'Calls the bgcolor function
    'Fixes the name captions
    lblName(0).Caption = "NAME"
    lblName(1).Caption = "NAME"
    lblName(2).Caption = "NAME"
    lblName(3).Caption = "NAME"
    imgPerson(0).Picture = LoadPicture("")
    imgPerson(1).Picture = LoadPicture("")
    imgPerson(2).Picture = LoadPicture("")
    imgPerson(3).Picture = LoadPicture("")
End Sub
