VERSION 5.00
Begin VB.Form frmInputBox 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Input Box Practice"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnOk 
      Caption         =   "Click to enter Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "You will be prompted to enter grades"
      Height          =   1815
      Left            =   960
      TabIndex        =   3
      Top             =   5400
      Width           =   3375
   End
   Begin VB.Label lblName 
      BackColor       =   &H0000FFFF&
      Height          =   615
      Left            =   5160
      TabIndex        =   2
      Top             =   1800
      Width           =   7935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "My name is:"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1800
      Width           =   3375
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author:    Jorge Monzon Diaz
'DATE:      June 4, 2014
'DESC:      Test using Input Box

' define variables
Dim strName As String
Dim noOfGrades As Integer, ctr As Integer, grade As Integer, totGrades As Integer
Dim gradeAverage As Double

Private Sub btnOk_Click()
    
    ' display INPUTBOX to get name
    strName = InputBox("Enter Name", "This is an Input Box")
    
    If strName = "" Then
        lblName = "No name entered"
    Else
        lblName = strName
    End If
    
    noOfGrades = InputBox("How many grades do you have?", "Number of grades")
    
    'Do While ctr < noOfGrades
        'ctr = ctr + 1
        'grade = InputBox("Enter the grade", "Grades")
        'totGrades = totGrades + grade
    'Loop
    'Do Until ctr > noOfGrades
        'ctr = ctr + 1
        'grade = InputBox("Enter the grade", "Grades")
        'totGrades = totGrades + grade
    'Loop
    For ctr = 1 To noOfGrades Step 1
        grade = InputBox("Enter the grade", "Grades")
        totGrades = totGrades + grade
    Next
    gradeAverage = totGrades / noOfGrades
    MsgBox "Your grade average is " & gradeAverage & "."
End Sub

Private Sub Form_Load()
    lblName = ""
End Sub
