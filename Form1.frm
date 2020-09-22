VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Fishing Math"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5655
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   5655
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TxtHigh 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Text            =   "15"
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Highest number used"
      Height          =   195
      Left            =   1680
      TabIndex        =   5
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Divide"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   3
      Left            =   4200
      TabIndex        =   3
      Top             =   2640
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Multiply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Subtract"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   3720
      TabIndex        =   1
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin VB.Image Fish 
      Height          =   1335
      Index           =   1
      Left            =   3720
      Picture         =   "Form1.frx":030A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1530
   End
   Begin VB.Image Fish 
      Height          =   1185
      Index           =   2
      Left            =   0
      Picture         =   "Form1.frx":2614
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1395
   End
   Begin VB.Image Fish 
      Height          =   1080
      Index           =   0
      Left            =   0
      Picture         =   "Form1.frx":426E
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   1530
   End
   Begin VB.Image Fish 
      Height          =   1305
      Index           =   3
      Left            =   3720
      Picture         =   "Form1.frx":4A48
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Fish_Click(Index As Integer)
  iHighv = TxtHigh.Text
  iMathcase = Index + 1
  frmFish.Show
  Unload Me
End Sub

Private Sub Form_Load()
  iHighv = 15
  iLowV = 0
  iMathcase = 1
  cSymbol(1) = "+"
  cSymbol(2) = "-"
  cSymbol(3) = "*"
  cSymbol(4) = "/"
  frmFish.Show
  Unload Me
End Sub

Private Sub Label1_Click(Index As Integer)
  iMathcase = Index + 1
  frmFish.Show
  Unload Me
End Sub
