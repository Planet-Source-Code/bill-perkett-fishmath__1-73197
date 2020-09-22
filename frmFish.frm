VERSION 5.00
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "Vtext.dll"
Begin VB.Form frmFish 
   Caption         =   "Fish"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10290
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFish.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmFish.frx":030A
   ScaleHeight     =   7155
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimGameFly 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7320
      Top             =   960
   End
   Begin VB.Timer TimGame 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   6600
      Top             =   1080
   End
   Begin VB.Timer TimFoe2 
      Enabled         =   0   'False
      Interval        =   180
      Left            =   6000
      Top             =   120
   End
   Begin VB.Timer TimerFly2 
      Left            =   4440
      Top             =   1440
   End
   Begin HTTSLibCtl.TextToSpeech TextToSpeech1 
      Height          =   495
      Left            =   360
      OleObjectBlob   =   "frmFish.frx":F07D4
      TabIndex        =   20
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer TimFoe 
      Enabled         =   0   'False
      Interval        =   220
      Left            =   5880
      Top             =   840
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.Timer Timer5 
      Left            =   6960
      Top             =   6120
   End
   Begin VB.Timer Timer4 
      Left            =   4080
      Top             =   480
   End
   Begin VB.Timer TimerLoose 
      Left            =   1080
      Top             =   1560
   End
   Begin VB.Timer Timerpause 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   3720
      Top             =   1560
   End
   Begin VB.Timer TimerFly 
      Enabled         =   0   'False
      Left            =   1920
      Top             =   1560
   End
   Begin VB.Timer Timer3 
      Left            =   3120
      Top             =   480
   End
   Begin VB.Timer Timer2 
      Left            =   2280
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Left            =   1560
      Top             =   480
   End
   Begin VB.Label LblRed 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Ouch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   19
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   495
      Left            =   5400
      TabIndex        =   18
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label LblYesNo 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8280
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label LblDisplay 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8280
      TabIndex        =   16
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   495
      Left            =   6840
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image ImgFoe 
      Height          =   1830
      Index           =   3
      Left            =   7080
      Picture         =   "frmFish.frx":F07F8
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Image ImgFoe 
      Height          =   1830
      Index           =   2
      Left            =   4920
      Picture         =   "frmFish.frx":F7482
      Stretch         =   -1  'True
      Top             =   6240
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Image ImgFoe 
      Height          =   1830
      Index           =   1
      Left            =   2880
      Picture         =   "frmFish.frx":FE10C
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Image ImgFoe 
      Height          =   1830
      Index           =   0
      Left            =   600
      Picture         =   "frmFish.frx":104D96
      Stretch         =   -1  'True
      Top             =   6000
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Lblf 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   3
      Left            =   8400
      TabIndex        =   13
      Top             =   6120
      Width           =   945
   End
   Begin VB.Label Lblf 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   4440
      TabIndex        =   12
      Top             =   4080
      Width           =   945
   End
   Begin VB.Label Lblf 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2760
      TabIndex        =   11
      Top             =   3480
      Width           =   945
   End
   Begin VB.Label Lblf 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF80FF&
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   6360
      TabIndex        =   10
      Top             =   5040
      Width           =   945
   End
   Begin VB.Image Fish 
      Height          =   1305
      Index           =   3
      Left            =   7680
      Picture         =   "frmFish.frx":10BA20
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   2040
   End
   Begin VB.Image Fish 
      Height          =   1185
      Index           =   2
      Left            =   3960
      Picture         =   "frmFish.frx":10C492
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1395
   End
   Begin VB.Image Fish 
      Height          =   1335
      Index           =   1
      Left            =   2160
      Picture         =   "frmFish.frx":10E0EC
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1530
   End
   Begin VB.Image Fish 
      Height          =   1080
      Index           =   0
      Left            =   5640
      Picture         =   "frmFish.frx":1103F6
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1530
   End
   Begin VB.Label LblGreen 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label7"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Hook 
      Height          =   225
      Left            =   4850
      Picture         =   "frmFish.frx":110BD0
      Top             =   2520
      Width           =   195
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      X1              =   4875
      X2              =   4875
      Y1              =   2220
      Y2              =   2550
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   6240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image Fish 
      Height          =   900
      Index           =   8
      Left            =   7920
      Picture         =   "frmFish.frx":110E6A
      Top             =   4440
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label LblScore 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Score"
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
      Left            =   8280
      TabIndex        =   4
      Top             =   0
      Width           =   720
   End
   Begin VB.Image ImageBoat 
      Height          =   2805
      Left            =   480
      Picture         =   "frmFish.frx":11798C
      Stretch         =   -1  'True
      Top             =   165
      Width           =   4485
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   5640
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image Fish 
      Height          =   465
      Index           =   7
      Left            =   5160
      Picture         =   "frmFish.frx":140B3A
      Top             =   5760
      Width           =   1155
   End
   Begin VB.Image Fish 
      Height          =   465
      Index           =   6
      Left            =   7320
      Picture         =   "frmFish.frx":1412EA
      Top             =   3480
      Width           =   1155
   End
   Begin VB.Image Fish 
      Height          =   600
      Index           =   5
      Left            =   2280
      Picture         =   "frmFish.frx":141ACF
      Top             =   4680
      Width           =   1170
   End
   Begin VB.Image Fish 
      Height          =   600
      Index           =   4
      Left            =   8400
      Picture         =   "frmFish.frx":1422B4
      Top             =   4920
      Width           =   1170
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   2775
   End
End
Attribute VB_Name = "frmFish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iCurrentPlayer As Integer
Private iAnswer As Integer
Private iHooked As Integer
Private jHooked As Integer
Private iFishMove As Integer
Private iScore(8) As Integer
Private iLooseFish(8) As Integer
Private iMaxWins As Integer
Dim IFish(8) As Integer
Dim iFISHTop(8) As Integer
Dim iFoecnt As Integer
Private Sub Game()
'Dim i, j, ival, iMax, iTry, iUse(5) As Integer
'  'Command1.Visible = False
'   'CmdPlayer.Visible = True
'   Timer5.Enabled = False
'   Frame1.Visible = True
'   CmdEnd.Visible = True
'PickQuestion:
' LblQuestion.Visible = True
' LblDisplay = "Player " & iCurrentPlayer '& Chr(13) & " Enter Answer"
' ival = Int((Rnd * iQuizcnt) + 1)
' If ival > iQuizcnt Then ival = 1
' j = Len(cQuiz(ival, 0))
' LblQuestion = Mid(cQuiz(ival, 0), 2, j - 1)
' If cQuiz(ival, 6) = "Y" Then GoTo PickQuestion
' cQuiz(ival, 6) = "Y"
' iDone = iDone + 1
' For i = 1 To 4
'   ComLetter(i - 1).Visible = False
'   LblAnswer(i - 1).Visible = False
'   iUse(i) = 0
'  If Len(cQuiz(ival, i)) > 1 Then iMax = i
' Next
' '
' 'Scramble answers up
' iTry = Int((Rnd * iMax) + 1)
' If iTry > iMax Then ival = 1
' iUse(iTry) = 1
' i = 1
' j = Len(cQuiz(ival, iTry))
' LblAnswer(i - 1).Visible = True
' ComLetter(i - 1).Visible = True
' LblAnswer(i - 1) = Mid(cQuiz(ival, iTry), 2, j - 1)
' If UCase(Mid(cQuiz(ival, iTry), 1, 1)) = "Y" Then
'      iAnswer = i - 1
' End If
' 'CmdStart.SetFocus
'Nextanswer:
' i = i + 1
'TryAgain:
' iTry = Int((Rnd * iMax) + 1)
' If iTry > iMax Then ival = 1
' If iUse(iTry) > 0 Then GoTo TryAgain
' iUse(iTry) = 1
' LblAnswer(i - 1).Visible = True
' ComLetter(i - 1).Visible = True
' j = Len(cQuiz(ival, iTry))
' LblAnswer(i - 1) = Mid(cQuiz(ival, iTry), 2, j - 1)
' If UCase(Mid(cQuiz(ival, iTry), 1, 1)) = "Y" Then
'      iAnswer = i - 1
' End If
' If i < iMax Then GoTo Nextanswer
' LblBible.Caption = ival & " " & cQuiz(ival, 5)
 
End Sub

Public Sub PickFoe()
   Dim i As Integer
   For i = 0 To 3
    ImgFoe(i).Visible = False
    ImgFoe(i).Top = 6120
   Next
DoAgain:
   i = Rnd * 1800
   iFoe = (i Mod 4) + 1
   If iFoeold = iFoe Then GoTo DoAgain
   iFoeold = iFoe
   TimFoe.Enabled = True
   ImgFoe(iFoe - 1).Visible = True
End Sub

Public Sub PickFoe2()
   Dim i As Integer
   iGamecnt = iGamecnt + 1
   Label6.Caption = iGamecnt
   For i = 0 To 3
    ImgFoe(i).Visible = False
    ImgFoe(i).Top = 6120
   Next
DoAgain:
   i = Rnd * 1800
   iFoe = (i Mod 4) + 1
   If iFoeold = iFoe Then GoTo DoAgain
   iFoeold = iFoe
   TimFoe2.Enabled = True
   ImgFoe(iFoe - 1).Visible = True
End Sub



Private Sub Command1_Click()
   Dim i As Integer
   For i = 0 To 3
    ImgFoe(i).Visible = False
   Next
   iFoecnt = 0
   Command1.Visible = False
'   LblScore.Visible = False
'   LblDisplay.Visible = False
'   Exit Sub
   iMaxWins = 0
   jHooked = 0
   iHooked = 0
   iFoe = 0
   Timer1.Enabled = True
   Timer1.Interval = 500
   Command5.Visible = False
   Command5_Click
'   Timer2.Enabled = True
'   Timer2.Interval = 200
'   Timer3.Enabled = True
'   Timer3.Interval = 100
'   Timer4.Enabled = True
'   Timer4.Interval = 100
'   Timer5.Interval = 150
'   Timer5.Enabled = True
   Fish(0).Visible = True
   Label3 = Hook.Left & " " & Hook.Top
   'Command1.Visible = False
   ' Command2.Visible = False
End Sub



Private Sub Command2_Click()
 '  'PickFoe
'  TimerFly2.Enabled = True
'  TimerFly2.Interval = 600
'  Timer1.Enabled = False
   TimGame.Enabled = True
   Command2.Visible = False
   'If TimFoe2.Enabled = False Then PickFoe2
   'If iFoecnt > 15 And TimFoe2.Enabled = False Then PickFoe2
    'iFoecnt = iFoecnt + 1
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyL And ImageBoat.Left < 5400 Then
      ImageBoat.Left = ImageBoat.Left + 100
      Label1.Caption = ImageBoat.Left
      Line1.X1 = Line1.X1 + 100
      Line1.X2 = Line1.X1
      Hook.Left = Hook.Left + 100
    End If
    If KeyCode = vbKeyK And ImageBoat.Left > 10 Then
      ImageBoat.Left = ImageBoat.Left - 100
      Line1.X1 = Line1.X1 - 100
      Line1.X2 = Line1.X1
      Hook.Left = Hook.Left - 100
    End If
    Label7.Caption = ImgFoe(iFoe - 1).Left & " " & (ImageBoat.Left + ImageBoat.Width - 1700) & " " & ImageBoat.Left
    Label8.Caption = "ok"
    If ImgFoe(iFoe - 1).Top < 2700 Then
      If (ImageBoat.Left + ImageBoat.Width - 1700) > ImgFoe(iFoe - 1).Left And _
       ImageBoat.Left < ImgFoe(iFoe - 1).Left Then Label8.Caption = "HIT"
    End If
End Sub








Private Sub Command5_Click()
  Dim dNum As Double
  Dim iMess As Integer
  Dim ilow As Integer
  Dim iHi As Integer
 '
 ' Standard problem
  iFoe = 0
  TimFoe.Enabled = False
  iMess = iHighv - iLowV
  dNum = (iMess * Rnd) + 0.4
  ilow = Int((dNum) + iLowV)
  iMess = iHighv - iLowV
  dNum = (iMess * Rnd) + 0.4
  iHi = Int((dNum) + iLowV)
 ' Problem from a math file
 
 Select Case iMathcase
 Case 1
   iAnswer = ilow + iHi
  Case 2
   If iHi > ilow Then
     iMess = iHi
     iHi = ilow
     ilow = iMess
   End If
   iAnswer = ilow - iHi
  Case 3
   iAnswer = ilow * iHi
  Case 4
   iMess = ilow * iHi
   iAnswer = iMess / iHi
   ilow = iMess
 End Select
 Label2.Caption = " " & ilow & " " & cSymbol(iMathcase) _
   & " " & iHi & " = " & iAnswer
 LblDisplay = " " & ilow & " " & cSymbol(iMathcase) _
   & " " & iHi & " = "
  Label4.Caption = LblDisplay.Caption
  TextToSpeech1.Speak (Label4.Caption)
 'TxtAnswer.SetFocus
 iMess = 600 * Rnd
 ilow = iMess Mod 8
 Select Case ilow
  Case 0
   Lblf(0) = " " & iAnswer & " "
   Lblf(1) = " " & iAnswer + 2 & " "
   Lblf(2) = " " & iAnswer + 1 & " "
   Lblf(3) = " " & iAnswer - 1 & " "
  Case 1
   Lblf(1) = " " & iAnswer & " "
   Lblf(0) = " " & iAnswer + 2 & " "
   Lblf(3) = " " & iAnswer + 1 & " "
   Lblf(2) = " " & iAnswer - 1 & " "
  Case 2
   Lblf(3) = " " & iAnswer & " "
   Lblf(1) = " " & iAnswer + 2 & " "
   Lblf(2) = " " & iAnswer + 1 & " "
   Lblf(0) = " " & iAnswer - 1 & " "
  Case 3
   Lblf(2) = " " & iAnswer & " "
   Lblf(3) = " " & iAnswer + 2 & " "
   Lblf(0) = " " & iAnswer + 1 & " "
   Lblf(1) = " " & iAnswer - 1 & " "
  Case 4
   Lblf(0) = " " & iAnswer & " "
   Lblf(1) = " " & iAnswer + 2 & " "
   Lblf(3) = " " & iAnswer + 1 & " "
   Lblf(2) = " " & iAnswer + 3 & " "
  Case 5
   Lblf(1) = " " & iAnswer & " "
   Lblf(2) = " " & iAnswer + 2 & " "
   Lblf(3) = " " & iAnswer + 1 & " "
   Lblf(0) = " " & iAnswer + 3 & " "
  Case 6
   Lblf(2) = " " & iAnswer & " "
   Lblf(0) = " " & iAnswer + 2 & " "
   Lblf(1) = " " & iAnswer + 1 & " "
   Lblf(3) = " " & iAnswer + 3 & " "
  Case 7
   Lblf(3) = " " & iAnswer & " "
   Lblf(0) = " " & iAnswer + 2 & " "
   Lblf(1) = " " & iAnswer + 1 & " "
   Lblf(2) = " " & iAnswer - 1 & " "
 End Select
 
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cMsg As String
    Dim j As Integer
    Dim k As Integer
    Dim i As Integer
    If Timerpause.Enabled Then Exit Sub
    If TimerFly2.Enabled Then Exit Sub
    If TimGameFly.Enabled Then Exit Sub
    If TimerFly.Enabled = False Then
      
      Label1.Caption = KeyCode
    '
    ' up=38  down=40 8520
    '
    If KeyCode = 38 And 2520 < Line1.Y2 Then Line1.Y2 = Line1.Y2 - 80
    If KeyCode = 40 And Line1.Y2 < 6160 Then Line1.Y2 = Line1.Y2 + 80
    If KeyCode = 39 And ImageBoat.Left < 5400 Then
      ImageBoat.Left = ImageBoat.Left + 100
      Label1.Caption = ImageBoat.Left
      Line1.X1 = Line1.X1 + 100
      Line1.X2 = Line1.X1
      Hook.Left = Hook.Left + 100
      LblRed.Left = ImageBoat.Left + 1170
    End If
    If KeyCode = 37 And ImageBoat.Left > 10 Then
      ImageBoat.Left = ImageBoat.Left - 100
      Line1.X1 = Line1.X1 - 100
      Line1.X2 = Line1.X1
      Hook.Left = Hook.Left - 100
      LblRed.Left = ImageBoat.Left + 1170
    End If
    
'     If KeyCode = 65 And 3120 < Fish(0).Top Then Fish(0).Top = Fish(0).Top - 80
'    If KeyCode = 90 And Fish(0).Top < 6120 Then Fish(0).Top = Fish(0).Top + 80
'    'Label3.Caption = Line1.y2
    For i = 0 To 3
     j = 640 + Fish(i).Top - Hook.Top
     'Label2.Caption = "TOP = " & j '& " " & k
     'LblScore.Caption = j
     If j < 481 And j > -1 Then
       k = 1350 + Fish(i).Left - Hook.Left
       Label2.Caption = j & " " & k
       If k > -1 And k < 1220 Then
         'Label2.Caption = " H I T"
         iHooked = i + 1
         Timer1.Enabled = False
         'TimerFly.Enabled = True
       End If
     End If
   Next
     Hook.Top = Line1.Y2
    If iHooked > 0 And jHooked = 0 Then
      Fish(iHooked - 1).Top = Hook.Top
      Lblf(iHooked - 1).Top = Hook.Top + 500
      LblDisplay.Caption = Label4.Caption & Lblf(iHooked - 1).Caption
      'If TimerLoose.Enabled = False Then TimerLoose.Enabled = True
      If Fish(iHooked - 1).Top < 2825 Then
           Hook.Top = 2520
           Line1.Y2 = 2550
           If TimGame.Enabled Then
             TimGameFly.Enabled = True
             Exit Sub
           End If
         If Trim(Lblf(iHooked - 1).Caption) = iAnswer Then
           j = Rnd * 800
           j = j Mod 4
           TimerFly.Enabled = True
           TimerLoose.Enabled = False
            LblYesNo.Caption = "Correct"
           If j = 1 Then LblYesNo.Caption = "Yes"
           If j = 2 Then LblYesNo.Caption = "Great"
           If j = 3 Then LblYesNo.Caption = "You are right"
            LblYesNo.Visible = True
            LblYesNo.BackColor = LblGreen.BackColor
            TextToSpeech1.Speak (LblDisplay.Caption)
            iWincnt = iWincnt + 1
         Else
            iHooked = -1
            LblDisplay = Label4.Caption
            j = Rnd * 800
            j = j Mod 4
            Timerpause.Enabled = True
            Timerpause.Interval = 2600
            LblYesNo.Caption = "Sorry"
            If j = 1 Then LblYesNo.Caption = "Wrong"
            If j = 2 Then LblYesNo.Caption = "Try again"
            If j = 3 Then LblYesNo.Caption = " No "
            LblYesNo.Visible = True
            LblYesNo.BackColor = LblRed.BackColor
            iScore(iCurrentPlayer) = iScore(iCurrentPlayer) - 1
            LblScore.Caption = "Score= " & iScore(iCurrentPlayer)
         End If
      End If
    End If
   End If
End Sub
Private Sub Form_Load()
   Dim i As Integer
   Dim cMsg As String
   'Imageo.Visible = False
   For i = 0 To 8
     Fish(i).Visible = False
     IFish(i) = 0
     iFISHTop(i) = 0
     iScore(i) = 0
     iLooseFish(i) = 35
   Next
   iMaxWins = 0
   iFishMove = 0
   iHooked = 0
   iFoeold = 0
   ImageBoat.Left = 510
   '
   '
   LblScore = "Score " & iScore(iCurrentPlayer)
   iCurrentPlayer = 1
   Randomize
   TimerFly2.Enabled = False
   TimerFly.Enabled = False
   TimerFly.Interval = 100
   TimerFly2.Interval = 100
   TimerLoose.Enabled = False
   TimerLoose.Interval = 200
   i = 8
   Command1_Click
End Sub

Private Sub Label4_Click()
If Timer5.Enabled Then
  Timer5.Enabled = False
Else
  Timer5.Interval = 150
  Timer5.Enabled = True
  Command1.Visible = False
  End If
End Sub

Private Sub LblDisplay_Click()
  Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim cMsg As String
   'Timer1.Enabled = False
   i = Rnd * 28
   If iFoecnt > 15 And TimFoe.Enabled = False And i > 25 Then PickFoe
   If iFoecnt > 35 And TimFoe.Enabled = False Then PickFoe
    iFoecnt = iFoecnt + 1
    Label8.Caption = iFoecnt
   For i = 0 To 3
    Fish(i).Visible = True
     If Fish(i).Left < 9000 Then '9000 Then '9000 Then
         Fish(i).Left = Fish(i).Left + 300 '3000
         Lblf(i).Left = Fish(i).Left + 500
     Else
      Fish(i).Left = 1000
      Lblf(i).Left = Fish(i).Left + 500
     End If
     j = 640 + Fish(i).Top - Hook.Top
     'LblScore.Caption = j
     If j < 481 And j > -1 Then
       k = 1350 + Fish(i).Left - Hook.Left
       'Label2.Caption = j & " " & k
       If k > -1 And k < 1220 Then
         'Label2.Caption = " H I T"
         iHooked = i + 1
         Timer1.Enabled = False
         'TimerFly.Enabled = True
       End If
     End If
     'if j
   Next
'     cMsg = ""
'     i = Fish(0).Top - Hook.Top + 8000
'     Label2.Caption = Fish(0).Left & " " & Fish(0).Top & " " & i
'     If i > 7120 And i < 8120 Then
'       j = Fish(0).Left - Hook.Left + 8000
'       If j > 6900 And j < 8400 Then
'          cMsg = " hit"
'          Timer1.Enabled = False
'          iHooked = 1
'           Timer2.Enabled = False
'          Timer3.Enabled = False
'          Timer4.Enabled = False
'          'TimerFly.Enabled = True
'          'Call GAME
'       End If
'          Label2.Caption = Fish(0).Left & " " & Fish(0).Top & " " & IFish(0) & " " & j & cMsg
'     End If
'   Else
'     IFish(0) = IFish(0) - 1
'     If IFish(0) = 0 Then Fish(0).Visible = True
'      'Label2.Caption = " ** " & iFISH(0) & " " & j & cMsg
'   End If
   'Label2.Caption = Fish(0).Left & " " & Fish(0).Top & " " & iFISH(0)
   
End Sub

Private Sub Timer2_Timer()
' Dim i As Integer
'   Dim j As Integer
'   Dim cMsg As String
'   If IFish(1) = 0 Then
'      Fish(1).Visible = True
'     If Fish(1).Left < 7000 Then '9000 Then '9000 Then
'         Fish(1).Left = Fish(1).Left + 300 '3000
'     Else
'      Fish(1).Left = 2000
''      i = Rnd * 1300
''      i = i Mod 15
'       iFISHTop(1) = iFISHTop(0) + 4
'       If iFISHTop(1) > 15 Then iFISHTop(1) = iFISHTop(1) - 15
'       Fish(1).Top = 3100 + 200 * iFISHTop(1)
'      i = Rnd * 1300
'      i = i Mod 20
'      IFish(1) = i
'      Fish(1).Visible = False
'     End If
'     cMsg = ""
'     i = Fish(1).Top - Hook.Top + 8000
'     'Label2.Caption = Fish(1).Left & " " & Fish(0).Top & " " & i
'     If i > 7120 And i < 8120 Then
'       j = Fish(1).Left - Hook.Left + 8000
'       If j > 6900 And j < 8400 Then
'          cMsg = " hit"
'          Timer2.Enabled = False
'           Timer1.Enabled = False
'          Timer3.Enabled = False
'          Timer4.Enabled = False
'           iHooked = 2
'          'Call GAME
'       End If
'          'Label2.Caption = Fish(0).Left & " " & Fish(0).Top & " " & iFISH(0) & " " & j & cMsg
'     End If
'   Else
'     IFish(1) = IFish(1) - 1
'     If IFish(1) = 0 Then Fish(1).Visible = True
'
'   End If
   'Label2.Caption = Fish(0).Left & " " & Fish(0).Top & " " & iFISH(0)
End Sub

Private Sub Timer3_Timer()
Dim i As Integer
   Dim j As Integer
   Dim cMsg As String
   If IFish(2) = 0 Then
      Fish(2).Visible = True
     If Fish(2).Left > 2000 Then '9000 Then '9000 Then
         Fish(2).Left = Fish(2).Left - 300 '3000
     Else
      Fish(2).Left = 7000
'      i = Rnd * 1300
'      i = i Mod 15
       iFISHTop(2) = iFISHTop(0) + 8
       If iFISHTop(2) > 15 Then iFISHTop(2) = iFISHTop(2) - 15
       Fish(2).Top = 3100 + 200 * iFISHTop(2)
      i = Rnd * 1300
      i = i Mod 20
      IFish(2) = i
      Fish(2).Visible = False
     End If
     cMsg = ""
     i = Fish(2).Top - Hook.Top + 8000
     'Label2.Caption = Fish(2).Left & " " & Fish(0).Top & " " & i
     If i > 7120 And i < 8120 Then
       j = Fish(2).Left - Hook.Left + 8000
       If j > 6900 And j < 8400 Then
          cMsg = " hit"
          Timer1.Enabled = False
          Timer2.Enabled = False
          Timer3.Enabled = False
          Timer4.Enabled = False
           iHooked = 3
          'Call GAME
       End If
          'Label2.Caption = Fish(0).Left & " " & Fish(0).Top & " " & iFISH(0) & " " & j & cMsg
     End If
   Else
     IFish(2) = IFish(2) - 1
     If IFish(2) = 0 Then Fish(2).Visible = True

   End If

End Sub

Private Sub Timer4_Timer()
  '
  ' Octopus
  '
   Dim i As Integer
   Dim j As Integer
   Dim cMsg As String
   If IFish(3) = 0 Then
      Fish(3).Visible = True
     If Fish(3).Left < 7000 Then   '9000 Then '9000 Then
         Fish(3).Left = Fish(3).Left + 300 '3000
         Imageo.Visible = True
     Else
      Fish(3).Left = 2000
'      i = Rnd * 1300
'      i = i Mod 15
       iFISHTop(3) = iFISHTop(0) + 12
       If iFISHTop(3) > 15 Then iFISHTop(3) = iFISHTop(3) - 15
       Fish(3).Top = 3100 + 200 * iFISHTop(3)
      i = Rnd * 1300
      i = i Mod 20
      IFish(3) = 4 * (i + 8)
      Label3.Caption = IFish(3)
      Fish(3).Visible = False
       Imageo.Visible = False
     End If
     cMsg = ""
     i = Fish(3).Top - Hook.Top + 8000
     'Label2.Caption = Fish(3).Left & " " & Fish(0).Top & " " & i
     If i > 7120 And i < 8120 Then
       j = Fish(3).Left - Hook.Left + 8000
       If j > 6900 And j < 8400 Then
          cMsg = " hit"
          Timer2.Enabled = False
          Timer1.Enabled = False
          Timer3.Enabled = False
          Timer4.Enabled = False
           iHooked = 4
          'Call GAME
       End If
          'Label2.Caption = Fish(0).Left & " " & Fish(0).Top & " " & iFISH(0) & " " & j & cMsg
     End If
   Else
     IFish(3) = IFish(3) - 1
     Label3.Caption = IFish(3)
     If IFish(3) = 0 Then Fish(3).Visible = True

   End If

End Sub

Private Sub Timer5_Timer()
  '
  '  Alligator
  '
   Dim i As Integer
   Dim j As Integer
   Dim cMsg As String
   Label6.Caption = IFish(8)
  If IFish(8) = 0 Then
      Fish(8).Visible = True
     If Fish(8).Left > 1000 Then   '9000 Then '9000 Then
         Fish(8).Left = Fish(8).Left - 300 '3000
         
          i = Fish(8).Top - Hook.Top + 8000
      ' Label5.Caption = "I=" & i
     'Label2.Caption = Fish(3).Left & " " & Fish(0).Top & " " & i
     If jHooked > 0 Then
        Fish(jHooked - 1).Left = Fish(8).Left - 1000
         Label5.Caption = "jHooked=" & i & Fish(jHooked - 1).Left
     Else
        If i > 7120 And i < 8120 Then
          j = Fish(8).Left - Hook.Left + 8000
          ' Label5.Caption = "I=" & i & " j=" & j
          If j > 6900 And j < 8400 Then
            If (iHooked + jHooked) = 0 Then
             cMsg = " hit"
             Timer2.Enabled = False
             Timer1.Enabled = False
             Timer3.Enabled = False
             Timer4.Enabled = False
             Timer5.Enabled = False
              iHooked = 9
             Else
               If jHooked = 0 Then jHooked = iHooked
               Fish(jHooked - 1).Top = Fish(8).Top
               Fish(jHooked - 1).Left = Fish(8).Left - 1000
               Label5.Caption = "jHooked=" & i & Fish(jHooked - 1).Left
               iHooked = 0
             End If
            
          End If
        End If
        If iHooked > 0 And iHooked <> 9 Then
           Fish(8).Top = Hook.Top
        End If
       End If
     Else
        Fish(8).Left = 7000
        If jHooked > 0 Then
          Timer2.Enabled = True
          Timer1.Enabled = True
          Timer3.Enabled = True
          Timer4.Enabled = True
          Timer5.Enabled = True
          iScore(iCurrentPlayer) = iScore(iCurrentPlayer) - 3
          LblScore.Caption = "Score = " & iScore(iCurrentPlayer)
        End If
        jHooked = 0
        iFISHTop(8) = iFISHTop(8) + 12
        If iFISHTop(8) > 15 Then iFISHTop(8) = iFISHTop(8) - 15
       Fish(8).Top = 3100 + 200 * iFISHTop(8)
       i = Rnd * 1300
       i = i Mod 20
       IFish(8) = 4 * (i + 8)
       Fish(8).Visible = False
       '
'       i = Fish(8).Top - Hook.Top + 8000
'       Label5.Caption = "I=" & i
'     'Label2.Caption = Fish(3).Left & " " & Fish(0).Top & " " & i
'     If i > 7120 And i < 8120 Then
'       j = Fish(8).Left - Hook.Left + 8000
'       If j > 6900 And j < 8400 Then
'          cMsg = " hit"
'          Timer2.Enabled = False
'          Timer1.Enabled = False
'          Timer3.Enabled = False
'          Timer4.Enabled = False
'           iHooked = 8
'          'Call GAME
'       End If
'       End If
     End If
     cMsg = ""
   Else
   IFish(8) = IFish(8) - 1
   End If
End Sub

Private Sub TimerFly_Timer()
  Dim i As Integer
  '
  '  Catch fish
  '
'  If iHooked = 0 Then
'    Timerpause.Enabled = True
'    Timerpause.Interval = 400
'    TimerFly.Enabled = False
'  Else
   TimFoe.Enabled = False
   For i = 0 To 3
     ImgFoe(i).Enabled = False
   Next
   If iFishMove = 0 And Fish(iHooked - 1).Top > 400 Then
    Fish(iHooked - 1).Top = Fish(iHooked - 1).Top - 200
     'Label1.Caption = Fish(iHooked - 1).Top
     Lblf(iHooked - 1).Top = Fish(iHooked - 1).Top + 500
     Exit Sub
   Else
     iFishMove = 1
   End If
   If iFishMove = 1 And Fish(iHooked - 1).Left > ImageBoat.Left + 500 Then
      Fish(iHooked - 1).Left = Fish(iHooked - 1).Left - 200
      Lblf(iHooked - 1).Left = Fish(iHooked - 1).Left + 500
     ' Label1.Caption = Fish(iHooked - 1).Left
      Exit Sub
   Else
    iFishMove = 2
   End If
   If iFishMove = 2 And Fish(iHooked - 1).Top < 1600 Then
    Fish(iHooked - 1).Top = Fish(iHooked - 1).Top + 200
     'Label1.Caption = Fish(iHooked - 1).Top
     Lblf(iHooked - 1).Top = Fish(iHooked - 1).Top + 500
     Exit Sub
   Else
'     i = Rnd * 1300
'      i = i Mod 20
'      IFish(3) = 2 * (i + 8)
'      Label3.Caption = "FLY" & IFish(3)
'      Fish(3).Visible = False
'       Imageo.Visible = False
    iFishMove = 0
    If iWincnt < 4 Then
       Timerpause.Enabled = True
       Timerpause.Interval = 600
    Else
     TimerFly.Enabled = False
     ImageBoat.Left = 510 '4485 165
     ImageBoat.Width = 4485
     ImageBoat.Top = 165
     Line1.X1 = 4875
     Line1.X2 = 4875
     Line1.Y2 = 2550
     Hook.Top = 2520
     Hook.Left = 4850
     Lblf(0).Visible = False
     Lblf(1).Visible = False
     Lblf(2).Visible = False
     Lblf(3).Visible = False
     Line1.Visible = True
     Hook.Visible = True
     Fish(0).Top = 4680
     Fish(1).Top = 2880
     Fish(2).Top = 3840
     Fish(3).Top = 5760
     Fish(0).Left = 5280
     Fish(1).Left = 3120
     Fish(2).Left = 1080
     Fish(3).Left = 7920
     iWincnt = 0
     TimGame.Enabled = True
    End If
    TimerFly.Enabled = False
    iScore(iCurrentPlayer) = iScore(iCurrentPlayer) + iHooked * 2
    LblScore = "Score= " & iScore(iCurrentPlayer)
   End If
 
End Sub

Private Sub TimerFly2_Timer()
  '
  ' Fly the boat
  '
  ImageBoat.Top = ImageBoat.Top - 200
  ImageBoat.Width = ImageBoat.Width - 400
  Line1.Visible = False
  Hook.Visible = False
  Label6.Caption = ImageBoat.Top
  If ImageBoat.Top < -1000 Then
    Timerpause.Interval = 1400
    Timerpause.Enabled = True
    TimerFly2.Enabled = False
  End If
End Sub

Private Sub TimerLoose_Timer()
'
' Loose the fisH
'
  Dim i As Integer
    iLooseFish(iCurrentPlayer) = iLooseFish(iCurrentPlayer) - 1
    Label4.Caption = iLooseFish(iCurrentPlayer)
    If iLooseFish(iCurrentPlayer) < 1 Then
       Timer1.Enabled = True
       Timer2.Enabled = True
       Timer3.Enabled = True
       Timer4.Enabled = True
        Fish(0).Left = Fish(0).Left + 300 '3000
        Fish(1).Left = Fish(1).Left + 300 '3000
       iHooked = 0
         TimerLoose.Enabled = False
       iLooseFish(iCurrentPlayer) = 35
        For i = 0 To 7
               Fish(i).Visible = False
           Fish(i).Top = 3100 + 200 * iFISHTop(i)
         Next
    End If
End Sub

Private Sub Timerpause_Timer()
  Dim i As Integer
  iHooked = 0
  jHooked = 0
  iGamecnt = 0
  TimFoe.Enabled = False
  TimFoe2.Enabled = False
  TimGame.Enabled = False
  TimGameFly.Enabled = False
  Timerpause.Enabled = False
  LblRed.Visible = False
  LblDisplay.Visible = True
  iFoe = 0
  For i = 0 To 3
    ImgFoe(i).Visible = False
  Next
  LblYesNo.Visible = False
     If LblYesNo.BackColor = LblGreen.BackColor Then Command5_Click
     Timer1.Enabled = True
     ImageBoat.Left = 510 '4485 165
     ImageBoat.Width = 4485
     ImageBoat.Top = 165
     Lblf(0).Visible = True
     Lblf(1).Visible = True
     Lblf(2).Visible = True
     Lblf(3).Visible = True
     Line1.X1 = 4875
     Line1.X2 = 4875
     Line1.Y2 = 2550
     Hook.Top = 2520
     Hook.Left = 4850
     Line1.Visible = True
     Hook.Visible = True
     Fish(0).Top = 4680
     Fish(1).Top = 2880
     Fish(2).Top = 3840
     Fish(3).Top = 5760
     Fish(0).Left = 5280
     Fish(1).Left = 3120
     Fish(2).Left = 1080
     Fish(3).Left = 7920
     Lblf(0).Top = Fish(0).Top + 500
     Lblf(1).Top = Fish(1).Top + 500
     Lblf(2).Top = Fish(2).Top + 500
     Lblf(3).Top = Fish(3).Top + 500
     Lblf(0).Left = Fish(0).Left + 500
     Lblf(1).Left = Fish(1).Left + 500
     Lblf(2).Left = Fish(2).Left + 500
     Lblf(3).Left = Fish(3).Left + 500
     LblRed.Left = ImageBoat.Left + 1170
    
'   Timer2.Enabled = True
'   Timer3.Enabled = True
'   Timer4.Enabled = True
'   Timer5.Enabled = True
    'End If
End Sub

Private Sub Timerpause2_Timer()
    
End Sub


Private Sub TimFoe_Timer()
   ImgFoe(iFoe - 1).Top = ImgFoe(iFoe - 1).Top - 200
   Label6.Caption = iFoe
   If ImgFoe(iFoe - 1).Top < 1800 Then
     TimFoe.Enabled = False
     ImgFoe(iFoe - 1).Visible = False
     iFoecnt = 0
   End If
   'TimFoe.Enabled = False
   'Label8.Caption = "ok"
    If ImgFoe(iFoe - 1).Top < 2700 Then
      If (ImageBoat.Left + ImageBoat.Width - 1700) > ImgFoe(iFoe - 1).Left And _
       ImageBoat.Left < ImgFoe(iFoe - 1).Left Then
          'Label8.Caption = "HIT"
          'LblRed.Visible = True
          LblYesNo.BackColor = LblRed.BackColor
          Timer1.Enabled = False
          TimerFly.Enabled = False
          TimerFly2.Interval = 600
          TimerFly2.Enabled = True
          'Timerpause.Interval = 2400
          'Timerpause.Enabled = True
          TimFoe.Enabled = False
          iFoecnt = 0
          iScore(iCurrentPlayer) = iScore(iCurrentPlayer) - 1
          LblScore.Caption = "Score= " & iScore(iCurrentPlayer)
       End If
    End If
  ' Label7.Caption = ImgFoe(iFoe - 1).Left & " " & ImageBoat.Left
End Sub


Private Sub TimFoe2_Timer()
  ImgFoe(iFoe - 1).Top = ImgFoe(iFoe - 1).Top - 200
    iWincnt = 0
    LblDisplay.Visible = False
   'Label4.Caption = ImgFoe(iFoe - 1).Top & " " & ImageBoat.Top
   If ImgFoe(iFoe - 1).Top < 1800 Then
     TimFoe2.Enabled = False
     ImgFoe(iFoe - 1).Visible = False
     iFoecnt = 0
     iScore(iCurrentPlayer) = iScore(iCurrentPlayer) + 2
     LblScore.Caption = "Score= " & iScore(iCurrentPlayer)
     If iGamecnt > 5 Then
       TimFoe2.Enabled = False
       TimGame.Enabled = False
       LblYesNo.BackColor = LblGreen.BackColor
       'Command5_Click
       Timerpause.Interval = 1200
       Timerpause.Enabled = True
     End If
   End If
   'TimFoe.Enabled = False
   'Label8.Caption = "ok"
    If ImgFoe(iFoe - 1).Top < 2700 Then
      If (ImageBoat.Left + ImageBoat.Width - 1700) > ImgFoe(iFoe - 1).Left And _
       ImageBoat.Left < ImgFoe(iFoe - 1).Left Then
          'Label8.Caption = "HIT"
          'LblRed.Visible = True
          LblYesNo.BackColor = LblGreen.BackColor
          'Command5_Click
          Timer1.Enabled = False
          TimerFly.Enabled = False
          TimerFly2.Interval = 600
          TimerFly2.Enabled = True
          TimGame.Enabled = False
          'Timerpause.Interval = 2400
          'Timerpause.Enabled = True
          TimFoe2.Enabled = False
          iFoecnt = 0
          iScore(iCurrentPlayer) = iScore(iCurrentPlayer) - 1
          LblScore.Caption = "Score= " & iScore(iCurrentPlayer)
       End If
    End If
End Sub

Private Sub TimGame_Timer()
  Lblf(0).Visible = False
  Lblf(1).Visible = False
  Lblf(2).Visible = False
  Lblf(3).Visible = False
  If TimFoe2.Enabled = False Then PickFoe2
End Sub

Private Sub TimGameFly_Timer()
  Dim i As Integer
  '
  '  Catch fish
  '
'  If iHooked = 0 Then
'    Timerpause.Enabled = True
'    Timerpause.Interval = 400
'    TimerFly.Enabled = False
'  Else
'   TimFoe.Enabled = False
'   For i = 0 To 3
'     ImgFoe(i).Enabled = False
'   Next
    TimGame.Enabled = False
   If iFishMove = 0 And Fish(iHooked - 1).Top > 400 Then
    Fish(iHooked - 1).Top = Fish(iHooked - 1).Top - 200
     'Label1.Caption = Fish(iHooked - 1).Top
     Lblf(iHooked - 1).Top = Fish(iHooked - 1).Top + 500
     Exit Sub
   Else
     iFishMove = 1
   End If
   If iFishMove = 1 And Fish(iHooked - 1).Left > ImageBoat.Left + 500 Then
      Fish(iHooked - 1).Left = Fish(iHooked - 1).Left - 200
      Lblf(iHooked - 1).Left = Fish(iHooked - 1).Left + 500
     ' Label1.Caption = Fish(iHooked - 1).Left
      Exit Sub
   Else
    iFishMove = 2
   End If
   If iFishMove = 2 And Fish(iHooked - 1).Top < 1600 Then
    Fish(iHooked - 1).Top = Fish(iHooked - 1).Top + 200
     'Label1.Caption = Fish(iHooked - 1).Top
     Lblf(iHooked - 1).Top = Fish(iHooked - 1).Top + 500
     Exit Sub
   Else
'     i = Rnd * 1300
'      i = i Mod 20
'      IFish(3) = 2 * (i + 8)
'      Label3.Caption = "FLY" & IFish(3)
'      Fish(3).Visible = False
'       Imageo.Visible = False
    iFishMove = 0
    'Timerpause.Enabled = True
    'Timerpause.Interval = 600
     TimGameFly.Enabled = False
     Timer1.Enabled = True
     TimGame.Enabled = True
     iHooked = 0
     Fish(0).Top = 4680
     Fish(1).Top = 2880
     Fish(2).Top = 3840
     Fish(3).Top = 5760
     Fish(0).Left = 5280
     Fish(1).Left = 3120
     Fish(2).Left = 1080
     Fish(3).Left = 7920
    'iScore(iCurrentPlayer) = iScore(iCurrentPlayer) + iHooked * 2
    'LblScore = "Score= " & iScore(iCurrentPlayer)
   End If
End Sub


