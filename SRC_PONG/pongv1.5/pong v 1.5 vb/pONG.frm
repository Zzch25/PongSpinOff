VERSION 5.00
Begin VB.Form frmPong 
   BackColor       =   &H80000007&
   Caption         =   "Pong?"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   Icon            =   "pONG.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "pONG.frx":014A
   ScaleHeight     =   6060
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDifficulty 
      BackColor       =   &H00808080&
      Caption         =   "Difficulty"
      Height          =   1695
      Left            =   6960
      TabIndex        =   8
      Top             =   3840
      Width           =   2055
      Begin VB.OptionButton optVH 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Very Hard"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1815
      End
      Begin VB.OptionButton optH 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hard"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton optM 
         BackColor       =   &H80000010&
         Caption         =   "Medium"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton optE 
         BackColor       =   &H80000015&
         Caption         =   "Easy"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdN 
      Caption         =   "New Game"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Timer tmrTwitch 
      Interval        =   40
      Left            =   2160
      Top             =   5520
   End
   Begin VB.Timer tmrPcAI 
      Interval        =   1
      Left            =   1680
      Top             =   5520
   End
   Begin VB.Timer tmrBall 
      Interval        =   10
      Left            =   1200
      Top             =   5520
   End
   Begin VB.Timer tmrMoved 
      Interval        =   1
      Left            =   720
      Top             =   5520
   End
   Begin VB.Timer tmrMoveu 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   5520
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "Play/Resume"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdH 
      Caption         =   "Help"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdE 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Image imgBG 
      Height          =   6075
      Left            =   0
      Picture         =   "pONG.frx":1E0C
      Top             =   0
      Width           =   9105
   End
   Begin VB.Shape shpBall 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblE 
      BackColor       =   &H80000012&
      Caption         =   "Pause"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8520
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image PongP 
      Height          =   195
      Index           =   1
      Left            =   8400
      Picture         =   "pONG.frx":B63F0
      Top             =   2760
      Width           =   255
   End
   Begin VB.Image PongP 
      Height          =   195
      Index           =   0
      Left            =   360
      Picture         =   "pONG.frx":B66D6
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   120
      Top             =   0
      Width           =   615
   End
   Begin VB.Label lblBarrier 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   5760
      Width           =   6375
   End
   Begin VB.Label lblW 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Ready Up!"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblA 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmPong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'barrier collision, increment by one saying this is useless diff colors
Option Explicit

 Private intShift As Integer 'Is used to hold the current keypress value
 Private intAdder As Integer 'Counts the times the user hits the barrier
 Private intRandom As Integer 'Holds 1 or 0 which determine whether or not _
                              to call the load part of tmrBall
 Private intRand As Integer 'holds value that determines the which direction _
                             the ball will travel
 Private intUserscore As Integer 'holds the user score
 Private intPCScore As Integer 'holds the computer score
 Private intSA As Integer 'holds string incrementor value for strpointless
 Private strSc As String 'holds score values with text for display
 Private intAiIntel As Integer
 Private intAiStop As Integer
 Private intDifficulty 'sets value to determine difficulty
 
 Dim strPointless(0 To 27) As String 'Creates an array used for animation

 Private Declare Function GetAsyncKeyState Lib _
 "user32" (ByVal vKey As Long) As Integer 'Sets the current keypress value to
                                         'the pressed buttons ascii value

Private Sub Form_Load()
 Randomize 'used to ensure the ball starts in different directions


 Call Endvis 'sets all items to their proper visual state before the user _
 begins the actual game
 cmdN.Enabled = False
 
 'animation assigned to strpointless array
 strPointless(0) = ""
 strPointless(1) = " B"
 strPointless(2) = " By"
 strPointless(3) = " By "
 strPointless(4) = " By Z"
 strPointless(5) = " By Za"
 strPointless(6) = " By Zac"
 strPointless(7) = " By Zach"
 strPointless(8) = " By Zacha"
 strPointless(9) = " By Zachar"
 strPointless(10) = " By Zachary"
 strPointless(11) = " By Zachary "
 strPointless(12) = " By Zachary J"
 strPointless(13) = " By Zachary Jo"
 strPointless(14) = " By Zachary Job"
 strPointless(15) = " By Zachary Job "
 strPointless(16) = ""
 strPointless(17) = " By Zachary Job "
 strPointless(18) = ""
 strPointless(19) = " By Zachary Job "
 strPointless(20) = ""
 strPointless(21) = " By Zachary Job "
 strPointless(22) = ""
 strPointless(23) = " By Zachary Job "
 strPointless(24) = ""
 strPointless(25) = " By Zachary Job "
 strPointless(26) = ""
 strPointless(27) = " By Zachary Job "

 'this sets strSc equal to text and scores needed to display game progress
 strSc = "Your Score: " & intUserscore & " The Computer's Score is: " & intPCScore
 
 
 intDifficulty = 45 'sets difficulty to easy
 

 Call Disabler 'sets all moving objects to their proper positions
 
End Sub

Private Sub cmdN_Click() 'begins game at default settings

 intAdder = 0 'resets barrier hit count
 intSA = 0 'sets lblbarrier to default

 Call BarrierC
 Call BeginVis 'sets all items to their proper visual state for when the _
 actual game begins
 
 'resets data so that the game state resets to its default
 intUserscore = 0
 intPCScore = 0
 intRandom = 0 'allows tmrball to start at its default
 
 intRand = Int(4 * Rnd + 1) 'ensures that the ball moves in different _
 directions by causing different cases to be called in the begining
 
 'sets strSc to the current score
 strSc = "Your Score: " & intUserscore & " The Computer's Score is: " & intPCScore
 lblA.Caption = strSc 'sets the score label to strSc
End Sub

Private Sub cmdC_Click() 'starts or resumes game

 Call BeginVis 'sets all items to their proper visual state for when the _
 actual game begins
 
 intSA = 0 'allows BarrierC to start at its default
 Call BarrierC 'resets barrier counter
 
 intRand = Int(4 * Rnd + 1) 'ensures that the ball moves in a random _
 start direction
 
 lblA.Caption = strSc 'sets lblA to the current game score
 
End Sub

Private Sub cmdE_Click() 'ends program
 Unload Me 'exit
End Sub

Private Sub cmdH_Click() 'instructions appear
 'instructions appear in a message box
 MsgBox ("Use the arrows to hit the ball and keep it from reaching behind you. When you go to hit the ball, ensure its in front of your paddle, or else the ball will not get pushed away. Don't let it touch the back of your paddle either or you lose!")
End Sub

Private Sub imgBG_Click()
 MsgBox ("Background by Zachary Job") 'author of the background appears
End Sub

Private Sub lblE_Click()
 Call Endvis  'sets all items to their proper visual state after the user ends _
 the actual game
  
 intAdder = 0 'sets the number of barrier collisions = 0
 intRandom = 0 'allows tmrBall to start at its default
 
 Call Disabler 'sets moving items to default position
 
 If intUserscore < 1 And intPCScore < 1 Then 'determines whether or not it is _
 appropriate to enable the new game button
  cmdN.Enabled = False
 Else
  cmdN.Enabled = True
 End If
End Sub

Sub Disabler() 'sets moving objects to their default positions
 PongP(1).top = 2760
 PongP(0).top = 2760
 PongP(0).Left = 240
 PongP(1).Left = 8400
 shpBall.top = 2760
 shpBall.Left = 4440
End Sub

Public Sub Pause(Duration As Double) 'causes the game to pause for a given time
 Dim Current As Long
 Current = Timer
 Do Until Timer - Current >= Duration 'constantly checks whther or not current _
 has become equal to Timer
 Loop
End Sub

Sub vKey_keydown(key As Integer) 'checks for keypress
 If GetAsyncKeyState(38) <> 0 Then 'enables up if up arrow value is present
  Call Up
 End If
 If GetAsyncKeyState(40) <> 0 Then 'enables down if down arrow value is present
  Call Down
 End If
End Sub
Sub BeginVis()
 'sets items to their proper viusal state when the actual game runs
 cmdC.Visible = False
 cmdE.Visible = False
 cmdH.Visible = False
 cmdN.Visible = False
 lblE.Visible = True
 imgBG.Visible = False
 tmrMoveu.Enabled = True
 tmrMoved.Enabled = True
 tmrBall.Enabled = True
 lblBarrier.Visible = True
 tmrPcAI.Enabled = True
 lblA.Visible = True
 lblW.Visible = True
 fraDifficulty.Visible = False
End Sub
Sub Endvis()
 'sets items to their proper visual state when the the actual game does not _
 run
 cmdC.Visible = True
 cmdE.Visible = True
 cmdH.Visible = True
 cmdN.Visible = True
 lblE.Visible = False
 tmrMoveu.Enabled = False
 tmrMoved.Enabled = False
 tmrBall.Enabled = False
 lblBarrier.Visible = False
 tmrPcAI.Enabled = False
 tmrTwitch.Enabled = False
 lblA.Visible = False
 imgBG.Visible = True
 fraDifficulty.Visible = True
End Sub

Sub Up() 'causes player pong to move up when sub is called
 PongP(0).Move PongP(0).Left, PongP(0).top - 40
End Sub

Sub Down() 'causes player pong to move down when sub is called
 PongP(0).Move PongP(0).Left, PongP(0).top + 40
End Sub

Sub BarrierC()
 Dim strA As String 'creates string to hold output

 If intSA = 0 Then
   If intAdder >= 1 Then 'if it is 1 it adds one before outputting to the _
    answer
    intAdder = intAdder + 1
    strA = strPointless(intSA)
    intSA = intSA + 1
   Else 'if it is 0, it adds 1 after outputting to the answer
    strA = strPointless(intSA)
    intAdder = intAdder + 1
    intSA = intSA + 1
   End If
 Else
  If intSA > 0 And intSA < 27 Then
   If intAdder >= 1 Then
    intSA = intSA + 1
    intAdder = intAdder + 1
    'The barrier counter with animation are set to strA
    strA = strPointless(intSA)
   Else
    strA = strPointless(intSA)
    intSA = intSA + 1
    intAdder = intAdder + 1
   End If
  Else
   intSA = 0
   intAdder = intAdder + 1
   strA = strPointless(intSA)
  End If
 End If
    lblBarrier = strA 'sets the label to strA If intAdder > 1 Then
End Sub
'looks for paddle collision in the front area of the paddle
Function PbCollision(top1 As Integer, top2 As Integer, left1 As Integer, _
left2 As Integer, RL As Integer, UD As Integer) As Integer
 
 If UD = 1 Then '1 ball = bottom of screen, 0 = top
  If RL = 1 Then '1 ball = right of screen, 0 = left
  'if ball is at the bottom right
   If top2 - top1 > -100 And top2 - top1 < 200 Then 'checks height similarities
    If left1 - left2 > -100 And left1 - left2 < 200 Then 'checks length similarities
     PbCollision = 1 'collision occurred
    Else
     PbCollision = 0 'no collision (top collide, back side of paddle or _
     no left collide)
    End If
   Else
    PbCollision = 0 'no collision
   End If
  Else 'if ball is at bottom left
   If top2 - top1 > -100 And top2 - top1 < 200 Then
    If left2 - left1 > -100 And left2 - left1 < 200 Then
     PbCollision = 1
    Else
     PbCollision = 0
    End If
   Else
    PbCollision = 0
   End If
  End If
   Else
    If RL = 1 Then
    'if ball is at top right
     If top1 - top2 > -100 And top1 - top2 < 200 Then
      If left1 - left2 > -100 And left1 - left2 < 200 Then
       PbCollision = 1
      Else
       PbCollision = 0
      End If
     Else
      PbCollision = 0
     End If
     Else 'if ball is at top left
       If top1 - top2 > -100 And top1 - top2 < 200 Then
        If left2 - left1 > -100 And left2 - left1 < 200 Then
         PbCollision = 1
        Else
         PbCollision = 0
        End If
       Else
        PbCollision = 0
       End If
      End If
     End If
End Function
 'checks for ball collision with barriers
Function BCollision(top As Integer, top2 As Integer, pP As Integer) As Boolean
 If pP >= top Or pP <= top2 Then
  BCollision = True
 Else
  BCollision = False
 End If
End Function
 'checks for ball position - top or bottom of screen
Function TopBottom(Ball As Integer, Mid As Integer) As Integer
 If Ball > Mid Then
  TopBottom = 1 'bottom
 Else
  TopBottom = 0 'top
 End If
End Function
 'checks for ball poistion - left or right of screen
Function RightLeft(Ball As Integer, Mid As Integer) As Integer
 If Ball > Mid Then
  RightLeft = 1 'right
 Else
  RightLeft = 0 'left
 End If
End Function

Private Sub optE_Click() 'easy
 intDifficulty = 45
End Sub

Private Sub optH_Click() 'hard
 intDifficulty = 20
End Sub

Private Sub optM_Click() 'medium
intDifficulty = 30
End Sub

Private Sub optVH_Click()
 intDifficulty = 10 'very hard
End Sub

Private Sub tmrBall_Timer() 'Begins ball movement and collision detection
  
  intRandom = intRandom + 1 'increments so it disables a block from running _
  after the program begins the actual game

   If shpBall.Left < 120 Then 'Detects if ball reaches goal area
    intPCScore = intPCScore + 1 'adds to Ai score
    If intPCScore > 24 Then 'Determines if Ai wins
     'resets to the main menu and resets all current data to their defaults
     Call Endvis
     intUserscore = 0
     intPCScore = 0
     intRandom = 0 'allows this sub to start from beggining
     intAdder = 0 'sets barrier counter value to 0
     Call BarrierC 'sets lblBarrier to default
     cmdN.Enabled = False
     Call Disabler 'moves moving items to their default positions
     
     MsgBox ("PC Wins") 'displays winner
    End If
    'registers score
    strSc = "Your Score: " & intUserscore & " The Computer's Score is: " & intPCScore
    lblA.Caption = strSc
    intSA = 0 'allows BarrierC to start at its default
    'resets/pauses game
    lblW.Visible = True 'brings up wait label
    Call Disabler 'moves moving objects to default position
    Pause (2)
    lblW.Visible = False 'removes wait label
   Else
    If shpBall.Left > 8640 Then 'Detects if ball reaches goal area
     intUserscore = intUserscore + 1 'adds to player score
     If intUserscore > 24 Then 'determines if you and PC lose
     'resets to the main menu and resets all current data to their defaults
     Call Endvis
     intUserscore = 0
     intPCScore = 0
     intRandom = 0 'allows this sub to start from beggining
     cmdN.Enabled = False
     Call Disabler 'moves moving items to their default positions
     
     MsgBox ("Chuck Norris wins, you and PC lose!") 'displays winner
     End If
    'registers score
    strSc = "Your Score: " & intUserscore & " The Computer's Score is: " & intPCScore
    lblA.Caption = strSc
    'resets/pauses game
    intSA = 0 'allows BarrierC to start at its default
    Call BarrierC 'resets barrier counter
    lblW.Visible = True 'brings up wait label
    Call Disabler 'moves moving objects to default position
    Pause (2)
    lblW.Visible = False 'removes wait label
    Else
     If intRandom = 1 Then 'if the game just started or resumed = true
      Pause (2) 'pauses game
      lblW.Visible = False 'brings up wait label
      Select Case intRand 'used to determine random direction
       Case 1 'move up right
        shpBall.Move shpBall.Left + 45, shpBall.top - 85
       Case 2 'move down right
        shpBall.Move shpBall.Left + 45, shpBall.top + 85
       Case 3 'move down left
        shpBall.Move shpBall.Left - 45, shpBall.top + 85
       Case 4 'move up left
        shpBall.Move shpBall.Left - 45, shpBall.top - 85
      End Select
     Else
      Select Case intRand
       Case 1 'looks for collision while ball moves up right
        shpBall.Move shpBall.Left + 45, shpBall.top - 85
        'looks for paddle-ball collision at paddles middle
        If PbCollision(PongP(1).top, shpBall.top, PongP(1).Left, _
        shpBall.Left, RightLeft(shpBall.Left, 3840), _
        TopBottom(shpBall.top, 2800)) = 1 Then
         intRand = 4
         Else
          'looks for ball-wall collision
          If BCollision(5800, 0, shpBall.top) = True Then
           'looks for top-bottom position to determine proper bounce direction
           If TopBottom(shpBall.top, 2800) = 0 Then
            intRand = 2 'determines next case to be called
          End If
         End If
        End If
       Case 2  'looks for collision while ball moves down right
        shpBall.Move shpBall.Left + 45, shpBall.top + 85
     
        If PbCollision(PongP(1).top, shpBall.top, PongP(1).Left, _
        shpBall.Left, RightLeft(shpBall.Left, 3840), _
        TopBottom(shpBall.top, 2800)) = 1 Then
         intRand = 3
        Else
          If BCollision(5800, 0, shpBall.top) = True Then
           If TopBottom(shpBall.top, 2800) = 1 Then
            intRand = 1
          End If
         End If
        End If
       Case 3 'looks for collision while ball moves down left
        shpBall.Move shpBall.Left - 45, shpBall.top + 85
     
        If PbCollision(PongP(0).top, shpBall.top, PongP(0).Left, _
        shpBall.Left, RightLeft(shpBall.Left, 3840), _
        TopBottom(shpBall.top, 2800)) = 1 Then
         intRand = 2
        Else
          If BCollision(5800, 0, shpBall.top) = True Then
           If TopBottom(shpBall.top, 2800) = 1 Then
            intRand = 4
          End If
         End If
        End If
       Case 4 'looks for collision while ball moves up left
        shpBall.Move shpBall.Left - 45, shpBall.top - 85
     
        If PbCollision(PongP(0).top, shpBall.top, PongP(0).Left, _
        shpBall.Left, RightLeft(shpBall.Left, 3840), _
        TopBottom(shpBall.top, 2800)) = 1 Then
         intRand = 1
         Else
          If BCollision(5800, 0, shpBall.top) = True Then
           If TopBottom(shpBall.top, 2800) = 0 Then
            intRand = 3
           End If
          End If
         End If
       End Select
      End If
     End If
    End If
End Sub

Private Sub tmrMoved_Timer()
 If PongP(0).top > 5800 Then 'checks for bottom collision then resets pong _
 position
  PongP(0).top = 5799 'resets pos
  Call BarrierC 'adds to barrier hits count
 Else
  Call vKey_keydown(intShift) 'allows paddle movement
 End If
End Sub

Private Sub tmrMoveu_Timer()
 If PongP(0).top < 0 Then 'checks for bottom collision then resets pong _
 position
  PongP(0).top = 1 'ressets pos
  Call BarrierC 'adds to barrier hit count
 Else
  Call vKey_keydown(intShift) 'allows paddle movement
 End If
End Sub

Private Sub tmrPcAI_Timer()
 'coordinates PC's Ai
 
 If RightLeft(shpBall.Left, 4560) = 1 Then 'if ball is right of screen
  intAiStop = intAiStop + 1
  If intAiStop > 2 Then
   intAiStop = 2
  End If
 Else
  intAiStop = 0
 End If

 If intAiStop = 1 Then 'allows intAiIntel to only set to one value while _
 the ball is on the right side of the screen
  intAiIntel = 100 * Rnd + 5
 End If
 
 If PongP(1).top > 5800 Then 'checks for bottom collision then resets pong _
 position
  PongP(1).top = 5799 'resets pos
 Else
  If PongP(1).top < 0 Then 'checks for bottom collision then resets pong _
  position
   PongP(1).top = 1 'ressets pos
  Else
  Select Case intRand
   Case 1 'while ball moves up right
    tmrTwitch.Enabled = False
    If intAiIntel > 4 And intAiIntel < intDifficulty Then 'random number _
    determines activation and probabilty is affected by difficulty
     PongP(1).Move PongP(1).Left, PongP(1).top - 70 'causes ai to move freely
    Else
     PongP(1).top = shpBall.top 'ai does not miss
    End If
  Case 2 'while ball moves down right
   tmrTwitch.Enabled = False
   If intAiIntel > 4 And intAiIntel < intDifficulty Then 'random number _
    determines activation and probabilty is affected by difficulty
    PongP(1).Move PongP(1).Left, PongP(1).top + 70 'causes ai to move freely
   Else
     PongP(1).top = shpBall.top 'ai does not miss
   End If
  Case 3
   tmrTwitch.Enabled = True 'ai move to middle if ball moves away
  Case 4
   tmrTwitch.Enabled = True 'ai move to middle if ball moves away
  End Select
  End If
 End If
End Sub

Private Sub tmrTwitch_Timer()
 Randomize
 
 Dim intWaffle As Integer 'distance to be moved by PC paddle
 
 If TopBottom(PongP(1).top, 2760) = 1 Then 'if paddle is bellow mid, move up X
  intWaffle = (PongP(1).top - 2760) / 15
  PongP(1).Move PongP(1).Left, PongP(1).top - intWaffle
 Else
  intWaffle = (2760 - PongP(1).top) / 15 'if paddle is ab
  PongP(1).Move PongP(1).Left, PongP(1).top + intWaffle
 End If
End Sub




