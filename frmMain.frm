VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CollideBall - A truly annoying game :p"
   ClientHeight    =   7200
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   7200
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCollide 
      Height          =   255
      Left            =   1920
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox picBall 
      Height          =   255
      Left            =   4200
      Picture         =   "frmMain.frx":03B2
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   4560
      Width           =   255
   End
   Begin VB.Timer tmrBalltimer 
      Interval        =   15
      Left            =   9480
      Top             =   0
   End
   Begin VB.PictureBox picPadL 
      Height          =   1215
      Left            =   120
      Picture         =   "frmMain.frx":0764
      ScaleHeight     =   1155
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   5880
      Width           =   255
   End
   Begin VB.PictureBox picPadT 
      Height          =   255
      Left            =   480
      Picture         =   "frmMain.frx":0A86
      ScaleHeight     =   195
      ScaleWidth      =   1155
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picPadR 
      Height          =   1215
      Left            =   9480
      Picture         =   "frmMain.frx":0DC8
      ScaleHeight     =   1155
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   5880
      Width           =   255
   End
   Begin VB.PictureBox picPadB 
      Height          =   255
      Left            =   480
      Picture         =   "frmMain.frx":10EA
      ScaleHeight     =   195
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label lblAboot2 
      Caption         =   "Evil ideas by Skeggo"
      ForeColor       =   &H80000010&
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label lblAboot1 
      Caption         =   "Made by Simon"
      ForeColor       =   &H80000010&
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   6480
      Width           =   2535
   End
   Begin VB.Label lblScore 
      Caption         =   "Score: 0p"
      ForeColor       =   &H80000010&
      Height          =   255
      Left            =   8040
      TabIndex        =   6
      Top             =   480
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuBlank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim destX, destY, spaceL, spaceR, spaceT, spaceB As Integer
Dim pL, pR, pT, pB As Integer
Dim pLR, pLT, pLB As Integer ' padLeftRight padLeftTop padLeftBottom
Dim pRL, pRT, pRB As Integer ' padRightLeft padRightTop padRightBottom
Dim pTL, pTR, pTB As Integer ' padTopLeft padTopRight padTopBottom
Dim pBL, pBR, pBT As Integer ' padBottomLeft padBottomRight padBottomTop
Dim score As Integer
Dim collideX, collideY As Integer

' Destination Legend:
'   destX=-1 == Left
'   destX=1 == Right
'   destY=-1 == Up
'   destY=1 == Down

Private Sub Form_Load()
    destX = 1
    destY = -1
    
    collideX = 1
    collideY = -1
    
    score = 0
    
    spaceL = 0
    spaceR = frmMain.ScaleWidth
    spaceT = 0
    spaceB = frmMain.ScaleHeight
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picPadB.Left = X
    picPadT.Left = X
    picPadL.Top = Y
    picPadR.Top = Y
    
    ' Cordinates for padLeft
    pLR = picPadL.Left + picPadL.Width
    pLT = picPadL.Top
    pLB = picPadL.Top + picPadL.Height
    
    ' Cordinates for padRight
    pRL = picPadR.Left
    pRT = picPadR.Top
    pRB = picPadR.Top + picPadR.Height
    
    ' Cordinates for padTop
    pTL = picPadT.Left
    pTR = picPadT.Left + picPadT.Width
    pTB = picPadT.Top + picPadT.Height
    
    ' Cordinates for padBottom
    pBL = picPadB.Left
    pBR = picPadB.Left + picPadB.Width
    pBT = picPadB.Top

End Sub

Private Sub lblCreator_Click()

End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuNew_Click()
    picBall.Left = 5000
    picBall.Top = 5000
    picCollide.Left = 2000
    picCollide.Top = 2000
    tmrBalltimer.Enabled = True
    score = 0
    lblScore = "Score: 0"
End Sub


Private Sub tmrBalltimer_Timer()
    
    ''''''''''''''''''
    'Normal ball
    
    ' Bounce on the top pad
    If (picBall.Top <= pTB) And (picBall.Left <= pTR) And (picBall.Left >= pTL) Then
        destY = destY * -1
        score = score + 1
        lblScore.Caption = "Score: " & score
    End If
    
    ' Bounce on the bottom pad
    If (picBall.Top + picBall.Height >= pBT) And (picBall.Left <= pBR) And (picBall.Left >= pBL) Then
        destY = destY * -1
        score = score + 1
        lblScore.Caption = "Score: " & score
    End If
    
    ' Bounce on the left pad
    If (picBall.Left <= pLR) And (picBall.Top <= pLB) And (picBall.Top + picBall.Height >= pLT) Then
        destX = destX * -1
        score = score + 1
        lblScore.Caption = "Score: " & score
    End If

    ' Bounce on the right pad
    If (picBall.Left + picBall.Width >= pRL) And (picBall.Top <= pRB) And (picBall.Top + picBall.Height >= pRT) Then
        destX = destX * -1
        score = score + 1
        lblScore.Caption = "Score: " & score
    End If
    
    If (picBall.Left <= spaceL) Or ((picBall.Left + picBall.Width) >= spaceR) Then 'when the ball touches the rectangle horizontally
        'die
        tmrBalltimer.Enabled = False
    End If

    If (picBall.Top <= spaceT) Or ((picBall.Top + picBall.Height) >= spaceB) Then 'when the ball touches the rectangle vertically
        'die
        tmrBalltimer.Enabled = False
    End If

    picBall.Left = picBall.Left + (destX * 50)
    picBall.Top = picBall.Top + (destY * 50)
    
    
    ''''''''''''''''''''
    ' COLLIDEBALL
    
    
    ' Bounce on the top pad
    If (picCollide.Top <= pTB) And (picCollide.Left <= pTR) And (picCollide.Left >= pTL) Then
        collideY = collideY * -1
        score = score + 1
        lblScore.Caption = "Score: " & score
    End If
    
    ' Bounce on the bottom pad
    If (picCollide.Top + picCollide.Height >= pBT) And (picCollide.Left <= pBR) And (picCollide.Left >= pBL) Then
        collideY = collideY * -1
        score = score + 1
        lblScore.Caption = "Score: " & score
    End If
    
    ' Bounce on the left pad
    If (picCollide.Left <= pLR) And (picCollide.Top <= pLB) And (picCollide.Top + picCollide.Height >= pLT) Then
        collideX = collideX * -1
        score = score + 1
        lblScore.Caption = "Score: " & score
    End If

    ' Bounce on the right pad
    If (picCollide.Left + picCollide.Width >= pRL) And (picCollide.Top <= pRB) And (picCollide.Top + picCollide.Height >= pRT) Then
        collideX = collideX * -1
        score = score + 1
        lblScore.Caption = "Score: " & score
    End If
    
    If (picCollide.Left <= spaceL) Or ((picCollide.Left + picCollide.Width) >= spaceR) Then 'when the ball touches the rectangle horizontally
        'die
        tmrBalltimer.Enabled = False
    End If

    If (picCollide.Top <= spaceT) Or ((picCollide.Top + picCollide.Height) >= spaceB) Then 'when the ball touches the rectangle vertically
        'die
        tmrBalltimer.Enabled = False
    End If

    picCollide.Left = picCollide.Left + (collideX * 50)
    picCollide.Top = picCollide.Top + (collideY * 50)
    
End Sub
