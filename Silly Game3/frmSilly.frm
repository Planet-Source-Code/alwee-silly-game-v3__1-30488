VERSION 5.00
Begin VB.Form frmSilly 
   BackColor       =   &H00C0FFFF&
   Caption         =   "The Stupidest Game in the World v3.0"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Start"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Click me to start"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4200
      Top             =   1080
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   600
   End
   Begin VB.Timer tmrRandomAlien 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   120
   End
   Begin VB.Label lblComment 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "***************************"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label lblPercent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Top             =   2760
      Width           =   90
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   4680
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Label lblOutOf 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   1680
      TabIndex        =   6
      Top             =   2760
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Out of"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   1080
      TabIndex        =   5
      Top             =   2760
      Width           =   435
   End
   Begin VB.Label lblEnd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GAME OVER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   585
      Left            =   840
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   2910
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   300
      Left            =   4200
      TabIndex        =   3
      Top             =   2700
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time Left:"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   3360
      TabIndex        =   2
      Top             =   2760
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caught:"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   555
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   2760
      Width           =   90
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   960
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF00FF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   2280
      Shape           =   1  'Square
      Top             =   840
      Width           =   150
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   4680
      Y1              =   2520
      Y2              =   2520
   End
End
Attribute VB_Name = "frmSilly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variables for our points(Moved em out of the subs, so program executes faster, no special variable making per each sub)
    Dim X As Integer
    Dim Y As Integer

Dim gTime As Byte 'Amount of time to play
Dim MoveAlien As Byte 'Gets a random number
Dim Go As Boolean 'Start game variable

'Shaun Rogerson's made these variables for calculating the percent and score
    Dim OutOf As Byte
    Dim Perc As Integer



Private Sub cmdStart_Click()
    'Sets the needed options for a new game
    Go = True
    OutOf = 0
    Perc = 0
    lblEnd.Visible = False
    cmdStart.Visible = False
    lblComment.Visible = False
    lblPercent.Caption = "0"
    lblOutOf.Caption = "0"
    lblScore.Caption = "0"
    gTime = 60
    tmrRandomAlien.Enabled = True
    Timer2.Enabled = True
    Timer3.Enabled = True
    
'This part randomizes where the game shapes start up at
    Randomize Timer
    Y = (Rnd * (2520 - Shape1.Width)) + 1
    X = (Rnd * (4680 - Shape1.Height)) + 1
    Shape1.Left = X
    Shape1.Top = Y
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Controls the scrolling
    If Go = True Then
        Select Case KeyCode
            Case 37 And Shape1.Left > 0 'Left
                    Shape1.Left = Shape1.Left - 100
            Case 38 And Shape1.Top > 0 'Down
                Shape1.Top = Shape1.Top - 100
            Case 39 And Shape1.Left < (frmSilly.Width - Shape1.Width) 'Right
                Shape1.Left = Shape1.Left + 100
            Case 40 And Shape1.Top < ((frmSilly.Height - (frmSilly.Height - Line1.Y1)) - Shape1.Height) 'Up
                Shape1.Top = Shape1.Top + 100
        End Select
    End If
End Sub

Private Sub tmrRandomAlien_Timer()
'This timer moves the alien
    Randomize Timer
    MoveAlien = Rnd * 4 ' Pick a random number
    
    'Work out the percentage, and then display it on screen
    If CInt(lblOutOf.Caption) >= 1 And CInt(lblScore.Caption) >= 1 Then
        Perc = (CInt(lblScore.Caption) / CInt(lblOutOf.Caption)) * 100
    End If
    lblPercent.Caption = Str$(Perc) + " %"
    
    'Do Mover Function
    Call Mover(Shape2, lblOutOf, MoveAlien, lblOutOf)
End Sub

Private Sub Timer2_Timer()
'This timer calculates time left, and takes care of after game procedures
    If gTime > 0 Then 'If games still going
        gTime = gTime - 1
        lblTime = gTime
    Else 'If game ended
        'Stop all things going on in the game
        tmrRandomAlien.Enabled = False
        Timer3.Enabled = False
        Go = False
        lblEnd.Visible = True
        
        ' Show user the Comment from his/her percentage
        lblComment.Visible = True
        lblComment.Caption = Comments(CInt(Perc))
        
        'Give them the option to start a new game
        cmdStart.Visible = True
        
        'Play Game Over Wav
        Call Play_Wav("GameOver")
        
        'Now stop this timer
        Timer2.Enabled = False
    End If
End Sub

Private Sub Timer3_Timer()
'This timer will be used for making one shape run away from the other, and collision detection
    If Shape1.Left < (Shape2.Left + Shape2.Width) And Shape1.Left > (Shape2.Left - Shape2.Width) Then
        If Shape1.Top < (Shape2.Top + Shape2.Height) And Shape1.Top > (Shape2.Top - Shape2.Height) Then
            Call UpScore(lblScore)                              ' Update Score
            MoveAlien = 3                                       ' Set Move to 3
            Call Play_Wav("Caught")                             'Play the Caught Wav
            Call Mover(Shape2, lblOutOf, MoveAlien, lblOutOf)   ' Do Mover Function
        End If
    End If
End Sub
