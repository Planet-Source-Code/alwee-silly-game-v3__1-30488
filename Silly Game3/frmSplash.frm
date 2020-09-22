VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stupid Game"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C000C0&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   3480
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   240
      Shape           =   3  'Circle
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "By:  Awyn Berkeley (Chrome)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "By:  Shaun Rogerson (Purple Gerbil)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Label lblName1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "By:  Pube The Undead"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   3735
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************
'Eventually someone will add a picture down the line
'and this will become a real splash form
'***************************************************


Private Sub Form_Initialize()
'Play the Splash Wav
    Call Play_Wav("Splash")
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On mouse down proceed to the actual game
    frmSplash.Hide
    frmSilly.Show
End Sub
