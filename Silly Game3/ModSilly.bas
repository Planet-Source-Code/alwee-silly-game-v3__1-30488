Attribute VB_Name = "ModSilly"
Option Explicit

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const Flags& = SND_ASYNC Or SND_NODEFAULT


'*********************************************************************
'Any codes in a module run faster than codes on a form,
'or so I hear, so I'm going to convert alot of things to this module
'If someone wants to run test I'd appreciate it
'Alot of the codes I've added here I changed for faster execution
'*********************************************************************

Public Function Comments(PercNum As Integer) As String
'Thanks to Shaun this function returns comments
    Select Case PercNum
        Case Is <= 20
            Comments$ = "That was terrible"
        Case Is > 20 And PercNum <= 50
            Comments$ = "You could do alot better"
        Case Is > 50 And PercNum <= 60
            Comments$ = "A bit better I think..."
        Case Is > 60 And PercNum <= 75
            Comments$ = "The Gerbil is smiling..."
        Case Is > 75 And PercNum <= 90
            Comments$ = "Not Bad, Not Bad At All"
        Case Is > 90 And PercNum <= 100
            Comments$ = "You Are The Gerbil King"
    End Select
End Function

Public Function UpScore(Scorelabel As Label)
'Add one to score and update label
    Scorelabel.Caption = CStr(CInt(Scorelabel.Caption) + 1)
End Function

Public Function Mover(TargetShape As Shape, LabelOut As Label, varMoveAlien As Byte, LabelOutOf As Label)
'This sub moves the shape around
    Dim X As Integer
    Dim Y As Integer

    ' Only move the alien if 'MoveAlien' = 3
    ' This is either random or set
    Select Case varMoveAlien
        Case 3
            Randomize Timer
            X = (Rnd * (4680 - TargetShape.Height)) + 1
            Y = (Rnd * (2520 - TargetShape.Width)) + 1
            TargetShape.Left = X
            TargetShape.Top = Y
            LabelOutOf.Caption = CStr(CInt(LabelOutOf.Caption) + 1)
        Case Else
            'Do nothing if MoveAlien is not 3
    End Select
End Function

Public Sub Play_Wav(Command As String)
'In case you wondered why I didn't add a pause,loop or stop sub
'It's because the wavs aren't long enough to need them
    Select Case Command
        Case Is = "Splash"
            Call sndPlaySound(GetSoundName("Splash"), Flags&)
        Case Is = "Caught"
            Call sndPlaySound(GetSoundName("Caught"), Flags&)
        Case Is = "GameOver"
            Call sndPlaySound(GetSoundName("Over"), Flags&)
    End Select
End Sub

Public Function GetSoundName(Command) As String
'Sub gets the name if the file we want to play and
'takes care of that little "\" problem when getting a programs path
Dim Path As String
GetSoundName = ""

    'This part gets the path
    Path$ = App.Path
    If Right$(Path$, 1) = "\" Then
        GetSoundName$ = Left$(Path$, Len(Path$) - 1)
    ElseIf Right$(Path$, 1) <> "\" Then
        'Do nothing it's already right
    End If
    
    'This part gets the filename and puts em together
    If Command = "Splash" Then
        GetSoundName = Path$ & "\Splash.wav"
    End If
    If Command = "Caught" Then
        GetSoundName = Path$ & "\Caught.wav"
    End If
    If Command = "Over" Then
        GetSoundName = Path$ & "\Over.wav"
    End If
End Function
