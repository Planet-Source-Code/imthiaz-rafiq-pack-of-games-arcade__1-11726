VERSION 5.00
Begin VB.UserControl Player 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4020
   ScaleHeight     =   3675
   ScaleWidth      =   4020
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.Timer TimerAtEndFile 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   3270
   End
   Begin VB.Timer TimerMisc 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   330
      Top             =   3300
   End
   Begin VB.Frame Frame1 
      Height          =   3075
      Left            =   270
      TabIndex        =   0
      Top             =   240
      Width           =   3495
      Begin VB.PictureBox FrameVideo 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   750
         ScaleHeight     =   1500
         ScaleWidth      =   1995
         TabIndex        =   2
         Top             =   270
         Width           =   2000
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Left            =   60
         TabIndex        =   1
         Top             =   2220
         Width           =   3375
         Begin Project1.ProgYbar Slider 
            Height          =   225
            Left            =   810
            TabIndex        =   4
            Top             =   180
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   397
            ForeColor       =   255
            BackColor       =   0
            Max             =   100
            Mode            =   0
            Border          =   1
            Mark            =   -1  'True
            MarkThicness    =   3
            MarkColor       =   65535
         End
         Begin VB.Image playb 
            Height          =   255
            Left            =   150
            Picture         =   "UserControl1.ctx":0312
            Top             =   150
            Width           =   555
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0080FF80&
         Height          =   225
         Left            =   60
         TabIndex        =   3
         Top             =   2700
         Width           =   3375
      End
   End
   Begin VB.Image Pausei 
      Height          =   255
      Left            =   2940
      Picture         =   "UserControl1.ctx":0AAB
      Top             =   3780
      Width           =   555
   End
   Begin VB.Image playi 
      Height          =   255
      Left            =   1950
      Picture         =   "UserControl1.ctx":1228
      Top             =   3780
      Width           =   555
   End
End
Attribute VB_Name = "Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim typeDevice As String
Dim AliasName As String
Dim Result As String

Dim Child As Integer


Dim Filename As String

Dim LbActualCx As Long
Dim LbActualCy As Long
Dim LbFramesPerSecond As Long
Dim LbTotalFrames As Long
Dim LbTotalTime As Long
Dim LbCurrPos As Long
Dim Started As Boolean
Dim PlayMode As String
Dim OpenedSucess As Boolean
Dim Autostart As Boolean


Public Event OnTimer(PlayerPlayMode As String, PlayerCurrentTime As Long, PlayerOldTime As Long, PlayerTotalTime As Long)

Public Event OnPlayerError(PlayerError As String, PlayerErrorNo As Integer)

Public Event OnSelfResize()

Public Event OnPlayFinish()

Public Event OnMouseMove(Button As Integer, X As Single, Y As Single)


Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OnMouseMove(Button, X, Y)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OnMouseMove(Button, X, Y)
End Sub

Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OnMouseMove(Button, X, Y)
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OnMouseMove(Button, X, Y)
End Sub

Private Sub FrameVideo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OnMouseMove(Button, X, Y)
End Sub

Private Sub FrameVideo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OnMouseMove(Button, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OnMouseMove(Button, X, Y)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OnMouseMove(Button, X, Y)
End Sub

Private Sub playb_Click()
If OpenedSucess = False Then Label1.Caption = "File not loaded..": Exit Sub
If Started = False And PlayMode = "" Then
  Moviesize
  
  Result = PlayMultimedia(AliasName, "", "")
  If Result = "Success" Then
    Label1.Caption = "Playing.."
    Slider.DrawBar 0
    TimerAtEndFile.Enabled = True
    TimerMisc.Enabled = True
    PlayMode = "Play"
    playb.Picture = Pausei.Picture
    Started = True
  End If
ElseIf Started = True And PlayMode = "Play" Then
  
  Result = PauseMultimedia(AliasName)
  Label1.Caption = "Paused.."
  If Result = "Success" Then
    TimerAtEndFile.Enabled = False
    playb.Picture = playi.Picture
    TimerMisc.Enabled = False
    PlayMode = "Pause"
  End If
ElseIf Started = True And PlayMode = "Pause" Then
  
  Result = ResumeMultimedia(AliasName)
  Label1.Caption = "Playing.."
  If Result = "Success" Then
    TimerAtEndFile.Enabled = True
    playb.Picture = Pausei.Picture
    TimerMisc.Enabled = True
    PlayMode = "Play"
  End If
End If

End Sub

Private Sub Slider_click(Value As Double)

Dim SValue As Long

SValue = Value
If OpenedSucess = False Then Slider.DrawBar 0: Label1.Caption = "File not loaded..": Exit Sub

If LbFramesPerSecond = 0 And SValue = -1 Then Exit Sub 'if this alias not opened then exit (improtant)


Dim pos As Long

 'this is the main improtant point to select the file which you want change position for it


pos = SValue * LbFramesPerSecond

If Started = False And PlayMode = "" Then
  Result = MoveMultimedia(AliasName, pos)      'call now function MoveMultimedia
  Result = PauseMultimedia(AliasName)
  Started = True
  PlayMode = "Pause"
ElseIf Started = True And PlayMode = "Play" Then
  Result = MoveMultimedia(AliasName, pos)
ElseIf Started = True And PlayMode = "Pause" Then
  Result = MoveMultimedia(AliasName, pos)      'call now function MoveMultimedia
  Result = PauseMultimedia(AliasName)
End If

If Result = "Success" Then 'this mean MoveMultimedia success
  Label1.Caption = "Moved to " + CalTime(SValue) + " of " + CalTime(LbTotalTime)
Else 'not success
  RaiseEvent OnPlayerError(Result, ErrorNo(Result))
  Slider.DrawBar 0
End If

End Sub

Private Sub Slider_ValueChange(Newval As Double, Oldval As Double)
Dim xx As Long
Dim xr As Long
Dim xg As Long
Dim xc As Long
xx = Newval
xr = Oldval
xg = Slider.Max

RaiseEvent OnTimer(PlayMode, xx, xr, xg)

End Sub

Private Sub TimerMisc_Timer()

Dim Percent As Double

If Started = True Then
  LbCurrPos = GetCurrentMultimediaPos(AliasName)
  Percent = GetPercent(AliasName)
  Slider.DrawBar LbCurrPos / LbFramesPerSecond
  Label1.Caption = "Played " + CalTime(LbCurrPos / LbFramesPerSecond) + " of " + CalTime(LbTotalTime)
Else
  Label1.Caption = "Idle"
End If

End Sub

Private Sub TimerAtEndFile_Timer()

 'this is the main improtant point to select the file which you want change position for it

If AreMultimediaAtEnd(AliasName, LbTotalFrames) = True Then ' alias name for e.g.:"movie"
     playb.Picture = playi.Picture
     PlayMode = ""
     Started = False
     Slider.DrawBar Slider.Max
     TimerMisc.Enabled = False
     TimerAtEndFile.Enabled = False
     RaiseEvent OnPlayFinish
End If


End Sub



Private Sub UserControl_Initialize()



If Not GetDefaultDevice("MPEGVideo") = "mciqtz.drv" Then
  SetDefaultDevice "MPEGVideo", "mciqtz.drv"
End If

If Not GetDefaultDevice("sequencer") = "mciseq.drv" Then
  SetDefaultDevice "sequencer", "mciseq.drv"
End If

If Not GetDefaultDevice("avivideo") = "mciavi.drv" Then
  SetDefaultDevice "avivideo", "mciavi.drv"
End If

Started = True
PlayMode = ""


End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OnMouseMove(Button, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OnMouseMove(Button, X, Y)
End Sub

Private Sub UserControl_Resize()
RaiseEvent OnSelfResize
If FrameVideo.Enabled = True Then
  UserControl.Height = 2500
  UserControl.Width = 3525
  Frame1.top = 0
  Frame1.left = 0
  Frame1.Height = 2475
  Frame1.Width = 3495
  FrameVideo.top = 180
  FrameVideo.left = 720
  Frame2.top = 1650
  Frame2.left = 60
  Label1.top = 2195
  Label1.left = 60
Else
  UserControl.Height = 945
  UserControl.Width = 3525
  Frame1.top = 0
  Frame1.left = 0
  Frame1.Height = 885
  Frame1.Width = 3495
  Frame2.top = 120
  Frame2.left = 60
  Label1.top = 600
  Label1.left = 60
  FrameVideo.top = 10000
  FrameVideo.left = 10000
End If

End Sub

Private Sub UserControl_Terminate()

TimerMisc.Enabled = False
TimerAtEndFile.Enabled = False

DoEvents
If OpenedSucess = True Then
  
  AliasName = "movie" & Child
  
  Result = CloseMultimedia(AliasName)
  
End If

If Result = "Success" Then 'this mean CloseAll success
'Write your commands here
Else 'not success
  RaiseEvent OnPlayerError(Result, ErrorNo(Result))
  Result = CloseMultimedia(AliasName)
End If

End Sub
Sub CloseAllPlayers()
  Result = CloseAll
  RaiseEvent OnPlayerError(Result, ErrorNo(Result))
End Sub

Sub StopPlay()
If Started = True Then
  Started = False
  PlayMode = ""
  Result = StopMultimedia(AliasName)
  Slider_click 0
  Label1.Caption = "Idle"
End If
End Sub


Sub Play()
If PlayMode = "" And Started = False Then
  playb_Click
End If
End Sub
Sub Pause()
If Started = True And PlayMode = "Play" Then
  playb_Click
End If
End Sub
Sub ResumePlay()
If Started = True And PlayMode = "Pause" Then
  playb_Click
End If
End Sub

Sub Moviesize()
Attribute Moviesize.VB_MemberFlags = "40"

 'this is the main improtant point to select the file which you want to resize it

Result = PutMultimedia(FrameVideo.hWnd, AliasName, 0, 0, 0, 0)         'call now function PutMultimedia

End Sub

Function CalTime(Timerx As Long) As String
Attribute CalTime.VB_MemberFlags = "40"
Dim X As Long
Dim Y As Long


X = Int(Timerx / 60)

Y = Int(Timerx - (X * 60))

CalTime = LTrim(RTrim(Str(X))) + ":" + LTrim(RTrim(Str(Y)))

End Function

Sub Loadfile()
Attribute Loadfile.VB_MemberFlags = "40"

Started = False
PlayMode = ""

If Filename = "" Then
  OpenedSucess = False
  Label1.Caption = ""
  Started = False
  PlayMode = ""
  Exit Sub
End If


Label1.Caption = "Loading.."

DoEvents


If Right(Filename, 4) = ".avi" Then
    typeDevice = "AviVideo"
ElseIf Right(Filename, 4) = ".rmi" Or Right(Filename, 4) = ".mid" Then
    typeDevice = "sequencer"
Else
    typeDevice = "MPEGVideo"
End If

OpenedSucess = False

Again:

AliasName = "movie" & Child


Result = OpenMultimedia(FrameVideo.hWnd, AliasName, Filename, typeDevice)      'call now function OpenMultimedia

If ErrorNo(Result) = 0 Then 'this mean OpenMultimedia success
  LbActualCx = GetSize(AliasName, "cx")
  LbActualCy = GetSize(AliasName, "cy")
  
  If LbActualCx <> -1 And LbActualCy <> -1 Then
    FrameVideo.Enabled = True
  Else
    FrameVideo.Enabled = False
  End If
  
  UserControl_Resize
  PlayMode = ""
  Started = False
  playb.Picture = playi.Picture
  TimerMisc.Enabled = False
  TimerAtEndFile.Enabled = False
  LbFramesPerSecond = GetFramesPerSecond(AliasName)
  LbTotalFrames = GetTotalframes(AliasName)  'Get total frames
  LbTotalTime = GetTotalTimeByMS(AliasName) / 1000   'Get Total Time
  Slider.Max = LbTotalFrames / LbFramesPerSecond
  Label1.Caption = "Loaded Sucessfully"
  OpenedSucess = True
  If Autostart = True Then playb_Click
Else
  Select Case ErrorNo(Result)
    Case 277
      Label1.Caption = "Cannot Initialize Sound.."
    Case 263
      Label1.Caption = "Invalid Media File.."
    Case 289
      Result = CloseMultimedia(AliasName)
      GoTo Again
    Case Else
      Label1.Caption = "Unknown Error.."
      RaiseEvent OnPlayerError(Result, ErrorNo(Result))
    End Select
  OpenedSucess = False
End If

End Sub
Function ErrorNo(Errstring As String) As Integer
  ErrorNo = Val(Mid$(Errstring, 9, 3))
End Function


'____________________________________________


Public Property Get File() As String
     File = Filename
End Property

Public Property Let File(ByVal N_File As String)
    Filename = N_File
    Loadfile
    PropertyChanged "File"
End Property

Public Property Get CurrentPosition() As Long
Attribute CurrentPosition.VB_MemberFlags = "400"
     CurrentPosition = LbCurrPos
End Property

Public Property Let CurrentPosition(ByVal Npos As Long)
If Npos <> 0 Then
    LbCurrPos = Npos
    Slider_click (Npos / LbFramesPerSecond)
    PropertyChanged "CurrentPosition"
End If
End Property


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("File", File, "")
    Call PropBag.WriteProperty("CurrentPosition", LbCurrPos, "")
    Call PropBag.WriteProperty("ChildNo", Child, Rnd * 1000)
    Call PropBag.WriteProperty("PlayerAutoStart", Autostart, False)
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    File = PropBag.ReadProperty("File", "")
    CurrentPosition = PropBag.ReadProperty("CurrentPosition", 0)
    ChildNo = PropBag.ReadProperty("ChildNo", Rnd * 1000)
    PlayerAutoStart = PropBag.ReadProperty("PlayerAutoStart", False)
End Sub


Public Property Get ChildNo() As Integer
     ChildNo = Child
End Property

Public Property Let ChildNo(ByVal NChild As Integer)
    Child = NChild
    PropertyChanged "Child"
End Property

Public Property Get PlayerAutoStart() As Boolean
     PlayerAutoStart = Autostart
End Property

Public Property Let PlayerAutoStart(ByVal NS As Boolean)
    Autostart = NS
    PropertyChanged "PlayerAutoStart"
End Property

