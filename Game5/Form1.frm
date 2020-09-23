VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7275
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9930
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Project1.Player Player2 
      Height          =   945
      Left            =   -60
      TabIndex        =   28
      Top             =   1170
      Visible         =   0   'False
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   1667
      CurrentPosition =   0
      ChildNo         =   0
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6705
      Left            =   690
      TabIndex        =   0
      Top             =   630
      Width           =   8865
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6400
         Left            =   150
         ScaleHeight     =   6405
         ScaleWidth      =   4725
         TabIndex        =   25
         Top             =   120
         Width           =   4720
         Begin VB.Image Road 
            Height          =   480
            Index           =   0
            Left            =   840
            Picture         =   "Form1.frx":0442
            Top             =   2070
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2130
         Top             =   1980
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   330
         Top             =   4920
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Car Acceleration"
         Enabled         =   0   'False
         Height          =   675
         Left            =   4980
         TabIndex        =   21
         Top             =   240
         Width           =   3705
         Begin Project1.ProgYbar Speedometer 
            Height          =   225
            Left            =   150
            TabIndex        =   22
            Top             =   300
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   397
            ForeColor       =   16744576
            BackColor       =   0
            Max             =   1000
            Mode            =   0
            Border          =   0
            Mark            =   -1  'True
            MarkThicness    =   3
            MarkColor       =   65535
         End
         Begin VB.Label Speedl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   3060
            TabIndex        =   24
            Top             =   270
            Width           =   45
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Speed"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Left            =   120
            TabIndex        =   23
            Top             =   0
            Width           =   2085
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "S&tart"
         Height          =   525
         Left            =   5010
         TabIndex        =   20
         Top             =   4830
         Width           =   1245
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Car Damage"
         Enabled         =   0   'False
         Height          =   675
         Left            =   4980
         TabIndex        =   16
         Top             =   960
         Width           =   3705
         Begin Project1.ProgYbar Damage 
            Height          =   225
            Left            =   180
            TabIndex        =   17
            Top             =   300
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   397
            ForeColor       =   16711680
            BackColor       =   0
            Max             =   250
            Mode            =   0
            Border          =   0
            Mark            =   -1  'True
            MarkThicness    =   3
            MarkColor       =   65535
         End
         Begin VB.Label Damagel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   3060
            TabIndex        =   19
            Top             =   270
            Width           =   45
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Damage"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Left            =   120
            TabIndex        =   18
            Top             =   0
            Width           =   2085
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Track Complete By"
         Enabled         =   0   'False
         Height          =   675
         Left            =   4980
         TabIndex        =   12
         Top             =   1710
         Width           =   3705
         Begin Project1.ProgYbar Roadleft 
            Height          =   225
            Left            =   150
            TabIndex        =   13
            Top             =   300
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   397
            ForeColor       =   12582912
            BackColor       =   0
            Max             =   1000
            Mode            =   0
            Border          =   0
            Mark            =   -1  'True
            MarkThicness    =   3
            MarkColor       =   65535
         End
         Begin VB.Label Roadleftl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   3060
            TabIndex        =   15
            Top             =   270
            Width           =   45
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Track Status"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Left            =   120
            TabIndex        =   14
            Top             =   0
            Width           =   2085
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1275
         Left            =   5010
         TabIndex        =   7
         Top             =   2520
         Width           =   3705
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Score :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   90
            TabIndex        =   11
            Top             =   210
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   900
            TabIndex        =   10
            Top             =   240
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   90
            TabIndex        =   9
            Top             =   780
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   900
            TabIndex        =   8
            Top             =   780
            Width           =   45
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Level"
         Height          =   885
         Left            =   5010
         TabIndex        =   4
         Top             =   3870
         Width           =   3735
         Begin VB.ComboBox Level 
            Height          =   315
            Left            =   180
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   330
            Width           =   3435
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Level"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Left            =   120
            TabIndex        =   6
            Top             =   0
            Width           =   2085
         End
      End
      Begin VB.Timer SoundX 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   2070
         Top             =   3150
      End
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3210
         Top             =   3090
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Pause"
         Enabled         =   0   'False
         Height          =   525
         Left            =   6255
         TabIndex        =   3
         Top             =   4830
         Width           =   1245
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Stop"
         Enabled         =   0   'False
         Height          =   525
         Left            =   7500
         TabIndex        =   2
         Top             =   4830
         Width           =   1245
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Exit"
         Height          =   525
         Left            =   6240
         TabIndex        =   1
         Top             =   5460
         Width           =   1245
      End
      Begin VB.Image R 
         Height          =   630
         Index           =   0
         Left            =   5040
         Picture         =   "Form1.frx":074C
         Top             =   5400
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Image R 
         Height          =   630
         Index           =   2
         Left            =   6540
         Picture         =   "Form1.frx":1D36
         Top             =   5400
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Image R 
         Height          =   630
         Index           =   3
         Left            =   8040
         Picture         =   "Form1.frx":3320
         Top             =   5400
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Lastkey 
         Height          =   315
         Left            =   6930
         TabIndex        =   26
         Top             =   4080
         Width           =   1545
      End
      Begin VB.Image R 
         Height          =   630
         Index           =   6
         Left            =   7305
         Picture         =   "Form1.frx":490A
         Top             =   6090
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Image R 
         Height          =   630
         Index           =   7
         Left            =   6540
         Picture         =   "Form1.frx":5EF4
         Top             =   6090
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Image R 
         Height          =   630
         Index           =   4
         Left            =   7290
         Picture         =   "Form1.frx":74DE
         Top             =   5400
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Image R 
         Height          =   630
         Index           =   1
         Left            =   5790
         Picture         =   "Form1.frx":8AC8
         Top             =   5400
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Image R 
         Height          =   630
         Index           =   5
         Left            =   8070
         Picture         =   "Form1.frx":A0B2
         Top             =   6090
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Image R 
         Height          =   630
         Index           =   8
         Left            =   5010
         Picture         =   "Form1.frx":B69C
         Top             =   6090
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Image R 
         Height          =   630
         Index           =   9
         Left            =   5775
         Picture         =   "Form1.frx":CC86
         Top             =   6090
         Visible         =   0   'False
         Width           =   660
      End
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   8730
      Top             =   8070
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6360
      Top             =   7470
   End
   Begin Project1.Player Player1 
      Height          =   945
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   1667
      CurrentPosition =   0
      ChildNo         =   1
   End
   Begin VB.Image R2 
      Height          =   630
      Index           =   0
      Left            =   8730
      Picture         =   "Form1.frx":E270
      Top             =   8070
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image R2 
      Height          =   630
      Index           =   1
      Left            =   8730
      Picture         =   "Form1.frx":F85A
      Top             =   8070
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image R4 
      Height          =   630
      Index           =   1
      Left            =   8730
      Picture         =   "Form1.frx":10E44
      Top             =   8070
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image R4 
      Height          =   630
      Index           =   0
      Left            =   8730
      Picture         =   "Form1.frx":1242E
      Top             =   8070
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image R3 
      Height          =   630
      Index           =   0
      Left            =   8730
      Picture         =   "Form1.frx":13A18
      Top             =   8070
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image R3 
      Height          =   630
      Index           =   1
      Left            =   8730
      Picture         =   "Form1.frx":15002
      Top             =   8070
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image R5 
      Height          =   630
      Index           =   1
      Left            =   8730
      Picture         =   "Form1.frx":165EC
      Top             =   8070
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image R5 
      Height          =   630
      Index           =   0
      Left            =   8730
      Picture         =   "Form1.frx":17BD6
      Top             =   8070
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image R9 
      Height          =   630
      Index           =   1
      Left            =   8730
      Picture         =   "Form1.frx":191C0
      Top             =   8070
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image R9 
      Height          =   630
      Index           =   0
      Left            =   8730
      Picture         =   "Form1.frx":1A7AA
      Top             =   8070
      Visible         =   0   'False
      Width           =   660
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Dim Mat(10, 7) As Integer
Dim CarPos As Integer
Dim CarDamage As Double
Dim RoadLen As Double
Dim Tsec As Double
Dim Score As Long
Dim LevelR As Double

Private Sub Command1_Click()
X = 0
For i = 1 To 10
  s = 1
For j = 1 To 7
  X = X + 1
  If Rnd > LevelR And s <> 3 Then
    s = s + 1
    If Rnd < 0.9 Then
      Mat(1, j) = 2 + Rnd * 2
    Else
      Mat(1, j) = 3
    End If
  Else
     If Mat(2, j) = 1 Then
      Mat(1, j) = 0
    Else
      Mat(1, j) = 1
    End If
  End If
  Road(X).Picture = R(Mat(i, j)).Picture
  Road(X).Visible = True
  'MsgBox X
Next j
Next i
Timer1.Interval = 1000
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Level.Enabled = False
CarDamage = 0
RoadLen = 0
Tsec = 0
Score = 0
Command2.Enabled = True
Command1.Enabled = False
Command3.Enabled = True
End Sub



Private Sub Command2_Click()
If Command2.Caption = "&Pause" Then
  Command2.Caption = "&Resume"
  Timer1.Enabled = False
  Timer2.Enabled = False
  Timer3.Enabled = False
  Timer4.Enabled = False
  Timer5.Enabled = False
Else
  Command2.Caption = "&Pause"
  Timer1.Enabled = True
  Timer2.Enabled = True
  Timer3.Enabled = True
  Timer4.Enabled = True
  Timer5.Enabled = True
End If
End Sub

Private Sub Command3_Click()
  Timer1.Enabled = False
  Timer2.Enabled = False
  Timer3.Enabled = False
  Timer4.Enabled = False
  Timer5.Enabled = False
  Level.Enabled = True
  Command1.Enabled = True
  Command2.Enabled = False
  Command3.Enabled = False
End Sub

Private Sub Command4_Click()
  End
End Sub

Private Sub Damage_ValueChange(Newval As Double, Oldval As Double)
  Damagel.Caption = Str(LTrim(RTrim(Int((Newval / 250) * 100)))) + " %"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Timer1.Enabled = True Then
  If KeyCode = 39 Then
    Lastkey.Caption = "Right"
    If CarPos < 7 Then
      CarPos = CarPos + 1
    End If
  End If
  If KeyCode = 37 Then
    Lastkey.Caption = "Left"
    If CarPos > 1 Then
      CarPos = CarPos - 1
    End If
  End If
  If KeyCode = 38 Then
    If Timer1.Interval - 10 > 0 Then Timer1.Interval = Timer1.Interval - 10
    Lastkey.Caption = "Up"
  End If
  If KeyCode = 40 Then
    If Timer1.Interval + 10 < 1000 Then Timer1.Interval = Timer1.Interval + 10
    Lastkey.Caption = "Down"
  End If
End If
End Sub

Private Sub Form_Load()
Level.AddItem "Level 1 - Easy Only For Kinder Gardens"
Level.AddItem "Level 2 - Medium for child Above 5 years "
Level.AddItem "Level 3 - Ok For chidrens Above 10 years"
Level.AddItem "Level 4 - Hard No Probelm"
Level.AddItem "Level 5 - Very Hard Try This Without Damage"


Level.ListIndex = 0

Player1.File = App.Path + "\a.wav"
Player2.File = App.Path + "\b.wav"
X = 0
For i = 1 To 10
For j = 1 To 7
  X = X + 1
  Load Road(X)
  Road(X).top = (i - 1) * 630
  Road(X).left = (j - 1) * 660
  If Rnd > 0.9 And s <> 3 Then
    s = s + 1
    If Rnd < 0.9 Then
      Mat(1, j) = 2 + Rnd * 2
    Else
      Mat(1, j) = 3
    End If
  Else
     If Mat(2, j) = 1 Then
      Mat(1, j) = 0
    Else
      Mat(1, j) = 1
    End If
  End If
  Road(X).Picture = R(Mat(i, j)).Picture
  Road(X).Visible = True
Next j
Next i
Picture1.Height = Road(X).top + Road(X).Height + 100
Picture1.Width = Road(X).left + Road(X).Width + 100
Form1.Height = Picture1.top + Picture1.Height
CarPos = 3
End Sub

Private Sub Form_Resize()
Frame1.top = (Screen.Height - Frame1.Height) / 2
Frame1.left = (Screen.Width - Frame1.Width) / 2
End Sub

Private Sub Level_Change()

Select Case Level.ListIndex
  Case 0
    LevelR = 0.9
  Case 1
    LevelR = 0.8
  Case 2
    LevelR = 0.7
  Case 3
    LevelR = 0.6
  Case 4
    LevelR = 0.4
End Select

End Sub


Private Sub Level_Click()

Select Case Level.ListIndex
  Case 0
    LevelR = 0.9
  Case 1
    LevelR = 0.8
  Case 2
    LevelR = 0.7
  Case 3
    LevelR = 0.6
  Case 4
    LevelR = 0.4
End Select

End Sub

Private Sub Roadleft_ValueChange(Newval As Double, Oldval As Double)
  Roadleftl.Caption = Str(LTrim(RTrim(Int((Newval / 1000) * 100)))) + " %"
End Sub

Private Sub SoundX_Timer()
    SoundX.Enabled = False
    sndPlaySound App.Path + "\a.wav", 0
End Sub

Private Sub Speedometer_ValueChange(Newval As Double, Oldval As Double)
  Speedl.Caption = Str(LTrim(RTrim(Int((Newval / 1000) * 100)))) + " %"
End Sub

Private Sub Timer1_Timer()
Dim s As String
Score = Score + 1
RoadLen = RoadLen + 1
For j = 10 To 2 Step -1
  For i = 1 To 7
    Mat(j, i) = Mat(j - 1, i)
  Next i
Next j
s = 1

For i = 1 To 7
  If Rnd > LevelR And s <> 5 And Mat(1, Rnd * 5) = 0 Then
    s = s + 1
    If Rnd < 0.95 Then
      Mat(1, i) = 2 + Rnd * 2
    Else
      Mat(1, i) = 5
    End If
  Else
    If Mat(2, i) = 1 Then
      Mat(1, i) = 0
    Else
      Mat(1, i) = 1
    End If
  End If
Next i

Mat(9, CarPos) = 9

Mat(10, CarPos) = Rnd * 1


For i = 1 To 7
  
  If Mat(9, i) = 6 Or Mat(9, i) = 5 Then
    Mat(9, i) = Rnd * 1
    Mat(10, i) = Rnd * 1
  End If
  
Next i


If Mat(8, CarPos) = 0 Or Mat(8, CarPos) = 1 Then
  Mat(8, CarPos) = 6
Else
  If Mat(8, CarPos) = 2 Then
    If Rnd > 0.5 Then
      If CarPos = 7 Then
         CarPos = CarPos - 1
      Else
         CarPos = CarPos + 1
      End If
    Else
      If CarPos = 1 Then
         CarPos = CarPos + 1
      Else
         CarPos = CarPos - 1
      End If
    End If
  Else
    If Mat(8, CarPos) = 5 Then
      Mat(8, CarPos) = 8
      If CarDamage > 0 Then CarDamage = CarDamage - 5
      Score = Score + Rnd * 50
    Else
      Mat(8, CarPos) = 7
      If Timer1.Interval + 50 < 1000 Then Timer1.Interval = Timer1.Interval + 50
      CarDamage = CarDamage + 10
      If Score > 0 Then Score = Score - Rnd * 100
      Player1.Play
      'SoundX.Enabled = True
    End If
  End If
End If

LoadRoadPic

End Sub
Sub LoadRoadPic()
X = 0
For i = 1 To 10
  For j = 1 To 7
    X = X + 1
    Road(X).Picture = R(Mat(i, j)).Picture
  Next j
Next i
End Sub

Private Sub Timer2_Timer()
  Randomize Timer
  Speedometer.DrawBar 1000 - Timer1.Interval
  Damage.DrawBar CarDamage
  Roadleft.DrawBar RoadLen
  Label2.Caption = Score
  If CarDamage > 250 Then
    MsgBox "Game Over " + vbCrLf + vbCrLf + "Score : " + Label2.Caption + vbCrLf + vbCrLf + "Time : " + Label4.Caption
    Open App.Path + "\scores.dat" For Append As #1
    Print #1, Label2.Caption, Label4.Caption
    Close #1
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Level.Enabled = True
    Command3_Click
  End If
  If RoadLen > 1000 Then
    MsgBox "Level Over " + vbCrLf + vbCrLf + "Score : " + Label2.Caption + vbCrLf + vbCrLf + "Time : " + Label4.Caption
    Open App.Path + "\scores.dat" For Append As #1
    Print #1, Label2.Caption, Label4.Caption
    Close #1
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Level.Enabled = True
    Command3_Click
  End If
End Sub

Private Sub Timer3_Timer()
Tsec = Tsec + 1
Label4.Caption = CalTime(Tsec) + " sec"
End Sub
Function CalTime(Timerx As Double) As String
Dim X As Long
Dim Y As Long
Dim Z As Double
Dim d As Integer
X = Int(Timerx / 60)

Y = Int(Timerx - (X * 60))

CalTime = LTrim(RTrim(Str(X))) + ":" + LTrim(RTrim(Str(Y)))
End Function

Private Sub Timer4_Timer()
Randomize Timer
R(9).Picture = R9(Rnd * 1).Picture
R(5).Picture = R5(Rnd * 1).Picture
R(3).Picture = R3(Rnd * 1).Picture
R(4).Picture = R4(Rnd * 1).Picture
R(2).Picture = R2(Rnd * 1).Picture
End Sub

Sub SoundMe()
End Sub

Private Sub Timer5_Timer()
  'sndPlaySound App.Path + "\b.wav", 0
  Player2.Play
End Sub
