VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Random Clicks.. By a Kid"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   90
      Top             =   3840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6990
      Top             =   3270
   End
   Begin VB.Frame Frame2 
      Height          =   2985
      Left            =   7080
      TabIndex        =   7
      Top             =   30
      Width           =   2205
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   720
         TabIndex        =   13
         Top             =   2250
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Time Elapsed :"
         Height          =   195
         Left            =   210
         TabIndex        =   12
         Top             =   1110
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   720
         TabIndex        =   11
         Top             =   1455
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No Of Tries :"
         Height          =   195
         Left            =   210
         TabIndex        =   10
         Top             =   1890
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   720
         TabIndex        =   9
         Top             =   660
         Width           =   45
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Time Left :"
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   330
         Width           =   750
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "New Game"
      Height          =   795
      Left            =   7470
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1545
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pause Game"
      Enabled         =   0   'False
      Height          =   795
      Left            =   7470
      Picture         =   "Form1.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4095
      Width           =   1545
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit Game"
      Height          =   795
      Left            =   7470
      Picture         =   "Form1.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5070
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   5835
      Left            =   840
      TabIndex        =   0
      Top             =   30
      Width           =   6135
      Begin VB.Image EventX 
         Height          =   480
         Index           =   0
         Left            =   150
         Picture         =   "Form1.frx":1108
         Top             =   300
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Image IX 
      Height          =   480
      Index           =   4
      Left            =   2610
      Picture         =   "Form1.frx":154A
      Top             =   6300
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IX 
      Height          =   480
      Index           =   3
      Left            =   2010
      Picture         =   "Form1.frx":1854
      Top             =   6300
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IX 
      Height          =   480
      Index           =   2
      Left            =   1410
      Picture         =   "Form1.frx":1C96
      Top             =   6330
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IX 
      Height          =   480
      Index           =   1
      Left            =   930
      Picture         =   "Form1.frx":20D8
      Top             =   6270
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IX 
      Height          =   480
      Index           =   0
      Left            =   330
      Picture         =   "Form1.frx":251A
      Top             =   6210
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Winner"
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   3270
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Try Again"
      Height          =   195
      Left            =   68
      TabIndex        =   2
      Top             =   2199
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Click This"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   1128
      Width           =   690
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   165
      Picture         =   "Form1.frx":2824
      Top             =   2592
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   165
      Picture         =   "Form1.frx":2C66
      Top             =   450
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   165
      Picture         =   "Form1.frx":30A8
      Top             =   1521
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Dim Winner As Integer
Dim Prev As Integer
Dim Tsec As Integer
Dim Tesec As Integer
Dim Nclick As Integer

Private Sub Command1_Click()
Timer1.Enabled = False
End
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Pause Game" Then
  Frame1.Enabled = False
  Timer1.Enabled = False
  Command2.Caption = "Resume Game"
Else
  Frame1.Enabled = True
  Timer1.Enabled = True
  Command2.Caption = "Pause Game"
End If
End Sub

Private Sub Command3_Click()
X = 0
X = 0
For I = 1 To 10
  For j = 1 To 10
    X = X + 1
    EventX(X).Picture = Image2.Picture
  Next j
Next I

Again:
Randomize Timer
Winner = 1 + Int(Rnd * 100)
If Winner > 100 Or Winner < 0 Then GoTo Again
Tsec = 30
Tesec = 0
Nclick = 0
Timer2.Enabled = False
Command2.Enabled = True
Frame1.Enabled = True
Timer1.Enabled = True
End Sub

Private Sub EventX_Click(Index As Integer)
EventX(Prev).Picture = Image2.Picture
Nclick = Nclick + 1
sndPlaySound App.Path + "\a.wav", 0
If Index <> Winner Then
  EventX(Index).Picture = Image1.Picture
  Label9.Caption = Nclick
Else
  EventX(Index).Picture = Image3.Picture
  Frame1.Enabled = False
  Timer1.Enabled = False
  Command2.Enabled = False
  Timer2.Enabled = True
  MsgBox "You Got it right in" + Str(Nclick) + " tries" + vbCrLf + vbCrLf + "Done by a kid "
End If
Prev = Index
End Sub

Private Sub Form_Load()
X = 0
For I = 1 To 10
For j = 1 To 10
  X = X + 1
  Load EventX(X)
  EventX(X).Top = (500 * I)
  EventX(X).Left = (500 * j)
  EventX(X).Visible = True
Next j
Next I
End Sub

Private Sub Timer1_Timer()

Label5.Caption = Str(Tsec) + " Sec"
Label7.Caption = Str(Tesec) + " Sec"
If Tsec = 0 Then
  Timer1.Enabled = False
  Frame1.Enabled = False
  Command2.Enabled = False
  MsgBox "You Got it wrong in" + Str(Nclick) + " tries" + vbCrLf + vbCrLf + "Done by a kid " + vbCrLf + vbCrLf + vbCrLf + vbCrLf + "Game Over", vbCritical
  EventX(Winner).Picture = Image3.Picture
  Timer2.Enabled = True
End If

Tsec = Tsec - 1
Tesec = Tesec + 1
End Sub

Private Sub Timer2_Timer()
Static sd As Integer
EventX(Winner).Picture = IX(sd)
sd = sd + 1
If sd = 5 Then sd = 0
End Sub
