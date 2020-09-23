VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8220
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   300
      TabIndex        =   0
      Top             =   4710
      Width           =   6975
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   5220
         TabIndex        =   5
         Top             =   210
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "60"
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
         Left            =   3660
         TabIndex        =   4
         Top             =   210
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Time Left :"
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
         Left            =   2010
         TabIndex        =   3
         Top             =   210
         Width           =   1485
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Score :"
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
         Left            =   240
         TabIndex        =   2
         Top             =   180
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "0"
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
         Left            =   1350
         TabIndex        =   1
         Top             =   210
         Width           =   180
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   4410
      Top             =   1140
   End
   Begin VB.Timer ClickT 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   500
      Left            =   3630
      Top             =   2520
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   4500
      Picture         =   "Form1.frx":0442
      Top             =   2940
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2640
      Picture         =   "Form1.frx":0884
      Top             =   2640
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3030
      Picture         =   "Form1.frx":0CC6
      Top             =   1260
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Click1 
      Height          =   840
      Index           =   0
      Left            =   1380
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":0FD0
      Stretch         =   -1  'True
      Top             =   450
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image mouseX 
      Height          =   480
      Left            =   1380
      OLEDropMode     =   2  'Automatic
      Picture         =   "Form1.frx":1412
      Top             =   1860
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PlayeRname As String
Dim TSec As Integer

Dim ClickX As Long
Dim ClickY As Long

Sub LoadClick1()
  ClickX = ClickX + 1
  Load Click1(ClickX)
  Load ClickT(ClickX)
  Click1(ClickX).Top = Rnd * (Form1.Height - (Frame1.Height * 1.8))
  Click1(ClickX).Left = Rnd * Form1.Width * 0.95
  Click1(ClickX).MouseIcon = mouseX
  Click1(ClickX).Visible = True
  ClickT(ClickX).Enabled = True
  If ClickX = 500 Then ClickX = 1
  
End Sub

Private Sub Click1_Click(Index As Integer)
  Click1(Index).Picture = Image3.Picture
  Label1.Caption = Val(Label1.Caption) + 1
  ClickT(Index).Enabled = False
  Unload Click1(Index)
  Unload ClickT(Index)

End Sub

Private Sub ClickT_Timer(Index As Integer)
If Click1(Index).Picture = Image1.Picture Or Click1(Index).Picture = Image3.Picture Then
  Unload Click1(Index)
  Unload ClickT(Index)
Else
  Click1(Index).Picture = Image1.Picture
End If

End Sub

Private Sub Command1_Click()
If Form1.WindowState = 0 Then
  Form1.WindowState = 2
  Form1.Caption = ""
Else
  Form1.WindowState = 0
  Form1.Caption = "Hit Me If You Can"
End If
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  MsgBox "Got Bugged Up. Hehehehehe"
  End
End If
End Sub

Private Sub Form_Load()
  TSec = 30
  ClickX = 1
  PlayeRname = InputBox("Please enter you name", "Player Name", "Yelam Programmers")
End Sub

Private Sub Form_Resize()
  Frame1.Top = Form1.Height - Frame1.Height * 1.5
  Frame1.Left = 0
  Frame1.Width = Form1.Width
End Sub

Private Sub Timer1_Timer()
LoadClick1
End Sub

Private Sub Timer2_Timer()
TSec = TSec - 1
Label4.Caption = LTrim(RTrim(Str(TSec))) + " Sec"
If TSec = 0 Then
  Dim x As VbMsgBoxResult
  Timer1.Enabled = False
  Timer2.Enabled = False
  MsgBox PlayeRname + ". Your Score is " + Label1.Caption, vbInformation
  x = MsgBox("Do you want to play again", vbInformation + vbYesNo)
  If x = vbYes Then
    TSec = 30
    Timer1.Enabled = True
    Timer2.Enabled = True
    Label1.Caption = ""
  Else
    End
  End If
End If
End Sub
