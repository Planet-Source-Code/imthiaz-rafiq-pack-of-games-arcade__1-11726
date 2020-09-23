VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Get Jiggy With It !!!!!! By Imthiaz"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8190
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Exit"
      Height          =   525
      Left            =   5580
      TabIndex        =   21
      Top             =   6360
      Width           =   2265
   End
   Begin VB.Frame Frame3 
      Height          =   1245
      Left            =   2670
      TabIndex        =   18
      Top             =   5670
      Width           =   2535
      Begin VB.ComboBox NoX 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   810
         Width           =   2385
      End
      Begin VB.Label Label3 
         Caption         =   "No of divisions to be made on the picture in X and Y Direction"
         Height          =   555
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2355
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Pause"
      Enabled         =   0   'False
      Height          =   525
      Left            =   5580
      TabIndex        =   17
      Top             =   5660
      Width           =   2265
   End
   Begin VB.Frame Frame2 
      Height          =   1515
      Left            =   2670
      TabIndex        =   11
      Top             =   4080
      Width           =   2535
      Begin VB.Label Tries 
         AutoSize        =   -1  'True
         Caption         =   "  "
         Height          =   195
         Left            =   1260
         TabIndex        =   16
         Top             =   270
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No Of Tries :"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   270
         Width           =   900
      End
      Begin VB.Label Timet 
         AutoSize        =   -1  'True
         Caption         =   "    "
         Height          =   195
         Left            =   1260
         TabIndex        =   14
         Top             =   690
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Time Taken :"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   690
         Width           =   945
      End
      Begin VB.Label Lastkey 
         AutoSize        =   -1  'True
         Caption         =   "  "
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   1080
         Width           =   90
      End
   End
   Begin VB.PictureBox Check 
      AutoSize        =   -1  'True
      Height          =   555
      Left            =   8130
      ScaleHeight     =   495
      ScaleWidth      =   465
      TabIndex        =   10
      Top             =   4530
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4000
      Left            =   30
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   4005
      ScaleWidth      =   4005
      TabIndex        =   9
      Top             =   0
      Width           =   4000
   End
   Begin MSComDlg.CommonDialog cdq 
      Left            =   9090
      Top             =   5100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load Picture"
      Height          =   525
      Left            =   5580
      TabIndex        =   8
      Top             =   4260
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   2805
      Left            =   60
      TabIndex        =   3
      Top             =   4080
      Width           =   2535
      Begin VB.CommandButton CmdDown 
         Height          =   705
         Left            =   900
         Picture         =   "Form1.frx":B6BC
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1980
         Width           =   705
      End
      Begin VB.CommandButton CmdLeft 
         Height          =   705
         Left            =   120
         Picture         =   "Form1.frx":BAFE
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1110
         Width           =   705
      End
      Begin VB.CommandButton CmdRight 
         Height          =   705
         Left            =   1710
         Picture         =   "Form1.frx":BF40
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1110
         Width           =   705
      End
      Begin VB.CommandButton CmdUp 
         Height          =   705
         Left            =   900
         Picture         =   "Form1.frx":C382
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   300
         Width           =   705
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7590
      Top             =   6930
   End
   Begin VB.PictureBox fs 
      AutoSize        =   -1  'True
      Height          =   1410
      Left            =   8640
      Picture         =   "Form1.frx":C7C4
      ScaleHeight     =   1350
      ScaleWidth      =   1305
      TabIndex        =   2
      Top             =   4500
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.PictureBox s 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4000
      Left            =   4140
      ScaleHeight     =   4005
      ScaleWidth      =   4005
      TabIndex        =   1
      Top             =   0
      Width           =   4000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   525
      Left            =   5580
      TabIndex        =   0
      Top             =   4960
      Width           =   2265
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mat(10000) As Integer
Dim Matrix(100, 100) As Integer

Dim CurrentPosX As Integer
Dim CurrentPosY As Integer

Dim No As Integer

Dim Nt As Long
Dim Ts As Long


Dim Cx As Integer
Dim Cy As Integer

Sub Loadmat()
Dim X As Integer
Dim Y As Integer
Dim I As Integer
Dim Found As Boolean

X = 1
Do Until X = (No * No) + 1
  Randomize Timer
  Y = Rnd * (No * No)
  Found = False
  For I = 1 To X
    If Mat(I) = Y Then
      Found = True
    End If
  Next I
  If Found = False Then
    Mat(X) = Y
    X = X + 1
  End If
Loop

End Sub


Private Sub CmdUp_Click()
  If CurrentPosY <> No Then
    X = Matrix(CurrentPosX, CurrentPosY)
    Matrix(CurrentPosX, CurrentPosY) = Matrix(CurrentPosX, (CurrentPosY + 1))
    Matrix(CurrentPosX, (CurrentPosY + 1)) = X
    Change
  End If
End Sub

Private Sub CmdDown_Click()
  If CurrentPosY <> 1 Then
    X = Matrix(CurrentPosX, CurrentPosY)
    Matrix(CurrentPosX, CurrentPosY) = Matrix(CurrentPosX, (CurrentPosY - 1))
    Matrix(CurrentPosX, (CurrentPosY - 1)) = X
    Change
  End If
End Sub

Private Sub CmdLeft_Click()
  If CurrentPosX <> No Then
    X = Matrix(CurrentPosX, CurrentPosY)
    Matrix(CurrentPosX, CurrentPosY) = Matrix((CurrentPosX + 1), CurrentPosY)
    Matrix((CurrentPosX + 1), CurrentPosY) = X
    Change
  End If
End Sub

Private Sub CmdRight_Click()
  If CurrentPosX <> 1 Then
    X = Matrix(CurrentPosX, CurrentPosY)
    Matrix(CurrentPosX, CurrentPosY) = Matrix((CurrentPosX - 1), CurrentPosY)
    Matrix((CurrentPosX - 1), CurrentPosY) = X
    Change
  End If
End Sub

Private Sub Command1_Click()
Dim Xc As Boolean
Xc = True
Do While Xc = True
  Loadmat
  For I = 1 To No * No
    If Mat(I) = 0 Then
      Xc = False
    End If
  Next I
  If Xc = False Then Xc = True Else Exit Do
Loop

X = p.Width / No
Y = p.Height / No

xz = 0
For I = 0 To No - 1
  For J = 0 To No - 1
    xz = xz + 1
    FcXcY Mat(xz)
    Matrix(I + 1, J + 1) = Mat(xz)
    If Mat(xz) = No * No Then
      s.PaintPicture fs.Picture, I * X, J * Y, X, Y, 0 * X, 0 * Y, X, Y, vbSrcCopy
      CurrentPosX = I + 1
      CurrentPosY = J + 1
    Else
      s.PaintPicture p.Picture, I * X, J * Y, X, Y, Cx * X, Cy * Y, X, Y, vbSrcCopy
    End If
  Next J
Next I
Nt = 0
Ts = 0
Frame1.Enabled = True
Timer1.Enabled = True
Command3.Enabled = True
s.SetFocus
End Sub


Sub FcXcY(AD As Integer)
ss = 0
For IC = 1 To No
  For jC = 1 To No
    ss = ss + 1
    If ss = AD Then
      Cx = IC - 1
      Cy = jC - 1
      Exit Sub
    End If
  Next jC
Next IC
End Sub

Sub Change()
Nt = Nt + 1
'sndPlaySound App.Path + "\a.wav", 0
X = p.Width / No
Y = p.Height / No
xz = 0
s.Cls
For I = 0 To No - 1
For J = 0 To No - 1
  xz = xz + 1
  Mat(xz) = Matrix(I + 1, J + 1)
  FcXcY Mat(xz)
  If Mat(xz) = No * No Then
    s.PaintPicture fs.Picture, I * X, J * Y, X, Y, 0 * X, 0 * Y, X, Y, vbSrcCopy
    CurrentPosX = I + 1
    CurrentPosY = J + 1
  Else
    s.PaintPicture p.Picture, I * X, J * Y, X, Y, Cx * X, Cy * Y, X, Y, vbSrcCopy
  End If
Next J
Next I

Tries.Caption = Nt
CheckWIn
End Sub


Sub CheckWIn()
Dim Win As Boolean
X = 0

Win = False
For I = 1 To No
For J = 1 To No
  X = X + 1
  If Mat(X) = X Then
    Win = True
  Else
    Win = False
    GoTo EndX
  End If
Next J
Next I
EndX:
If Win = True Then
  MsgBox "You have won the Game....." + vbCrLf + vbCrLf + "Keep Smiling .... :-)"
  Timer1.Enabled = False
  Command3.Enabled = False
  Frame1.Enabled = False
End If
s.SetFocus
End Sub

Private Sub Command2_Click()
On Error GoTo esub
cdq.Filter = "Bitmap Files (*.bmp)|*.bmp|Gif Files (*.gif)|*.gif|Jpeg Files (*.jpg)|*.jpg|Pictures (*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif"
cdq.FilterIndex = 4
cdq.ShowOpen
Check.Picture = LoadPicture(cdq.FileName)
If Check.Height > 4000 And Check.Width > 4000 Then
  p.Picture = LoadPicture(cdq.FileName)
Else
  MsgBox "The Picture should be 4000 X 4000 " + vbCrLf + vbCrLf + "Selected picture size was Height :" + Str(Check.Height) + " Width" + Str(Check.Width), vbInformation, "Picture Rejected"
End If
s.SetFocus
Exit Sub
esub:
End Sub

Private Sub Command3_Click()
If Command3.Caption = "&Pause" Then
  Timer1.Enabled = False
  Frame1.Enabled = False
  Command3.Caption = "&Resume"
Else
  Timer1.Enabled = True
  Frame1.Enabled = True
  Command3.Caption = "&Pause"
End If
s.SetFocus
End Sub

Private Sub Command4_Click()
  End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Timer1.Enabled = True Then
  If KeyCode = 39 Then
    CmdRight_Click
    Lastkey.Caption = "Right"
  End If
  If KeyCode = 37 Then
    CmdLeft_Click
    Lastkey.Caption = "Left"
  End If
  If KeyCode = 38 Then
    CmdUp_Click
    Lastkey.Caption = "Up"
  End If
  If KeyCode = 40 Then
    CmdDown_Click
    Lastkey.Caption = "Down"
  End If
End If
End Sub

Private Sub Form_Load()
NoX.AddItem "3"
NoX.AddItem "4"
NoX.AddItem "5"
NoX.AddItem "6"
NoX.AddItem "7"
No = 4
End Sub

Private Sub NoX_Change()
No = Val(NoX.Text)
Command1_Click
End Sub

Private Sub NoX_Click()
No = Val(NoX.Text)
Command1_Click
End Sub

Private Sub Timer1_Timer()
Ts = Ts + 1
Timet.Caption = CalTime(Ts) + " Minutes"
End Sub

Function CalTime(Timerx As Long) As String
Dim X As Long
Dim Y As Long
X = Int(Timerx / 60)
Y = Int(Timerx - (X * 60))
CalTime = LTrim(RTrim(Str(X))) + ":" + LTrim(RTrim(Str(Y)))
End Function

