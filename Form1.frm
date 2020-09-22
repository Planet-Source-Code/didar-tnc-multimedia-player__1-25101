VERSION 5.00
Object = "{288F1520-FAC4-11CE-B16F-00AA0060D93D}#1.0#0"; "MCIWNDX.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   ClientHeight    =   5550
   ClientLeft      =   1860
   ClientTop       =   60
   ClientWidth     =   6165
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   DrawStyle       =   1  'Dash
   HasDC           =   0   'False
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   5550
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5040
      Picture         =   "Form1.frx":A236
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5520
      Picture         =   "Form1.frx":A8A9
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   5520
      TabIndex        =   0
      Top             =   360
      Width           =   75
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Play Bar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      Picture         =   "Form1.frx":AF1C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "TNC2 Digital Driver Control"
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Timer Timer2 
      Left            =   1320
      Top             =   6240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   -120
      Top             =   -120
   End
   Begin MCIWndX.MCIWnd mov 
      Height          =   4335
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   7646
      _StockProps     =   96
      Menu            =   0   'False
      Playbar         =   0   'False
      Record          =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   5775
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TNC2 Digital Driver Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      ToolTipText     =   "TNC2 Driver Control"
      Top             =   5040
      Visible         =   0   'False
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Dim y As Long
Dim f As Integer
Dim g As Integer


        







Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 32 Then
Form1.mov.Volume = pp * 33
mov.Position = mov.Position + 50
mov.Command = "play"
End If

If KeyCode = 78 Then
Form1.mov.Volume = pp * 33
mov.Position = mov.Position - 70
mov.Command = "play"
End If



If KeyCode = 188 Then
Command4_Click
Command1.SetFocus
End If

If KeyCode = 190 Then
Command3_Click
Command1.SetFocus
End If




If KeyCode = 86 Then
Form1.WindowState = 0
Form1.mov.Left = 430
Form1.mov.Top = 360
Form1.mov.Zoom = 100
Command2_Click
Form1.WindowState = 2
Form1.mov.Top = 0
Form1.mov.Left = 0
Form1.mov.Height = Form1.Height
Form1.mov.Width = Form1.Width
End If



If KeyCode = 67 Then
If f = 1 Then
Form1.mov.Volume = pp * 33
mov.Command = "play"
f = 0
Else
Form1.mov.Volume = pp * 33
mov.Command = "pause"
f = 1
End If
End If

If KeyCode = 83 Then
Form1.mov.Volume = pp * 33
mov.Command = "pause"
mov.Command = "copy"
SavePicture Clipboard.GetData, App.Path & "\tnc" & cap & ".bmp"
cap = cap + 1
mov.Command = "play"
End If



If KeyCode = 88 Then
If g = 1 Then
Form1.mov.Volume = pp * 33
mov.Command = "play"
g = 0
Else
Form1.mov.Volume = pp * 33
mov.Command = "stop"
mov.Position = 0
g = 1
End If
End If




If KeyCode = 66 Then
mov.Position = mov.Length
z = lSetVolume(pp, pp, 0)
End If
If KeyCode = 90 Then
Form9.Text2.Text = Form9.Text2.Text - 2
mov.Position = mov.Length
z = lSetVolume(pp, pp, 0)
End If


If KeyCode = 70 Then
Form1.WindowState = 2
Form1.mov.Top = 0
Form1.mov.Left = 0
Form1.mov.Height = Form1.Height
Form1.mov.Width = Form1.Width
End If

If KeyCode = 81 Then
End
End If

If KeyCode = 27 Then
Form5.Show
Form7.Show
Form1.WindowState = 0
Form1.Top = Form7.Top - (Form1.Height + 20)
Form1.mov.Left = 430
Form1.mov.Top = 360
Form1.mov.Zoom = 100
Form1.Command2.Visible = True
Form7.Label4.Enabled = True
Form7.Label5.Enabled = True
Command1.SetFocus
End If

End Sub


Private Sub Command2_Click()
If Command2.Caption = "Hide" Then
mov.Playbar = False
Label2.Visible = False
Command2.Caption = "Play Bar"
Command1.SetFocus
Else
If Command2.Caption = "Play Bar" Then
mov.Playbar = True
Label2.Visible = True
Command1.SetFocus
Command2.Caption = "Hide"
End If
End If
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command5.Height = 480
End Sub



Private Sub Command3_Click()
On Error Resume Next
pp = pp + 2
If pp > 29 Then
pp = 30
End If
z = lSetVolume(pp, pp, 0)
Command1.SetFocus
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command1.SetFocus
End Sub

Private Sub Command4_Click()
On Error Resume Next
pp = pp - 2
If pp < 1 Then
pp = 0
End If
z = lSetVolume(pp, pp, 0)
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command1.SetFocus
End Sub

Private Sub Form_Click()
Form5.Show
Form7.Show
Command1.SetFocus
End Sub

Private Sub form_load()
On Error Resume Next
Unload Form2
Command2.Visible = False
x = 0
f = 0
g = 0
mov.Volume = pp * 34
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Command1.SetFocus
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Command1.SetFocus
End Sub

Private Sub mov_DblClick()
Form5.Show
Form7.Show
Form1.WindowState = 0
Form1.mov.Left = 430
Form1.mov.Top = 360
Form1.mov.Zoom = 100
Form1.Command2.Visible = True
Form7.Label4.Enabled = True
Form7.Label5.Enabled = True
Command1.SetFocus
End Sub



Private Sub mov_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Command1.SetFocus
End Sub

Private Sub Timer1_Timer()
Command1.SetFocus
End Sub

