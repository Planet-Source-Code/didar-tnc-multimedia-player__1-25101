VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   ClientHeight    =   900
   ClientLeft      =   1920
   ClientTop       =   5700
   ClientWidth     =   6015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HasDC           =   0   'False
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form7.frx":08CA
   ScaleHeight     =   900
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   0
      Top             =   840
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   18
      ToolTipText     =   "Capture Still Image"
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   960
      TabIndex        =   17
      ToolTipText     =   "Video CD"
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "::"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   3000
      TabIndex        =   16
      Top             =   50
      Width           =   195
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5130
      TabIndex        =   14
      ToolTipText     =   "Option"
      Top             =   40
      Width           =   165
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   4880
      TabIndex        =   13
      Top             =   120
      Width           =   150
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      ToolTipText     =   "Help ?"
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   4680
      TabIndex        =   11
      ToolTipText     =   "Auto Video  Searching Mode.."
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   150
      Left            =   4560
      TabIndex        =   10
      ToolTipText     =   "Full Screen"
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   5520
      TabIndex        =   8
      ToolTipText     =   "Information About This Player."
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "E x i t"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   5400
      TabIndex        =   7
      ToolTipText     =   "Press  Exit + ALT To Quit."
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Remote"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Remote Control"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   4800
      TabIndex        =   5
      ToolTipText     =   "Video Control Menu"
      Top             =   555
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      ToolTipText     =   "Next Track"
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   3
      ToolTipText     =   "Prev Track"
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      ToolTipText     =   "Stop"
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "Pause"
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "Play"
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Dim dolby As Integer
Dim play As Integer


Private Sub Form_Click()
On Error Resume Next
Form1.Command1.SetFocus
End Sub

Private Sub Form_DblClick()
On Error Resume Next

Form1.Command1.SetFocus
End Sub

Private Sub form_load()
x = 0
dolby = 0
play = 0
Load Form9
Form9.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Form1.Command1.SetFocus
End Sub

Private Sub Label1_Click()
On Error Resume Next
If play = 1 Then
Unload Form1
Form1.mov.FileName = List1.List(Text2.Text)
Form1.mov.Command = "Play"
Form1.mov.Visible = True
Form1.Command1.SetFocus
Form9.Hide
Form9.Timer1.Enabled = True
Form9.Timer2.Enabled = True
play = 0
Exit Sub
End If

Label4.Visible = True
Label5.Visible = True
Unload Form1
Form1.mov.Visible = True
Form9.Dir1.Path = Form9.Drive1.Drive
Form9.File1.Path = Form9.Dir1.Path
Form1.mov.FileName = Form9.List1.List(x)
Form1.mov.Command = "Play"
Form9.Text2.Text = x
Form1.Command1.SetFocus

If Form9.Combo1.Text = "*.mp3" Then
Form1.Height = 540
Form1.Width = 5060
Form1.Show
Form1.WindowState = 0
Form1.mov.Top = 0
Form1.mov.Left = 0
Form1.mov.Playbar = True
Form1.Top = Form7.Top - 540
Form1.Left = Form7.Left + 900
Form7.Label4.Enabled = True
Form7.Label5.Enabled = True
Form9.Timer1.Enabled = True
Form9.Timer2.Enabled = True
Else




Form1.Show
Form1.WindowState = 2
Form1.mov.Top = 0
Form1.mov.Left = 0
Form1.mov.Height = Form1.Height
Form1.mov.Width = Form1.Width
Label4.Enabled = False
Label5.Enabled = False




Form9.Hide
Form9.Timer1.Enabled = True
Form9.Timer2.Enabled = True
End If
Form1.Command1.SetFocus
End Sub

Private Sub Label10_Click()
On Error Resume Next
Form1.Command1.SetFocus
End Sub

Private Sub Label11_Click()
On Error Resume Next
Form1.WindowState = 2
Form1.mov.Top = 0
Form1.mov.Left = 0
Form1.mov.Height = Form1.Height
Form1.mov.Width = Form1.Width
Form1.Command1.SetFocus
End Sub

Private Sub Label12_Click()
Form9.Hide
Form10.Show
End Sub

Private Sub Label13_Click()
On Error Resume Next

MsgBox "                      <<<<<<<<<Welcome>>>>>>>>>>>" & Chr(10) & Chr(13) & " " & Chr(10) & Chr(13) & " " & Chr(10) & Chr(13) & "                 TNC Driver. A Real Video Imagination.  " & Chr(10) & Chr(13) & " " & Chr(10) & Chr(13) & "Common Function:" & Chr(10) & Chr(13) & " " & Chr(10) & Chr(13) & "Click 'P.E' To Enable The Playlist Advance." & Chr(10) & Chr(13) & "Click 'P.D' To Disable The Playlist Advance." & Chr(10) & Chr(13) & "To Play In MultiSelect Mode--Select The Files And Press'Play*X'." & Chr(10) & Chr(13) & "To Play In Normal Mode or To Play All Songs In The List Press--'Play' Button. " & Chr(10) & Chr(13) & "To Play In Auto Search Video Method---Click 'A' And Select The Default Video MB And Filetype." _
& " " & Chr(10) & Chr(13) & "" & Chr(10) & Chr(13) & "HOTKEY's " & Chr(10) & Chr(13) & "" & Chr(10) & Chr(13) & "'Z'==Prev Track,   'X'==Stop/Play,  'C'==Pause/Play,   'V'==Playbar Enable/disable,   'B'==Next Track,   'F'==Full Screen,   'Esc'==Close Full Screen,   'Space'==Forward,   'N'==Rewind,   '>'==Volume Increase,   '<'==Volume Decrease, 'S'==Capture Still Image, 'Q'==Quit.  " & Chr(10) & Chr(13) & "" & Chr(10) & Chr(13) & " Please See The Read Me File For Details.", 32, "Help"
Form1.Command1.SetFocus
End Sub

Private Sub Label14_Click()
On Error Resume Next

If Label14.Caption = "L" Then
Form9.Width = 6720
Label14.Caption = "N"
Else
If Label14.Caption = "N" Then
Form9.Width = 2325
Label14.Caption = "L"
Else
Label1.Caption = "L"
End If
End If
Form1.Command1.SetFocus
End Sub

Private Sub Label15_Click()
frmINI.Show
End Sub

Private Sub Label18_Click()
Dim p, i As Integer
Dim q As String

On Error GoTo error
Form9.Timer1.Enabled = False
Form9.Timer2.Enabled = False
Form9.Timer3.Enabled = False
Form7.Label5.Enabled = False
Form4.Label4.Enabled = False
Unload Form1
Form9.List1.Clear
Form9.Drive1.Drive = "c:\"
Form1.mov.Visible = True
Form9.Combo1.Text = "*.dat"
Form9.File1.Pattern = Form9.Combo1.Text
p = Form9.Drive1.ListCount
q = p + 65
Form9.Text1.Text = Chr(q)
Form9.Drive1.Drive = Form9.Text1.Text
Form9.Dir1.Path = Form9.Drive1.Drive

Form9.List1.Clear
Form9.File1.Path = Form9.Dir1.Path
For i = 0 To Form9.File1.ListCount - 1
 Form9.List1.AddItem Form9.File1.List(i)
Next


Form9.Dir1.Path = "\mpegav"
Form1.mov.FileName = Form9.List1.List(0)
Form1.mov.Command = "Play"
Form9.Text2.Text = 0
Form1.Show
Form1.WindowState = 2
Form1.mov.Top = 0
Form1.mov.Left = 0
Form1.mov.Height = Form1.Height
Form1.mov.Width = Form1.Width
Form7.Label4.Enabled = False
Form7.Label5.Enabled = False
Form9.Hide
Form9.Timer1.Enabled = True
Form9.Timer2.Enabled = True
Exit Sub
error:
MsgBox "No Video Cd Found..", 16, "Invalid VCD"

End Sub

Private Sub Label19_Click()
On Error Resume Next
Form1.mov.Command = "Pause"
Form1.mov.Command = "copy"
SavePicture Clipboard.GetData, App.Path & "\tnc" & cap & ".bmp"
cap = cap + 1
Form1.mov.Command = "Play"
Form1.Command1.SetFocus
End Sub

Private Sub Label2_Click()
On Error Resume Next
If dolby = 1 Then
Form1.mov.Command = "play"
dolby = 0
Form1.Command1.SetFocus
Else
Form1.mov.Command = "Pause"
dolby = 1
Form1.Command1.SetFocus
End If
End Sub

Private Sub Label3_Click()
On Error Resume Next
Form1.mov.Command = "Stop"
play = 1
Form9.Timer1.Enabled = False
Form9.Timer2.Enabled = False
Form1.mov.Visible = False
Unload Form1
End Sub

Private Sub Label4_Click()
On Error GoTo err





Form9.Text2.Text = Form9.Text2.Text
Unload Form1
Form9.Text2.Text = Form9.Text2.Text - 1


If Form9.Text2.Text = -1 Then
Unload Form1
Form9.Text2.Text = -1
Else


Form9.Dir1.Path = Form9.Drive1.Drive
Form9.File1.Path = Form9.Dir1.Path
Form1.mov.FileName = Form9.List1.List(Form9.Text2.Text)
Form1.mov.Command = "Play"



If Form9.Combo1.Text = "*.mp3" Then
Form1.Height = 540
Form1.Width = 5060
Form1.Show
Form1.WindowState = 0
Form1.mov.Top = 0
Form1.mov.Left = 0
Form1.mov.Playbar = True
Form1.Top = Form7.Top - 540
Form1.Left = Form7.Left + 900
Form7.Label4.Enabled = True
Form7.Label5.Enabled = True
Else







Form1.Show
Form1.WindowState = 2
Form1.mov.Top = 0
Form1.mov.Left = 0
Form1.mov.Height = Form1.Height
Form1.mov.Width = Form1.Width
Label4.Enabled = False
Label5.Enabled = False



Form9.Hide
Form9.Timer1.Enabled = True
Form9.Timer2.Enabled = True

Form9.Timer3.Enabled = False    'Updated


Exit Sub
err:
Timer1.Enabled = False
Timer1.Enabled = False
Form1.Command1.SetFocus
End If
End If
End Sub

Private Sub Label5_Click()
On Error GoTo err


Form9.Text2.Text = Form9.Text2.Text
Unload Form1
Form9.Text2.Text = Form9.Text2.Text + 1




If Form9.Text2.Text = Form9.List1.ListCount Then
Unload Form1
Form9.Text2.Text = -1
Else




Form9.Dir1.Path = Form9.Drive1.Drive
Form9.File1.Path = Form9.Dir1.Path
Form1.mov.FileName = Form9.List1.List(Form9.Text2.Text)
Form1.mov.Command = "Play"


If Form9.Combo1.Text = "*.mp3" Then
Form1.Height = 540
Form1.Width = 5060
Form1.Show
Form1.WindowState = 0
Form1.Top = Form7.Top - 540
Form1.Left = Form7.Left + 900
Form7.Label4.Enabled = True
Form7.Label5.Enabled = True
Form1.mov.Top = 0
Form1.mov.Left = 0
Form1.mov.Playbar = True
Else







Form1.Show
Form1.WindowState = 2
Form1.mov.Top = 0
Form1.mov.Left = 0
Form1.mov.Height = Form1.Height
Form1.mov.Width = Form1.Width
Label4.Enabled = False
Label5.Enabled = False



Form9.Hide
Form9.Timer1.Enabled = True
Form9.Timer2.Enabled = True
Form9.Timer3.Enabled = False    'Updated
Exit Sub
err:
Form1.Command1.SetFocus
Timer1.Enabled = False
Timer1.Enabled = False
End If
End If

End Sub

Private Sub Label6_Click()
On Error Resume Next

Form9.Show
Form1.Command1.SetFocus
End Sub

Private Sub Label7_Click()
On Error Resume Next
Form5.Show
Form1.Command1.SetFocus
End Sub

Private Sub Label8_Click()
End
End Sub

Private Sub Label9_Click()
Form4.Show
End Sub

Private Sub Timer1_Timer()
Dim h As Integer
On Error Resume Next
h = ((Form1.mov.Position) / 24) - (Label16.Caption * 60)
Label10.Caption = h
If h > 59 Then
Label16.Caption = Label16.Caption + 1
End If
End Sub
