VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5520
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   2190
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HasDC           =   0   'False
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form9.frx":08CA
   ScaleHeight     =   5520
   ScaleWidth      =   2190
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   3240
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   2280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Auto Search"
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
      Left            =   80
      Picture         =   "Form9.frx":50B6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5200
      Width           =   2040
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000E&
      Caption         =   "Exit"
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
      Left            =   1130
      Picture         =   "Form9.frx":67EE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4940
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   3240
      Top             =   960
   End
   Begin VB.ListBox List1 
      Height          =   5520
      Left            =   2280
      TabIndex        =   6
      Top             =   0
      Width           =   4335
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00808080&
      Height          =   315
      Left            =   70
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000E&
      Caption         =   "Play*X"
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
      Left            =   70
      Picture         =   "Form9.frx":7F26
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "To Play In MultiSelect Mode.."
      Top             =   4940
      Width           =   1050
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      ItemData        =   "Form9.frx":965E
      Left            =   70
      List            =   "Form9.frx":9677
      TabIndex        =   2
      Text            =   "*.dat"
      Top             =   240
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   3000
      Top             =   1440
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00808080&
      Height          =   2040
      Left            =   70
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   2880
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00808080&
      Height          =   1890
      Left            =   70
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Open Files"
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
      Left            =   600
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Dim x, y As Integer


Private Sub Combo1_Change()
On Error Resume Next
File1.Pattern = Combo1.Text
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
Drive1.Refresh
Dir1.Refresh
File1.Refresh
End Sub






Private Sub Combo1_Click()
On Error Resume Next
File1.Pattern = Combo1.Text
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
Drive1.Refresh
Dir1.Refresh
File1.Refresh
dir1_Change
End Sub

Private Sub Command1_Click()
Me.Hide
Form10.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next
Timer1.Enabled = False
Timer1.Enabled = False
List1.Clear
Text2.Text = 0
Unload Form1
Form1.mov.Command = "Stop"
For i = 0 To File1.ListCount - 1
If File1.Selected(i) Then
If Len(Dir1.Path) < 4 Then
List1.AddItem Dir1.Path & File1.List(i)
Else
List1.AddItem Dir1.Path & "\" & File1.List(i)
End If
End If
Next


'If Not Multiselect
If List1.ListCount = 0 Then
For k = 0 To File1.ListCount - 1
If Len(Dir1.Path) < 4 Then
List1.AddItem Dir1.Path & File1.List(k)
Else
List1.AddItem Dir1.Path & "\" & File1.List(k)
End If
Next
File1.Refresh
End If



File1.Path = Dir1.Path
Form1.mov.FileName = List1.List(Text2.Text)
Text2.Text = x
Form1.mov.Command = "Play"
Form1.mov.Visible = True





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
Form7.Label4.Enabled = False
Form7.Label5.Enabled = False
Form9.Hide
Timer1.Enabled = True
Timer2.Enabled = True
End If
Form1.Command1.SetFocus
End Sub

Private Sub Command3_Click()
On Error Resume Next
Me.Hide
Form1.Command1.SetFocus
End Sub


Private Sub dir1_Change()
Dim i As Integer
On Error GoTo err
List1.Clear
Dir1.Path = Drive1.Drive
File1.Pattern = Combo1.Text


If Combo1.Text = "*.*" Then

File1.Pattern = "*.dat"
File1.Path = Dir1.Path
For j = 0 To File1.ListCount - 1
If Len(Dir1.Path) < 4 Then
List1.AddItem Dir1.Path & File1.List(j)
Else
List1.AddItem Dir1.Path & "\" & File1.List(j)
End If
Next



File1.Pattern = "*.mpg"
File1.Path = Dir1.Path
For k = 0 To File1.ListCount - 1
If Len(Dir1.Path) < 4 Then
List1.AddItem Dir1.Path & File1.List(k)
Else
List1.AddItem Dir1.Path & "\" & File1.List(k)
End If
Next



File1.Pattern = "*.avi"
File1.Path = Dir1.Path
For l = 0 To File1.ListCount - 1
If Len(Dir1.Path) < 4 Then
List1.AddItem Dir1.Path & File1.List(l)
Else
List1.AddItem Dir1.Path & "\" & File1.List(l)
End If
Next
Form9.Width = 6500
Form7.Label14.Caption = "N"
Exit Sub
End If



















File1.Path = Dir1.Path
For i = 0 To File1.ListCount - 1
If Len(Dir1.Path) < 4 Then  'Update
List1.AddItem Dir1.Path & File1.List(i) 'Update
Else                                'Update
 List1.AddItem Dir1.Path & "\" & File1.List(i) 'Update
 End If 'Update
Next

File1.Refresh
Exit Sub
err:
MsgBox "Invalid File Path , Check The Drive ", 16, "Reading Error"

End Sub

Private Sub drive1_Change()
Dim i As Integer
On Error GoTo err
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
List1.Clear
File1.Pattern = Combo1.Text

If Combo1.Text = "*.*" Then
File1.Pattern = "*.dat"
File1.Path = Dir1.Path
For i = 0 To File1.ListCount - 1
 List1.AddItem (Dir1.Path & File1.List(i))
Next

File1.Pattern = "*.mpg"
File1.Path = Dir1.Path
For i = 0 To File1.ListCount - 1
 List1.AddItem (Dir1.Path & File1.List(i))
Next

File1.Pattern = "*.avi"
File1.Path = Dir1.Path
For i = 0 To File1.ListCount - 1
 List1.AddItem (Dir1.Path & File1.List(i))
Next
Form9.Width = 6500
Form7.Label14.Caption = "N"
Exit Sub
End If





File1.Path = Dir1.Path
For i = 0 To File1.ListCount - 1
 List1.AddItem (Dir1.Path & File1.List(i))
Next
Exit Sub
err:
MsgBox "Disk Is Not Ready Or Failed to Reading", 16, "Reading Error"
Dir1.Path = "c:\"
End Sub





Private Sub Form_Click()
On Error Resume Next
Form1.Command1.SetFocus
End Sub

Private Sub Form_DblClick()
On Error Resume Next
Form1.Command1.SetFocus
End Sub

Private Sub form_load()
x = 1
pp = 5
cap = 0
Text2.Text = x
File1.Pattern = Combo1.Text
File1.Hidden = True
File1.ReadOnly = True
Dir1.Path = "c:\"
End Sub

Private Sub List1_Click()
On Error Resume Next
Unload Form1
List2.Clear
For i = 0 To List1.ListCount - 1
If List1.Selected(i) Then
List2.AddItem List1.List(i)
End If
Next
Load Form1
Form1.mov.FileName = List2.List(0)
Form1.mov.Command = "Play"
Text2.Text = -1
Form1.mov.Visible = True
Form1.Show
Form1.WindowState = 2
Form1.mov.Top = 0
Form1.mov.Left = 0
Form1.mov.Height = Form1.Height
Form1.mov.Width = Form1.Width
Form9.Hide
Form1.Command1.SetFocus
End Sub

Private Sub List1_DblClick()
On Error Resume Next
Me.Width = 1740
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
Form9.Width = 2325
Form7.Label14.Caption = "L"
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next

If List1.ListCount = 0 Then
Form1.WindowState = 0
Form1.mov.Left = 430
Form1.mov.Top = 360
Form1.mov.Zoom = 100
Unload Form1
Timer3.Enabled = False
Timer2.Enabled = False
Timer1.Enabled = False
Exit Sub
End If




If y = List1.ListCount + 1 Then
Unload Form1

y = 0
Form9.Show
Timer2.Enabled = False
Timer1.Enabled = False
Else

If Form1.mov.Position = Form1.mov.Length Then
Text2.Text = Text2.Text + 1
y = Text2.Text

Form9.Hide
Unload Form1
Form1.mov.FileName = List1.List(Text2.Text)
Form1.mov.Command = "Play"
Form7.Label10.Caption = 0
Form7.Label16.Caption = 0
Form1.mov.Visible = True
Form1.mov.Volume = pp * 33
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
Else







Form1.Show
Form1.WindowState = 2
Form1.mov.Top = 0
Form1.mov.Left = 0
Form1.mov.Height = Form1.Height
Form1.mov.Width = Form1.Width
Form7.Label4.Enabled = False
Form7.Label5.Enabled = False


Form9.Hide
End If
End If
End If
If Form1.WindowState = 2 Then
Form1.Timer1.Enabled = True
Else
If Form1.WindowState = 0 Then
Form1.Timer1.Enabled = False
End If
End If
End Sub

Private Sub Timer2_Timer()
On Error GoTo error
Form1.mov.Position = 0
error:
On Error Resume Next
Form1.mov.FileName = List1.List(Text2.Text)
Form1.mov.Command = "Play"
Form7.Label10.Caption = 0
Form7.Label16.Caption = 0
Form1.mov.Volume = pp * 33
Form1.mov.Visible = True


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
Timer1.Enabled = True
Timer2.Enabled = False
Else







Form1.Show
Form1.WindowState = 2
Form1.mov.Top = 0
Form1.mov.Left = 0
Form1.mov.Height = Form1.Height
Form1.mov.Width = Form1.Width
Form7.Label4.Enabled = False
Form7.Label5.Enabled = False





Form9.Hide
Timer1.Enabled = True
Timer2.Enabled = False
End If
End Sub


Private Sub Timer3_Timer()
On Error Resume Next
If y = List1.ListCount + 1 Then
Unload Form1
y = 0
Text2.Text = y

Form1.mov.FileName = List1.List(Text2.Text)
Form1.mov.Command = "Play"
Form7.Label10.Caption = 0
Form7.Label16.Caption = 0
Form1.mov.Volume = pp * 33
Form1.mov.Visible = True




Form1.Show
Form1.WindowState = 2
Form1.mov.Top = 0
Form1.mov.Left = 0
Form1.mov.Height = Form1.Height
Form1.mov.Width = Form1.Width
Form7.Label1.Enabled = False
Form7.Label2.Enabled = False
Form7.Label3.Enabled = False
Form7.Label4.Enabled = False
Form7.Label5.Enabled = False

Form9.Hide
Form1.Command1.SetFocus
Else

If Form1.mov.Position = Form1.mov.Length Then
Text2.Text = Text2.Text + 1
y = Text2.Text

Form9.Hide
Unload Form1
Form1.mov.FileName = List1.List(Text2.Text)
Form1.mov.Command = "Play"
Form1.mov.Visible = True




Form1.Show
Form1.WindowState = 2
Form1.mov.Top = 0
Form1.mov.Left = 0
Form1.mov.Height = Form1.Height
Form1.mov.Width = Form1.Width
Form7.Label1.Enabled = False
Form7.Label2.Enabled = False
Form7.Label3.Enabled = False
Form7.Label4.Enabled = False
Form7.Label5.Enabled = False
Form9.Hide
Form1.Command1.SetFocus
End If
End If
End Sub

