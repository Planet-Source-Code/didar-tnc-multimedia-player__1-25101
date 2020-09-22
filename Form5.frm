VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   0  'None
   Caption         =   "Remote "
   ClientHeight    =   3015
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   1350
   HasDC           =   0   'False
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form5.frx":08CA
   ScaleHeight     =   3015
   ScaleWidth      =   1350
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command9 
      Caption         =   "Drv"
      Height          =   255
      Left            =   120
      Picture         =   "Form5.frx":DFAE
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      Caption         =   "D.A."
      Height          =   255
      Left            =   720
      Picture         =   "Form5.frx":F4F7
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Disable Auto Repeat"
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      Caption         =   "A. R"
      Height          =   255
      Left            =   120
      Picture         =   "Form5.frx":10A40
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Auto Repeat Mode"
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "P.E"
      Height          =   255
      Left            =   720
      Picture         =   "Form5.frx":11F89
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Enable PlayList Advance"
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "P.D"
      Height          =   255
      Left            =   120
      Picture         =   "Form5.frx":134D2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Disable PlayList Advance"
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "List"
      Height          =   255
      Left            =   720
      Picture         =   "Form5.frx":14A1B
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Control Gear"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Exit"
      Height          =   255
      Left            =   720
      Picture         =   "Form5.frx":15F64
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Exit To Windows System"
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Refre"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   720
      Picture         =   "Form5.frx":174AD
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "To Refresh The PlayList File."
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Vcd"
      Height          =   255
      Left            =   120
      Picture         =   "Form5.frx":189F6
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Video Cd Mode"
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Vol"
      Height          =   255
      Left            =   120
      Picture         =   "Form5.frx":19F3F
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Volume"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Auto"
      Height          =   255
      Left            =   720
      Picture         =   "Form5.frx":1B488
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Auto Searching Video.."
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Full "
      Height          =   255
      Left            =   120
      Picture         =   "Form5.frx":1C9D1
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Full Screen "
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
On Error Resume Next
Form1.WindowState = 2
Form1.mov.Top = 0
Form1.mov.Left = 0
Form1.mov.Height = Form1.Height
Form1.mov.Width = Form1.Width
Form7.Label4.Enabled = False
Form7.Label5.Enabled = False
Form1.Command1.SetFocus
End Sub

Private Sub Command10_Click()
On Error Resume Next
Form1.Command1.SetFocus
Unload Me
Form9.Show
End Sub


Private Sub Command12_Click()
On Error Resume Next
Form7.Label1.Enabled = False
Form7.Label2.Enabled = False
Form7.Label3.Enabled = False
Form7.Label4.Enabled = False
Form7.Label5.Enabled = False

Form9.Timer1.Enabled = False
Form9.Timer3.Enabled = True
Form9.Timer2.Enabled = False
Form1.Command1.SetFocus
End Sub

Private Sub Command13_Click()
On Error Resume Next
Form9.Timer3.Enabled = False
Form7.Label1.Enabled = True
Form7.Label2.Enabled = True
Form7.Label3.Enabled = True
Form7.Label4.Enabled = True
Form7.Label5.Enabled = True
Form9.Timer1.Enabled = True
Form1.Command1.SetFocus

End Sub

Private Sub Command2_Click()
Unload Me
Load Form10
Form10.Show
End Sub

Private Sub Command3_Click()
i = Shell("C:\windows\sndvol32", vbNormalFocus)
End Sub

Private Sub Command4_Click()
On Error Resume Next
Form9.Timer1.Interval = 0
Form1.Command1.SetFocus
End Sub

Private Sub Command5_Click()
On Error Resume Next
Form9.Timer1.Interval = 5
Form1.Command1.SetFocus
End Sub

Private Sub Command6_Click()
Dim p, i As Integer
Dim q As String

On Error GoTo error
Command7_Click

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

Private Sub Command7_Click()
On Error Resume Next
Form9.Timer1.Enabled = False
Form9.Timer2.Enabled = False
Form9.Timer3.Enabled = False
Form7.Label5.Enabled = False
Form4.Label4.Enabled = False
Unload Form1
Form9.List1.Clear
Form9.Drive1.Drive = "c:\"
End Sub

Private Sub Command8_Click()
End
End Sub


Private Sub Command9_Click()
frmINI.Show
End Sub

Private Sub Form_Click()
On Error Resume Next
Form1.Command1.SetFocus
End Sub



Private Sub Label2_Click()
Unload Form5
End Sub

Private Sub Label3_Click()
Me.Hide
End Sub

