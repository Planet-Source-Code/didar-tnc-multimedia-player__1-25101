VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000012&
   ClientHeight    =   7080
   ClientLeft      =   -120
   ClientTop       =   60
   ClientWidth     =   9480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   7080
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer7 
      Interval        =   180
      Left            =   1080
      Top             =   3240
   End
   Begin VB.Timer Timer6 
      Interval        =   200
      Left            =   7320
      Top             =   5640
   End
   Begin VB.Timer Timer5 
      Interval        =   200
      Left            =   7920
      Top             =   2760
   End
   Begin VB.Timer Timer4 
      Interval        =   400
      Left            =   4800
      Top             =   3960
   End
   Begin VB.Timer Timer3 
      Interval        =   300
      Left            =   3840
      Top             =   3120
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   5040
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Left            =   2160
      Top             =   5520
   End
   Begin VB.Image Shape7 
      Height          =   720
      Left            =   960
      Picture         =   "Form3.frx":08CA
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   720
   End
   Begin VB.Image Shape6 
      Height          =   720
      Left            =   7200
      Picture         =   "Form3.frx":1194
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   720
   End
   Begin VB.Image shape5 
      Height          =   720
      Left            =   7680
      Picture         =   "Form3.frx":1A5E
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   720
   End
   Begin VB.Image SHAPE4 
      Height          =   780
      Left            =   4440
      Picture         =   "Form3.frx":2328
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   780
   End
   Begin VB.Image SHAPE3 
      Height          =   720
      Left            =   4560
      Picture         =   "Form3.frx":2BF2
      Stretch         =   -1  'True
      Top             =   960
      Width           =   720
   End
   Begin VB.Image SHAPE2 
      Height          =   720
      Left            =   1440
      Picture         =   "Form3.frx":34BC
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   720
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Product ID:0011-ADH22011-595444"
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
      Left            =   3120
      TabIndex        =   2
      Top             =   6720
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "General Corporation.Bangladesh"
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
      Left            =   3360
      TabIndex        =   1
      Top             =   6480
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "TNC2 Digital Video Driver"
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
      Left            =   3600
      TabIndex        =   0
      Top             =   6240
      Width           =   2415
   End
   Begin VB.Image shape1 
      Height          =   1965
      Left            =   1440
      Picture         =   "Form3.frx":4386
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2115
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, y, p, q As Integer

Private Sub Form_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Unload Me
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
Timer1.Interval = 300
Timer2.Interval = 150
x = 75
y = 75
End Sub

Private Sub Timer1_Timer()
shape1.Move shape1.Left + x, shape1.Top + y

If shape1.Left < ScaleLeft Then x = 75
If (shape1.Left) > (ScaleWidth) Then x = -75


If shape1.Top < ScaleTop Then y = 75
If (shape1.Top) > (ScaleHeight) Then y = -75
'
End Sub

Private Sub Timer2_Timer()
If p > 720 Then
p = 0
Else
SHAPE2.Width = p
SHAPE2.Height = p
p = p + 50
End If
End Sub

Private Sub Timer3_Timer()
If p > 820 Then
p = 0
Else
SHAPE3.Width = p
SHAPE3.Height = p
p = p + 50
End If
End Sub

Private Sub Timer4_Timer()
If p > 1120 Then
p = 0
Else
SHAPE4.Width = p
SHAPE4.Height = p
p = p + 100
End If
End Sub

Private Sub Timer5_Timer()
If p > 720 Then
p = 0
Else
shape5.Width = p
shape5.Height = p
p = p + 50
End If
End Sub

Private Sub Timer6_Timer()
If p > 720 Then
p = 0
Else
Shape6.Width = p
Shape6.Height = p
p = p + 50
End If

End Sub

Private Sub Timer7_Timer()
If p > 720 Then
p = 0
Else
Shape7.Width = p
Shape7.Height = p
p = p + 50
End If

End Sub
