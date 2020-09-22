VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   6435
   ClientLeft      =   1665
   ClientTop       =   255
   ClientWidth     =   6240
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      Picture         =   "Form2.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3960
      Top             =   2640
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   600
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4680
      Top             =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Checking For Default Video Driver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   5400
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   480
      Top             =   5640
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label label2 
      BackStyle       =   0  'Transparent
      Caption         =   "A Program By Didar"
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
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Video Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   2280
      TabIndex        =   0
      Top             =   3840
      Width           =   1605
   End
   Begin VB.Image Image1 
      Height          =   6450
      Left            =   0
      Picture         =   "Form2.frx":3206
      Top             =   0
      Width           =   6750
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command1_Click()
Unload Me
Load Form7
Form7.Show
Timer1.Enabled = False
End Sub

Private Sub Image1_Click()
frmINI.Show
End Sub

Private Sub Timer1_Timer()
  iniPath$ = "c:\windows\system.ini"
Text1.Text = GetFromINI("mci", "MPEGVideo", iniPath$)
If Text1.Text = "mciqtz.drv" Then
Unload Me
Load Form7
Form7.Show
Timer1.Enabled = False
Else
If Text1.Text = "C:\WINDOWS\system\xmdrv95.dll" Then
Unload Me
Load Form7
Form7.Show
Timer1.Enabled = False
Else
Timer2.Enabled = True
Label4.Visible = True
Shape1.Visible = True
Timer1.Enabled = False
End If
End If
End Sub

Private Sub Timer2_Timer()
If Shape1.Width = 5055 Then
Command1.Visible = True
Label4.Caption = "No Default Video Driver Found!! Select Any Driver.."
frmINI.Show
Timer2.Enabled = False
Else
Shape1.Width = Shape1.Width + 30
End If
End Sub
