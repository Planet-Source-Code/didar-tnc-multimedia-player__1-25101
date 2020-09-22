VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About TNC Multimedia"
   ClientHeight    =   4440
   ClientLeft      =   2340
   ClientTop       =   1500
   ClientWidth     =   5910
   HasDC           =   0   'False
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.PictureBox image6 
         AutoSize        =   -1  'True
         Height          =   3945
         Left            =   600
         Picture         =   "Form4.frx":08CA
         ScaleHeight     =   3885
         ScaleWidth      =   4500
         TabIndex        =   10
         Top             =   240
         Width           =   4560
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   255
         Left            =   4680
         TabIndex        =   8
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "General Corporation. Number One In The Software Technology"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   2880
         Width           =   4695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Dedicated To Tahmina Nur Chowdhuri"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   3000
         Width           =   3615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Product No:0001-1244HM-44500"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "CopyRight By General Corporation"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   3600
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "A multimedia Production"
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
         TabIndex        =   4
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   360
         Picture         =   "Form4.frx":12423
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form4.frx":12CED
         Height          =   1455
         Left            =   600
         TabIndex        =   3
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Version 4.2"
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "About TNC2 Driver"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   480
         X2              =   5160
         Y1              =   2520
         Y2              =   2520
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next

Form1.Command1.SetFocus
Unload Me
End Sub

Private Sub form_load()
Label7.Visible = False
image6.Visible = False
End Sub

Private Sub image6_Click()
image6.Visible = False
Command1.Visible = True
End Sub

Private Sub Label5_DblClick()
image6.Visible = True
Command1.Visible = False
Label7.Visible = True
Label8.Visible = False
End Sub
