VERSION 5.00
Begin VB.Form Form10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   Search Video"
   ClientHeight    =   3780
   ClientLeft      =   4410
   ClientTop       =   1260
   ClientWidth     =   1605
   ControlBox      =   0   'False
   HasDC           =   0   'False
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form10.frx":0ECA
   ScaleHeight     =   3780
   ScaleWidth      =   1605
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Txtsearchspec 
      Height          =   315
      ItemData        =   "Form10.frx":1E1EE
      Left            =   810
      List            =   "Form10.frx":1E204
      TabIndex        =   12
      Text            =   "*.dat"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   900
      TabIndex        =   11
      Text            =   "9"
      ToolTipText     =   "Default Searching File Length.You Can Change It For Different MB File..."
      Top             =   2600
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      Picture         =   "Form10.frx":1E230
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Search For Selected Type And MB Length File.."
      Top             =   3300
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   780
      Picture         =   "Form10.frx":1F779
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Go Back.."
      Top             =   3300
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   4320
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      Height          =   1215
      Left            =   2400
      ScaleHeight     =   1155
      ScaleWidth      =   915
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   4080
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Default MB"
      Height          =   255
      Left            =   80
      TabIndex        =   10
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblcount 
      BackStyle       =   0  'Transparent
      Caption         =   " 0"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      ToolTipText     =   "File Number(Searching Result)"
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblcriteria 
      BackStyle       =   0  'Transparent
      Caption         =   "File Type"
      Height          =   255
      Left            =   75
      TabIndex        =   8
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblfound 
      BackStyle       =   0  'Transparent
      Caption         =   "File Found"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   735
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Integer
Dim SearchFlag As Integer

Private Sub Command5_Click()
On Error Resume Next

    If Command5.Caption = "E&xit" Then
    Form1.mov.Visible = True
    Form1.Show
    Form1.mov.FileName = Form9.List1.List(x)
Form1.mov.Command = "Play"
Form1.Show
Form1.WindowState = 2
Form1.mov.Top = 0
Form1.mov.Left = 0
Form1.mov.Height = Form1.Height
Form1.mov.Width = Form1.Width
Form9.Timer1.Enabled = True
Form9.Timer2.Enabled = True



Form7.Show
Form1.Command1.SetFocus
Unload Me
    Else
        SearchFlag = False
    End If
End Sub


Private Sub Command4_Click()
Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
Dim result As Integer
On Error Resume Next

Form7.Show
Form9.List1.Clear

    If Command4.Caption = "&Reset" Then
        ResetSearch
        Txtsearchspec.SetFocus
        Exit Sub
    End If

    
    
    If Dir1.Path <> Dir1.List(Dir1.ListIndex) Then
        Dir1.Path = Dir1.List(Dir1.ListIndex)
        Exit Sub
    End If

    
    Picture2.Move 0, 0
    Picture1.Visible = False
    Picture2.Visible = False

    Command5.Caption = "Cancel"

    
     If Txtsearchspec.Text = "*.*" Then
    File1.Pattern = "*.dat"
    FirstPath = Dir1.Path
    DirCount = Dir1.ListCount
    NumFiles = 0
    result = DirDiver(FirstPath, DirCount, "")
    File1.Path = Dir1.Path
    
    
    File1.Pattern = "*.mpg"
    FirstPath = Dir1.Path
    DirCount = Dir1.ListCount
    NumFiles = 0
    result = DirDiver(FirstPath, DirCount, "")
    File1.Path = Dir1.Path
    
    
    File1.Pattern = "*.avi"
    FirstPath = Dir1.Path
    DirCount = Dir1.ListCount
    NumFiles = 0
    result = DirDiver(FirstPath, DirCount, "")
    File1.Path = Dir1.Path

    
     File1.Pattern = "*.mov"
    FirstPath = Dir1.Path
    DirCount = Dir1.ListCount
    NumFiles = 0
    result = DirDiver(FirstPath, DirCount, "")
    File1.Path = Dir1.Path

    Command4.Caption = "&Reset"
    Command4.SetFocus
    Command5.Caption = "E&xit"

    
    Else
    
    File1.Pattern = Txtsearchspec.Text
    FirstPath = Dir1.Path
    DirCount = Dir1.ListCount
     NumFiles = 0
    result = DirDiver(FirstPath, DirCount, "")
    File1.Path = Dir1.Path
    Command4.Caption = "&Reset"
    Command4.SetFocus
    Command5.Caption = "E&xit"
    End If
    End Sub

Private Function DirDiver(NewPath As String, DirCount As Integer, BackUp As String) As Integer
Static FirstErr As Integer
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String
Dim retval As Integer
    SearchFlag = True
    DirDiver = False
    retval = DoEvents()
    If SearchFlag = False Then
        DirDiver = True
        Exit Function
    End If
    On Local Error GoTo DirDriverHandler
    DirsToPeek = Dir1.ListCount
    Do While DirsToPeek > 0 And SearchFlag = True
            OldPath = Dir1.Path
        Dir1.Path = NewPath
        
        If Dir1.ListCount > 0 Then
            
            Dir1.Path = Dir1.List(DirsToPeek - 1)
            AbandonSearch = DirDiver((Dir1.Path), DirCount%, OldPath)
        End If
tahmina:
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop

    If File1.ListCount Then
        If Len(Dir1.Path) <= 3 Then
            ThePath = Dir1.Path
        Else
            ThePath = Dir1.Path + "\"
        End If
        For ind = 0 To File1.ListCount - 1
            entry = ThePath + File1.List(ind)
            If (FileLen(entry) / 1048000) > Text1.Text Then
           Form9.List1.AddItem entry
            lblcount.Caption = Str(Val(lblcount.Caption) + 1)
            Form9.Show
            
             End If
        Next ind
    
    End If
    If BackUp <> "" Then
        Dir1.Path = BackUp
    End If
    Exit Function
DirDriverHandler:
    If err = 7 Then
        DirDiver = True
        MsgBox "You've filled the list box. Abandoning search..."
        Exit Function
    Else
        GoTo tahmina
    End If
End Function


Private Sub dir1_LostFocus()
    Dir1.Path = Dir1.List(Dir1.ListIndex)
End Sub



Private Sub ResetSearch()

    Form9.List1.Clear
    lblcount.Caption = 0
    SearchFlag = False
    Picture2.Visible = False
    Command4.Caption = "&Search"
    Command5.Caption = "E&xit"
    Picture1.Visible = True
    Dir1.Path = CurDir: Drive1.Drive = Dir1.Path
End Sub

Private Sub form_load()
x = 0
File1.Hidden = True
File1.ReadOnly = True
End Sub

Private Sub txtSearchSpec_Change()
        File1.Pattern = Txtsearchspec.Text
End Sub

Private Sub txtSearchSpec_GotFocus()
    Txtsearchspec.SelStart = 0
    Txtsearchspec.SelLength = Len(Txtsearchspec.Text)
End Sub


Private Sub dir1_Change()
        File1.Path = Dir1.Path
End Sub

Private Sub drive1_Change()
    On Error GoTo DriveHandler
    Dir1.Path = Drive1.Drive
    Exit Sub

DriveHandler:
    Drive1.Drive = Dir1.Path
    Exit Sub
End Sub



