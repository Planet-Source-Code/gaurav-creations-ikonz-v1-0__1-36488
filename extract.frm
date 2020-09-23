VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ikonz v1.0"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10005
   Icon            =   "extract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   10005
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Extract Selected Icon"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Top             =   4680
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   5640
      TabIndex        =   6
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton Command2 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   16
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   13
         Text            =   "0"
         Top             =   2640
         Width           =   495
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   4095
      End
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   2280
         Pattern         =   "*.dll;*.ocx;*.exe"
         TabIndex        =   9
         Top             =   600
         Width           =   1935
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4095
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2055
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   4560
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Show files with min. icons"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "icons"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   14
         Top             =   2640
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   4455
      Left            =   120
      ScaleHeight     =   4395
      ScaleWidth      =   4995
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   33375
         Left            =   0
         ScaleHeight     =   33315
         ScaleWidth      =   4995
         TabIndex        =   3
         Top             =   0
         Width           =   5055
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   0
            Left            =   -540
            ScaleHeight     =   426.667
            ScaleMode       =   0  'User
            ScaleWidth      =   426.667
            TabIndex        =   4
            Top             =   10
            Width           =   480
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "www.gauravcreations.cjb.net"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   360
            TabIndex        =   17
            Top             =   1920
            Width           =   4455
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4455
      LargeChange     =   10
      Left            =   5280
      Max             =   100
      Min             =   1
      SmallChange     =   5
      TabIndex        =   1
      Top             =   120
      Value           =   1
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0"
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Icons Displayed"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Ikonz '
' Source Code By Gaurav dhup '
' www.gauravcreations.cjb.net '

Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Path As String
Dim X As Integer ' Variable for index of picture box 2
Dim setload As Integer ' Variable for determining whether icons are loaded
Dim oldtop As Integer ' variable for scrolling icons to the next line
Dim portind As Integer ' variable for determining the index of the selected icons

' Save Command '
Private Sub Command1_Click()
On Error GoTo ClassHandler1
ClassHandler1:
    If Err.Number = 53 Then
    End If
    Resume Next
CommonDialog1.Filter = "Bmp File|*.bmp|Icon File|*.ico"
CommonDialog1.ShowSave
SavePicture Picture2(portind).Image, CommonDialog1.FileName
End Sub

' Searching for specified number of icons '
Private Sub Command2_Click()

' Erasing the previously displayed list '
If List1.ListCount <> 0 Then
listc = List1.ListCount - 1
For cfile = listc To 0 Step -1
    List1.RemoveItem cfile
Next cfile
End If
ProgressBar1.Visible = True
If File1.ListCount <> 0 Then
    ProgressBar1.Max = File1.ListCount
End If

' Checking for specified value and transfering the appropriate file paths to list box from file list box
For cfile = 0 To File1.ListCount - 1
    File1.ListIndex = cfile
    pathcheck = File1.Path + "\" + File1.FileName
    return1& = ExtractIcon(Me.hWnd, pathcheck, -1)
    If return1& >= Val(Text3.Text) Then
       List1.AddItem pathcheck
    End If
    ProgressBar1.Value = cfile
Next cfile
ProgressBar1.Visible = False
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo ClassHandler1
ClassHandler1:
    If Err.Number = 68 Then
        response1 = MsgBox("Drive not ready", vbOKOnly, "Ikons")
    End If
    Resume Next
Dir1.Path = Drive1.Drive
End Sub

' Extract Icons from File list box on Dbl Click '
Private Sub File1_dblClick()
Path = File1.Path + "\" + File1.FileName
Call extract
End Sub

Private Sub Form_Load()
setload = 0
oldtop = Picture2(X).Top
End Sub

Private Sub Label4_Click()
Shell "Explorer http://www.gauravcreations.cjb.net"
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFF0000
End Sub

Private Sub List1_Click()
Path = List1.List(List1.ListIndex)
Call extract
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.ForeColor = &HFF&
End Sub

' Selecting Icons displayed for extraction '
Private Sub Picture2_Click(Index As Integer)
For X = 1 To Val(Text1.Text)
    Picture2(X).BorderStyle = 0
Next X
Picture2(Index).BorderStyle = 1
portind = Index
End Sub

Private Sub VScroll1_Change()
   Picture1.Top = -VScroll1.Value
End Sub

' Main function to extract the icons from Dll, Exes and Ocx files
Private Sub extract()
 return1& = ExtractIcon(Me.hWnd, Path, -1)
If setload = 1 Then
   For X = 1 To Val(Text1.Text)
       Unload Picture2(X)
   Next X
   setload = 0
End If
Text1.Text = return1&
If Val(Text1.Text) > 27 Then
   Label4.Visible = False
End If
If Val(Text1.Text) < 27 Then
   Label4.Visible = True
End If

For X = 1 To Val(Text1.Text)
    Load Picture2(X)
    Picture2(X).Left = Picture2(X - 1).Left + 560
    Picture2(X).Top = oldtop
    If Picture2(X).Left > (9 * 560) Then
       Picture2(X).Left = 10
       Picture2(X).Top = Picture2(X).Top + 560
       oldtop = Picture2(X).Top
    End If
    Picture2(X).Visible = True
    Picture2(X).Picture = LoadPicture()
    return2& = ExtractIcon(Me.hWnd, Path, return1& - X)
    return3& = DrawIcon(Picture2(X).hdc, 0, 0, return2&)
Next X
setload = 1
oldtop = 10
VScroll1.Max = (Val(Text1.Text) \ 11) * 560
End Sub

