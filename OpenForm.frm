VERSION 5.00
Begin VB.Form OpenForm 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PBase 1.0"
   ClientHeight    =   8310
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11880
   ControlBox      =   0   'False
   Icon            =   "OpenForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Enter Picture Description"
      Top             =   6960
      Visible         =   0   'False
      Width           =   8895
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "Prev. Thumbs"
      Height          =   495
      Left            =   840
      TabIndex        =   11
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Thumbs"
      Height          =   495
      Left            =   840
      TabIndex        =   12
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   6745
      Index           =   12
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   120
      Width           =   8895
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   1440
      TabIndex        =   10
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   9
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   1440
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Index           =   9
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Index           =   8
      Left            =   120
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Index           =   7
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Index           =   6
      Left            =   120
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Index           =   5
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Index           =   4
      Left            =   120
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Index           =   3
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Index           =   2
      Left            =   120
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Index           =   1
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   911
      Index           =   0
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "Files"
      Begin VB.Menu mnuFOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit Description"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "OpenForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I needed a picture display program that would show me the
'thumbnails of all pictures in a given directory, enlarge the
'chosen thumnail and give me a description of what I'm looking
'at.
'So, here is my answer. It probably could be elaborated upon,
'but it'll do for my purposes. I am a coach, so I can give this
'program to my athletes and then, periodically, send them
'annotated pictures of their performances which they can then
'review along with my thoughts. Cool!
'Most of the code is self explanatory. Comments have been added
'only where necessary.
'I wrote this code for a 800*600 screen. No apologies!
'If anybody can improve this code, don't bitch... dot it and let
'me know.
'peter_rabbit@shaw.ca - Dec. 20, 2003

Private Sub Form_Load()
Dim x As Byte

For x = 0 To 9
    Label1(x).ForeColor = vbBlack
Next

Edited = False
cmdPrev.Enabled = False
cmdNext.Enabled = False
LargeSize = False
End Sub

Private Sub cmdNext_Click()
LoadingForm.Show
LoadingForm.Refresh

SaveText CurrDex + OldDex   'save current notes
CurrDex = CurrDex + 10      'next pics on the list
Populate
LoadingForm.Hide
End Sub

Private Sub cmdPrev_Click()
LoadingForm.Show
LoadingForm.Refresh

SaveText CurrDex + OldDex   'save current notes
CurrDex = CurrDex - 10      'previous pics on the list
If CurrDex < 0 Then CurrDex = 0
If CurrDex >= 0 Then Populate   'show them
LoadingForm.Hide
End Sub

Private Sub Image1_Click(i As Integer)
If i <> OldDex And i < 12 Then
    SizePics i, 0
    Label1(OldDex).ForeColor = vbWhite
    Label1(OldDex).BackColor = vbBlack
    Label1(i).ForeColor = vbBlack
    Label1(i).BackColor = vbYellow
    SaveText CurrDex + OldDex   'save current notes
    OldDex = i
    Edited = False
    Text1.Locked = True
    LoadText CurrDex + OldDex   'get notes for curr. pic
End If

If i = 12 Then                  'enlarge or normal size
    Select Case LargeSize
        Case True
            SizePics OldDex, 0
            setControls True
            LargeSize = False
        Case False
            setControls False
            SizePics OldDex, 1
            LargeSize = True
    End Select
End If
End Sub

Private Sub mnuEdit_Click()
    Text1.Locked = False
    Edited = True
    Text1.SelStart = Len(Text1)
    Text1.SetFocus
End Sub

Private Sub mnuFExit_Click()    'shut it dawn
    SaveText CurrDex + OldDex
    Unload DirForm
    Unload Me
    End
End Sub

Private Sub mnuFOpen_Click()
    DirForm.Show
End Sub

Public Sub Populate()           'show 10 thumbs & 1 large
Dim i As Byte, j As Byte
Dim x As Integer

For i = 0 To 9
    Label1(i).ForeColor = vbWhite
    Label1(i).BackColor = vbBlack
Next

cmdPrev.Enabled = False
cmdNext.Enabled = False

x = 0
j = CurrDex + 9
If ListCount < j Then j = ListCount

For i = CurrDex To j
    FileName = DirForm.File1.Path + "\" + DirForm.File1.List(i)
    Image1(x).Picture = LoadPicture(FileName)
    Image1(x).Enabled = True
    Label1(x).Caption = Right$(DirForm.File1.List(i), 15)
    x = x + 1
Next
For i = x To 9          'hide empty images
    Image1(i).Picture = LoadPicture("")
    Image1(i).Enabled = False
    Label1(i).Caption = ""
Next

SizePics 0, 0           'display first image as large
Label1(0).ForeColor = vbBlack
Label1(0).BackColor = vbYellow

If CurrDex > 9 Then
    cmdPrev.Enabled = True
Else
    cmdPrev.Enabled = False
End If
If j < ListCount Then
    cmdNext.Enabled = True
Else
    cmdNext.Enabled = False
End If
OldDex = 0
LoadText CurrDex        'display associated notes
End Sub

Private Sub SizePics(i As Integer, Size As Byte)
Dim PicX As Long, PicY As Long
Dim XYRatio As Single
Dim Wide As Integer, High As Integer, Top As Integer, Left As Integer

Select Case Size
    Case 0  'small
        Left = 2880
        Wide = 8995
        Top = 120
        High = 6745
    Case 1  'large
        Left = 470
        Wide = 11060
        Top = 0
        High = 8295
End Select

PicX = Image1(i).Picture.Width
PicY = Image1(i).Picture.Height
XYRatio = PicX / PicY       'keep aspect ratio

With Image1(12)
    .Visible = False        'avoid flicker

    If PicX > PicY Then
        .Width = Wide
        .Height = Wide / XYRatio
        .Top = Top + (High - .Height) / 2   'centered
        .Left = Left
    Else
        .Height = High
        .Width = High * XYRatio
        .Top = Top
        .Left = Left + (Wide - .Width) / 2  'centered
    End If

    .Picture = Image1(i).Picture
    .Visible = True
End With
End Sub

Private Sub SaveText(i As Integer)
Dim DirName As String, Fname As String

If Edited = True Then
    DirName = DirForm.Dir1
    If Right$(DirName, 1) <> "\" Then DirName = DirName + "\"
    Fname = DirForm.File1.List(i)
    Fname = Left$(Fname, Len(Fname) - 3) + "txt"

    Open DirName + Fname For Output As 1
    Write #1, Text1
   Close #1
End If
End Sub

Private Sub LoadText(i As Integer)
On Error Resume Next
Dim DirName As String, Fname As String, Temp As String

DirName = DirForm.Dir1
If Right$(DirName, 1) <> "\" Then DirName = DirName + "\"
Fname = DirForm.File1.List(i)
Fname = Left$(Fname, Len(Fname) - 3) + "txt"

Open DirName + Fname For Input As 1
Input #1, Temp
Close #1

Text1 = Temp
End Sub

Private Sub setControls(i As Boolean)   'hide/show controls
Dim x As Byte

For x = 0 To 9
    Image1(x).Enabled = i
Next

mnuFiles.Enabled = i
mnuEdit.Enabled = i
cmdNext.Visible = i
cmdPrev.Visible = i
Text1.Visible = i
End Sub

Private Sub mnuHelp_Click()
    Load frmBrowser
    frmBrowser.WebBrowser1.Navigate App.Path & "\help.htm"
    frmBrowser.Show 1
End Sub

