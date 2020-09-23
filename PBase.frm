VERSION 5.00
Begin VB.Form DirForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Directory Open"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "PBase.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "DirForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
ListCount = File1.ListCount - 1
If ListCount >= 0 Then
    LoadingForm.Show
    LoadingForm.Refresh
    OldPath = Dir1.Path
    CurrDex = 0
    With OpenForm
        .Enabled = True
        .Populate
        LoadingForm.Hide
        .Text1.Visible = True
        .cmdNext.Visible = True
        .cmdPrev.Visible = True
    End With
    cmdCancel.Enabled = True
    DirForm.Hide
End If
End Sub

Private Sub cmdCancel_Click()
If File1.ListCount Then
    Dir1.Path = OldPath
    OpenForm.Enabled = True
    DirForm.Hide
Else
    OpenForm.Enabled = True
    DirForm.Hide
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
OpenForm.Show
OpenForm.Enabled = False
File1.Pattern = "*.BMP;*.Gif;*.JPG"
End Sub
