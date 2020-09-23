VERSION 5.00
Begin VB.Form frm_Scores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Score table"
   ClientHeight    =   2940
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Scores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   480
      Left            =   4110
      TabIndex        =   2
      Top             =   615
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Clear"
      Height          =   480
      Left            =   4110
      TabIndex        =   1
      Top             =   90
      Width           =   1290
   End
   Begin VB.ListBox lst_Scores 
      Height          =   2760
      Left            =   75
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   3960
   End
End
Attribute VB_Name = "frm_Scores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
    Dim msg
    msg = MsgBox("You dont want to do this, do you ?", vbYesNo Or vbQuestion, "Oh, no...")
    If msg = vbYes Then Kill App.Path & "\scores.txt"
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next '// if file is erased
    Dim L
    '// Open the file, and list its contents
    Open App.Path & "\scores.dat" For Input As #1
        Do Until EOF(1)
            Line Input #1, L
            lst_Scores.AddItem L
        Loop
    Close #1
End Sub
