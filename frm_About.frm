VERSION 5.00
Begin VB.Form frm_About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Cuballs"
   ClientHeight    =   3180
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   435
      Left            =   3255
      TabIndex        =   5
      Top             =   105
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Warning: "
      Height          =   1695
      Left            =   105
      TabIndex        =   2
      Top             =   1365
      Width           =   4320
      Begin VB.Label Label3 
         Caption         =   $"frm_About.frx":000C
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   105
         TabIndex        =   3
         Top             =   315
         Width           =   4110
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Version: 1.0"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "alas@eunet.yu"
      Height          =   225
      Left            =   945
      TabIndex        =   1
      Top             =   525
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Author: Sala Bojan"
      Height          =   225
      Left            =   840
      TabIndex        =   0
      Top             =   195
      Width           =   1515
   End
End
Attribute VB_Name = "frm_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If Not frm_Game.vPause Then frm_Game.mnu_Pause_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If frm_Game.vPause Then frm_Game.mnu_Pause_Click
End Sub
