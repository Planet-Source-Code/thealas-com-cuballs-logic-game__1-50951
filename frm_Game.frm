VERSION 5.00
Begin VB.Form frm_Game 
   BackColor       =   &H00A51414&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuballs"
   ClientHeight    =   3900
   ClientLeft      =   150
   ClientTop       =   765
   ClientWidth     =   8085
   Icon            =   "frm_Game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frm_Game.frx":5BDA
   ScaleHeight     =   260
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   539
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr_Add 
      Interval        =   8000
      Left            =   7170
      Tag             =   "You can shange this timer from options window"
      Top             =   2325
   End
   Begin VB.PictureBox pic_Memo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   3
      Left            =   5400
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   6
      Top             =   3000
      Width           =   600
      Begin VB.Image img_Memo 
         Height          =   600
         Index           =   3
         Left            =   0
         Picture         =   "frm_Game.frx":6C96C
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox pic_Memo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   2
      Left            =   5400
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   5
      Top             =   1800
      Width           =   600
      Begin VB.Image img_Memo 
         Height          =   600
         Index           =   2
         Left            =   0
         Picture         =   "frm_Game.frx":6CD16
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox pic_Memo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   1
      Left            =   2100
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   4
      Top             =   3000
      Width           =   600
      Begin VB.Image img_Memo 
         Height          =   600
         Index           =   1
         Left            =   0
         Picture         =   "frm_Game.frx":6D0C0
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox pic_Memo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   0
      Left            =   2100
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   3
      Top             =   1800
      Width           =   600
      Begin VB.Image img_Memo 
         Height          =   600
         Index           =   0
         Left            =   0
         Picture         =   "frm_Game.frx":6D46A
         Top             =   0
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin VB.PictureBox pic_TableFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   300
      Picture         =   "frm_Game.frx":6D814
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   0
      Top             =   300
      Width           =   7500
      Begin VB.PictureBox pic_Table 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   3600
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   7
         Top             =   0
         Width           =   300
      End
   End
   Begin VB.PictureBox pic_DropFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   3000
      Picture         =   "frm_Game.frx":8AD16
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   140
      TabIndex        =   1
      Top             =   1800
      Width           =   2100
      Begin VB.PictureBox pic_Drop 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1200
         Left            =   900
         ScaleHeight     =   80
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   80
         TabIndex        =   2
         Top             =   300
         Width           =   1200
      End
   End
   Begin VB.Shape sh_Sel 
      BorderColor     =   &H00B8B8B8&
      BorderWidth     =   2
      Height          =   690
      Left            =   2055
      Top             =   1755
      Width           =   690
   End
   Begin VB.Image img_DropRight 
      Height          =   1815
      Left            =   5100
      Top             =   1800
      Width           =   315
   End
   Begin VB.Image img_DropLeft 
      Height          =   1815
      Left            =   2700
      Top             =   1800
      Width           =   315
   End
   Begin VB.Image Image4 
      Height          =   315
      Left            =   3000
      Top             =   1500
      Width           =   2115
   End
   Begin VB.Image img_Left 
      Height          =   1515
      Left            =   0
      Top             =   0
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   3600
      Top             =   0
      Width           =   915
   End
   Begin VB.Image img_Right 
      Height          =   1515
      Left            =   7800
      Top             =   0
      Width           =   315
   End
   Begin VB.Menu mnu_Game 
      Caption         =   "&Game"
      Begin VB.Menu mnu_Game_Score 
         Caption         =   "&Score list"
      End
      Begin VB.Menu mnu_Game_S2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Game_1 
         Caption         =   "&Pretty easy"
      End
      Begin VB.Menu mnu_Game_2 
         Caption         =   "&Normal"
      End
      Begin VB.Menu mnu_Game_3 
         Caption         =   "&Hard"
      End
      Begin VB.Menu mnu_Game_Custom 
         Caption         =   "&Custom..."
      End
      Begin VB.Menu mnu_Game_S3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Game_Clear 
         Caption         =   "&Use bonus (Clear selected box)"
      End
      Begin VB.Menu mnu_Game_S1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Game_Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnu_Pause 
      Caption         =   "&Pause"
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "&Help"
      Begin VB.Menu mnu_Help_About 
         Caption         =   "&About"
      End
      Begin VB.Menu mnu_Help_Help 
         Caption         =   "&What the heck is this ???"
      End
   End
End
Attribute VB_Name = "frm_Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// Cuballs game, version 1.0
'// Based on Microsoft's "Finty Flush" game.
'//
'//  JUST READ THIS:
'// Do not distribute or compile this code without permition of the author.
'// You are only allowed to use it for personal needs. Any image file that
'// came with this program, and any grafic part of this program my be used
'// for any kind of distribution/use. Design&Programming - Sala Bojan
'// Copyright (C) 2003 Hallsoft team, Yugoslavia, www.univerzalsoft.com, www.hallsoft.tk .
'//
'// Please email me if you have any questions or want to report a bug, to: alas@euney.yu !
'// PLEASE:
'// VOTE  VOTE  VOTE  VOTE  VOTE  VOTE  VOTE  VOTE  VOTE  VOTE  VOTE  VOTE  VOTE  VOTE  !
'// Came from www.planetsourcecode.com .



Option Explicit

'// Image processing
Private cPicture(1 To 20) As New StdPicture
Private cBmp As New cls_Bitmap

'// Some table vars
Private vCol As Long '// Selected Col on the table, zero based
Private vCols As Long '// Count, one based
Private vColor(1 To 25, 1 To 4) As Long '// 1 for empty

'// Players craps, if you wanna cheat, then this is for you :)
Private vScore As Long
Private vColors As Long '// How many colors are in the game
Private vLevel As Long '// Current level
Private vCubes As Long '// Solved cubes
Public vPause As Boolean

'1 | 5 | ...
'2 | 6
'3 | 7
'4 | 8

'// Dropbox stuffs here
Private Type tMemo
    Memo(1 To 4, 1 To 4)
End Type

                  'Column   Row
Private vDropColor(1 To 4, 1 To 4) As Long '// 1 for empty
Private vDropCol As Long '// Selected col on the droptable
Private vSelectedMemoBox As Long
Private vMemoBox(0 To 3) As tMemo '// from zero, cuz the images are zero based

Private Sub Form_Load()
    '// Load pictures, using simple vb StdPicture
    Set cPicture(1) = LoadPicture(App.Path & "\images\ball1.bmp")
    Set cPicture(2) = LoadPicture(App.Path & "\images\ball2.bmp")
    Set cPicture(3) = LoadPicture(App.Path & "\images\ball3.bmp")
    Set cPicture(4) = LoadPicture(App.Path & "\images\ball4.bmp")
    Set cPicture(5) = LoadPicture(App.Path & "\images\ball5.bmp")
    Set cPicture(6) = LoadPicture(App.Path & "\images\ball6.bmp")
    Set cPicture(7) = LoadPicture(App.Path & "\images\black.bmp")
    Set cPicture(8) = LoadPicture(App.Path & "\images\black_mask.bmp")
    
    Set cPicture(15) = LoadPicture(App.Path & "\images\rotating.bmp")
    Set cPicture(16) = LoadPicture(App.Path & "\images\rotating_mask.bmp")
    
    '// How many colors
    vColors = 2
    
    Dim I&
    For I = 0 To 6
        AddLine vColors, I
    Next I
    ClearDrop
    For I = 0 To 3
        MemorizeBox CInt(I)
    Next I
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '// Clear memory
    Dim I&
    For I = 1 To 10
        Set cPicture(I) = Nothing
    Next I
    Set cBmp = Nothing
    MsgBox "Thanks for playing my game" & vbCrLf & "VOTE FOR ME !", vbInformation
End Sub


Private Sub img_DropLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With pic_Drop
        If Not .Left = 0 Then .Left = .Left - 20: vDropCol = vDropCol + 1
    End With
End Sub

Private Sub img_DropRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With pic_Drop
        If Not .Left + .Width = 140 Then .Left = .Left + 20: vDropCol = vDropCol - 1
    End With
End Sub

Private Sub img_Left_Click()
    '// Move the table to the left, for 20 pixels
    With pic_Table
        If Not .Left = 0 Then
            '// Not to pass the arrow
            If Not .Left + .Width = 260 Then
                .Move .Left - 20
                vCol = vCol + 1 '// Move Col
            End If
        End If
    End With
End Sub

Private Sub img_Memo_Click(Index As Integer)
    With img_Memo(Index)
        sh_Sel.Move pic_Memo(Index).Left - 3, pic_Memo(Index).Top - 3 '// Select the box
        img_Memo(vSelectedMemoBox).Visible = True
        vSelectedMemoBox = Index
        .Visible = False
    End With
    
    '// Transfer from small to big :)
    Dim X&, Y&
    For X = 1 To 4
        For Y = 1 To 4
            vDropColor(X, Y) = vMemoBox(Index).Memo(X, Y)
        Next Y
    Next X
    FillDrop '// Draw the balls
End Sub

Private Sub img_Right_Click()
    '// Move the table to the right, for 20 pixels
    With pic_Table
        If Not .Left + .Width = 500 Then
            '// Not to pass the arrow
            If Not .Left = 240 Then
                .Move .Left + 20
                vCol = vCol - 1 '// Select Col
            End If
        End If
    End With
End Sub

Public Sub AddLine(Colors As Long, Col&)
    '// Adds one new line to the table
    Dim C&, I&, Y&, T&, RowColor&
    If vCols = 25 Then
        tmr_Add.Enabled = False
        Audio "gmover"
        '// Add our player to the score list
        If vScore > 10 Then
            Dim msg
            msg = InputBox("Game over dude, too many balls for today. You can be listed on the score list if you want to, just enter your name here, and you can compare your score with others: ", "Enter name")
            If Not msg = "" Then
                Open App.Path & "\scores.dat" For Append As #1
                    Print #1, Zeros(vScore, 4) & " , " & msg
                Close #1
            End If
            Unload Me
            Exit Sub
        Else
            MsgBox "Game over dude, too many balls for today."
            Unload Me
            Exit Sub
        End If
    End If
again:
    '// Set random color, or empty
    Y = 0
    T = 0
    Randomize
    '// Move the table
    If pic_Table.Left + pic_Table.Width = 500 Then pic_Table.Left = pic_Table.Left - 20
    pic_Table.Width = Col * 20 + 20
    '// Choose color
    RowColor = Int((Rnd * Colors) + 2)
    For I = 1 To 4
new_color:
        C = Int((Rnd * Colors) + 1)
        If Not C = 1 Then If Not C = RowColor Then GoTo new_color
        T = T + C
        cBmp.Clone cPicture(C)
        cBmp.PaintTo pic_Table.hDC, Col * 20, Y
        Y = Y + 20
        vColor(Col + 1, I) = C '// Set the color
    Next I
    
    '// If there are no balls
    If T = 4 Then
        pic_Table.Width = pic_Table.Width - 20
        GoTo again
    End If
    
    vCols = vCols + 1
End Sub

Public Sub ClearDrop()
   '// Clears the drop and memo boxes
    Dim I&, X&, Y&
    cBmp.Clone cPicture(1)
    For X = 0 To 3
        For Y = 0 To 3
            cBmp.PaintTo pic_Drop.hDC, X * 20, Y * 20
            vDropColor(X + 1, Y + 1) = 1
        Next Y
    Next X
End Sub

Public Sub MemorizeBox(Index As Integer)
    SetStretchBltMode pic_Memo(Index).hDC, 4  '// Halftone mode
    '// Just replicate the image of box
    StretchBlt pic_Memo(Index).hDC, 0, 0, 40, 40, pic_Drop.hDC, 0, 0, 80, 80, vbSrcCopy
    pic_Memo(Index).Refresh
        
'    '// And cover other boxes
'    Dim I&
'    cBmp.Clone cPicture(7)
'    For I = 0 To 3
'        If Not I = Index Then
'            BitBlt pic_Memo(I).hDC, 0, 0, 40, 40, cBmp.hDC, 0, 0, vbSrcPaint
'        End If
'    Next I
    
    '// Now, we have to transfer bytes
    Dim X&, Y&
    For X = 1 To 4
        For Y = 1 To 4
            vMemoBox(Index).Memo(X, Y) = vDropColor(X, Y)
        Next Y
    Next X
End Sub


Private Sub mnu_Game_1_Click()
    tmr_Add.Interval = 15000
End Sub

Private Sub mnu_Game_2_Click()
    tmr_Add.Interval = 8000
End Sub

Private Sub mnu_Game_3_Click()
    tmr_Add.Interval = 4000
End Sub

Private Sub mnu_Game_Clear_Click()
    Dim msg
    msg = MsgBox("You can use this bonus only once, are you sure you want to clear the drop-box ?", vbYesNo Or vbQuestion)
    If msg = vbYes Then
        Audio "cube"
        ClearDrop
        FillDrop
        MemorizeBox (vSelectedMemoBox)
        mnu_Game_Clear.Enabled = False
    End If
End Sub

Private Sub mnu_Game_Custom_Click()
On Error Resume Next
    Dim N
    N = InputBox("Enter the interval for inserting rows in miliseconds (1000=1 sec.): ", "Totaly insane", tmr_Add.Interval)
    If IsNumeric(CLng(N)) Then tmr_Add.Interval = N
End Sub

Private Sub mnu_Game_Exit_Click()
    Dim msg
    '// If the game has started
    If Len(Me.Caption) <> 7 Then msg = MsgBox("You realy want to leave now ? ... dont go !", vbYesNo, "Exit")
    If Not msg = vbNo Then Unload Me
End Sub

Private Sub mnu_Game_Score_Click()
    frm_Scores.Show vbModal, Me
End Sub

Private Sub mnu_Help_About_Click()
    frm_About.Show vbModal, Me
End Sub

Private Sub mnu_Help_Help_Click()
    Dim msg
    msg = "The objective is to fill all empty grids (four small boxes) with the balls of the same color. You must drop the columns from the upper grid to the lower grid by clicking on the upper table (grid), you can rotate the lower grid by clicking the left or right mouse button. There are five levels, each level adds new color in the game, the last level is almost impossible to complete. You must fill out three grids to complete a level, you can manipulate with the four grids at the same time. If you have any problems or questions, then contact me please."
    MsgBox msg, vbInformation, "Help"
End Sub

Public Sub mnu_Pause_Click()
    vPause = Not vPause
    pic_Table.Visible = Not vPause
    tmr_Add.Enabled = Not vPause
End Sub



Private Sub pic_Drop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim C&, F&, BMP As New cls_Bitmap, Mask As New cls_Bitmap
    BMP.Clone cPicture(15)
    Mask.Clone cPicture(16)
    
    If Button = 2 Then
        F = 1
        Audio "max"
    Else
        Audio "min"
    End If
    '// Now rotate the cube
    Do
        DoEvents
        C = C + 1
        If C = 6000 Then
            pic_DropFrame.Cls
            '// Remove the real box, and blt the false one
            pic_Drop.Visible = False
            
            BitBlt pic_DropFrame.hDC, pic_Drop.Left - 20, pic_Drop.Top - 20, 120, 120, Mask.hDC, F * 120, 0, vbSrcPaint
            BitBlt pic_DropFrame.hDC, pic_Drop.Left - 20, pic_Drop.Top - 20, 120, 120, BMP.hDC, F * 120, 0, vbSrcAnd
            
            If Button = 1 Then
                F = F + 1: If F = 3 Then Exit Do
            Else
                F = F - 1: If F = -2 Then Exit Do
            End If
            C = 0
        End If
    Loop
    pic_DropFrame.Cls
    pic_Drop.Visible = True
    
    If Button = 1 Then
        RotateDrop False
    Else
        RotateDrop True
    End If
    FillDrop '// Refresh the droptable with new values
    MemorizeBox CInt(vSelectedMemoBox)
    Set BMP = Nothing
    Set Mask = Nothing
End Sub



Public Sub DropCol(Col As Long)
    Dim lLeft&, lRight&, I&, C&
    With pic_Table
        '// Check for match
        For C = 1 To 4
            If vColor(vCol + 1, C) > 1 Then
                If vDropColor(vDropCol + 1, C) > 1 Then AddLine vColors, vCols: Audio "error": Exit Sub
            End If
        Next C
        
        Audio "click"
        
        '// Replicate row
        For C = 1 To 4
            If vColor(vCol + 1, C) > 1 Then
                 vDropColor(vDropCol + 1, C) = vColor(vCol + 1, C)
            End If
        Next C
        
        For C = 1 To 4
            If vColor(vCol + 1, C) > 1 Then
                BitBlt pic_Drop.hDC, vDropCol * 20, (C - 1) * 20, 20, 20, pic_Table.hDC, Col * 20, (C - 1) * 20, vbSrcCopy
            End If
        Next C
        BitBlt .hDC, Col * 20, 0, .Width - (Col + 1) * 20, 80, .hDC, (Col + 1) * 20, 0, vbSrcCopy
        
        '// Move the table, delete row
        For I = vCol + 1 To vCols - 1
            For C = 1 To 4
                 vColor(I, C) = vColor(I + 1, C)
            Next C
        Next I
        If .Left + .Width = 260 Then .Left = .Left + 20: vCol = vCol - 1 '// Move it, if it is the last col
        .Width = .Width - 20
        .Refresh
        pic_Drop.Refresh
        vCols = vCols - 1
        MemorizeBox CInt(vSelectedMemoBox)
        CheckDrop
        Score 1
    End With
End Sub

Private Sub pic_Table_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then DropCol vCol
End Sub


Private Sub pic_Table_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then tmr_Add_Timer
End Sub

Private Sub pic_TableFrame_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < 240 Then img_Left_Click Else img_Right_Click
End Sub

Private Sub pic_DropFrame_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < 60 Then img_DropLeft_MouseUp 1, 0, 0, 0 Else img_DropRight_MouseUp 1, 0, 0, 0
End Sub

Private Sub tmr_Add_Timer()
    AddLine vColors, vCols
End Sub

Public Sub FillDrop()
    '// Fills the drop
    Dim I&, X&, Y&
    For X = 0 To 3
        For Y = 0 To 3
            cBmp.Clone cPicture(vDropColor(X + 1, Y + 1))
            cBmp.PaintTo pic_Drop.hDC, X * 20, Y * 20
        Next Y
    Next X
    pic_Drop.Refresh
End Sub

Public Sub RotateDrop(Right As Boolean)
    Dim vTemp(1 To 4, 1 To 4) As Long, I&, C&, Y&
    Y = 4
    '// Rotate the values for 90 degrees:
    '// Bellow is the first code, this shit is too complicated...
    For I = 1 To 4
        For C = 1 To 4
            If Right Then
                vTemp(I, C) = vDropColor(C, Y)
            Else
                vTemp(C, Y) = vDropColor(I, C)
            End If
        Next C
        Y = Y - 1
    Next I
    For I = 1 To 4
        For C = 1 To 4
            vDropColor(I, C) = vTemp(I, C)
        Next C
    Next I
'// Here is the basic stuff, no need for math :) :
'        vTemp(1, 1) = vDropColor(1, 4)
'        vTemp(1, 2) = vDropColor(2, 4)
'        vTemp(1, 3) = vDropColor(3, 4)
'        vTemp(1, 4) = vDropColor(4, 4)
'
'        vTemp(2, 1) = vDropColor(1, 3)
'        vTemp(2, 2) = vDropColor(2, 3)
'        vTemp(2, 3) = vDropColor(3, 3)
'        vTemp(2, 4) = vDropColor(4, 3)
'       ... etc.
End Sub



Public Sub CheckDrop()
    '// If we have the full box
    Dim X&, Y&, C&, Temp(1 To 16) As Long, I&
    For X = 1 To 4
        For Y = 1 To 4
            I = I + 1
            Temp(I) = vDropColor(X, Y)
        Next Y
    Next X
    
    For I = 2 To 16
        If Temp(I) = Temp(I - 1) Then C = C + 1
    Next I
    Debug.Print vCubes
    If C = 15 Then
        Audio "cube"
        vCubes = vCubes + 1
        Score 10 '// Count the score, this is what we all want :)
        If vCubes = 3 Then Level 1: vCubes = 0
        ClearDrop   '// Clear it
        FillDrop    '// Draw it
        MemorizeBox CInt(vSelectedMemoBox) '// Memorize it, this CInt is DRIVING me crazy...
    End If
End Sub

Public Sub Score(S As Long)
    vScore = vScore + S
    Me.Caption = "Cuballs ::: Level: " & vLevel & " , Score: " & vScore
End Sub

Public Sub Level(L As Long)
    Audio "level"
    Score 30
    vLevel = vLevel + L
    If vLevel = 5 Then MsgBox "You have won the game now, please email me your score !", vbExclamation
    If vLevel < 6 Then vColors = vColors + 1
        Me.Caption = "Cuballs ::: Level: " & vLevel & " , Score: " & vScore
End Sub


