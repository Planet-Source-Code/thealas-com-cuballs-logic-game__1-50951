VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Bitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'// This amazing class is made by Damian (who ever he is), it will teach you
'// how to make real, and quality games, it is small and simple, yet amazing.
'// Instead of pixel craps, we will use DIB image (dib bitmap)
'// (c) Damian, 2000
'// dmitrya@thewercs.com, contact him if you want to use MapArray (for image manipulation)

Public hDC As Long, lpRGB As Long
Private bmH As BITMAPINFOHEADER
Private hBMO As Long, hDIB As Long

Private sa As SAFEARRAY2, saPtr As Long

Sub Create(ByVal W As Long, ByVal H As Long)
    Class_Terminate
    With bmH
        .biSize = Len(bmH)
        .biWidth = W
        .biHeight = H
        .biPlanes = 1
        .biBitCount = 24
        .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
    End With
    hDC = CreateCompatibleDC(0)
    hDIB = CreateDIBSection(hDC, bmH, 0, lpRGB, 0, 0)
    If hDIB Then
        hBMO = SelectObject(hDC, hDIB)
        Cls
    Else
        Err.Raise -1, , "DIB failed to create"
    End If
End Sub

Sub Clone(pic As StdPicture, Optional ByVal FitSize As Boolean = True)
    Dim BMP As BITMAP
    GetObjectA pic.handle, Len(BMP), BMP
    If FitSize Then Create BMP.bmWidth, BMP.bmHeight
    If hDIB Then
        Dim hDCt As Long, hBMOt As Long
        hDCt = CreateCompatibleDC(hDC)
        hBMOt = SelectObject(hDCt, pic.handle)
        SetStretchBltMode hDC, 4
        StretchBlt hDC, 0, 0, Width, Height, hDCt, 0, 0, BMP.bmWidth, BMP.bmHeight, vbSrcCopy
        SelectObject hDCt, hBMOt
        DeleteDC hDCt
    Else
        Err.Raise -1, , "DIB has to be created first"
    End If
End Sub

Sub PaintTo(ByVal toDC As Long, ByVal atX As Long, ByVal atY As Long)
    BitBlt toDC, atX, atY, bmH.biWidth, bmH.biHeight, hDC, 0, 0, vbSrcCopy
End Sub

Sub Cls()
    If hDIB Then ZeroMemory ByVal lpRGB, bmH.biSizeImage
End Sub

Private Sub Class_Terminate()
    If hDC Then
        If hBMO Then DeleteObject SelectObject(hDC, hBMO): hBMO = 0
        DeleteObject hDC: hDC = 0
    End If
End Sub

Property Get RGBSize() As Long
    RGBSize = bmH.biSizeImage
End Property

Property Get Width() As Long
    Width = bmH.biWidth
End Property

Property Get Height() As Long
    Height = bmH.biHeight
End Property

Function MapArray(ByRef A As Variant) As Long
    sa.cDims = 1
    sa.cbElements = 1
    sa.pvData = lpRGB
    sa.CE0 = bmH.biSizeImage
    
    CopyMemory saPtr, ByVal VarPtr(A) + 8, 4
    CopyMemory ByVal saPtr, VarPtr(sa), 4
    MapArray = bmH.biSizeImage \ bmH.biHeight
End Function

Sub UnMapArray(ByRef A As Variant)
    CopyMemory saPtr, ByVal VarPtr(A) + 8, 4
    CopyMemory ByVal saPtr, 0&, Len(sa)
End Sub

