Attribute VB_Name = "mdl_SomeCrapsAndDecs"
Option Explicit

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Type SAFEARRAY2
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    CE0 As Long
    LB0 As Long
    CE1 As Long
    LB1 As Long
End Type

Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFOHEADER, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function GetObjectA Lib "gdi32" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


'// The ONLY consts here... not used much :), I'm weak with math...
Public Const PI = 3.14159265358979
Public Const SND_ASYNC = &H1


Public Function Zeros(Num, Count) As String
    '// Stuff for putting zeros as first digits
    '// it is good if you need to sort the numbers
    Dim I&, N$
    For I = 1 To Count - Len(CStr(Num))
        N = N & "0"
    Next I
    Zeros = N & Num
End Function

Public Sub Audio(FileName As String)
    sndPlaySound App.Path & "\" & FileName & ".wav", &H1
End Sub
