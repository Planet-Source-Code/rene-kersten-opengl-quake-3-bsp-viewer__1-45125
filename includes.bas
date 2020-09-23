Attribute VB_Name = "includes"

Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SaveDIBitmap Lib "gdi32" (ByVal path As String, ByVal info As BITMAPINFO, bits As Any) As Long
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Const DIB_RGB_COLORS = 0 '  color table in RGBs
Public Const BI_RGB = 0&

Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

Public Type BITMAP '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Public Function splitword(word As String) As String()
    On Error GoTo errhandler
    
    Dim words() As String
    Dim c As String
    Dim ptr As Long
    Dim inQuotes As Boolean
    Dim i As Long
    
    ReDim words(1)
    
    i = -1
    
    Do
        ptr = ptr + 1
        
        If ptr > Len(word) Then Exit Do
        
        c = Mid(word, ptr, 1)
        
        If inQuotes Then
            If c = """" Then
                inQuotes = False
            Else
                words(i) = words(i) & c
            End If
        Else
            If c = """" Then
                i = i + 1
                If i = 2 Then Exit Do
                inQuotes = True
            End If
        End If
    Loop
    
    splitword = words
    
    Exit Function
    
errhandler:
    ReDim words(0)
    splitword = words
End Function

Public Function pathexists(path As String) As Boolean
    On Error GoTo pathDoesntExist
    
    Dim filenr As Long
    
    filenr = FreeFile
    Open path For Input As filenr
    Close filenr
    
    pathexists = True
    
    Exit Function
    
pathDoesntExist:
    pathexists = False
End Function

Public Function makevect(x As Single, y As Single, z As Single) As vect
    makevect.x = x
    makevect.y = y
    makevect.z = z
End Function
