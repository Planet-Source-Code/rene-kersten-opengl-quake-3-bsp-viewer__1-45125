Attribute VB_Name = "ElementGL"
'************************************************
'*  ElementGL.bas                               *
'*                                              *
'* By: W-Buffer                                 *
'* Web: www.lunarpages.com/istudios/            *
'* Mail: chadruva@hotmail.com                   *
'*                                              *
'* Notes: Do whatever you want with this bas    *
'*        (Steal, Copy, Etc.), as long this     *
'*        note stays here.                      *
'************************************************

Option Explicit

'Display Settings
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_FULLSCREEN = &H4
Public Const CDS_RESET = &H40000000
Public Const CCDEVICENAME = 32

'Display Setting Return Values
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1
Public Const DISP_CHANGE_FAILED = -1
Public Const DISP_CHANGE_BADMODE = -2
Public Const DISP_CHANGE_NOTUPDATED = -3
Public Const DISP_CHANGE_BADFLAGS = -4
Public Const DISP_CHANGE_BADPARAM = -5

Public Const ENUM_CURRENT_SETTINGS = -1

'Extensions Supported
Public Const GL_EXT_MULTITEXTURE = 1
Public Const GL_EXT_LOCKARRAYS = 2

'Font
Public Const FW_BOLD = 700
Public Const FW_DONTCARE = 0
Public Const FW_EXTRABOLD = 800
Public Const FW_EXTRALIGHT = 200
Public Const FW_HEAVY = 900
Public Const FW_LIGHT = 300
Public Const FW_MEDIUM = 500
Public Const FW_NORMAL = 400
Public Const FW_SEMIBOLD = 600
Public Const FW_THIN = 100

Public Enum FontWeight
    FWE_BOLD = 700
    FWE_DONTCARE = 0
    FWE_EXTRABOLD = 800
    FWE_EXTRALIGHT = 200
    FWE_HEAVY = 900
    FWE_LIGHT = 300
    FWE_MEDIUM = 500
    FWE_NORMAL = 400
    FWE_SEMIBOLD = 600
    FWE_THIN = 100
End Enum

Private Declare Function GetLastError Lib "kernel32" () As Long

'Misc Font Stuff
Public Const DEFAULT_PITCH = 0

'Other Font Stuff
Public Const ANSI_CHARSET = 0
Public Const ANSI_FIXED_FONT = 11
Public Const ANSI_VAR_FONT = 12

'Font Precision
Public Const OUT_CHARACTER_PRECIS = 2
Public Const OUT_DEFAULT_PRECIS = 0
Public Const OUT_DEVICE_PRECIS = 5
Public Const OUT_OUTLINE_PRECIS = 8
Public Const OUT_RASTER_PRECIS = 6
Public Const OUT_STRING_PRECIS = 1
Public Const OUT_STROKE_PRECIS = 3
Public Const OUT_TT_ONLY_PRECIS = 7
Public Const OUT_TT_PRECIS = 4

'Font Cliping
Public Const CLIP_CHARACTER_PRECIS = 1
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const CLIP_EMBEDDED = 128
Public Const CLIP_LH_ANGLES = 16
Public Const CLIP_MASK = &HF
Public Const CLIP_STROKE_PRECIS = 2
Public Const CLIP_TO_PATH = 4097
Public Const CLIP_TT_ALWAYS = 32

'Pixel Format
Public Const PFD_DRAW_TO_WINDOW As Long = &H4
Public Const PFD_SUPPORT_OPENGL As Long = &H20
Public Const PFD_DOUBLEBUFFER As Long = &H1
Public Const PFD_TYPE_RGBA As Long = 0
Public Const PFD_MAIN_PLANE As Long = 0
Public Const PFD_DRAW_TO_BITMAP As Long = &H8

'PointLong
Public Type POINTL
    x As Long
    y As Long
End Type

'Device Mode with some Workarounds for VB
Public Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCDEVICENAME
    dmLogPixels As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

'Pixel Format   36 bytes
Public Type PIXELFORMATDESCRIPTOR
    nSize As Integer
    nVersion As Integer
    dwFlags As Long
    iPixelType As Byte
    cColorBits As Byte
    cRedBits As Byte
    cRedShift As Byte
    cGreenBits As Byte
    cGreenShift As Byte
    cBlueBits As Byte
    cBlueShift As Byte
    cAlphaBits As Byte
    cAlphaShift As Byte
    cAccumBits As Byte
    cAccumRedBits As Byte
    cAccumGreenBits As Byte
    cAccumBlueBits As Byte
    cAccumAlphaBits As Byte
    cDepthBits As Byte
    cStencilBits As Byte
    cAuxBuffers As Byte
    iLayerType As Byte
    bReserved As Byte
    dwLayerMask As Long
    dwVisibleMask As Long
    dwDamageMask As Long
End Type

Public Type POINTFLOAT  '8 Bytes
        x As Single
        y As Single
End Type
Public Type GLYPHMETRICSFLOAT   '24 Bytes
        gmBlackBoxX As Single
        gmBlackBoxY As Single
        gmptGlyphOrigin As POINTFLOAT
        gmCellIncX As Single
        gmCellIncY As Single
End Type

Public Type EGLFONT 'ElementGL Font
    fName As String
    fSize As Long
    fWeight As Long
    fBaseList As Long
    fDeviation As Single
    fExtrusion As Single
    fMetrics(255) As GLYPHMETRICSFLOAT
    f3D As Boolean
    hFont As Long
End Type

Public GHDC As Long
Public GLDC As Long
Public EXTS As Long
Public font As EGLFONT
Private dm As DEVMODE
Private fontTexture As Long

Public Sub setFontTGA(textureNR As Long)
    Dim cx As Single, cy As Single
    Dim i As Long
    
    If Not textureNR = -1 Then
        If shaderIndices(textureNR) > -1 Then
            If shaderShouldBeParsed(shaderIndices(textureNR)) Then
                fontTexture = parsedShaders(shaderIndices(textureNR)).stages(1).textureID(0)
            End If
        End If
    End If
    
    font.fBaseList = glGenLists(255)
    
    For i = 0 To 255
        cx = (i Mod 16) / 16
        cy = (i / 16) / 16
        
        glNewList font.fBaseList + i, GL_COMPILE
            glBegin GL_QUADS
                glTexCoord2f cx, 1 - cy - 0.0625
                glVertex3f 0, 0, 0
                
                glTexCoord2f cx + 0.0625, 1 - cy - 0.0625
                glVertex3f 0, (10 * 16 / 600), 0
                
                glTexCoord2f cx + 0.0625, 1 - cy
                glVertex3f (10 * 16 / 800), (10 * 16 / 600), 0
                
                glTexCoord2f cx, 1 - cy
                glVertex3f (10 * 16 / 800), 0, 0
            glEnd
        glEndList
    Next i
End Sub

Public Sub glStart(bpp As Byte, zbuffer As Byte, stencil As Byte, hdc As Long)
    Dim pf As PIXELFORMATDESCRIPTOR
    Dim npf As Long
    Dim i As Long
        
    pf.nSize = Len(pf)
    pf.nVersion = 1
    pf.dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
    pf.iPixelType = PFD_TYPE_RGBA
    pf.cColorBits = bpp
    pf.cDepthBits = zbuffer
    pf.cStencilBits = stencil
    pf.iLayerType = PFD_MAIN_PLANE
    
    For i = 0 To 4
        npf = ChoosePixelFormat(hdc, pf)
        SetPixelFormat hdc, npf, pf
    
        GHDC = hdc
        GLDC = wglCreateContext(hdc)
        If Not GLDC = 0 Then Exit For
    Next i
    If GLDC = 0 Then Log.Log "elementGL", "glStart", "Unable to initialize opengl"
    wglMakeCurrent GHDC, GLDC
    
    'EXTS = extSetupExts()
End Sub

Public Sub glQuit()
    wglMakeCurrent 0, 0
    wglDeleteContext GLDC
End Sub

Public Sub glFlip()
    SwapBuffers GHDC
End Sub

Public Sub SetDisplayMode(ByVal Width As Long, ByVal Height As Long, ByVal bpp As Long, ByVal fullscreen As Boolean)
    Dim ret
    
    dm.dmSize = LenB(dm)
    dm.dmPelsHeight = Height
    dm.dmPelsWidth = Width
    dm.dmBitsPerPel = bpp
    dm.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    
    ret = ChangeDisplaySettingsA(dm, 0)
    
    If fullscreen Then ret = ChangeDisplaySettingsA(dm, CDS_FULLSCREEN)

    Log.Log "Graphics", "initAll", "SetDisplayMode successfull: " & CBool(ret = ElementGL.DISP_CHANGE_SUCCESSFUL)  'set the resolution
End Sub

Public Function GetDisplayMode(Height As Long, Width As Long, bpp As Long) As Long
    Dim DispDC As Long

    DispDC = CreateIC("DISPLAY", "", "", ByVal 0)

    Width = GetDeviceCaps(DispDC, 8)
    Height = GetDeviceCaps(DispDC, 10)
    bpp = GetDeviceCaps(DispDC, 12)

    DeleteDC DispDC
End Function

Public Function glFont(FontName As String, FontSize As Long, FontWeight As FontWeight) As EGLFONT
    glFont.f3D = False
    glFont.fName = FontName
    glFont.fSize = FontSize
    glFont.fWeight = FontWeight
    Erase glFont.fMetrics
    
    glFont.fBaseList = glGenLists(256)
    glFont.hFont = CreateFontA(FontSize, 0, 0, 0, FontWeight, False, False, False, ANSI_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS, 0, 0 Or DEFAULT_PITCH, FontName)
    
    SelectObject GHDC, glFont.hFont
    wglUseFontBitmapsA GHDC, 0, 256, glFont.fBaseList
End Function

Public Sub glFont3D(FontName As String, FontSize As Long, FontWeight As FontWeight, deviation As Single, extrusion As Single)
    font.f3D = True
    font.fName = FontName
    font.fSize = FontSize
    font.fWeight = FontWeight
    font.fDeviation = deviation
    font.fExtrusion = extrusion
    
    font.fBaseList = glGenLists(256)
    font.hFont = CreateFontA(FontSize, 0, 0, 0, FontWeight, False, False, False, ANSI_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS, 0, 0 Or DEFAULT_PITCH, FontName)
    
    SelectObject GHDC, font.hFont
    wglUseFontOutlinesA GHDC, 0, 256, font.fBaseList, deviation, extrusion, 1, font.fMetrics(0)
End Sub

Public Sub glDestroyFont()
    glDeleteLists font.fBaseList, 256
End Sub

Public Sub glTextOut(str As String, x As Single, y As Single, Optional thescale As Single = 1)
    If SHOW_TEXT <> 1 Then Exit Sub
    
    glPushMatrix
    glLoadIdentity
    'glColor3f 1, 0, 0
    glTranslatef x, y, -10
    
    glScalef thescale, thescale, 1
    
    useTexture fontTexture
    glPushAttrib GL_LIST_BIT
        glListBase font.fBaseList
        glCallLists Len(str), GL_UNSIGNED_BYTE, ByVal str
    glPopAttrib
    glLoadIdentity
    'glColor3f 1, 1, 1
    glPopMatrix
End Sub

Public Function glGetTextWidth(theStr As String, Optional thescale As Single = 1) As Single
    Dim i
    For i = 1 To Len(theStr)
        glGetTextWidth = glGetTextWidth + font.fMetrics(AscB(Mid(theStr, i, 1))).gmCellIncX
    Next i
    glGetTextWidth = glGetTextWidth * thescale
End Function

Public Function glGetTextHeight(theStr As String, Optional thescale As Single = 1) As Single
    Dim i
    Dim tmp As Single
    For i = 1 To Len(theStr)
        tmp = font.fMetrics(AscB(Mid(theStr, i, 1))).gmBlackBoxY
        If tmp > glGetTextHeight Then glGetTextHeight = tmp
    Next i
    glGetTextHeight = glGetTextHeight * thescale
End Function
