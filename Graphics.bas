Attribute VB_Name = "Graphics"

Option Explicit

Public SHOW_TEXT As Long

Public Type vect
    x As Single
    y As Single
    z As Single
End Type

Public Type vect2
    x As Single
    y As Single
End Type

Public Type color
    red As Single
    green As Single
    blue As Single
End Type

Public Type animationFrame
    coords As RECT
    indexNumber As Long
End Type

Public Type picType
    vertexRect As RECT
    
    numTextures As Long
    textures() As Long
    frames() As animationFrame
    
    drawcolor As color
    
    useAlpha As Boolean
    alpha As Single
    
    numFrames As Long
End Type

Private oldTime As Long
Private newtime As Long
Private fpscounter As Long
Private fps As Long

Public Width As Long
Public Height As Long
Public bpp As Byte

Private oldwidth As Long
Private oldheight As Long
Private oldbpp As Long

Private realwidth As Single
Private realheight As Single

Public frequency As Currency
Private fullscreen As Boolean

'Public objEngine As New CEngine

Public Function makeRect(LEFT As Single, TOP As Single, RIGHT As Single, BOTTOM As Single) As RECT
    makeRect.LEFT = LEFT
    makeRect.RIGHT = RIGHT
    makeRect.TOP = TOP
    makeRect.BOTTOM = BOTTOM
End Function

Public Function initAll(fullscreenFlag As Boolean, desiredWidth As Long, desiredHeight As Long, desiredBpp As Byte) As Boolean
    On Error GoTo errhandler
    
    Log.Log "Graphics", "initAll", "Initializing"
    
    fullscreen = fullscreenFlag
    
    initAll = True
    
    Width = desiredWidth
    Height = desiredHeight
    bpp = desiredBpp
    
    If fullscreen Then
        GetDisplayMode oldheight, oldwidth, oldbpp 'get the displaymode
        SetDisplayMode Width, Height, bpp, True
    End If
    
    glStart CSng(bpp), 32, 0, Form1.hdc
    Form1.Show
    
    Form1.LEFT = 0
    Form1.TOP = 0
    
    Form1.Width = Width * Screen.TwipsPerPixelX
    Form1.Height = Height * Screen.TwipsPerPixelY
    
    QueryPerformanceFrequency frequency
    
    glShadeModel GL_SMOOTH               ' Enables Smooth Shading
    
    glClearColor 0, 0, 0, 0          ' Black Background
    'glClearColor 0, 0.5, 1, 0          ' blue Background
    
    glEnable GL_DEPTH_TEST
    glDepthFunc GL_LEQUAL
    glClearDepth 1000
    
    Form1.Show
    
    glHint GL_PERSPECTIVE_CORRECTION_HINT, GL_NICEST    ' Really Nice Perspective Calculations
    
    'glFont3D "arial", 8, FWE_BOLD, 0, 0
    
    
    Exit Function
    
errhandler:
    initAll = False
End Function

Public Function resize()
    If Height = 0 Then              ' Prevent A Divide By Zero By
        Height = 1                  ' Making Height Equal One
    End If
    glViewport 0, 0, Form1.ScaleWidth, Form1.ScaleHeight ' Reset The Current Viewport
    glMatrixMode GL_PROJECTION       ' Select The Projection Matrix
    glLoadIdentity                  ' Reset The Projection Matrix
    
    ' Calculate The Aspect Ratio Of The Window
    gluPerspective 70, 16 / 12, 10, 5000
    realwidth = 16
    realheight = 12
    
    glMatrixMode GL_MODELVIEW        ' Select The Modelview Matrix
    glLoadIdentity                  ' Reset The Modelview Matrix
End Function

Public Function nextFrame()
    Dim result As Long
    Dim x As Long
    
    newtime = GetTickCount
    result = newtime - oldTime
    
    If x = 0 Then
        x = 1
    Else
        x = x
    End If
    
    If result > 1000 Then
        oldTime = newtime
        fps = fpscounter
        fpscounter = 0
    End If
    
    fpscounter = fpscounter + 1
    
    glDisable GL_TEXTURE_2D
    
    glColor3f 1, 1, 1
    'Translatef glGetTextWidth(right(str(fps), Len(str(fps)) - 1)), glGetTextHeight(right(str(fps), Len(str(fps)) - 1)), 0, 0
    'glRotatef 180, 0, 0, 1
    
    'Translatef glGetTextWidth(right(str(fps), Len(str(fps)) - 1)), -2, 0, 3
    
    Dim fpsstr As String
    fpsstr = "FPS: " & str(RIGHT(str(fps), Len(str(fps)) - 1))
    ElementGL.glTextOut fpsstr, -5, -0.5, 0.3 '"FPS: " & fps
    
    glLoadIdentity
    
    ElementGL.glFlip
    glClear GL_DEPTH_BUFFER_BIT Or GL_COLOR_BUFFER_BIT
End Function

Public Function killAll()
    If fullscreen Then SetDisplayMode oldwidth, oldheight, oldbpp, False 'switch to old display settings
    
    ElementGL.glQuit
End Function
