Attribute VB_Name = "mainloop"
    
Public stopGame As Boolean
Public inGame As Boolean
Public initialized As Boolean

Public posX As Single
Public posY As Single
Public pic1 As Long

Public settings2 As animationSettings
Public settings As animationSettings

Sub initialize(fullscreen As Boolean)
    Log.initializeLog
    
    initialized = Graphics.initAll(fullscreen, 800, 600, 16)
    If initialized = False Then GoTo errhandler
    
    Exit Sub
    
errhandler:
    Log.Log "mainloop", "main", "Unable to initialize"
    endGame
End Sub

Sub main()
    Dim tmpGraphic As New CPicture
    Dim angle As Single
    Dim picNr As Long
    
    'Form1.Show
    initialize False
    initializeGui
    
    Log.Log "mainloop", "main", "Engine started"
    
    If Not initialized Then
        Log.Log "mainloop", "main", "Unable to initialize"
        endGame
    End If
    
    Dim tmpRects(1) As RECT
    
    tmpRects(0) = makeRect(0, 0, 0.5, 0.5)
    tmpRects(1) = makeRect(0.5, 0.5, 1, 1)
    With tmpGraphic
        .setAlpha True, 0.5
        .setColor 1, 1, 1
        .setAnimationTexture App.Path & "\test.bmp", tmpRects, True
        .setVerticleRect 0, 0, 1, 1
        settings.state = ANIM_PLAY
        settings.fps = 1
        settings.loop = True
        settings2.state = ANIM_PLAY
        settings2.loop = True
        settings2.fps = 0.5
    End With
    
    Do Until stopGame
        DoEvents
        
        tmpGraphic.setAlpha True, 1
        tmpGraphic.draw 2, 2, settings2, 0, 0, 1, 40
        
        tmpGraphic.setAlpha True, 0.5
        tmpGraphic.draw posX, posY, settings
        
        glLoadIdentity
        glDisable GL_TEXTURE_2D
        glColor4f 1, 1, 1, 1
        glBegin GL_LINES
            glVertex3f 0, 0, 0
            glVertex3f 3, 3, 0
        glEnd
        glEnable GL_TEXTURE_2D
        
        checkButtons
        drawall
        nextFrame
    Loop
    
    deleteGui
    endGame
End Sub

Public Sub endGame()
    killAll
    Unload Form1
    Log.Log "mainloop", "endGame", "Engine uninitialized"
    End
End Sub
