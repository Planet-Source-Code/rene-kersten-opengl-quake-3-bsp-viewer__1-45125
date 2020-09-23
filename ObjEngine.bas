Attribute VB_Name = "ObjEngine"
Option Explicit

Public running As Boolean

Public lighting As Boolean
Public texturesOn As Boolean
Public drawType As Byte

Public mousespeed As Single
Public scrollspeed As Single
Public cursorOn As Boolean

Public velocity As Single
Private oldTime As Long
Private time As Long

Public printscreen As Boolean

Public objInputEngine As New CInputEngine

Private Sub checkCursor()
    Dim newCursorOn As Boolean
    newCursorOn = (getCVar("mouseon") = "1")
    
    If cursorOn = newCursorOn Then Exit Sub
    
    cursorOn = newCursorOn
    
    If cursorOn Then
        Form1.MousePointer = 1
        ShowCursor 1
        cursorOn = True
    Else
        Form1.MousePointer = 4
        ShowCursor 0
        cursorOn = False
    End If
End Sub

Private Sub mouselook()
    Dim diffX As Long, diffY As Long
    
    If mouseX = 400 And mouseY = 300 Then Exit Sub
    
    diffX = mouseX - 400
    diffY = mouseY - 300
    
    angleX = angleX + CLng(diffX * mousespeed)
    angleY = angleY + CLng(diffY * mousespeed)
    
    If angleY < -90 Then angleY = -90
    If angleY > 90 Then angleY = 90
    
    If angleX < 0 Then angleX = 360
    If angleX > 360 Then angleX = 0
    
    SetCursorPos 400, 300
    
    player1.setupcam rotate
End Sub

Public Function getpos() As vect
    getpos = player1.getpos
End Function

Public Sub setpos(pos As vect)
    player1.setpos pos
End Sub

Public Sub initialize()
    Randomize
    
    Set Form1.MouseIcon = Nothing
    
    Log.initializeLog
    running = Graphics.initAll(False, 800, 600, 24)
    Form1.Width = 806 * Screen.TwipsPerPixelX
    Form1.Height = 625 * Screen.TwipsPerPixelY
    Graphics.resize
    
    initTextureLib
    init_shaderlib
    loadAllShaders
    
    init_console
    
    objBSP.BSP_init
    
    init_cvars
    loadCFG App.path & "\autoexec.cfg"
    
    player1.initialize
    
    'Dim fonttex As Long
    'fonttex = openQuakeTexture("gfx/2d/bigchars")
    
    'setFontTGA fonttex
    glFont3D "arial", 24, FWE_NORMAL, 0, 0
    
    objInputEngine.initialize
    
    Log.Log "", "", ""
    Log.Log "", "", "initialized and running"
    textureManager.logStatus
    
    texturesOn = True
    lighting = True
    
    mousespeed = 0.3
    scrollspeed = 3
    cursorOn = True
    
    velocity = 0.1
    
    oldTime = GetTickCount
    
    SetCursorPos 400, 300
End Sub

Public Sub start()
    initialize
    
    Do While running
        DoEvents
        
        checkCursor
        
        Graphics.nextFrame
        
        glLoadIdentity
        
        mouselook
        
        objShaderlib.setupShaders
        
        time = GetTickCount
        scrollspeed = velocity * (time - oldTime)
        oldTime = time
        
        Select Case drawType
            Case 0 'normal
                glPolygonMode GL_FRONT_AND_BACK, GL_FILL
            Case 1 'wireframe
                glPolygonMode GL_FRONT_AND_BACK, GL_LINE
            Case 2 'dots
                glPolygonMode GL_FRONT_AND_BACK, GL_POINT
            Case 3 'silhouette
                glPolygonMode GL_FRONT_AND_BACK, GLU_SILHOUETTE
        End Select
        
        glCullFace GL_FRONT
        glEnable GL_CULL_FACE
        glEnable GL_DEPTH_TEST
        
        setup_camera
        objBSP.drawMap 'this needs to be done first cause of the skybox!
        
        print_camera_position
        player1.update
        
        glPolygonMode GL_FRONT_AND_BACK, GL_FILL
        
        glDisable GL_CULL_FACE
        glDisable GL_DEPTH_TEST
        glDisable GL_TEXTURE_2D
        glLoadIdentity
        
        If getCVar("draw_crosshair") = "1" Then
            glPushMatrix
            glLoadIdentity
            glTranslatef 0, 0, -10
            glDisable GL_TEXTURE_2D
            glDisable GL_BLEND
            glDisable GL_ALPHA_TEST
            
            glScalef 0.5, 0.5, 1
            
            glColor3f 0, 1, 0
            glBegin GL_LINES
                glVertex3f -1, 0, 0
                glVertex3f 0, 0, 0
                
                glVertex3f 1, 0, 0
                glVertex3f 0, 0, 0
                
                glVertex3f 0, -1, 0
                glVertex3f 0, 0, 0
                
                glVertex3f 0, 1, 0
                glVertex3f 0, 0, 0
            glEnd
            glColor3f 1, 1, 1
            glPopMatrix
        End If
        
        console_draw
        
        If Not console_enabled Then getInput
    Loop
    
    quit
End Sub

Public Sub quit()
    ShowCursor True
    Form1.MousePointer = 0
    
    objBSP.BSP_quit
    objInputEngine.quit
    
    Graphics.killAll
    running = False
    End
End Sub
