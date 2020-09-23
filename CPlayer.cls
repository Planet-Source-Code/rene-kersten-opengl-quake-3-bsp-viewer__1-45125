VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Enum player_events
    idle = 1
    walk = 2
    run = 4
    shoot = 8
    rotate = 16
End Enum

Private Const CAM_DISTANCE As Single = 75

Private model As New CMD3Controller
Private pos As vect

Private state As Long

Private notIdle As Boolean

Private viewtype As Long '1 = first person 2 = locked chase 3 = free chase

Friend Function getpos() As vect
    getpos = pos
End Function

Public Sub setstate(thestate As MD3State, Optional finishCurrent As Boolean = False)
    state = thestate
    model.setAnim state, True
End Sub

Public Sub setupAnimations()
    Dim r As Long
    
    If state And walk Then
        setstate LEGS_WALK
        state = walk
        
    ElseIf state And run Then
        setstate LEGS_RUN
        state = run
        
    ElseIf state And shoot Then
        r = Rnd
        If r = 0 Then
            setstate TORSO_ATTACK, True
        Else
            setstate TORSO_ATTACK2, True
        End If
        
    ElseIf state And rotate Then
        setstate LEGS_TURN, True
        
    Else '(state And idle) Or (state = 0) Then
        setstate LEGS_IDLE
        'state = idle
        
    End If
End Sub

Public Sub setupcam(theEvent As player_events)
    Dim tmppos As vect
    
    state = state Or theEvent
    
    If viewtype = 1 Then
        tmppos = pos
        tmppos.y = tmppos.y + 32
        set_cam_to_point tmppos
    ElseIf viewtype = 3 Or viewtype = 2 Then
        tmppos = pos
        tmppos.y = tmppos.y + 16
        calc_cam_around_point tmppos, CAM_DISTANCE
    End If
End Sub

Public Sub key_up()
    pos.y = pos.y + scrollspeed
    setupcam state
End Sub

Public Sub key_down()
    pos.y = pos.y - scrollspeed
    setupcam state
End Sub

Public Sub key_left()
    angleX = angleX - 2
    If angleX < 0 Then angleX = 358
    setupcam rotate
End Sub

Public Sub key_right()
    angleX = angleX + 2
    If angleX > 358 Then angleX = 0
    setupcam rotate
End Sub

Public Sub key_a()
    If viewtype = 3 Then
        pos.x = pos.x '- Sin((angleX - 90) / 180 * pi) * scrollspeed
        pos.z = pos.z '- scrollspeed '+ Cos((angleX - 90) / 180 * pi) * scrollspeed
    Else
        pos.x = pos.x + Sin((angleX - 90) / 180 * pi) * scrollspeed
        pos.z = pos.z - Cos((angleX - 90) / 180 * pi) * scrollspeed
    End If
    
    setupcam walk
End Sub

Public Sub key_d()
    If viewtype = 3 Then
        pos.x = pos.x 'Sin((angleX + 90) / 180 * pi) * scrollspeed
        pos.z = pos.z '+ scrollspeed '+ Cos((angleX + 90) / 180 * pi) * scrollspeed
    Else
        pos.x = pos.x + Sin((angleX + 90) / 180 * pi) * scrollspeed
        pos.z = pos.z - Cos((angleX + 90) / 180 * pi) * scrollspeed
    End If
    
    setupcam walk
End Sub

Public Sub key_s()
    Dim tmpangleX As Single
    
    tmpangleX = angleY
    If tmpangleX < 0 Then tmpangleX = -tmpangleX
    
    If viewtype = 3 Then
        pos.x = pos.x '- scrollspeed '(Sin(angleX / 180 * pi) * (1 - Sin(tmpangleX / 180 * pi))) * scrollspeed
        pos.z = pos.z '- (Cos(angleX / 180 * pi) * (1 - Sin(tmpangleX / 180 * pi))) * scrollspeed
        pos.y = pos.y '- Sin(angleY / 180 * pi) * scrollspeed
    Else 'If viewtype = 2 Then
        pos.x = pos.x - Sin(angleX / 180 * pi) * scrollspeed
        pos.z = pos.z + Cos(angleX / 180 * pi) * scrollspeed
    'ElseIf viewtype = 1 Then
    '    pos.x = pos.x - (Sin(angleX / 180 * pi) * (1 - Sin(tmpangleX / 180 * pi))) * scrollspeed
    '    pos.z = pos.z + (Cos(angleX / 180 * pi) * (1 - Sin(tmpangleX / 180 * pi))) * scrollspeed
    '    'pos.y = pos.y + Sin(angleY / 180 * pi) * scrollspeed
    End If
    
    setupcam run
End Sub

Public Sub key_w()
    Dim tmpangleX As Single
    
    tmpangleX = angleY
    If tmpangleX < 0 Then tmpangleX = -tmpangleX
    
    If viewtype = 3 Then
        pos.x = pos.x '+ scrollspeed '(Sin(angleX / 180 * pi) * (1 - Sin(tmpangleX / 180 * pi))) * scrollspeed
        pos.z = pos.z '(Cos(angleX / 180 * pi) * (1 - Sin(tmpangleX / 180 * pi))) * scrollspeed
        pos.y = pos.y '+ Sin(angleY / 180 * pi) * scrollspeed
    Else 'If viewtype = 2 Then
        pos.x = pos.x + Sin(angleX / 180 * pi) * scrollspeed
        pos.z = pos.z - Cos(angleX / 180 * pi) * scrollspeed
    'ElseIf viewtype = 1 Then
    '    pos.x = pos.x + (Sin(angleX / 180 * pi) * (1 - Sin(tmpangleX / 180 * pi))) * scrollspeed
    '    pos.z = pos.z - (Cos(angleX / 180 * pi) * (1 - Sin(tmpangleX / 180 * pi))) * scrollspeed
    '    'pos.y = pos.y - Sin(angleY / 180 * pi) * scrollspeed
    End If
    
    setupcam run
End Sub

Friend Sub setpos(p As vect)
    pos = p
    setupcam state
End Sub

Public Sub change_view_type()
    If viewtype = 2 Then
        viewtype = 1 '3
    ElseIf viewtype = 1 Then
        viewtype = 2
    Else
        viewtype = 1
    End If
    
    setupcam state
End Sub

Public Sub initialize()
    model.load "models\players\visor\"
    setstate LEGS_IDLE
    setstate TORSO_STAND
    
    viewtype = 2
End Sub

Public Sub update()
    setupAnimations
    draw
    state = 0
End Sub

Public Sub draw()
    If viewtype = 1 Then Exit Sub
    
    objShaderlib.resetStates
    glDisable GL_BLEND
    glEnable GL_TEXTURE_2D
    
    glPushMatrix
        glTranslatef pos.x, pos.y, pos.z
        If viewtype = 2 Then
            glRotatef -angleX + 90, 0, 1, 0
        End If
        
        model.update angleY / 2
    glPopMatrix
End Sub

Public Sub quit()
    model.quit
End Sub
