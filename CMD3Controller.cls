VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMD3Controller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Enum MD3State
    BOTH_DEATH1 = 0
    BOTH_DEAD1
    BOTH_DEATH2
    BOTH_DEAD2
    BOTH_DEATH3
    BOTH_DEAD3
    
    TORSO_GESTURE
    TORSO_ATTACK
    TORSO_ATTACK2
    TORSO_DROP
    TORSO_RAISE
    TORSO_STAND
    TORSO_STAND2
    
    LEGS_WALKCR
    LEGS_WALK
    LEGS_RUN
    LEGS_BACK
    LEGS_SWIM
    LEGS_JUMP
    LEGS_LAND
    LEGS_JUMPB
    LEGS_LANDB
    LEGS_IDLE
    LEGS_IDLECR
    LEGS_TURN
End Enum

Private Type anim_table_element
    start As Long
    stop As Long
    loop As Boolean
    fps As Long
    part As Byte
End Type

Private upper As New CMD3
Private lower As New CMD3
Private head As New CMD3

Private numanims As Long
Private anims(25) As anim_table_element

Public Enum modelPart
    ANIM_TORSO = 1
    ANIM_LEGS = 2
    ANIM_BOTH = 3
End Enum

Private legsAnim As Long
Private torsoAnim As Long

Public Function animationDone(part As modelPart) As Boolean
    If part = ANIM_BOTH Then
        animationDone = upper.animDone And lower.animDone
    ElseIf part = ANIM_LEGS Then
        animationDone = lower.animDone
    ElseIf part = ANIM_TORSO Then
        animationDone = upper.animDone
    End If
End Function

Private Sub loadAnimTable(filename As String)
    Dim filenr As Long
    Dim l As String
    Dim parts() As String
    Dim i As Long
    
    i = -1
    filenr = FreeFile
    Open filename For Input As filenr
        Do Until EOF(filenr)
            Line Input #filenr, l
            
            l = Trim(l)
            If InStr(1, l, "//") > 0 Then l = LEFT(l, InStr(1, l, "//") - 1)
            If l <> "" Then
                parts = Split(l, Chr(9))
                
                If Not parts(0) = "sex" And Not LEFT(parts(0), Len("headoffset")) = "headoffset" And Not LEFT(parts(0), Len("footsteps")) = "footsteps" Then
                    i = i + 1
                    anims(i).start = CLng(parts(0))
                    anims(i).stop = anims(i).start + CLng(parts(1))
                    anims(i).loop = (CLng(parts(2)) > 0)
                    anims(i).fps = CLng(parts(3))
                End If
            End If
        Loop
    Close filenr
    
    For i = BOTH_DEATH1 To BOTH_DEAD3
        anims(i).part = ANIM_BOTH
    Next i
    For i = TORSO_GESTURE To TORSO_STAND2
        anims(i).part = ANIM_TORSO
    Next i
    
    Dim offset As Long
    offset = anims(LEGS_WALKCR).start - anims(TORSO_GESTURE).start
    For i = LEGS_WALKCR To LEGS_TURN
        anims(i).part = ANIM_LEGS
        anims(i).start = anims(i).start - offset
        anims(i).stop = anims(i).stop - offset
    Next i
End Sub

Public Sub setAnim(anim As MD3State, Optional shouldloop As Boolean)
    If anims(anim).part = ANIM_BOTH Then
        If Not torsoAnim = anim Then
            torsoAnim = anim
            legsAnim = anim
            upper.setFrame anims(anim).start, anims(anim).stop, anims(anim).fps, anims(anim).loop Or shouldloop
            lower.setFrame anims(anim).start, anims(anim).stop, anims(anim).fps, anims(anim).loop Or shouldloop
        End If
    ElseIf anims(anim).part = ANIM_LEGS Then
        If Not legsAnim = anim Then
            legsAnim = anim
            lower.setFrame anims(anim).start, anims(anim).stop, anims(anim).fps, anims(anim).loop Or shouldloop
        End If
    ElseIf anims(anim).part = ANIM_TORSO Then
        If Not torsoAnim = anim Then
            torsoAnim = anim
            upper.setFrame anims(anim).start, anims(anim).stop, anims(anim).fps, anims(anim).loop Or shouldloop
        End If
    End If
End Sub

Public Sub load(model As String)
    model = Replace(model, "/", "\")
    
    If Not InStr(1, model, ":") > 0 Then
        If LEFT(model, 1) = "\" Then
            model = App.path & model
        Else: model = App.path & "\" & model
        End If
    End If
    
    If RIGHT(model, 1) <> "\" Then model = model & "\"
    
    upper.load model & "upper.md3", True
    lower.load model & "lower.md3", True
    head.load model & "head.md3", False
    loadAnimTable model & "animation.cfg"
    
    setAnim LEGS_IDLE
    setAnim TORSO_STAND
End Sub

Public Sub update(torsoRotate As Single)
    Dim upperPos As vect
    Dim headPos As vect
    Dim tmp As vect
    Dim matrix(3, 3) As Single
    Dim tag1 As MD3Tag
    Dim tag2 As MD3Tag
    
    glPushMatrix
    
    'update all
    lower.update
    upper.update
    head.update
    
    glRotatef -90, 1, 0, 0
    
    'draw lower
    lower.draw
    
    'align upper with lower
    tag1 = lower.getTagOrigin("tag_torso")
    
    matrix(0, 0) = 1
    matrix(1, 1) = 1
    matrix(2, 2) = 1
    matrix(3, 3) = 1
    
    matrix(0, 0) = tag1.axis(0).x
    matrix(1, 0) = tag1.axis(0).y
    matrix(2, 0) = tag1.axis(0).z
    
    matrix(0, 1) = tag1.axis(1).x
    matrix(1, 1) = tag1.axis(1).y
    matrix(2, 1) = tag1.axis(1).z
    
    matrix(0, 2) = tag1.axis(2).x
    matrix(1, 2) = tag1.axis(2).y
    matrix(2, 2) = tag1.axis(2).z
    
    matrix(0, 3) = tag1.origin.x
    matrix(1, 3) = tag1.origin.y
    matrix(2, 3) = tag1.origin.z
        
    glMultMatrixf matrix(0, 0)
    
    glRotatef torsoRotate, 0, 1, 0
    
    'draw upper
    upper.draw
    
    'align head with upper
    tag1 = upper.getTagOrigin("tag_head")
    
    matrix(0, 0) = 1
    matrix(1, 1) = 1
    matrix(2, 2) = 1
    matrix(3, 3) = 1
    
    matrix(0, 0) = tag1.axis(0).x
    matrix(1, 0) = tag1.axis(0).y
    matrix(2, 0) = tag1.axis(0).z
    
    matrix(0, 1) = tag1.axis(1).x
    matrix(1, 1) = tag1.axis(1).y
    matrix(2, 1) = tag1.axis(1).z
    
    matrix(0, 2) = tag1.axis(2).x
    matrix(1, 2) = tag1.axis(2).y
    matrix(2, 2) = tag1.axis(2).z
    
    matrix(0, 3) = tag1.origin.x
    matrix(1, 3) = tag1.origin.y
    matrix(2, 3) = tag1.origin.z
        
    glMultMatrixf matrix(0, 0)
    
    'draw head
    head.draw
    
    glPopMatrix
End Sub

Public Sub quit()
    
End Sub
