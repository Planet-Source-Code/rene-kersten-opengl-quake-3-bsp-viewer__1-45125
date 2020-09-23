Attribute VB_Name = "Frustrum"
Option Explicit

Private Const a = 0
Private Const b = 1
Private Const c = 2
Private Const d = 3

Private Const RIGHT = 0
Private Const LEFT = 1
Private Const BOTTOM = 2
Private Const TOP = 3
Private Const BACK = 4
Private Const FRONT = 5

Private planes(5, 3)

Private Sub normalizePlane(side As Single)
    Dim m As Single
    m = Sqr(planes(side, a) * planes(side, a) + planes(side, b) * planes(side, b) + planes(side, c) * planes(side, c))
    
    planes(side, a) = planes(side, a) / m
    planes(side, b) = planes(side, b) / m
    planes(side, c) = planes(side, c) / m
    planes(side, d) = planes(side, d) / m
End Sub

Public Sub getfrustrum()
    Dim prj(15) As Single
    Dim mdl(15) As Single
    Dim clip(15) As Single
    
    glGetFloatv GL_PROJECTION_MATRIX, prj(0)
    glGetFloatv GL_MODELVIEW_MATRIX, mdl(0)
    
    clip(0) = mdl(0) * prj(0) + mdl(1) * prj(4) + mdl(2) * prj(8) + mdl(3) * prj(12)
    clip(1) = mdl(0) * prj(1) + mdl(1) * prj(5) + mdl(2) * prj(9) + mdl(3) * prj(13)
    clip(2) = mdl(0) * prj(2) + mdl(1) * prj(6) + mdl(2) * prj(10) + mdl(3) * prj(14)
    clip(3) = mdl(0) * prj(3) + mdl(1) * prj(7) + mdl(2) * prj(11) + mdl(3) * prj(15)
    
    clip(4) = mdl(4) * prj(0) + mdl(5) * prj(4) + mdl(6) * prj(8) + mdl(7) * prj(12)
    clip(5) = mdl(4) * prj(1) + mdl(5) * prj(5) + mdl(6) * prj(9) + mdl(7) * prj(13)
    clip(6) = mdl(4) * prj(2) + mdl(5) * prj(6) + mdl(6) * prj(10) + mdl(7) * prj(14)
    clip(7) = mdl(4) * prj(3) + mdl(5) * prj(7) + mdl(6) * prj(11) + mdl(7) * prj(15)

    clip(8) = mdl(8) * prj(0) + mdl(9) * prj(4) + mdl(10) * prj(8) + mdl(11) * prj(12)
    clip(9) = mdl(8) * prj(1) + mdl(9) * prj(5) + mdl(10) * prj(9) + mdl(11) * prj(13)
    clip(10) = mdl(8) * prj(2) + mdl(9) * prj(6) + mdl(10) * prj(10) + mdl(11) * prj(14)
    clip(11) = mdl(8) * prj(3) + mdl(9) * prj(7) + mdl(10) * prj(11) + mdl(11) * prj(15)

    clip(12) = mdl(12) * prj(0) + mdl(13) * prj(4) + mdl(14) * prj(8) + mdl(15) * prj(12)
    clip(13) = mdl(12) * prj(1) + mdl(13) * prj(5) + mdl(14) * prj(9) + mdl(15) * prj(13)
    clip(14) = mdl(12) * prj(2) + mdl(13) * prj(6) + mdl(14) * prj(10) + mdl(15) * prj(14)
    clip(15) = mdl(12) * prj(3) + mdl(13) * prj(7) + mdl(14) * prj(11) + mdl(15) * prj(15)
    
    planes(RIGHT, a) = clip(3) - clip(0)
    planes(RIGHT, b) = clip(7) - clip(4)
    planes(RIGHT, c) = clip(11) - clip(8)
    planes(RIGHT, d) = clip(15) - clip(12)
    normalizePlane RIGHT
    
    planes(LEFT, a) = clip(3) + clip(0)
    planes(LEFT, b) = clip(7) + clip(4)
    planes(LEFT, c) = clip(11) + clip(8)
    planes(LEFT, d) = clip(15) + clip(12)
    normalizePlane LEFT
    
    planes(BOTTOM, a) = clip(3) + clip(1)
    planes(BOTTOM, b) = clip(7) + clip(5)
    planes(BOTTOM, c) = clip(11) + clip(9)
    planes(BOTTOM, d) = clip(15) + clip(13)
    normalizePlane BOTTOM
    
    planes(TOP, a) = clip(3) - clip(1)
    planes(TOP, b) = clip(7) - clip(5)
    planes(TOP, c) = clip(11) - clip(9)
    planes(TOP, d) = clip(15) - clip(13)
    normalizePlane TOP
    
    planes(BACK, a) = clip(3) + clip(2)
    planes(BACK, b) = clip(7) + clip(6)
    planes(BACK, c) = clip(11) + clip(10)
    planes(BACK, d) = clip(15) + clip(14)
    normalizePlane BACK
    
    planes(FRONT, a) = clip(3) - clip(2)
    planes(FRONT, b) = clip(7) - clip(6)
    planes(FRONT, c) = clip(11) - clip(10)
    planes(FRONT, d) = clip(15) - clip(14)
    normalizePlane FRONT
End Sub

Public Function boxInFrustrum(x As Single, y As Single, z As Single, x2 As Single, y2 As Single, z2 As Single) As Boolean
    Dim i As Long
    
    For i = 0 To 5
        If (planes(i, 0) * x2 + planes(i, 1) * y2 + planes(i, 2) * z2 + planes(i, 3) > 0) Then GoTo NextSide:
        If (planes(i, 0) * x + planes(i, 1) * y2 + planes(i, 2) * z2 + planes(i, 3) > 0) Then GoTo NextSide:
        If (planes(i, 0) * x2 + planes(i, 1) * y + planes(i, 2) * z2 + planes(i, 3) > 0) Then GoTo NextSide:
        If (planes(i, 0) * x + planes(i, 1) * y + planes(i, 2) * z2 + planes(i, 3) > 0) Then GoTo NextSide:
        If (planes(i, 0) * x2 + planes(i, 1) * y2 + planes(i, 2) * z + planes(i, 3) > 0) Then GoTo NextSide:
        If (planes(i, 0) * x + planes(i, 1) * y2 + planes(i, 2) * z + planes(i, 3) > 0) Then GoTo NextSide:
        If (planes(i, 0) * x2 + planes(i, 1) * y + planes(i, 2) * z + planes(i, 3) > 0) Then GoTo NextSide:
        If (planes(i, 0) * x + planes(i, 1) * y + planes(i, 2) * z + planes(i, 3) > 0) Then GoTo NextSide:
        boxInFrustrum = False: Exit Function
NextSide:
    Next i
    
    boxInFrustrum = True
End Function

Public Function cubeInFrustrum(x As Single, y As Single, z As Single, size As Single) As Boolean

End Function

Public Function sphereInFrustrum(x As Single, y As Single, z As Single, radius As Single) As Boolean
    
End Function

Public Function pointInFrustrum(x As Single, y As Single, z As Single) As Boolean
    
End Function
