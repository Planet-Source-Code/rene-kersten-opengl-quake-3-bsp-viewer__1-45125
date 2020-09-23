Attribute VB_Name = "Math3d"

Public Type matrix
    m(15) As Single
End Type

Public Type quaternion
    quat(3) As Single
End Type

Public Type vector
    pos(3) As Single
End Type

Public Const pi = 3.14159265358979

'quaternion code

Public Function quat_setAngles(angles() As Single) As quaternion
    Dim sr As Double, sp As Double, sy As Double, cr As Double, cp As Double, cy As Double
    Dim angle As Single
    
    angle = angles(2)
    sy = Sin(angle)
    cy = Cos(angle)
    angle = angles(1)
    sp = Sin(angle)
    cp = Cos(angle)
    angle = angles(0)
    sr = Sin(angle)
    cr = Cos(angle)
    
    Dim crcp As Double
    Dim srsp As Double
    crcp = cr * cp
    srsp = sr * sp
    
    quat_setAngles.quat(0) = sr * cp * cy - cr * sp * sy
    quat_setAngles.quat(1) = cr * sp * cy + sr * cp * sy
    quat_setAngles.quat(2) = crcp * sy - srsp * cy
    quat_setAngles.quat(3) = crcp * sy + srsp * cy
End Function

Public Function quat_slerp(quat1 As quaternion, quat2 As quaternion, interp As Single) As quaternion
    Dim i As Long
    Dim a As Single
    Dim b As Single
    
    For i = 0 To 3
        a = a + (quat1.quat(i) - quat2.quat(i)) * (quat1.quat(i) - quat2.quat(i))
        a = a + (quat1.quat(i) + quat2.quat(i)) * (quat1.quat(i) + quat2.quat(i))
    Next i
    
    If a > b Then quat_inverse quat2
    
    Dim cosom As Single
    cosom = quat1.quat(0) * quat2.quat(0) + quat1.quat(1) * quat2.quat(1) + quat1.quat(2) * quat2.quat(2) + quat1.quat(3) * quat2.quat(3)
    
    Dim omega As Double, sinom As Double
    Dim sclq1 As Double, sclq2 As Double
    If 1 + cosom > 0.00000001 Then
        If 1 - cosom > 0.00000001 Then
            omega = Atn(-cosom / Sqr(-cosom * cosom + 1)) + 2 * Atn(1)
            sinom = Sin(omega)
            sclq1 = Sin((1 - interp) * omega) / sinom
            sclq2 = Sin(interp * omega) / sinom
        Else
            sclq1 = 1 - interp
            sclq2 = interp
        End If
        
        For i = 0 To 3
            quat_slerp.quat(i) = sclq1 * quat1.quat(i) + sclq2 * quat2.quat(i)
        Next i
    Else
        quat_slerp.quat(0) = -quat1.quat(1)
        quat_slerp.quat(1) = quat1.quat(0)
        quat_slerp.quat(2) = -quat1.quat(3)
        quat_slerp.quat(3) = quat1.quat(2)
        
        sclq1 = Sin((1 - interp) * 0.5 * pi)
        sclq2 = Sin(interp * 0.5 * pi)
        
        For i = 0 To 3
            quat_slerp.quat(i) = sclq1 * quat1.quat(i) + sclq2 * quat_slerp.quat(i)
        Next i
    End If
End Function

Public Sub quat_inverse(quat As quaternion)
    quat.quat(0) = -quat.quat(0)
    quat.quat(1) = -quat.quat(1)
    quat.quat(2) = -quat.quat(2)
    quat.quat(3) = -quat.quat(3)
End Sub

'vertex code

Public Sub vect_transform(v As vector, m As matrix)
    Dim vect(3) As Double
    vect(0) = v.pos(0) * m.m(0) + v.pos(1) * m.m(4) + v.pos(2) * m.m(8) + v.pos(3) * m.m(12)
    vect(1) = v.pos(0) * m.m(1) + v.pos(1) * m.m(5) + v.pos(2) * m.m(9) + v.pos(3) * m.m(13)
    vect(2) = v.pos(0) * m.m(2) + v.pos(1) * m.m(6) + v.pos(2) * m.m(10) + v.pos(3) * m.m(14)
    vect(3) = v.pos(0) * m.m(3) + v.pos(1) * m.m(7) + v.pos(2) * m.m(11) + v.pos(3) * m.m(15)
    
    v.pos(0) = vect(0)
    v.pos(1) = vect(1)
    v.pos(2) = vect(2)
    v.pos(3) = vect(3)
End Sub

Public Sub vect_transform3(vect As vector, mat As matrix)
    v.pos(0) = v.pos(0) * m.m(0) + v.pos(1) * m.m(4) + v.pos(2) * m.m(8) + v.pos(3) * m.m(12)
    v.pos(1) = v.pos(0) * m.m(1) + v.pos(1) * m.m(5) + v.pos(2) * m.m(9) + v.pos(3) * m.m(13)
    v.pos(2) = v.pos(0) * m.m(2) + v.pos(1) * m.m(6) + v.pos(2) * m.m(10) + v.pos(3) * m.m(14)
    v.pos(3) = 1
End Sub

Public Sub vect_add(vect1 As vector, vect2 As vector)
    vect.pos(0) = vect.pos(0) + vect2.pos(0)
    vect.pos(1) = vect.pos(1) + vect2.pos(1)
    vect.pos(2) = vect.pos(2) + vect2.pos(2)
    vect.pos(3) = vect.pos(3) + vect2.pos(3)
End Sub

Public Sub vect_reset(vect As vector)
    vect.pos(0) = 0
    vect.pos(1) = 0
    vect.pos(2) = 0
    vect.pos(3) = 1
End Sub

Public Sub vect_normalize(vect As vector)
    Dim length As Single
    length = vect_length(vect)
    
    vect.pos(0) = vect.pos(0) / length
    vect.pos(1) = vect.pos(1) / length
    vect.pos(2) = vect.pos(2) / length
End Sub

Public Function vect_length(vect As vector) As Single
    vect_length = vect.pos(0) * vect.pos(0) + vect.pos(1) * vect.pos(1) + vect.pos(2) * vect.pos(2)
End Function

'matrix code

Public Sub mat_loadIdentity(m As matrix)
    m.m(0) = 1
    m.m(1) = 0
    m.m(2) = 0
    m.m(3) = 0
    m.m(4) = 0
    m.m(5) = 1
    m.m(6) = 0
    m.m(7) = 0
    m.m(8) = 0
    m.m(9) = 0
    m.m(10) = 1
    m.m(11) = 0
    m.m(12) = 0
    m.m(13) = 0
    m.m(14) = 0
    m.m(15) = 1
End Sub

Public Sub mat_postMultiply(m1 As matrix, m2 As matrix)
    m1.m(0) = m1.m(0) * m2.m(0) + m1.m(4) * m2.m(1) + m1.m(8) * m2.m(2)
    m1.m(1) = m1.m(1) * m2.m(0) + m1.m(5) * m2.m(1) + m1.m(9) * m2.m(2)
    m1.m(2) = m1.m(2) * m2.m(0) + m1.m(6) * m2.m(1) + m1.m(10) * m2.m(2)
    m1.m(3) = 0
    
    m1.m(4) = m1.m(0) * m2.m(4) + m1.m(4) * m2.m(5) + m1.m(8) * m2.m(6)
    m1.m(5) = m1.m(1) * m2.m(4) + m1.m(5) * m2.m(5) + m1.m(9) * m2.m(6)
    m1.m(6) = m1.m(2) * m2.m(4) + m1.m(6) * m2.m(5) + m1.m(10) * m2.m(6)
    m1.m(7) = 0

    m1.m(8) = m1.m(0) * m2.m(8) + m1.m(4) * m2.m(9) + m1.m(8) * m2.m(10)
    m1.m(9) = m1.m(1) * m2.m(8) + m1.m(5) * m2.m(9) + m1.m(9) * m2.m(10)
    m1.m(10) = m1.m(2) * m2.m(8) + m1.m(6) * m2.m(9) + m1.m(10) * m2.m(10)
    m1.m(11) = 0

    m1.m(12) = m1.m(0) * m2.m(12) + m1.m(4) * m2.m(13) + m1.m(8) * m2.m(14)
    m1.m(13) = m1.m(1) * m2.m(12) + m1.m(5) * m2.m(13) + m1.m(9) * m2.m(14)
    m1.m(14) = m1.m(2) * m2.m(12) + m1.m(6) * m2.m(13) + m1.m(10) * m2.m(14)
    m1.m(15) = 1
End Sub

Public Sub mat_setTranslation(m As matrix, translation() As Single)
    m.m(12) = translation(0)
    m.m(13) = translation(1)
    m.m(14) = translation(2)
End Sub

Public Sub mat_setInverseTranslation(m As matrix, translation() As Single)
    m.m(12) = -translation(0)
    m.m(13) = -translation(1)
    m.m(14) = -translation(2)
End Sub

Public Sub mat_setRotationRadians(m As matrix, radians() As Single)
    Dim cr As Double, sr As Double
    Dim cp As Double, sp As Double
    Dim cy As Double, sy As Double
    
    cr = Cos(radians(0))
    sr = Sin(radians(0))
    cp = Cos(radians(1))
    sp = Sin(radians(1))
    cy = Cos(radians(2))
    sy = Sin(radians(2))
    
    m.m(0) = cp * cy
    m.m(1) = cp * sy
    m.m(2) = -sp
    
    Dim crsp As Double, srsp As Double
    crsp = cr * sp
    srsp = sr * sp
    
    m.m(4) = srsp * cy - cr * sy
    m.m(5) = srsp * sy + cr * cy
    m.m(6) = sr * cp
    
    m.m(8) = crsp * cy + sr * sy
    m.m(9) = crsp * sy - sr * cy
    m.m(10) = cr * cp
End Sub

Public Sub mat_setRotationDegrees(m As matrix, degrees() As Single)
    Dim vec(2) As Single
    vec(0) = degrees(0) * 180 / pi
    vec(1) = degrees(1) * 180 / pi
    vec(2) = degrees(2) * 180 / pi
    
    mat_setRotationRadians m, vec
End Sub

Public Sub mat_setInverseRotationRadians(m As matrix, radians() As Single)
    Dim cr As Double, sr As Double
    Dim cp As Double, sp As Double
    Dim cy As Double, sy As Double
    
    cr = Cos(radians(0))
    sr = Sin(radians(0))
    cp = Cos(radians(1))
    sp = Sin(radians(1))
    cy = Cos(radians(2))
    sy = Sin(radians(2))
    
    m.m(0) = (cp * cy)
    m.m(4) = (cp * sy)
    m.m(8) = (-sp)
    
    Dim crsp As Double, srsp As Double
    crsp = cr * sp
    srsp = sr * sp
    
    m.m(1) = srsp * cy - cr * sy
    m.m(5) = srsp * sy + cr * cy
    m.m(9) = sr * cp
    
    m.m(2) = crsp * cy + sr * sy
    m.m(6) = crsp * sy - sr * cy
    m.m(10) = cr * cp
End Sub

Public Sub mat_setInverseRotationDegrees(m As matrix, degrees() As Single)
    Dim vec(2) As Single
    vec(0) = degrees(0) * 180 / pi
    vec(1) = degrees(1) * 180 / pi
    vec(2) = degrees(2) * 180 / pi
    
    mat_setInverseRotationRadians m, vec
End Sub

Public Sub mat_setRotationQuaternion(m As matrix, Q As quaternion)
    m.m(0) = 1 - 2 * Q.quat(1) * Q.quat(1) - 2 * Q.quat(2) * Q.quat(2)
    m.m(1) = 2 * Q.quat(0) * Q.quat(1) + 2 * Q.quat(3) * Q.quat(2)
    m.m(2) = 2 * Q.quat(0) * Q.quat(2) - 2 * Q.quat(3) * Q.quat(1)

    m.m(4) = 2 * Q.quat(0) * Q.quat(1) - 2 * Q.quat(3) * Q.quat(2)
    m.m(5) = 1 - 2 * Q.quat(0) * Q.quat(0) - 2 * Q.quat(2) * Q.quat(2)
    m.m(6) = 2 * Q.quat(1) * Q.quat(2) - 2 * Q.quat(3) * Q.quat(0)

    m.m(8) = 2 * Q.quat(0) * Q.quat(2) - 2 * Q.quat(3) * Q.quat(1)
    m.m(9) = 2 * Q.quat(1) * Q.quat(2) - 2 * Q.quat(3) * Q.quat(0)
    m.m(10) = 1 - 2 * Q.quat(0) * Q.quat(0) - 2 * Q.quat(1) * Q.quat(1)
End Sub

Public Sub mat_inverseTranslateVect(m As matrix, v As vector)
    v.pos(0) = v.pos(0) - m.m(12)
    v.pos(1) = v.pos(1) - m.m(13)
    v.pos(2) = v.pos(2) - m.m(14)
End Sub

Public Sub mat_inverseRotateVect(m As matrix, v As vector)
    Dim vec(2) As Single
    
    vec(0) = v.pos(0) * m.m(0) + v.pos(1) * m.m(1) + v.pos(2) * m.m(2)
    vec(1) = v.pos(0) * m.m(4) + v.pos(1) * m.m(5) + v.pos(2) * m.m(6)
    vec(2) = v.pos(0) * m.m(8) + v.pos(1) * m.m(9) + v.pos(2) * m.m(10)
    
    v.pos(0) = vec(0)
    v.pos(1) = vec(1)
    v.pos(2) = vec(2)
End Sub
