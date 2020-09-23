Attribute VB_Name = "camera"

Public angleX As Single
Public angleY As Single

Private campos As vect

Public Function getcampos() As vect
    getcampos = campos
End Function

Public Sub calc_cam_around_point(pos As vect, distance As Single)
    Dim tmpangleY As Single
    tmpangleY = angleY
    If tmpangleY > 0 Then tmpangleY = -tmpangleY
    
    campos.x = -pos.x - (Sin(-angleX / 180 * pi) * (Cos(tmpangleY / 180 * pi))) * distance
    campos.y = -pos.y + Sin((-angleY) / 180 * pi) * distance
    campos.z = -pos.z - (Cos(-angleX / 180 * pi) * (Cos(tmpangleY / 180 * pi))) * distance
End Sub

Public Sub set_cam_to_point(pos As vect)
    campos.x = -pos.x
    campos.y = -pos.y
    campos.z = -pos.z
End Sub

Public Sub setup_camera()
    glLoadIdentity
    glRotatef angleY, 1, 0, 0
    glRotatef angleX, 0, 1, 0
    glTranslatef campos.x, campos.y, campos.z
End Sub

Public Sub print_camera_position()
    glDisable GL_TEXTURE_2D
    glDisable GL_CULL_FACE
    glTextOut "position: " & campos.x & " " & campos.y & " " & campos.z, -5, -0.8, 0.3
End Sub
