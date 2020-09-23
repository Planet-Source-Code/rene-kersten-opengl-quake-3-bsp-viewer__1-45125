Attribute VB_Name = "Physics"

Option Explicit

Public Type movement
    start As vect
    end As vect
End Type

Private Function interpolate(v1 As vect, v2 As vect, t As Single) As vect
    interpolate.x = v1.x * t + v2.x * (t - 1)
    interpolate.y = v1.y * t + v2.y * (t - 1)
    interpolate.z = v1.z * t + v2.z * (t - 1)
End Function

Private Function getdist(v1 As vect, v2 As vect) As Single
    Dim vSub As vect
    Dim dXY As Single
    Dim dYZ As Single
    
    vSub.x = v1.x - v2.x
    vSub.y = v1.y - v2.y
    vSub.z = v1.z - v2.z
    
    dXY = Sqr(vSub.x * vSub.x + vSub.y * vSub.y)
    dYZ = Sqr(vSub.y * vSub.y + vSub.z * vSub.z)
    getdist = Sqr(dXY * dXY + dYZ * dYZ)
End Function

Public Function collision(move As movement) As Boolean
    Dim leaf As Long
    Dim startleaf As Long
    Dim endleaf As Long
    Dim pos As vect
    Dim x As Single
    Dim t As Single
    Dim tmpmove As movement
    Dim prefLeaf As Long
    Dim result As Boolean
    Dim length As Single
    
    length = getdist(move.start, move.end)
    
    If length = 0 Then
        collision = False
        Exit Function
    End If
    
    startleaf = findleaf(move.start)
    endleaf = findleaf(move.end)
    
    If startleaf = endleaf Then
        collision = trace(move, startleaf)
        Exit Function
    Else
        tmpmove.start = move.start
        
        For x = 0 To length
            t = x / length
            pos = interpolate(move.start, move.end, t)
            leaf = findleaf(pos)
            
            If leaf <> prefLeaf Then
                tmpmove.end = pos
                
                If trace(tmpmove, prefLeaf) Then
                    collision = True
                    Exit Function
                End If
                
                prefLeaf = leaf
                tmpmove.end = tmpmove.start
            End If
        Next x
    End If
End Function
