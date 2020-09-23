Attribute VB_Name = "objEntityLib"

Option Explicit

Private numEntities As Long
Private entities() As New CEntity

Public player1 As New CPlayer

Private Sub parseInfoPlayerStart()
    Dim i As Long
    Dim pos As vect
    
    Dim words() As String
    
    For i = numEntities To 0 Step -1
        If entities(i).findkey("classname") = "info_player_deathmatch" Then
            pos = makeVectFromString(entities(i).findkey("origin"))
            
            ObjEngine.setpos pos
            angleX = CSng(entities(i).findkey("angle")) - 110
            angleY = 0
        End If
    Next i
End Sub

Public Function makeVectFromString(s As String) As vect
    Dim w() As String
        
    Dim m As matrix
    Dim r(2) As Single
    Dim v As vector
    r(0) = 270
    Math3d.mat_loadIdentity m
    Math3d.mat_setRotationDegrees m, r
    
    w = Split(s, " ")
    If UBound(w) <> 2 Then Exit Function
    
    v.pos(0) = CSng(w(0))
    v.pos(1) = CSng(w(1))
    v.pos(2) = CSng(w(2))
    Math3d.vect_transform v, m
    makeVectFromString.x = v.pos(0)
    makeVectFromString.y = v.pos(1)
    makeVectFromString.z = v.pos(2)
End Function

Public Sub drawInfoPlayerStart()
    If getCVar("drawentities") <> "1" Then Exit Sub
End Sub

Public Sub parseEntities(path As String)
    Dim filenr As Long
    Dim line As String
    
    filenr = FreeFile
    Open path For Input As filenr
        Do Until EOF(filenr)
            Line Input #filenr, line
            If line = "{" Then
                numEntities = numEntities + 1
                ReDim Preserve entities(numEntities)
                'Set entities(numEntities) = New CEntity
            ElseIf line = "}" Then
            Else
                entities(numEntities).parse line
            End If
        Loop
    Close filenr
    
    parseInfoPlayerStart
End Sub

Public Sub init_entityLib()
    ReDim entities(0)
    numEntities = -1
End Sub

