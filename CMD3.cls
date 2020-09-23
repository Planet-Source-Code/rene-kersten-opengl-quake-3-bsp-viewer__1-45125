VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMD3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Const MAX_QPATH = 64
Private Const MD3_SCALE = 1 / 64

Private Type MD3Header
   magic As String * 4
   version As Long
   path As String * MAX_QPATH
   flags As Long
   
   numFrames As Long
   numTags As Long
   numSurfaces As Long
   numSkins As Long
   
   offsetFrames As Long
   offsetTags As Long
   offsetSurfaces As Long
   offsetEnd As Long
End Type

Private Type MD3Frame
    BBCorner1 As vect
    BBCorner2 As vect
    origin As vect
    radius As Single
    name As String * 16
End Type

Private Type MD3SurfaceHeader
    magic As String * 4
    name As String * MAX_QPATH
    flags As Long
    
    numFrames As Long
    numShaders As Long
    numVertices As Long
    numTriangles As Long
    
    offsetTriangles As Long
    offsetShaders As Long
    offsetTexcoords As Long
    offsetVertices As Long
    offsetEnd As Long
End Type

Private Type MD3Shader
    path As String * MAX_QPATH
    index As Integer
    
    texture As Long
End Type

Private Type MD3Triangle
    indices(2) As Long
End Type

Private Type MD3Texcoord
    coords(1) As Single
End Type

Private Type MD3Vertex
    x As Integer
    y As Integer
    z As Integer
    lon As Byte
    lat As Byte
    
    nX As Single
    nY As Single
    nZ As Single
End Type

Private Type MD3Surface
    header As MD3SurfaceHeader
    shaders() As MD3Shader
    triangles() As MD3Triangle
    texcoords() As MD3Texcoord
    vertices() As MD3Vertex
End Type

Private header As MD3Header

Private frames() As MD3Frame
Private tags() As MD3Tag
Private surfaces() As MD3Surface

Public framenr As Long
Private nextFrame As Long
Private interpolation As Single

Private startTime As Long
Private shouldloop As Boolean

Private startFrame As Long
Private stopFrame As Long
Private fps As Long
Private hasanims As Boolean

Public Function animDone() As Boolean
    If Not hasanims Then
        animDone = True
        Exit Function
    End If
    
    animDone = (framenr = stopFrame)
End Function

Private Function parseSkin(shader As MD3Shader)
    Dim filename1 As String, filename2 As String, filenr As Long
    
    If LEFT(shader.path, 1) = Chr(0) Then shader.path = "m" & RIGHT(shader.path, Len(shader.path) - 1)
    
    filename1 = Replace(shader.path, Chr(0), "")
    filename2 = LEFT(filename1, Len(filename1) - 3) & "jpg"
    
    If pathexists(App.path & "\" & filename2) Then
        shader.texture = addTexture(filename2)
    ElseIf pathexists(App.path & "\" & filename1) Then
        shader.texture = addTexture(filename1)
    Else: shader.texture = -1
    End If
End Function

Friend Function getTagOrigin(tagname As String) As MD3Tag
    Dim i As Long
    Dim tagOffset As Long
    tagOffset = (framenr - 1) * (header.numTags + 1)
    For i = 0 To header.numTags
        If LEFT(tags(i).name, InStr(1, tags(i).name, Chr(0)) - 1) = tagname Then
            getTagOrigin = tags(i + tagOffset)
            Exit Function
        End If
    Next i
End Function

Private Function unpackNormals(vertex As MD3Vertex)
    Dim lon As Single
    Dim lat As Single
    With vertex
        lat = .lat / 255 * (2 * pi)
        lon = .lon / 255 * (2 * pi)
        .nX = Cos(lat) * Sin(lon)
        .nY = Sin(lat) * Sin(lon)
        .nZ = Cos(lon)
    End With
End Function

Public Function load(filename As String, hasAnimation As Boolean) As Boolean
    Dim filenr As Long
    Dim i As Long
    filenr = FreeFile
    
    hasanims = hasAnimation
    
    On Error GoTo errhandler
    
    Open filename For Binary As filenr
        Get filenr, , header
        
        header.numFrames = header.numFrames - 1
        header.numSkins = header.numSkins - 1
        header.numSurfaces = header.numSurfaces - 1
        header.numTags = header.numTags - 1
        
        If Not header.magic = "IDP3" Or Not header.version < 16 Then Err.Raise 10002, , "invalid md3 file"
        
        Seek filenr, header.offsetFrames + 1
        
        ReDim frames(header.numFrames)
        For i = 0 To header.numFrames
            Get filenr, , frames(i)
        Next i
        
        Seek filenr, header.offsetTags + 1
        
        ReDim tags((header.numTags + 1) * (header.numFrames + 1) - 1)
        For i = 0 To (header.numTags + 1) * (header.numFrames + 1) - 1
            Get filenr, , tags(i)
        Next i
        
        Seek filenr, header.offsetSurfaces + 1
        
        Dim ptr As Long
        ptr = header.offsetSurfaces
        
        ReDim surfaces(header.numSurfaces)
        For i = 0 To header.numSurfaces
            Seek filenr, ptr + 1
            Get filenr, , surfaces(i).header
            
            'parse the whole surface before
            'moving on to the next one
            If Not surfaces(i).header.magic = "IDP3" Then Err.Raise 10002, , "invalid md3 file"
            
            surfaces(i).header.numShaders = surfaces(i).header.numShaders - 1
            surfaces(i).header.numTriangles = surfaces(i).header.numTriangles - 1
            surfaces(i).header.numVertices = surfaces(i).header.numVertices - 1
            
            Seek filenr, ptr + surfaces(i).header.offsetShaders + 1
            
            ReDim surfaces(i).shaders(surfaces(i).header.numShaders)
            For j = 0 To surfaces(i).header.numShaders
                Get filenr, , surfaces(i).shaders(j).path
                Get filenr, , surfaces(i).shaders(j).index
                
                parseSkin surfaces(i).shaders(j)
            Next j
            
            Seek filenr, ptr + surfaces(i).header.offsetTriangles + 1
            
            ReDim surfaces(i).triangles(surfaces(i).header.numTriangles)
            For j = 0 To surfaces(i).header.numTriangles
                Get filenr, , surfaces(i).triangles(j)
            Next j
            
            Seek filenr, ptr + surfaces(i).header.offsetTexcoords + 1
            
            ReDim surfaces(i).texcoords(surfaces(i).header.numVertices)
            For j = 0 To surfaces(i).header.numVertices
                Get filenr, , surfaces(i).texcoords(j)
            Next j
            
            Seek filenr, ptr + surfaces(i).header.offsetVertices + 1
            
            ReDim surfaces(i).vertices(surfaces(i).header.numVertices * surfaces(i).header.numFrames)
            For j = 0 To surfaces(i).header.numVertices * surfaces(i).header.numFrames
                Get filenr, , surfaces(i).vertices(j).x
                Get filenr, , surfaces(i).vertices(j).y
                Get filenr, , surfaces(i).vertices(j).z
                Get filenr, , surfaces(i).vertices(j).lat
                Get filenr, , surfaces(i).vertices(j).lon
                unpackNormals surfaces(i).vertices(j)
            Next j
            
            ptr = ptr + surfaces(i).header.offsetEnd
        Next i
    Close filenr
    
    setFrame 1, 1, 1, False
    
    Exit Function
    
errhandler:
    Log.Log "CMD3", "load", "error"
End Function

Private Sub interpolateVertices(i As Long, j As Long, k As Long)
    Dim x1 As Single, x2 As Single, xt As Single
    Dim y1 As Single, y2 As Single, yt As Single
    Dim z1 As Single, z2 As Single, zt As Single
    Dim nx1 As Single, nx2 As Single, nxt As Single
    Dim ny1 As Single, ny2 As Single, nyt As Single
    Dim nz1 As Single, nz2 As Single, nzt As Single
    Dim offsetCurrent As Long, offsetNext As Long
    Dim index As Long
    
    'interpolation = 1
    
    offsetCurrent = (framenr - 1) * (surfaces(i).header.numVertices + 1)
    offsetNext = (framenr - 1) * (surfaces(i).header.numVertices + 1)
    index = surfaces(i).triangles(j).indices(k)
    
    If offsetCurrent + surfaces(i).header.numVertices > UBound(surfaces(i).vertices) Then
        Exit Sub
    End If
    If offsetNext + surfaces(i).header.numVertices > UBound(surfaces(i).vertices) Then
        Exit Sub
    End If
    
    x1 = surfaces(i).vertices(index + offsetCurrent).x
    y1 = surfaces(i).vertices(index + offsetCurrent).y
    z1 = surfaces(i).vertices(index + offsetCurrent).z
    
    x2 = surfaces(i).vertices(index + offsetNext).x
    y2 = surfaces(i).vertices(index + offsetNext).y
    z2 = surfaces(i).vertices(index + offsetNext).z
    
    xt = x1 * interpolation + x2 * (1 - interpolation)
    yt = y1 * interpolation + y2 * (1 - interpolation)
    zt = z1 * interpolation + z2 * (1 - interpolation)
    
    nx1 = surfaces(i).vertices(index + offsetCurrent).nX
    ny1 = surfaces(i).vertices(index + offsetCurrent).nY
    nz1 = surfaces(i).vertices(index + offsetCurrent).nZ
    
    nx2 = surfaces(i).vertices(index + offsetNext).nX
    ny2 = surfaces(i).vertices(index + offsetNext).nY
    nz2 = surfaces(i).vertices(index + offsetNext).nZ
    
    nxt = nx1 * interpolation + nx2 * (1 - interpolation)
    nyt = ny1 * interpolation + ny2 * (1 - interpolation)
    nzt = nz1 * interpolation + nz2 * (1 - interpolation)
    
    glNormal3f nzt, nyt, nxt
    glVertex3f xt * MD3_SCALE, yt * MD3_SCALE, zt * MD3_SCALE
End Sub

Public Sub draw()
    Dim i As Long, j As Long
    Dim frameOffset As Long
    
    If hasanims Then
        If framenr > header.numFrames Then
            framenr = header.numFrames
            'Log.Log "CMD3", "draw", "error: invalid frame number"
            'objEngine.quit
        End If
    End If

    For i = 0 To header.numSurfaces
        With surfaces(i)
            If (framenr - 1) * (.header.numVertices + 2) > UBound(.vertices) Then
                framenr = framenr - 1
            End If
            
            If hasanims Then
                frameOffset = (framenr - 1) * (.header.numVertices + 1)
            Else
                frameOffset = 0
            End If
            
            useTexture .shaders(0).texture
            glBegin GL_TRIANGLES
                If Not framenr = header.numFrames Then
                    For j = 0 To .header.numTriangles
                        glTexCoord2f .texcoords(.triangles(j).indices(0)).coords(0), 1 - .texcoords(.triangles(j).indices(0)).coords(1)
                        interpolateVertices i, j, 0
                        
                        glTexCoord2f .texcoords(.triangles(j).indices(1)).coords(0), 1 - .texcoords(.triangles(j).indices(1)).coords(1)
                        interpolateVertices i, j, 1
                        
                        glTexCoord2f .texcoords(.triangles(j).indices(2)).coords(0), 1 - .texcoords(.triangles(j).indices(2)).coords(1)
                        interpolateVertices i, j, 2
                    Next j
                Else
                    For j = 0 To .header.numTriangles
                        glNormal3f .vertices(.triangles(j).indices(0) + frameOffset).nZ, .vertices(.triangles(j).indices(0) + frameOffset).nY, .vertices(.triangles(j).indices(0) + frameOffset).nX
                        glTexCoord2f .texcoords(.triangles(j).indices(0)).coords(0), 1 - .texcoords(.triangles(j).indices(0)).coords(1)
                        glVertex3f .vertices(.triangles(j).indices(0) + frameOffset).x * MD3_SCALE, .vertices(.triangles(j).indices(0) + frameOffset).y * MD3_SCALE, .vertices(.triangles(j).indices(0) + frameOffset).z * MD3_SCALE
                        
                        glNormal3f .vertices(.triangles(j).indices(1) + frameOffset).nZ, .vertices(.triangles(j).indices(1) + frameOffset).nY, .vertices(.triangles(j).indices(1) + frameOffset).nX
                        glTexCoord2f .texcoords(.triangles(j).indices(1)).coords(0), 1 - .texcoords(.triangles(j).indices(1)).coords(1)
                        glVertex3f .vertices(.triangles(j).indices(1) + frameOffset).x * MD3_SCALE, .vertices(.triangles(j).indices(1) + frameOffset).y * MD3_SCALE, .vertices(.triangles(j).indices(1) + frameOffset).z * MD3_SCALE
                        
                        glNormal3f .vertices(.triangles(j).indices(2) + frameOffset).nZ, .vertices(.triangles(j).indices(2) + frameOffset).nY, .vertices(.triangles(j).indices(2) + frameOffset).nX
                        glTexCoord2f .texcoords(.triangles(j).indices(2)).coords(0), 1 - .texcoords(.triangles(j).indices(2)).coords(1)
                        glVertex3f .vertices(.triangles(j).indices(2) + frameOffset).x * MD3_SCALE, .vertices(.triangles(j).indices(2) + frameOffset).y * MD3_SCALE, .vertices(.triangles(j).indices(2) + frameOffset).z * MD3_SCALE
                    Next j
                End If
            glEnd
        End With
    Next i
End Sub

Public Sub setFrame(theStartframe As Long, theStopframe As Long, theFps As Long, shouldLoopflag As Boolean)
    startTime = GetTickCount
    framenr = startFrame
    startFrame = theStartframe
    stopFrame = theStopframe
    fps = theFps
    shouldloop = shouldLoopflag
    framenr = startFrame + IIf(startFrame < header.numFrames, 1, 0)
    startTime = GetTickCount
End Sub

Public Sub update()
    Dim time As Long
    
    time = GetTickCount
    If time - startTime > 1000 / fps Then
        startTime = time
        framenr = framenr + 1
        nextFrame = nextFrame + 1
        
        interpolation = 1
        
        If framenr > stopFrame Then
            If shouldloop Then
                startTime = GetTickCount
                framenr = startFrame + 1
            Else
                framenr = stopFrame
            End If
        End If
    Else
        interpolation = 1 - (time - startTime) / (1000 / fps)
    End If
    
    'framenr = ((time - startTime) / totalTime) * totalFrames
    
    'framenr = objEngine.state + 1
    'If framenr > header.numFrames + 1 Then framenr = header.numFrames + 1
End Sub
