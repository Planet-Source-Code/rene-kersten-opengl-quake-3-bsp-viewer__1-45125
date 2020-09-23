Attribute VB_Name = "textureManager"

Option Explicit

Private numTextures As Long
Private pathLib() As String
Public textures() As Long
Public shaderIndices() As Long

Public Sub logStatus()
    Dim i As Long
    Dim numshadersloaded As Long
    
    For i = 0 To numShaders
        If shaderShouldBeParsed(i) Then
            numshadersloaded = numshadersloaded + 1
        End If
    Next i
    
    Log.Log "", "", ""
    Log.Log "", "", "texturemanager status:"
    Log.Log "", "", "number of textures loaded: " & (numTextures + 1)
    Log.Log "", "", "number of shaders loaded: " & numshadersloaded
    Log.Log "", "", ""
End Sub

Public Sub logShaderList()
    Dim i As Long
    For i = 0 To numShaders
        If shaderShouldBeParsed(i) Then Log.Log "", "", getShaderName(i)
    Next i
End Sub

Public Sub useTexture(n As Long)
    If n = -1 Then
        glDisable GL_TEXTURE_2D
        glBindTexture GL_TEXTURE_2D, 0
    Else
        glEnable GL_TEXTURE_2D
        glBindTexture GL_TEXTURE_2D, textures(n)
    End If
End Sub

Public Function useTextureWithShader(n As Long, face As bsp_face) As Boolean
    objShaderlib.resetStates
    
    If n = -1 Then
        glDisable GL_TEXTURE_2D
        glBindTexture GL_TEXTURE_2D, 0
    Else
        If shaderIndices(n) > -1 Then
            objShaderlib.execShader shaderIndices(n), face
            useTextureWithShader = True
        Else
            glEnable GL_TEXTURE_2D
            glBindTexture GL_TEXTURE_2D, textures(n)
        End If
    End If
End Function

Public Function textureExists(path As String) As Boolean
    Dim i As Long
    
    For i = 0 To numTextures
        If pathLib(i) = path Then
            textureExists = True
            Exit Function
        End If
    Next i
End Function

Public Function getTextureReal(path As String) As Long
    Dim i As Long
    
    For i = 0 To numTextures
        If pathLib(i) = path Then
            getTextureReal = textures(i)
            Exit Function
        End If
    Next i
End Function

Public Function getTexture(path As String) As Long
    Dim i As Long
    
    For i = 0 To numTextures
        If pathLib(i) = path Then
            getTexture = i
            Exit Function
        End If
    Next i
End Function

Public Function openQuakeTexture(path As String, Optional noLog As Boolean = False) As Long
    Dim texdata As EGLTEXTUREDATA
    
    On Error Resume Next
    
    path = Replace(path, "/", "\")
    
    If Not InStr(1, path, ":") > 0 Then
        If LEFT(path, 1) = "\" Then
            path = App.path & path
        Else: path = App.path & "\" & path
        End If
    End If
    
    If textureExists(path) Then
        openQuakeTexture = getTexture(path)
        Exit Function
    End If
    
    If numTextures = 37 Then
        path = path
    End If
    
    If pathexists(path & ".jpg") Then
        Dim tmppic As IPictureDisp
        Set tmppic = LoadPicture(path & ".jpg")
        SavePicture tmppic, App.path & "\tmp.bmp"
        texdata = auxLoadBitmap(App.path & "\tmp.bmp")
        Kill App.path & "\tmp.bmp"
    ElseIf pathexists(path & ".bmp") Then
        texdata = auxLoadBitmap(path & ".bmp")
    ElseIf pathexists(path & ".tga") Then
        texdata = auxLoadTarga(path & ".tga")
    End If
    
    If Not texdata.imgSizes = 0 Then 'Or pathexists(path & ".tga") Then
        numTextures = numTextures + 1
        ReDim Preserve textures(numTextures)
        ReDim Preserve pathLib(numTextures)
        ReDim Preserve shaderIndices(numTextures)
        
        shaderIndices(numTextures) = objShaderlib.findShader(RIGHT(path, Len(path) - Len(App.path)))
        
        objShaderlib.useShader shaderIndices(numTextures)
        
        textures(numTextures) = auxMipmaps(texdata, GL_LINEAR, GL_LINEAR_MIPMAP_NEAREST)
        pathLib(numTextures) = path
        
        openQuakeTexture = numTextures
    Else
        Dim shadernr As Long
        
        shadernr = objShaderlib.findShader(RIGHT(path, Len(path) - Len(App.path)))
        If shadernr > -1 Then
            objShaderlib.useShader shadernr
            
            numTextures = numTextures + 1
            ReDim Preserve textures(numTextures)
            ReDim Preserve pathLib(numTextures)
            ReDim Preserve shaderIndices(numTextures)
        
            textures(numTextures) = 0
            pathLib(numTextures) = path
            shaderIndices(numTextures) = shadernr
            
            openQuakeTexture = numTextures
            
            If Not noLog Then Log.Log "", "", "shader loaded as texture: " & RIGHT(path, Len(path) - Len(App.path))
        Else
            If Not noLog Then Log.Log "", "", "invalid texture: " & RIGHT(path, Len(path) - Len(App.path))
            openQuakeTexture = -1
        End If
    End If
End Function

Public Function addTexture(filename As String) As Long
    Dim texdata As EGLTEXTUREDATA
    
    If textureExists(filename) Then
        addTexture = getTexture(filename)
        Exit Function
    End If
    
    filename = Replace(filename, "/", "\")
    
    If Not InStr(1, filename, ":") > 0 Then
        If LEFT(filename, 1) = "\" Then
            filename = App.path & filename
        Else: filename = App.path & "\" & filename
        End If
    End If
    
    If RIGHT(filename, 3) = "bmp" Then
        texdata = auxLoadBitmap(filename)
    ElseIf RIGHT(filename, 3) = "jpg" Then
        Dim tmppic As IPictureDisp
        Set tmppic = LoadPicture(filename)
        SavePicture tmppic, App.path & "\tmp.bmp"
        texdata = auxLoadBitmap(App.path & "\tmp.bmp")
        Kill App.path & "\tmp.bmp"
    ElseIf RIGHT(filename, 3) = "tga" Then
        texdata = auxLoadTarga(filename)
    Else
        addTexture = -1
        Exit Function
    End If
    
    If Not texdata.imgSizes = 0 Then 'Or (pathexists(filename) And RIGHT(filename, 3) = "tga") Then
        numTextures = numTextures + 1
        ReDim Preserve textures(numTextures)
        ReDim Preserve pathLib(numTextures)
        textures(numTextures) = auxMipmaps(texdata, GL_LINEAR, GL_LINEAR_MIPMAP_NEAREST)
        pathLib(numTextures) = filename
        
        addTexture = numTextures
    Else: addTexture = -1
    End If
End Function

Public Sub clearAll()
    glDeleteTextures numTextures + 1, textures(0)
    numTextures = -1
End Sub

Public Sub initTextureLib()
    numTextures = -1
    ReDim textures(0)
End Sub
