VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CShaderLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Type shader_line
    func As Long
    args() As String
End Type

Private Type shader_block
    numLines As Long
    lines() As shader_line
End Type

Private Type shader
    name As String
    
    numBlocks As Long
    blocks() As shader_block
End Type

Private Type shader_function
    name As String
    numArgs As Long
End Type

Private numShaders As Long
Private shaders() As shader
Private faces() As bsp_face

Private numFuncs As Long
Private funcs() As shader_function

Private paths() As String
Private numPaths As Long

Private currentFace As bsp_face

Private Sub parseline(line As shader_line)
    If line.func > -1 Then
        If Not UBound(line.args) <> funcs(line.func).numArgs Then
            CallByName Me, funcs(line.func).name, VbMethod, line.args
        End If
    End If
End Sub

Public Sub map(args() As String)
    Dim tex As Long
    If LEFT(args(0), 1) = "$" Then Exit Sub
    
    tex = addTexture(App.path & "\" & Replace(args(0), "/", "\"))
    
    If Not tex = -1 Then
        currentFace.textureID = textures(tex)
    Else
        currentFace.textureID = -1
    End If
End Sub

Public Sub tcMod(args() As String)
    If arg1 = "scale" Then
        
    End If
End Sub

Public Function shaderPrecached(path As String) As Boolean
    Dim i As Long
    
    For i = 0 To numPaths
        If paths(i) = path Then
            shaderPrecached = True
            Exit Function
        End If
    Next i
End Function

Friend Sub execShader(shaderNr As Long)
    If faces(shaderNr).textureID = -1 Then
        glDisable GL_TEXTURE_2D
        glBindTexture GL_TEXTURE_2D, 0
    Else
        glEnable GL_TEXTURE_2D
        glBindTexture GL_TEXTURE_2D, faces(shaderNr).textureID
    End If
End Sub

Private Sub execShaderReal(shaderNr As Long, face As bsp_face)
    Dim i As Long, j As Long
    
    currentFace = face
    
    currentFace.textureID = -1
    
    With shaders(shaderNr)
        For i = 0 To .numBlocks
            For j = 0 To .blocks(i).numLines
                parseline .blocks(i).lines(j)
                If currentFace.textureID > -1 Then GoTo endloop
            Next j
        Next i
    End With
    
endloop:
    If currentFace.textureID > -1 Then
        i = i
    End If
    face = currentFace
End Sub

Friend Sub findAndExecShader(ByVal shaderName As String, face As bsp_face)
    Dim i As Long
    
    shaderName = Replace(shaderName, "\", "/")
    
    For i = 0 To numShaders
        If shaders(i).name = shader Then
            execShader i
            Exit Sub
        End If
    Next i
End Sub

Public Function findShader(ByVal shaderName As String) As Long
    Dim i As Long
    
    shaderName = Replace(shaderName, "\", "/")
    
    If LEFT(shaderName, 1) = "/" Then shaderName = RIGHT(shaderName, Len(shaderName) - 1)
    
    For i = 0 To numShaders
        If shaders(i).name = shaderName Then
            findShader = i
            Exit Function
        End If
    Next i
    
    findShader = -1
End Function

Private Function findfunc(funcName As String) As Long
    Dim i As Long
    For i = 0 To numFuncs
        If funcs(i).name = funcName Then
            findfunc = i
            Exit Function
        End If
    Next i
    findfunc = -1
End Function

Public Sub calcShaders()
    Dim i As Long
    
    For i = 0 To numShaders
        execShaderReal i, faces(i)
    Next i
End Sub

Public Sub precacheShader(name As String)
    Dim filenr As Long
    Dim L As String
    Dim words() As String
    Dim brackets As Long
    Dim i As Long
    
    On Error GoTo errhandler
    
    If shaderPrecached(name) Then
        Exit Sub
    Else
        numPaths = numPaths + 1
        ReDim Preserve paths(numPaths)
        paths(numPaths) = name
    End If
    
    filenr = FreeFile
    Open name For Input As filenr
        Do Until EOF(filenr)
            Line Input #filenr, L
            
            L = Trim(LCase(Replace(L, Chr(9), "")))
            If L = "" Then GoTo next_do
            If LEFT(L, 2) = "//" Then GoTo next_do
            
            Select Case L
                Case "{"
                    brackets = brackets + 1
                    
                    shaders(numShaders).numBlocks = shaders(numShaders).numBlocks + 1
                    ReDim Preserve shaders(numShaders).blocks(shaders(numShaders).numBlocks)
                    shaders(numShaders).blocks(shaders(numShaders).numBlocks).numLines = -1
                    
                Case "}"
                    brackets = brackets - 1
                Case Else
                    If brackets = 0 Then
                        numShaders = numShaders + 1
                        ReDim Preserve shaders(numShaders)
                        ReDim Preserve faces(numShaders)
                        shaders(numShaders).name = L
                        shaders(numShaders).numBlocks = -1
                    Else
                        words = Split(L, " ")
                        
                        With shaders(numShaders).blocks(shaders(numShaders).numBlocks)
                            .numLines = .numLines + 1
                            ReDim Preserve .lines(.numLines)
                            .lines(.numLines).func = findfunc(words(0))
                            
                            If words(0) <> "" Then
                                i = i
                            End If
                            
                            If Not UBound(words) = 0 Then
                                ReDim .lines(.numLines).args(UBound(words) - 1)
                                
                                For i = 0 To UBound(.lines(.numLines).args)
                                    .lines(.numLines).args(i) = words(i + 1)
                                Next i
                            Else
                                ReDim .lines(.numLines).args(0)
                            End If
                        End With
                    End If
            End Select
next_do:
        Loop
    Close filenr
    
    Log.Log "", "", "shader loaded: " & name
    
    Exit Sub
    
errhandler:
    Log.Log "", "", "invalid shader: " & name
End Sub

Private Sub initfuncs()
    numFuncs = 1
    If Not numFuncs < 0 Then ReDim funcs(numFuncs)
    
    funcs(0).name = "map"
    funcs(0).numArgs = 0
    
    funcs(1).name = "tcMod"
    funcs(1).numArgs = 2
End Sub
