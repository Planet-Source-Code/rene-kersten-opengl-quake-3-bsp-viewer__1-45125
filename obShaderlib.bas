Attribute VB_Name = "objShaderlib"

Option Explicit

Public Const surfaceparm_alphashadow = 1
Public Const surfaceparm_areaportal = 2
Public Const surfaceparm_clusterportal = 4
Public Const surfaceparm_donotenter = 8

Public Const surfaceparm_flesh = 16
Public Const surfaceparm_fog = 32
Public Const surfaceparm_lava = 64
Public Const surfaceparm_metalsteps = 128

Public Const surfaceparm_nodamage = 2 ^ 8
Public Const surfaceparm_nodlight = 2 ^ 9
Public Const surfaceparm_nodraw = 2 ^ 10
Public Const surfaceparm_nodrop = 2 ^ 11

Public Const surfaceparm_noimpact = 2 ^ 12
Public Const surfaceparm_nomarks = 2 ^ 13
Public Const surfaceparm_nolightmaps = 2 ^ 14
Public Const surfaceparm_nosteps = 2 ^ 15

Public Const surfaceparm_nonsolid = 2 ^ 16
Public Const surfaceparm_origin = 2 ^ 17
Public Const surfaceparm_playerclip = 2 ^ 18
Public Const surfaceparm_slick = 2 ^ 19

Public Const surfaceparm_slime = 2 ^ 20
Public Const surfaceparm_structural = 2 ^ 21
Public Const surfaceparm_trans = 2 ^ 22
Public Const surfaceparm_water = 2 ^ 23

Public Const surfaceparm_sky = 2 ^ 24

Private surfaceparmNames(24) As String

Private Const func_map = 0
Private Const func_tcmod = 1
Private Const func_alphafunc = 2
Private Const func_cull = 3
Private Const func_blendfunc = 4
Private Const func_blendfunc_explicit = 5
Private Const func_surfaceparm = 6
Private Const func_animmap = 7
Private Const func_tcgen = 8

Private Const tcgen_base = 0
Private Const tcgen_lightmap = 1
Private Const tcgen_enviroment = 2

Public Const func_none = 0
Public Const func_GT0 = 1    'greater than zero
Public Const func_LT128 = 2    'less than 128
Public Const func_GE128 = 3    'greater than or equal to 128

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

Private Type bsp_shader_stage
    'alphafunc stuff
    alphafunc As Long
    textureID(7) As Long
    freqency As Single
    numTextures As Long
    
    shouldRotate As Boolean
    texRotate As Single
    
    'current texture info
    currentTex As Long
    oldTime As Long
    
    'blendfunc stuff
    blendL As Long
    blendR As Long
    shouldBlend As Boolean
    
    'tcMod stuff
    scrollX As Single
    scrollY As Single
    
    scalex As Single
    scaley As Single
    
    'tcgen
    texCoordType As Long
    
    'cull param
    cull As Long
End Type

Private Type bsp_shader
    numStages As Long
    stages() As bsp_shader_stage
    
    'globals
    surfaceparms As Long
End Type

Public numShaders As Long
Private shaders() As shader
Public parsedShaders() As bsp_shader
Public shaderShouldBeParsed() As Boolean

Private numFuncs As Long
Private funcs() As shader_function

Private paths() As String
Private numPaths As Long

Private currentStage As bsp_shader_stage
Private currentShader As bsp_shader

Private startTime As Long
Public t As Single 't = (gettickcount - starttime) / 1000

Public Function getShaderName(shadernr As Long) As String
    getShaderName = shaders(shadernr).name
End Function

Public Sub drawSkybox(shadernr As Long)
    Dim i As Long
    
    resetStates
    glDisable GL_CULL_FACE
    
    For i = 0 To parsedShaders(shadernr).numStages
        With parsedShaders(shadernr).stages(i)
            'culling
            If .cull > 0 Then
                glCullFace .cull
            Else: glDisable GL_CULL_FACE
            End If 'texture stuff
            
            If .textureID(0) = -1 Then
                GoTo next_i
            Else
                glEnable GL_TEXTURE_2D
                glBindTexture GL_TEXTURE_2D, .textureID(0)
                
                glTexEnvf GL_TEXTURE_2D, GL_TEXTURE_ENV_MODE, GL_DECAL
            End If
            
            'alphafunc stuff
            glEnable GL_ALPHA_TEST
            glEnable GL_ALPHA_TEST_FUNC
            If .alphafunc = func_GE128 Then
                glAlphaFunc GL_GEQUAL, 128
            ElseIf .alphafunc = func_GT0 Then
                glAlphaFunc GL_GREATER, 0
            ElseIf .alphafunc = func_LT128 Then
                glAlphaFunc GL_LESS, 128
            Else: glAlphaFunc GL_NONE, 0
            End If
            
            'blend stuff
            If .shouldBlend Then
                glBlendFunc .blendL, .blendR
                glEnable GL_BLEND
            End If
            
            glScalef 20, 20, 20
            
            glBegin GL_QUADS
                'front
                glTexCoord2f .scalex + .scrollX * t, .scrollY * t
                glVertex3f -1, -1, 1
                glTexCoord2f .scalex + .scrollX * t, .scaley + .scrollY * t
                glVertex3f -1, 1, 1
                glTexCoord2f .scrollX * t, .scaley + .scrollY * t
                glVertex3f 1, 1, 1
                glTexCoord2f .scrollX * t, .scrollY * t
                glVertex3f 1, -1, 1
                
                'back
                glTexCoord2f .scrollX * t, .scrollY * t
                glVertex3f -1, -1, -1
                glTexCoord2f .scrollX * t, .scaley + .scrollY * t
                glVertex3f -1, 1, -1
                glTexCoord2f .scalex + .scrollX * t, .scaley + .scrollY * t
                glVertex3f 1, 1, -1
                glTexCoord2f .scalex + .scrollX * t, .scrollY * t
                glVertex3f 1, -1, -1
                
                'left
                glTexCoord2f .scalex + .scrollX * t, .scrollY * t
                glVertex3f -1, -1, -1
                glTexCoord2f .scrollX * t, .scrollY * t
                glVertex3f -1, -1, 1
                glTexCoord2f .scrollX * t, .scaley + .scrollY * t
                glVertex3f -1, 1, 1
                glTexCoord2f .scalex + .scrollX * t, .scaley + .scrollY * t
                glVertex3f -1, 1, -1
                
                'right
                glTexCoord2f .scrollX * t, .scrollY * t
                glVertex3f 1, -1, -1
                glTexCoord2f .scalex + .scrollX * t, .scrollY * t
                glVertex3f 1, -1, 1
                glTexCoord2f .scalex + .scrollX * t, .scaley + .scrollY * t
                glVertex3f 1, 1, 1
                glTexCoord2f .scrollX * t, .scaley + .scrollY * t
                glVertex3f 1, 1, -1
                
                'up
                glTexCoord2f .scrollX * t, .scaley + .scrollY * t
                glVertex3f -1, 1, -1
                glTexCoord2f .scalex + .scrollX * t, .scaley + .scrollY * t
                glVertex3f -1, 1, 1
                glTexCoord2f .scalex + .scrollX * t, .scrollY * t
                glVertex3f 1, 1, 1
                glTexCoord2f .scrollX * t, .scrollY * t
                glVertex3f 1, 1, -1
                
                'down
                glTexCoord2f .scrollX * t, .scrollY * t
                glVertex3f -1, -1, -1
                glTexCoord2f .scrollX * t, .scaley + .scrollY * t
                glVertex3f -1, -1, 1
                glTexCoord2f .scalex + .scrollX * t, .scaley + .scrollY * t
                glVertex3f -1, 1, 1
                glTexCoord2f .scalex + .scrollX * t, .scrollY * t
                glVertex3f -1, 1, -1
            glEnd
        End With
next_i:
    Next i
End Sub

'this sets up the shaders for each frame
Public Sub setupShaders()
    Dim newtime As Long
    newtime = GetTickCount
    t = (newtime - startTime) / 1000
    't = t Mod 1
    
    Dim currenttime As Long
    Dim d As Single
    currenttime = GetTickCount
    
    Dim i As Long, j As Long
    For i = 0 To numShaders
        If shaderShouldBeParsed(i) Then
            For j = 0 To parsedShaders(i).numStages
                With parsedShaders(i).stages(j)
                    If .numTextures < 0 Then GoTo next_j
                    
                    d = CSng(currenttime - .oldTime)
                    If d > ((1 / .freqency) / (.numTextures + 1)) * 1000 Then
                        .currentTex = .currentTex + 1
                        .oldTime = GetTickCount
                    End If
                    
                    If .currentTex > .numTextures Then
                        .currentTex = 0
                    End If
                End With
next_j:
            Next j
        End If
    Next i
End Sub

Private Sub parseline(line As shader_line, i As Long)
    Dim j As Long
    Dim tex As Long
    
    If line.func > -1 Then
        Select Case line.func
            Case func_map
                If LEFT(line.args(0), 1) = "$" Then Exit Sub
                
                If Len(line.args(0)) < 5 Then Exit Sub
                
                tex = openQuakeTexture(LEFT(line.args(0), Len(line.args(0)) - 4), True) 'addTexture(line.args(0)) '
                
                currentStage.freqency = 1
                
                If Not tex = -1 Then
                    currentStage.textureID(0) = textures(tex)
                    currentStage.numTextures = 0
                    Log.Log "", "", "shader texture loaded: " & App.path & "\" & Replace(line.args(0), "/", "\") & " for shader " & shaders(i).name
                Else
                    currentStage.numTextures = -1
                    currentStage.textureID(0) = -1
                End If
            Case func_tcmod
                If UBound(line.args) = 2 Then
                    If line.args(0) = "scale" Then
                        currentStage.scalex = CSng(line.args(1))
                        currentStage.scaley = CSng(line.args(2))
                    ElseIf line.args(0) = "scroll" Then
                        currentStage.scrollX = CSng(line.args(1))
                        currentStage.scrollY = -CSng(line.args(2))
                    End If
                ElseIf UBound(line.args) = 1 Then
                    If line.args(0) = "rotate" Then
                        currentStage.texRotate = CSng(line.args(1)) / 180 * pi
                        currentStage.shouldRotate = True
                    End If
                End If
            Case func_tcgen
                If line.args(0) = "base" Then
                    currentStage.texCoordType = tcgen_base
                ElseIf line.args(0) = "lightmap" Then
                    currentStage.texCoordType = tcgen_lightmap
                ElseIf line.args(0) = "environment" Then
                    currentStage.texCoordType = tcgen_enviroment
                End If
            Case func_alphafunc
                Select Case line.args(0)
                    Case "gt0"
                        currentStage.alphafunc = 1
                    Case "lt128"
                        currentStage.alphafunc = 2
                    Case "ge128"
                        currentStage.alphafunc = 3
                End Select
            Case func_cull
                Select Case line.args(0)
                    Case "none", "disable"
                        currentStage.cull = 0
                    Case "front"
                        currentStage.cull = GL_FRONT
                    Case "back"
                        currentStage.cull = GL_BACK
                End Select
            
            Case func_blendfunc
                currentStage.shouldBlend = True
                
                Select Case line.args(0)
                    Case "add"
                        currentStage.blendL = GL_ONE
                        currentStage.blendR = GL_ONE
                    Case "filter"
                        currentStage.blendL = GL_ZERO
                        currentStage.blendR = GL_SRC_COLOR
                    Case "blend"
                        currentStage.blendL = GL_SRC_ALPHA
                        currentStage.blendR = GL_ONE_MINUS_SRC_ALPHA
                End Select
            Case func_blendfunc_explicit
                currentStage.shouldBlend = True
                
                Select Case line.args(0)
                    Case "gl_one"
                        currentStage.blendL = GL_ONE
                    Case "gl_zero"
                        currentStage.blendL = GL_ZERO
                    Case "gl_dst_color"
                        currentStage.blendL = GL_DST_COLOR
                    Case "gl_src_alpha"
                        currentStage.blendL = GL_SRC_ALPHA
                    Case "gl_one_minus_dst_color"
                        currentStage.blendL = GL_ONE_MINUS_DST_COLOR
                    Case "gl_one_minus_src_alpha"
                        currentStage.blendL = GL_ONE_MINUS_SRC_ALPHA
                End Select
                
                Select Case line.args(1)
                    Case "gl_one"
                        currentStage.blendR = GL_ONE
                    Case "gl_zero"
                        currentStage.blendR = GL_ZERO
                    Case "gl_src_color"
                        currentStage.blendR = GL_SRC_COLOR
                    Case "gl_src_alpha"
                        currentStage.blendR = GL_SRC_ALPHA
                    Case "gl_one_minus_src_alpha"
                        currentStage.blendR = GL_ONE_MINUS_SRC_ALPHA
                    Case "gl_one_minus_src_color"
                        currentStage.blendR = GL_ONE_MINUS_SRC_COLOR
                End Select
            Case func_surfaceparm
                For j = 0 To 24
                    If surfaceparmNames(j) = line.args(0) Then
                        currentShader.surfaceparms = 2 ^ j
                    End If
                Next j
            Case func_animmap
                If UBound(line.args) < 1 Then Exit Sub
                
                currentStage.freqency = CSng(line.args(0))
                
                For j = 0 To UBound(line.args) - 1
                    If LEFT(line.args(j + 1), 1) = "$" Then
                        currentStage.textureID(j) = -1
                        GoTo next_j
                    End If
                    
                    If Len(line.args(j + 1)) < 5 Then
                        currentStage.textureID(j) = -1
                        GoTo next_j
                    End If
                    
                    tex = openQuakeTexture(LEFT(line.args(j + 1), Len(line.args(j + 1)) - 4), True) 'addTexture(line.args(j + 1)) '
                    
                    If Not tex = -1 Then
                        currentStage.textureID(j) = textures(tex)
                        Log.Log "", "", "shader texture loaded: " & App.path & "\" & Replace(line.args(j + 1), "/", "\") & " for shader " & shaders(i).name
                    Else
                        currentStage.textureID(j) = -1
                    End If
next_j:
                Next j
                
                For j = UBound(line.args) To 7
                    currentStage.textureID(j) = -1
                Next j
                
                currentStage.numTextures = UBound(line.args) - 1
        End Select
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

Public Sub useShader(shadernr As Long)
    If shadernr = -1 Then Exit Sub
    shaderShouldBeParsed(shadernr) = True
End Sub

Public Sub clearAllShaders()
    Dim i As Long
    For i = 0 To numShaders
        shaderShouldBeParsed(i) = False
    Next i
End Sub

Private Sub drawGeometry(face As bsp_face, stage As bsp_shader_stage, shadernr As Long)
    Dim rotx As Single
    Dim roty As Single
    Dim j As Long
    
    rotx = (Sin(stage.texRotate) * 0.5) * (t - Int(t / 1))
    roty = (Cos(stage.texRotate) * 0.5) * (t - Int(t / 1))
    
    If face.type = 1 Then
        glBegin GL_TRIANGLE_FAN
            For j = face.vertexIndex To face.vertexIndex + face.numVertices - 1
                If parsedShaders(shadernr).surfaceparms And surfaceparm_sky Then
                    glTexCoord2f (vertices(j).texture(0) + stage.scrollX * t), (vertices(j).texture(1) + stage.scrollY * t)
                Else
                    If stage.shouldRotate Then
                        glTexCoord2f vertices(j).texture(0) * stage.scalex + stage.scrollX * t + rotx, vertices(j).texture(1) * stage.scaley + stage.scrollY * t + roty
                    Else: glTexCoord2f vertices(j).texture(0) * stage.scalex + stage.scrollX * t, vertices(j).texture(1) * stage.scaley + stage.scrollY * t
                    End If
                    glTexCoord2f vertices(j).texture(0) + stage.scrollX * t, vertices(j).texture(1) + stage.scrollY * t
                End If
                glVertex3fv vertices(j).position(0)
            Next j
        glEnd
    ElseIf face.type = 3 Then
        glBegin GL_TRIANGLES
            For j = face.meshVertexIndex To face.meshVertexIndex + face.meshNumVertices - 1
                If stage.shouldRotate Then
                    glTexCoord2f vertices(meshvertices(j) + face.vertexIndex).texture(0) * stage.scalex + stage.scrollX * t + rotx, vertices(meshvertices(j) + face.vertexIndex).texture(1) * stage.scaley + stage.scrollY * t + roty
                Else: glTexCoord2f vertices(meshvertices(j) + face.vertexIndex).texture(0) * stage.scalex + stage.scrollX * t, vertices(meshvertices(j) + face.vertexIndex).texture(1) * stage.scaley + stage.scrollY * t
                End If
                glVertex3fv vertices(meshvertices(j) + face.vertexIndex).position(0)
            Next j
        glEnd
    End If
End Sub

Public Sub execShader(shadernr As Long, face As bsp_face)
    Dim i As Long
    
    'GLOBALS
    'surface parms stuff
    
    If parsedShaders(shadernr).surfaceparms And surfaceparm_sky Then Exit Sub
    
    'STAGES
    For i = 0 To parsedShaders(shadernr).numStages
        With parsedShaders(shadernr).stages(i)
        'culling
            If .cull > 0 Then
                glCullFace .cull
                glEnable GL_CULL_FACE
            Else: glDisable GL_CULL_FACE
            End If 'texture stuff
            
            If .textureID(.currentTex) = -1 Then
                GoTo next_i
            ElseIf .textureID(.currentTex) = 0 Then
                GoTo next_i 'i = i
            Else
                glEnable GL_TEXTURE_2D
                glBindTexture GL_TEXTURE_2D, .textureID(.currentTex)
            End If
            
            'alphafunc stuff
            glEnable GL_ALPHA_TEST
            glEnable GL_ALPHA_TEST_FUNC
            If .alphafunc = func_GE128 Then
                glAlphaFunc GL_GEQUAL, 128
            ElseIf .alphafunc = func_GT0 Then
                glAlphaFunc GL_GREATER, 0
            ElseIf .alphafunc = func_LT128 Then
                glAlphaFunc GL_LESS, 128
            Else
                glDisable GL_ALPHA_TEST
                glDisable GL_ALPHA_TEST_FUNC
            End If
            
            'blend stuff
            If .shouldBlend Then
                glBlendFunc .blendL, .blendR
                glEnable GL_BLEND
            End If
        End With
        
        drawGeometry face, parsedShaders(shadernr).stages(i), shadernr
next_i:
    Next i
End Sub

Public Sub resetStates()
    glDisable GL_ALPHA_TEST
    glDisable GL_ALPHA_TEST_FUNC
    glCullFace GL_FRONT
    glEnable GL_CULL_FACE
    
    glDisable GL_BLEND
    glBlendFunc GL_ZERO, GL_SRC_COLOR
End Sub

Private Function execShaderReal(shadernr As Long, theShader As bsp_shader) As bsp_shader
    Dim i As Long, j As Long
    
    currentShader.surfaceparms = 0
    currentShader.numStages = -1
    ReDim currentShader.stages(0)
    
    currentShader.numStages = shaders(shadernr).numBlocks
    
    If Not currentShader.numStages < 0 Then
        ReDim currentShader.stages(currentShader.numStages)
    End If
    
    With shaders(shadernr)
        For i = 0 To .numBlocks
            currentStage.alphafunc = 0
            currentStage.shouldBlend = False
            
            For j = 0 To 7
                currentStage.textureID(j) = -1
            Next j
            currentStage.freqency = 1
            currentStage.numTextures = 0
            
            currentStage.scalex = 1
            currentStage.scaley = 1
            currentStage.scrollX = 0
            currentStage.scrollY = 0
            currentStage.cull = 0
            
            currentStage.texCoordType = tcgen_base
            
            currentStage.shouldRotate = False
            
            For j = 0 To .blocks(i).numLines
                parseline .blocks(i).lines(j), shadernr
            Next j
            
            currentShader.stages(i) = currentStage
        Next i
    End With
End Function

Public Sub findAndExecShader(ByVal shaderName As String, face As bsp_face)
    Dim i As Long
    
    shaderName = Replace(shaderName, "\", "/")
    
    For i = 0 To numShaders
        If shaders(i).name = shaderName Then
            execShader i, face
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

Private Function findfunc(funcName As String, numArgs As Long) As Long
    Dim i As Long
    For i = 0 To numFuncs
        If funcs(i).name = funcName Then
            If funcs(i).numArgs = numArgs Or funcs(i).numArgs = -1 Then
                findfunc = i
                Exit Function
            End If
        End If
    Next i
    findfunc = -1
End Function

Public Sub calcShaders()
    Dim i As Long, j As Long, k As Long
    
    Log.Log "", "", "number of shaders: " & numShaders
    For i = 0 To numShaders
        If i = 735 Then
            i = i
        End If
        If shaderShouldBeParsed(i) Then
            execShaderReal i, parsedShaders(i)
            
            parsedShaders(i).numStages = currentShader.numStages
            parsedShaders(i).surfaceparms = currentShader.surfaceparms
            
            If Not parsedShaders(i).numStages < 0 Then
                ReDim parsedShaders(i).stages(parsedShaders(i).numStages)
                For j = 0 To parsedShaders(i).numStages
                    parsedShaders(i).stages(j).cull = currentShader.stages(j).cull
                    parsedShaders(i).stages(j).alphafunc = currentShader.stages(j).alphafunc
                    parsedShaders(i).stages(j).blendL = currentShader.stages(j).blendL
                    parsedShaders(i).stages(j).blendR = currentShader.stages(j).blendR
                    parsedShaders(i).stages(j).shouldBlend = currentShader.stages(j).shouldBlend
                    parsedShaders(i).stages(j).texRotate = currentShader.stages(j).texRotate
                    parsedShaders(i).stages(j).shouldRotate = currentShader.stages(j).shouldRotate
                    
                    parsedShaders(i).stages(j).texCoordType = currentShader.stages(j).texCoordType
                    
                    parsedShaders(i).stages(j).freqency = currentShader.stages(j).freqency
                    parsedShaders(i).stages(j).numTextures = currentShader.stages(j).numTextures
                    For k = 0 To 7
                        parsedShaders(i).stages(j).textureID(k) = currentShader.stages(j).textureID(k)
                    Next k
                    
                    parsedShaders(i).stages(j).scalex = currentShader.stages(j).scalex
                    parsedShaders(i).stages(j).scaley = currentShader.stages(j).scaley
                    parsedShaders(i).stages(j).scrollX = currentShader.stages(j).scrollX
                    parsedShaders(i).stages(j).scrollY = -currentShader.stages(j).scrollY
                Next j
            End If
            
            Log.Log "", "", "shader loaded"
        End If
        
    Next i
    
    startTime = GetTickCount
End Sub

Public Sub loadAllShaders()
    Dim x As FileListBox
    
    Dim i As Long
    
    Log.Log "", "", ""
    Log.Log "", "", "loading all shaders"
    
    Set x = Form1.dir
    
    x.path = App.path & "\scripts\"
    
    For i = 0 To x.ListCount
        If RIGHT(LCase(x.list(i)), 6) = "shader" Then
            precacheShader App.path & "\scripts\" & x.list(i)
        End If
    Next i
End Sub

Public Sub precacheShader(name As String)
    Dim filenr As Long
    Dim l As String
    Dim words() As String
    Dim brackets As Long
    Dim i As Long
    
    On Error GoTo errhandler
    
    name = Replace(name, "/", "\")
    
    If Not InStr(1, name, ":") > 0 Then
        If LEFT(name, 1) = "\" Then
            name = App.path & name
        Else: name = App.path & "\" & name
        End If
    End If
    
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
            Line Input #filenr, l
            
            l = Trim(LCase(Replace(l, Chr(9), "")))
            If l = "" Then GoTo next_do
            If LEFT(l, 2) = "//" Then GoTo next_do
            
            Select Case l
                Case "{"
                    brackets = brackets + 1
                    
                    shaders(numShaders).numBlocks = shaders(numShaders).numBlocks + 1
                    ReDim Preserve shaders(numShaders).blocks(shaders(numShaders).numBlocks)
                    shaders(numShaders).blocks(shaders(numShaders).numBlocks).numLines = -1
                    
                Case "}"
                    brackets = brackets - 1
                    If brackets < 0 Then brackets = 0
                    
                Case Else
                    If brackets = 0 Then
                        numShaders = numShaders + 1
                        ReDim Preserve shaders(numShaders)
                        ReDim Preserve parsedShaders(numShaders)
                        ReDim Preserve shaderShouldBeParsed(numShaders)
                        shaders(numShaders).name = l
                        shaders(numShaders).numBlocks = -1
                    Else
                        Do
                            If Not (InStr(1, l, "  ") > 0) Then Exit Do
                            l = Replace(l, "  ", " ")
                        Loop
                        words = Split(l, " ")
                        
                        With shaders(numShaders).blocks(shaders(numShaders).numBlocks)
                            .numLines = .numLines + 1
                            ReDim Preserve .lines(.numLines)
                            
                            If Not UBound(words) = 0 Then
                                ReDim .lines(.numLines).args(UBound(words) - 1)
                                
                                For i = 0 To UBound(.lines(.numLines).args)
                                    .lines(.numLines).args(i) = words(i + 1)
                                Next i
                            Else
                                ReDim .lines(.numLines).args(0)
                            End If
                            
                            .lines(.numLines).func = findfunc(words(0), UBound(.lines(.numLines).args))
                        End With
                    End If
            End Select
next_do:
        Loop
    Close filenr
    
    Log.Log "", "", "shader file loaded: " & name
    
    Exit Sub
    
errhandler:
    Log.Log "", "", "invalid shader: " & name
    Log.Log "", "", "error in " & Err.Source & ", number " & Err.Number & ": " & Err.Description
End Sub

Private Sub initfuncs()
    numFuncs = 8
    If Not numFuncs < 0 Then ReDim funcs(numFuncs)
    
    funcs(0).name = "map"
    funcs(0).numArgs = 0
    
    funcs(1).name = "tcmod"
    funcs(1).numArgs = -1
    
    funcs(2).name = "alphafunc"
    funcs(2).numArgs = 0
    
    funcs(3).name = "cull"
    funcs(3).numArgs = 0
    
    funcs(4).name = "blendfunc"
    funcs(4).numArgs = 0
    
    funcs(5).name = "blendfunc"
    funcs(5).numArgs = 1
    
    funcs(6).name = "surfaceparm"
    funcs(6).numArgs = 0
    
    funcs(7).name = "animmap"
    funcs(7).numArgs = -1
    
    funcs(8).name = "tcgen"
    funcs(8).numArgs = 0
    
    surfaceparmNames(0) = "alphashadow"
    surfaceparmNames(1) = "areaportal"
    surfaceparmNames(2) = "clusterportal"
    surfaceparmNames(3) = "donotenter"
    
    surfaceparmNames(4) = "flesh"
    surfaceparmNames(5) = "fog"
    surfaceparmNames(6) = "lava"
    surfaceparmNames(7) = "metalsteps"
    
    surfaceparmNames(8) = "nodamage"
    surfaceparmNames(9) = "nolight"
    surfaceparmNames(10) = "nodraw"
    surfaceparmNames(11) = "nodrop"
    
    surfaceparmNames(12) = "noimpact"
    surfaceparmNames(13) = "nomarks"
    surfaceparmNames(14) = "nolightmap"
    surfaceparmNames(15) = "nosteps"
    
    surfaceparmNames(16) = "nonsolid"
    surfaceparmNames(17) = "origin"
    surfaceparmNames(18) = "playerclip"
    surfaceparmNames(19) = "slick"
    
    surfaceparmNames(20) = "slime"
    surfaceparmNames(21) = "structural"
    surfaceparmNames(22) = "trans"
    surfaceparmNames(23) = "water"
    
    surfaceparmNames(24) = "sky"
End Sub

Public Sub init_shaderlib()
    numShaders = -1
    numPaths = -1
    
    initfuncs
End Sub
