Attribute VB_Name = "console"

Option Explicit

Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, ByRef lpwTransKey As Byte, ByVal fuState As Long) As Long

Private Const ALOWED_KEYS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890&#,.-/\+*$<>:;@!?=()[]{}_|%^*`~ '""" & vbCr & vbTab & vbBack

Private Const consoleTextScale As Single = 0.5

Private Const func_set = 0
Private Const func_help = 1
Private Const func_cmdlist = 2
Private Const func_get = 3
Private Const func_texturestatus = 4
Private Const func_credits = 5
Private Const func_quit = 6
Private Const func_varlist = 7
Private Const func_exit = 8
Private Const func_shaderlist = 9
Private Const func_load = 10
Private Const func_echo = 11
Private Const func_rem = 12
Private Const func_exec = 13

Public console_enabled As Boolean

Private console_tex As Long

Private oldTime As Long

Public currentConsoleLine As String

Private Type func
    name As String
    numArgs As Long
    argsNames() As String
    help As String
End Type

Private numFuncs As Long
Private functions() As func

Private Sub initfuncs()
    numFuncs = 13
    ReDim functions(numFuncs)
    
    functions(0).name = "set"
    functions(0).numArgs = 1
    ReDim functions(0).argsNames(functions(0).numArgs)
    functions(0).argsNames(0) = "variable"
    functions(0).argsNames(1) = "value"
    functions(0).help = "sets a cvar to a specified value"
    
    functions(1).name = "help"
    functions(1).numArgs = 0
    ReDim functions(1).argsNames(0)
    functions(1).argsNames(0) = "function"
    functions(1).help = "returns a desciption of the specified function"
    
    functions(2).name = "cmdlist"
    functions(2).numArgs = -2
    ReDim functions(2).argsNames(0)
    functions(2).help = "returns a list of all commands"
    
    functions(3).name = "get"
    functions(3).numArgs = 0
    ReDim functions(3).argsNames(0)
    functions(3).argsNames(0) = "variable"
    functions(3).help = "returns the value of the specified cvar"
    
    functions(4).name = "texturestatus"
    functions(4).numArgs = -2
    ReDim functions(4).argsNames(0)
    functions(4).help = "returns the status of the texturemanager"
    
    functions(5).name = "credits"
    functions(5).numArgs = -2
    ReDim functions(5).argsNames(0)
    functions(5).help = "returns the credits"
    
    functions(6).name = "quit"
    functions(6).numArgs = -2
    ReDim functions(6).argsNames(0)
    functions(6).help = "quits the BSP viewer"
    
    functions(7).name = "varlist"
    functions(7).numArgs = -2
    ReDim functions(7).argsNames(0)
    functions(7).help = "displays a list of all cvars"
    
    functions(8).name = "exit"
    functions(8).numArgs = -2
    ReDim functions(8).argsNames(0)
    functions(8).help = "quits the BSP viewer"
    
    functions(9).name = "shaderlist"
    functions(9).numArgs = -2
    ReDim functions(9).argsNames(0)
    functions(9).help = "returns a list or names of loaded shaders"
    
    functions(10).name = "load"
    functions(10).numArgs = 0
    ReDim functions(10).argsNames(0)
    functions(10).argsNames(0) = "map"
    functions(10).help = "loads the specified map"
    
    functions(11).name = "echo"
    functions(11).numArgs = 0
    ReDim functions(11).argsNames(0)
    functions(11).argsNames(0) = "message"
    functions(11).help = "prints to the console"
    
    functions(12).name = "rem"
    functions(12).numArgs = 0
    ReDim functions(12).argsNames(0)
    functions(12).argsNames(0) = "message"
    functions(12).help = "prints to the console"
    
    functions(13).name = "exec"
    functions(13).numArgs = 0
    ReDim functions(13).argsNames(0)
    functions(13).argsNames(0) = "cfg file"
    functions(13).help = "executes a cfg file"
End Sub

Public Sub load_console_gfx()
    console_tex = openQuakeTexture("console")
    objShaderlib.calcShaders
End Sub

Public Sub init_console()
    initfuncs
    load_console_gfx
End Sub

Public Sub enable_console()
    input_handling_enabled = False
    oldTime = GetTickCount
End Sub

Public Sub disable_console()
    input_handling_enabled = True
    unregisterHandler vbKeyTab
    registerHandler vbKeyTab, objInputEngine, "console_change", 150, True
End Sub

Private Function checkArguments(funcnr As Long, words() As String) As Boolean
    If functions(funcnr).numArgs = -1 Then
        checkArguments = True
        Exit Function
    End If
    
    If functions(funcnr).numArgs = UBound(words) - 1 Then
        checkArguments = True
        Exit Function
    End If
    
    If functions(funcnr).numArgs = -2 And UBound(words) = 0 Then
        checkArguments = True
        Exit Function
    End If
    
    'error! invalid number of arguments
    Dim argstr As String
    Dim i As Long
    
    argstr = functions(funcnr).name & " "
    For i = 0 To functions(funcnr).numArgs
        argstr = argstr & "[" & functions(funcnr).argsNames(i) & "] "
    Next i
    
    argstr = LEFT(argstr, Len(argstr) - 1)
    
    Log.Log "", "", "invalid arguments. expected: " & argstr
    Log.Log "", "", "type help [function] for more info"
End Function

Private Function findfunc(name As String) As Long
    Dim i As Long
    For i = 0 To numFuncs
        If functions(i).name = name Then
            findfunc = i
            Exit Function
        End If
    Next i
    findfunc = -1
    
    'error! invalid command
    Log.Log "", "", "invalid command (type cmdlist for a list of all commands)"
End Function

Private Sub parseline(l As String)
    Dim words() As String
    Dim funcnr As Long
    Dim i As Long
    Dim tmpstr As String
    
    Do
        If Not (InStr(1, l, "  ") > 0) Then Exit Do
        l = Replace(l, "  ", " ")
    Loop
    
    If l = "" Then Exit Sub
    
    words = Split(l, " ")
    
    funcnr = findfunc(words(0))
    
    If funcnr = -1 Then Exit Sub
    If Not checkArguments(funcnr, words) Then Exit Sub
    
    Select Case funcnr
        Case func_set
            setCVar words(1), words(2)
            
            If getCVar("showtext") = "1" Then
                SHOW_TEXT = 1
            Else: SHOW_TEXT = 0
            End If
            
        Case func_help
            Dim argFuncnr As Long
            argFuncnr = findfunc(words(1))
            
            If Not argFuncnr = -1 Then
                Log.Log "", "", functions(argFuncnr).help
            End If
            
        Case func_cmdlist
            For i = 0 To numFuncs
                Log.Log "", "", functions(i).name
            Next i
            
        Case func_get
            tmpstr = """" & words(1) & """ = """ & getCVar(words(1)) & """"
            Log.Log "", "", tmpstr
            
        Case func_texturestatus
            logStatus
            
        Case func_credits
            Log.Log "", "", "Concept and code by Rene Kersten"
            Log.Log "", "", ""
            Log.Log "", "", "A big thanks goes out to Eric coleman and Lupin for their help"
            Log.Log "", "", "Also a big thanks to ID software for making quake 3"
            Log.Log "", "", "and gametutorials.com for the BSP file structure info"
            Log.Log "", "", ""
            Log.Log "", "", "Copyright (c) 2002 Rene Kersten"
            Log.Log "", "", "All artwork is property of ID Software"
            
        Case func_quit, func_exit
            quit
            
        Case func_varlist
            printVarlist
            
        Case func_shaderlist
            logShaderList
            
        Case func_load
            setCVar "map", words(1)
            
            BSP_quit  'this will cause all graphics to be deleted
            BSP_init
            
            clearAllShaders
            
            parseline "texturestatus"
            
            console.load_console_gfx
            
            If getCVar("gametype") = "mohaa" Then
                Log.Log "", "", "loading map"
                loadBSP getCVar("map"), "2015", 19
            ElseIf getCVar("gametype") = "q3a" Then
                Log.Log "", "", "loading map"
                loadBSP getCVar("map")
            Else: Log.Log "", "", "invalid game type"
            End If
            
            objShaderlib.calcShaders
            
            objBSP.parseSurfaceInfo
            
        Case func_echo, func_rem
            For i = 1 To UBound(words)
                tmpstr = tmpstr & words(i) & " "
            Next i
            
            If tmpstr = "" Then Exit Sub
            tmpstr = LEFT(tmpstr, Len(tmpstr) - 1)
            
            Log.Log "", "", tmpstr
            
        Case func_exec
            cvars.loadCFG words(1)
            
    End Select
End Sub

Public Sub parseConsoleLine(Optional drawLine As Boolean = True)
    If drawLine Then Log.Log "", "", ">" & currentConsoleLine
    
    parseline currentConsoleLine
    
    currentConsoleLine = ""
End Sub

Private Function keypress_valid(i As String) As Boolean
    keypress_valid = InStr(1, LCase(ALOWED_KEYS), i)
End Function

Private Sub parseKeyPress(i As Long, state As Long)
    Select Case i
        Case vbKeyBack
            If Len(currentConsoleLine) = 1 Then
                currentConsoleLine = ""
            ElseIf currentConsoleLine <> "" Then
                currentConsoleLine = LEFT(currentConsoleLine, Len(currentConsoleLine) - 1)
            End If
            
        Case vbKeyReturn
            parseConsoleLine
            
        Case vbKeyTab
            'do nothing
            
        Case Else
            Dim c(3) As Byte
            Dim tmpkeys(255)
            Dim x As Long
            Dim tmpstate As Long
            Dim resultStr As String * 1
            
            For x = 0 To 255
                If keys(x) = key_down Then
                    tmpstate = 255
                ElseIf keys(x) = key_toggled Then
                    tmpstate = 255
                Else: tmpstate = 0
                End If
                
                tmpkeys(x) = tmpstate
            Next x
            
            ToAscii i, 255, tmpkeys(0), c(0), 0
            
            resultStr = Chr$(c(0))
            
            If keypress_valid(resultStr) Then
                If (keys(vbKeyShift) And key_down) Or (keys(vbKeyCapital) And key_toggled) Then
                    currentConsoleLine = currentConsoleLine & UCase(resultStr)
                Else
                    currentConsoleLine = currentConsoleLine & LCase(resultStr)
                End If
            End If
            
    End Select
End Sub

Public Sub console_update()
    Dim i As Long
    
    getInput
    
    For i = 0 To 255
        If (keys(i) And key_down) And Not (oldkeys(i) And key_down) Then
            parseKeyPress i, CLng(keys(i))
        End If
    Next i
    
    If GetTickCount - oldTime < 150 Then Exit Sub
    
    If keys(vbKeyTab) And key_down Then
        console_enabled = False
        disable_console
    End If
End Sub

Private Sub console_draw_text()
    Dim filenr As Long
    Dim numL As Long
    Dim l() As String
    Dim tmpLine As String
    Dim oldshowtext As Long
    
    On Error GoTo errhandler
    
    oldshowtext = SHOW_TEXT
    SHOW_TEXT = 1
    
    filenr = FreeFile
    Open App.path & "\Log.txt" For Input As filenr
        Do Until EOF(filenr)
            Line Input #filenr, tmpLine
            
            numL = numL + 1
            ReDim Preserve l(numL)
            l(numL) = tmpLine
        Loop
    Close filenr
    
    glColor3f 0.75, 0.75, 0.75
    glLoadIdentity
    
    Dim i As Long
    For i = 0 To 10 / consoleTextScale
        If numL > 10 / consoleTextScale Then
            'i = 18 - i
            glTextOut l(numL - i), -9, CSng(i) * consoleTextScale + 0.2 + consoleTextScale, consoleTextScale
        Else
            glTextOut l(i), -9, CSng(i) / 2 + 0.2 + consoleTextScale, consoleTextScale
        End If
    Next i
    
    glTextOut ">" & currentConsoleLine, -9, 0.2, consoleTextScale
    
    SHOW_TEXT = oldshowtext
    Exit Sub
    
errhandler:
    SHOW_TEXT = oldshowtext
End Sub

Public Sub console_draw()
    Dim i As Long
    
    If Not console_enabled Then Exit Sub
    
    console_update
    
    If console_tex = -1 Then Exit Sub
    If shaderIndices(console_tex) = -1 Then Exit Sub
    
    glLoadIdentity
    
    resetStates
    glDisable GL_CULL_FACE
    
    glColor4f 1, 1, 1, 1
    glTranslatef 0, 0, -10
    
    For i = 0 To parsedShaders(shaderIndices(console_tex)).numStages
        With parsedShaders(shaderIndices(console_tex)).stages(i)
            'we ignore the culling part
            
            If .textureID(0) = -1 Then
                GoTo next_i
            Else
                glEnable GL_TEXTURE_2D
                glBindTexture GL_TEXTURE_2D, .textureID(0)
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
            Else: glDisable GL_BLEND
            End If
            
            'draw the geometry
            glBegin GL_QUADS
                glTexCoord2f .scrollX * t, .scrollY * t
                glVertex3f -10, 10, 0
                
                glTexCoord2f .scrollX * t, .scaley + .scrollY * t
                glVertex3f -10, 0, 0
                
                glTexCoord2f .scalex + .scrollX * t, .scaley + .scrollY * t
                glVertex3f 10, 0, 0
                
                glTexCoord2f .scalex + .scrollX * t, .scrollY * t
                glVertex3f 10, 10, 0
            glEnd
        End With
next_i:
    Next i
    
    resetStates
    glDisable GL_TEXTURE_2D
    glDisable GL_CULL_FACE
    
    glColor3f 1, 1, 0
    glBegin GL_LINES
        glVertex3f -10, 0, 0
        glVertex3f 10, 0, 0
    glEnd
    
    console_draw_text
End Sub
