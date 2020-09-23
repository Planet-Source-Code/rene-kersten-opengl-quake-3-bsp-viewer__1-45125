Attribute VB_Name = "cvars"

Option Explicit

Public Type var
    name As String
    value As String
End Type

Private numVars As Long
Private vars() As var

Public Sub printVarlist()
    Dim i As Long
    Dim varstr As String
    For i = 0 To numVars
        varstr = """" & vars(i).name & """ = """ & vars(i).value & """"
        Log.Log "", "", varstr
    Next i
End Sub

Public Function getCVar(name As String) As String
    Dim i As Long
    For i = 0 To numVars
        If vars(i).name = name Then
            getCVar = vars(i).value
            Exit Function
        End If
    Next i
End Function

Public Sub setCVar(name As String, value As String)
    Dim i As Long
    For i = 0 To numVars
        If vars(i).name = LCase(name) Then
            vars(i).value = value
            Exit Sub
        End If
    Next i
    
    'cvar is not found
    numVars = numVars + 1
    ReDim Preserve vars(numVars)
    vars(numVars).name = LCase(name)
    vars(numVars).value = value
End Sub

Public Sub loadCFG(path As String)
    Dim filenr As Long
    Dim l As String
    Dim w() As String
    
    'On Error GoTo errhandler
    
    If path = "" Then Exit Sub
    
    path = Replace(path, "/", "\")
    
    If InStr(1, path, ":") < 1 Then
        If LEFT(path, 1) <> "\" Then
            path = App.path & "\" & path
        Else: path = App.path & path
        End If
    End If
    
    Log.Log "", "", ""
    Log.Log "", "", "executing " & RIGHT(path, Len(path) - Len(App.path) - 1)
    
    filenr = FreeFile
    Open path For Input As filenr
        Do Until EOF(filenr)
            Line Input #filenr, l
            
            l = Trim(l)
            If InStr(1, l, "//") > 1 Then
                l = LEFT(l, InStr(1, l, "//"))
            ElseIf LEFT(l, 2) = "//" Then
                l = ""
            End If
            
            If l = "" Then GoTo next_do
            
            currentConsoleLine = l
            console.parseConsoleLine False
next_do:
        Loop
    Close filenr
    
    Exit Sub
    
errhandler:
    If Len(path) > Len(App.path) + 1 Then
        path = RIGHT(path, Len(path) - Len(App.path) - 1)
        Log.Log "", "", "Unable to load " & path
    Else
        Log.Log "", "", "File not found"
    End If
End Sub

Public Sub init_cvars()
    numVars = -1
End Sub
