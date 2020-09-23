Attribute VB_Name = "Input"

Option Explicit

Private Type handler
    handlerObject As Object
    handlerFunction As String
    pause As Single
    lastPressed As Single
End Type

Private handlers(255) As handler
Public keys(255) As Integer
Public oldkeys(255) As Integer

Public Const key_toggled = &H1
Public Const key_down = &H80

Public mouseX As Long
Public mouseY As Long

Public mousestate As Long

Public input_handling_enabled As Boolean

Public Function registerHandler(key As Byte, handlerObject As Object, handlerFunction As String, Optional pause As Single = 0, Optional startnow As Boolean = True) As Boolean
    If Not handlers(key).handlerObject Is Nothing Then Exit Function
    
    registerHandler = True
    Set handlers(key).handlerObject = handlerObject
    handlers(key).handlerFunction = handlerFunction
    handlers(key).pause = pause
    If startnow Then handlers(key).lastPressed = GetTickCount
End Function

Public Function unregisterHandler(key As Byte)
    Set handlers(key).handlerObject = Nothing
    handlers(key).handlerFunction = ""
End Function

Public Sub unregisterAll()
    Dim i As Long
    For i = 0 To 255
        unregisterHandler CByte(i)
    Next i
End Sub

Public Sub getInput()
    Dim i As Long
    
    For i = 0 To 255
        oldkeys(i) = keys(i)
        keys(i) = GetKeyState(i)
    Next i
    
    If input_handling_enabled Then
        For i = 0 To 255
            If Not handlers(i).handlerObject Is Nothing Then
                If checkKeyDown(i) Then
                    If GetTickCount - handlers(i).lastPressed > handlers(i).pause Then
                        CallByName handlers(i).handlerObject, handlers(i).handlerFunction, VbMethod
                        handlers(i).lastPressed = GetTickCount
                    End If
                End If
            End If
        Next i
    End If
    
    Dim mousepos As POINTAPI
    GetCursorPos mousepos
    
    mouseX = mousepos.x
    mouseY = mousepos.y
End Sub

Public Function checkKeyDown(KeyCode As Long) As Boolean
    checkKeyDown = (keys(KeyCode) < 0)
End Function

Public Function checkKeyUp(KeyCode As Long) As Boolean
    checkKeyUp = (oldkeys(KeyCode) <> 1) And (keys(KeyCode) = 0)
End Function
