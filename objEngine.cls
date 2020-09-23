VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInputEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub console_change()
    console_enabled = Not console_enabled
    
    If console_enabled Then
        enable_console
    Else: disable_console
    End If
End Sub

Public Sub key_l()
    lighting = Not lighting
End Sub

Public Sub key_t()
    texturesOn = Not texturesOn
End Sub

Public Sub key_p()
    Select Case drawType
        Case 0 'normal
            drawType = 1 'wireframe
        Case 1 'wireframe
            drawType = 2 'dots
        Case 2 'dots
            drawType = 3 'normal
        Case 3 'sillouete
            drawType = 0 'normal
    End Select
End Sub

Public Sub escape_pressed()
    ObjEngine.quit
End Sub

Public Sub initialize()
    input_handling_enabled = True
    unregisterAll
    
    registerHandler vbKeyC, player1, "change_view_type", 150
    
    registerHandler vbKeyUp, player1, "key_up"
    registerHandler vbKeyDown, player1, "key_down"
    registerHandler vbKeyLeft, player1, "key_left"
    registerHandler vbKeyRight, player1, "key_right"
    registerHandler vbKeyA, player1, "key_a"
    registerHandler vbKeyD, player1, "key_d"
    registerHandler vbKeyW, player1, "key_w"
    registerHandler vbKeyS, player1, "key_s"
    
    registerHandler vbKeyL, Me, "key_l", 150
    registerHandler vbKeyT, Me, "key_t", 150
    registerHandler vbKeyP, Me, "key_p", 150
    
    registerHandler vbKeyTab, Me, "console_change", 150
    
    registerHandler vbKeyEscape, Me, "escape_pressed", 150
End Sub

Public Sub quit()
    unregisterAll
End Sub
