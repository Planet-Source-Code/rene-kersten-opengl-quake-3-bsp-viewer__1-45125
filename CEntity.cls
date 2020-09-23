VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Type key
    name As String
    value As String
End Type

Private numKeys As Long
Private keys() As key

Public Function findkey(name As String) As String
    Dim i As Long
    
    For i = 0 To numKeys
        If keys(i).name = name Then
            findkey = keys(i).value
            Exit Function
        End If
    Next i
End Function

Public Sub parse(line As String)
    Dim words() As String
    
    line = LCase(Trim(line))
    
    If line = "" Then Exit Sub
    
    words = splitword(line)
    
    numKeys = numKeys + 1
    ReDim Preserve keys(numKeys)
    keys(numKeys).name = words(0)
    keys(numKeys).value = words(1)
End Sub

Private Sub Class_Initialize()
    ReDim keys(0)
    numKeys = -1
End Sub
