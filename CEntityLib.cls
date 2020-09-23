VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEntityLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private numEntities As Long
Private entities() As New CEntity

Private Sub parseInfoPlayerStart()
    Dim i As Long
    Dim pos As vect
    Dim str As String
    
    Dim words() As String
    
    For i = numEntities To 0 Step -1
        If entities(i).findkey("classname") = "info_player_deathmatch" Then
            str = entities(i).findkey("origin")
            words = Split(str, " ")
            pos.x = CSng(words(0))
            pos.y = CSng(words(1))
            pos.z = CSng(words(2))
            
            ObjEngine.setpos pos
        End If
    Next i
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

Private Sub init_entityLib()
    ReDim entities(0)
    numEntities = -1
End Sub
