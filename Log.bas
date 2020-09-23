Attribute VB_Name = "Log"
Option Explicit

Public Sub initializeLog()
    Dim file As String
    Dim filenr As Long
    
    filenr = FreeFile
    
    file = App.path & "\log.txt"
    
    Open file For Output As filenr
    Close filenr
    
    Kill App.path & "\log.txt"
End Sub

Public Sub Log(strModule As String, strSubFunc As String, strEvent As String)
    Dim strCModule As String * 20, strCSubFunc As String * 25
    Dim strLog As String
    
    strCModule = Space(21 - Len(strModule))
    strCSubFunc = Space(26 - Len(strSubFunc))
    
    strCModule = strModule
    strCSubFunc = strSubFunc
    
    If Err.Number > 0 Then
        strLog = "Time: " & Time & ".   " & strCModule & "- " & strCSubFunc & "- Error Number: " & Err.Number & ". Error Description: " & Err.Description & "."
    Else
        ''There is no error, now log the event.
        strLog = "Time: " & Time & ".   " & strCModule & "- " & strCSubFunc & "- " & strEvent
    End If
    saveLog strEvent
End Sub

Private Sub saveLog(strLog As String)
    Dim filenr As Long
    filenr = FreeFile
    Open (App.path & "\Log.txt") For Append As filenr
        Print #filenr, strLog
    Close filenr
End Sub
