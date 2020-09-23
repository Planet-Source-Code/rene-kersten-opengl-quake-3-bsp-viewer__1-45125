VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   412
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox dir 
      Height          =   480
      Left            =   2040
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public focus As Boolean

Private Sub Form_GotFocus()
    focus = True
End Sub

Private Sub Form_Load()
    Show
    ObjEngine.start
End Sub

Private Sub Form_LostFocus()
    focus = False
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mousestate = Button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ObjEngine.quit
End Sub
