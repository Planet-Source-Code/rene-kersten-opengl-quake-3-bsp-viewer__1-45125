Attribute VB_Name = "typelib"

Option Explicit

Public Enum animationState
    ANIM_PLAY = 1
    ANIM_PAUSE = 2
    ANIM_STOP = 4
End Enum

Public Enum headerinfo
    enum_addButton
    enum_showButton
    enum_hideButton
    enum_changeButton
    enum_deleteButton
    
    enum_addLabel
    enum_showLabel
    enum_hideLabel
    enum_changeLabel
    enum_deleteLabel
    
    enum_addPicture
    enum_showPicture
    enum_hidePicture
    enum_changePicture
    enum_deletePicture
End Enum

Public Type labeldesc
    text As String
    x As Single
    y As Single
    visible As Boolean
    
    usecolor As Boolean
    
    colorR As Single
    colorG As Single
    colorB As Single
End Type

Public Type animationSettings
    fps As Single
    currentFrame As Long
    state As animationState
    lastupdated As Long
    loop As Boolean
End Type

Public Type picInfo
    animsettings As animationSettings
    rotx As Single
    roty As Single
    rotz As Single
    angleX As Single
    scalex As Single
    scaley As Single
    x As Single
    y As Single
    z As Single
    Width As Single
    Height As Single
    picNr As Long
End Type

Public Type laser
    x As Single
    y As Single
    
    vectX As Single
    vectY As Single
    
    life As Long
End Type
