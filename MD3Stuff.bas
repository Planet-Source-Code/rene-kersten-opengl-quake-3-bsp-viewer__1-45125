Attribute VB_Name = "MD3Stuff"

Public Type MD3Tag
    name As String * 64
    origin As vect
    axis(2) As vect
End Type
