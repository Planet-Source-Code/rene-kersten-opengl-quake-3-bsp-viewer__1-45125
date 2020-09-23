Attribute VB_Name = "TextureLib"
'************************************************
'*  TextureLib.bas                              *
'*                                              *
'* By: W-Buffer                                 *
'* Web: www.lunarpages.com/istudios/            *
'* Mail: chadruva@hotmail.com                   *
'*                                              *
'* Notes: Do whatever you want with this bas    *
'*        (Steal, Copy, Etc.), as long this     *
'*        note stays here.                      *
'************************************************

Option Explicit

Public Const TFT_URGB = 2       'Uncompressed RGB
Public Const TFT_UBW = 3        'Uncompressed Black & White
Public Const TFT_RLERGB = 10    'RLE Compressed RGB
Public Const TFT_CBW = 11       'Compressed Black & White

Public Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type
Public Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Public Type TARGAFILEHEADER
    imageIDLength As Byte
    colorMapType As Byte
    imageTypeCode As Byte
    
    colorMapOrigin As Integer
    colorMapLenght As Integer
    colorMapEntrySize As Byte
    
    imageXOrigin As Integer
    imageYOrigin As Integer
    imageWidth As Integer
    imageHeight As Integer
    bitCount As Byte
    imageDescriptor As Byte
End Type

Public Type EGLTEXTUREDATA
    Height As Long
    Width As Long
    bpp As Integer
    imgSizes As Long
    tdata() As Byte
End Type

Public Function auxTexture(texture As EGLTEXTUREDATA, Optional ByVal MagFilter As Long = GL_NEAREST, Optional ByVal MinFilter As Long = GL_NEAREST) As Long
    Dim tID As Long
    Dim tformat As Long
    
    If texture.bpp = 32 Then
        tformat = GL_RGBA
    ElseIf texture.bpp = 24 Then
        tformat = GL_RGB
    ElseIf texture.bpp = 8 Then
        tformat = GL_LUMINANCE
    End If
    
    glGenTextures 1, tID
    glBindTexture GL_TEXTURE_2D, tID
    
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, MagFilter
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, MinFilter
    
    glTexImage2D GL_TEXTURE_2D, 0, tformat, texture.Width, texture.Height, 0, tformat, GL_UNSIGNED_BYTE, texture.tdata(0)
    auxTexture = tID
End Function

Public Function auxMipmaps(texture As EGLTEXTUREDATA, Optional ByVal MagFilter As Long = GL_NEAREST, Optional ByVal MinFilter As Long = GL_NEAREST_MIPMAP_NEAREST, Optional repeat As Boolean = True) As Long
    Dim tID As Long
    Dim tformat As Long
    
    If texture.bpp = 32 Then
        tformat = GL_RGBA
    ElseIf texture.bpp = 24 Then
        tformat = GL_RGB
    ElseIf texture.bpp = 8 Then
        tformat = GL_LUMINANCE
    End If
    
    glGenTextures 1, tID
    glBindTexture GL_TEXTURE_2D, tID
    
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, MagFilter
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, MinFilter
    
    If Not repeat Then
        glTexEnvf GL_TEXTURE_2D, GL_TEXTURE_ENV_MODE, GL_MODULATE
    End If
    
    gluBuild2DMipmaps GL_TEXTURE_2D, tformat, texture.Width, texture.Height, tformat, GL_UNSIGNED_BYTE, texture.tdata(0)
    auxMipmaps = tID
End Function

Public Sub auxAlphaCK(texture As EGLTEXTUREDATA, ByVal KeyR As Byte, ByVal KeyG As Byte, ByVal KeyB As Byte)
    Dim i As Long
    Dim t As Long
    Dim tarray() As Byte
    
    'Check if the texture data applyes for this function
    If texture.bpp <> 24 Then Exit Sub
    
    'Now the texture will be 32 bits per pixel
    texture.bpp = 32
    'Allocate memory for our temporal array
    ReDim tarray(texture.Height * texture.Width * 4)
    
    For i = 0 To (texture.imgSizes) - 1 Step 3
        tarray(t) = texture.tdata(i)
        tarray(t + 1) = texture.tdata(i + 1)
        tarray(t + 2) = texture.tdata(i + 2)
        
        If texture.tdata(i) = KeyR And texture.tdata(i + 1) = KeyG And texture.tdata(i + 2) = KeyB Then
            'If the pixel color is equal to color key then
            'set the Alpha Channel to 0 (transparent)
            tarray(t + 3) = 0
        Else
            'Else set the Alpha Channel to 255 (Opaque)
            tarray(t + 3) = 255
        End If
        
        t = t + 4
    Next i
    
    'Set the new texture Sizes
    texture.imgSizes = texture.Height * texture.Width * 4
    
    'Allocate memory for the texture and copy from temp array
    ReDim texture.tdata(texture.imgSizes)
    CopyMemory texture.tdata(0), tarray(0), texture.imgSizes
    
    'Erase out temporal array
    Erase tarray
End Sub

Public Function auxLoadBitmap(filename As String) As EGLTEXTUREDATA
    Dim filet As Long
    Dim bm As BITMAPINFOHEADER
    Dim bf As BITMAPFILEHEADER
    Dim i As Long
    Dim c As Long
    Dim tbyte As Byte
    
    filet = FreeFile
    Open filename For Binary Access Read As filet
    
    Get filet, , bf
    Get filet, , bm
    
    auxLoadBitmap.bpp = bm.biBitCount
    auxLoadBitmap.Height = bm.biHeight
    auxLoadBitmap.Width = bm.biWidth
    
    If bm.biSizeImage = 0 Then
        auxLoadBitmap.imgSizes = bm.biHeight * bm.biWidth * (bm.biBitCount / 8)
    Else
        auxLoadBitmap.imgSizes = bm.biSizeImage
    End If
    
    ReDim auxLoadBitmap.tdata(auxLoadBitmap.imgSizes)
    Get filet, , auxLoadBitmap.tdata
    
    If auxLoadBitmap.bpp = 24 Then
        c = auxLoadBitmap.imgSizes - 1
        For i = 0 To c Step 3
            tbyte = auxLoadBitmap.tdata(i)
            auxLoadBitmap.tdata(i) = auxLoadBitmap.tdata(i + 2)
            auxLoadBitmap.tdata(i + 2) = tbyte
        Next i
    End If
    
    Close filet
End Function

Private Function DecompressTGA(RLEStream() As Byte, bits As Long, Höhe As Long, Breite As Long) As Byte()
    Dim InitSize As Long
    Dim Temp() As Byte
    Dim n As Long
    Dim k As Boolean
    Dim b As Long
    Dim l As Long
    Dim fertig As Long
    Dim z As Long
    Dim Länge As Long
    Dim Byteanzahl As Long
    
    Byteanzahl = bits / 8
       InitSize = CLng(Höhe * Breite * Byteanzahl)
    ReDim Temp(0 To InitSize)
    Do While fertig < InitSize
    If l > UBound(RLEStream) Then GoTo Ende
        z = 0
        If RLEStream(l) > 127 Then
            n = RLEStream(l) - 127
            For b = 0 To n - 1
            
            CopyMemory Temp(fertig), RLEStream(l + 1), Byteanzahl
                           fertig = fertig + Byteanzahl
            
            Next b
        k = True
        Else
           n = RLEStream(l) + 1
           Länge = n * Byteanzahl

            CopyMemory Temp(fertig), RLEStream(l + 1), Länge
            
            k = False
            z = z + Länge

            fertig = fertig + z
            End If
            If k = True Then
                l = Byteanzahl + 1 + l
                n = z + (n * Byteanzahl) + 1
            Else
                l = (n * Byteanzahl + 1) + l
                n = z + n
            End If
    Loop
Ende:

    DecompressTGA = Temp
    
End Function

Public Function auxLoadTarga(filename As String, Optional FlipColors As Boolean = True, Optional turnUpsidedown As Boolean = True) As EGLTEXTUREDATA
    Dim filet As Long
    Dim tf As TARGAFILEHEADER
    Dim i As Long
    Dim e As Long
    Dim d As Long
    Dim tbyte As Byte
    Dim ColorMap() As Byte
    Dim tdata() As Byte
    
    On Error GoTo errhandler
    
    filet = FreeFile
    Open filename For Binary Access Read As filet
    
    Get filet, , tf
    
    auxLoadTarga.bpp = tf.bitCount
    auxLoadTarga.Height = tf.imageHeight
    auxLoadTarga.Width = tf.imageWidth
    
    Dim s As Double
    s = (CDbl(tf.imageHeight) * CDbl(tf.imageWidth) * CDbl((tf.bitCount / 8)))
    
    auxLoadTarga.imgSizes = CLng(s)
    
    ReDim auxLoadTarga.tdata(auxLoadTarga.imgSizes)
    ReDim tdata(auxLoadTarga.imgSizes)
    'Get filet, , tdata
    Get filet, , auxLoadTarga.tdata
    
    If tf.imageTypeCode = 10 Then
        auxLoadTarga.tdata = DecompressTGA(auxLoadTarga.tdata, CLng(tf.bitCount), CLng(tf.imageHeight), CLng(tf.imageWidth))
        auxLoadTarga.imgSizes = (tf.bitCount / 8) * tf.imageWidth * tf.imageHeight - 1
    End If
    
    Dim x As Long, y As Long, k As Long
        
    For i = 0 To auxLoadTarga.imgSizes - 1 Step auxLoadTarga.bpp / 8
        tbyte = auxLoadTarga.tdata(i)
        auxLoadTarga.tdata(i) = auxLoadTarga.tdata(i + 2)
        auxLoadTarga.tdata(i + 2) = tbyte
    Next i
    Close filet
    
    Exit Function
    
errhandler:
End Function
