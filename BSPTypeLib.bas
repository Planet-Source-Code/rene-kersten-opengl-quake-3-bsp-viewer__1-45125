Attribute VB_Name = "BSPTypeLib"

Public Type bsp_face
    textureID As Long
    effect As Long
    type As Long
    vertexIndex As Long
    numVertices As Long
    meshVertexIndex As Long
    meshNumVertices As Long
    lightmapID As Long
    lmapCorner(1) As Long
    lmapSize(1) As Long
    lmapPosition(2) As Single
    lmapBitset(1, 2) As Single
    normal(2) As Single
    size(1) As Long
End Type
