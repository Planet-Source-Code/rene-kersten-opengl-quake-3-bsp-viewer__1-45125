Attribute VB_Name = "objBSP"

Option Explicit

Const DRAW_BB = 1

'declarations
Private Enum enumBSPLumpType
    lump_textures = 1 'done
    lump_vertices = 10 'done
    lump_meshvertices = 11 'done
    lump_lightmaps = 14 'done
    lump_faces = 13 'done including bezier patches
    lump_models = 7 'done
    
    lump_shaders = 12 'some functions need implementation
    lump_lightvolumes = 15 'how do i implement this?
    
    lump_entities = 0 'framework is done but needs implementation
    
    'painter's algorithm implemented
    lump_planes = 2 'done
    lump_nodes = 3 'done
    lump_leafs = 4 'done
    lump_leafsfaces = 5 'done
    lump_LeafsBrushes = 6 'done
    lump_brushes = 8 'done
    lump_brushessides = 9 'done
    lump_visdata = 16 'done
End Enum

Private Type bsp_lump
    offset As Long
    length As Long
End Type

Private Type bsp_header
    id As String * 4
    version As Long
End Type

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

Public Type bsp_vertex
    position(2) As Single
    texture(1) As Single
    lightmap(1) As Single
    normal(2) As Single
    color(3) As Byte
End Type

Private Type bsp_texture
    path As String * 64
    flags As Long
    contents As Long
End Type

Private Const LIGHTMAPSIZE As Long = 49151

Private Type bsp_lightmap
    color(LIGHTMAPSIZE) As Byte
End Type

Private Type bsp_node
    plane As Long
    FRONT As Long
    BACK As Long
    mins(2) As Long
    maxs(2) As Long
End Type

Private Type bsp_leaf
    cluster As Long
    area As Long
    mins(2) As Long
    maxs(2) As Long
    leafFace As Long
    numLeaffaces As Long
    leafBrush As Long
    numLeafbrushes As Long
End Type

Private Type bsp_plane
    normals(2) As Single
    d As Single
End Type

Private Type bsp_visdata
    numClusters As Long
    bytesPerCluster As Long
    bytes() As Byte
End Type

Private Type bsp_entities
    entities As String
End Type

Private Type bsp_brush
    brushside As Long
    numBrushsides As Long
    textureID As Long
End Type

Private Type bsp_brushSide
    plane As Long
    textureID As Long
End Type

Private Type bsp_model
    mins(2) As Single
    maxs(2) As Single
    faceIndex As Long
    numFaceIndex As Long
    brushIndex As Long
    numBrushIndex As Long
End Type

Private Type bsp_shader
    path As String * 64
    brushIndex As Long
    unknown As Long
End Type

Private Type bsp_lightdata
    ambient(2) As Byte
    directional(2) As Byte
    direction(1) As Byte
End Type

Private Type bsp_patch_part
    l As Long
    w As Long
    verts() As vect
    textures() As vect2
    lightmaps() As vect2
End Type

Private Type bsp_patch
    patches() As bsp_patch_part
    patchesX As Long
    patchesY As Long
End Type

Private filenr As Long
Private lumpData(16) As bsp_lump

Private numTextures As Long
Private textureData() As bsp_texture
Private textureIDs() As Long

Private numPlanes As Long
Private planes() As bsp_plane

Private numNodes As Long
Private nodes() As bsp_node

Private numLeafs As Long
Private leafs() As bsp_leaf

Private numLeafsfaces As Long
Private leafsfaces() As Long

Private numLeafsbrushes As Long
Private leafsbrushes() As Long

Private numModels As Long
Private models() As bsp_model

Private numBrushes As Long
Private brushes() As bsp_brush

Private numBrushsides As Long
Private brushside() As bsp_brushSide

Public numVertices As Long
Public vertices() As bsp_vertex

Public numMeshVertices As Long
Public meshvertices() As Long

Private numShaders As Long
Private shaders() As bsp_shader

Private numFaces As Long
Private faces() As bsp_face
Private faceDrawn() As Boolean
Private surfaceFlags() As Long
Private facetransparant() As Boolean

Private numLightvolumes As Long
Private lightvolumes() As bsp_lightdata

Private numLightmaps As Long
Private lightmaps() As Long

Private visdata As bsp_visdata

Private factor As Single

Private levelToWidth(6) As Long

Private numPatches As Long
Private patches() As bsp_patch
Private patchIndex() As Long
    
Private loadscreen As Long

Private currentcluster As Long

Private nodeDepth As Long

Public mapLoaded As Boolean

'Option Explicit

'patch building code

Private Function Factorial(n As Long) As Single
    Dim i As Long
    
    Factorial = 1
    For i = 1 To n
        Factorial = Factorial / i
    Next i
End Function

'return double because of standard index form...
Private Function nCr(n As Long, r As Long) As Single
    '1/r!(n-r)! / (1/n!)
    '(1/r! * 1/(n-r)!) / (1/n!)
    nCr = (Factorial(r) * Factorial(n - r)) / (Factorial(n))
End Function

Private Function NthDPatchLTC(x As Single, y As Single, cp() As bsp_vertex) As vect2
    Dim r As Long, i As Long
    Dim Expansion1 As vect2
    Dim MainExpansion As vect2
    Dim a As Single, b As Single, c As Single, d As Single
    
    a = x: b = 1 - x
    c = y: d = 1 - y
    
    For r = 0 To 2
        Expansion1.x = 0
        Expansion1.y = 0
        For i = 0 To 2
            'calculate the  expansion, and store the result
            Expansion1.x = Expansion1.x + (cp(r, i).lightmap(0) * nCr(2, i) * (c ^ (2 - i)) * (d ^ i))
            Expansion1.y = Expansion1.y + (cp(r, i).lightmap(1) * nCr(2, i) * (c ^ (2 - i)) * (d ^ i))
        Next i
        MainExpansion.x = MainExpansion.x + (nCr(2, r) * (a ^ (2 - r)) * (b ^ r)) * Expansion1.x
        MainExpansion.y = MainExpansion.y + (nCr(2, r) * (a ^ (2 - r)) * (b ^ r)) * Expansion1.y
    Next r
    
    NthDPatchLTC = MainExpansion
End Function

Private Function NthDPatchTC(x As Single, y As Single, cp() As bsp_vertex) As vect2
    Dim r As Long, i As Long
    Dim Expansion1 As vect2
    Dim MainExpansion As vect2
    Dim a As Single, b As Single, c As Single, d As Single
    
    a = x: b = 1 - x
    c = y: d = 1 - y
    
    For r = 0 To 2
        Expansion1.x = 0
        Expansion1.y = 0
        For i = 0 To 2
            'calculate the  expansion, and store the result
            Expansion1.x = Expansion1.x + (cp(r, i).texture(0) * nCr(2, i) * (c ^ (2 - i)) * (d ^ i))
            Expansion1.y = Expansion1.y + (cp(r, i).texture(1) * nCr(2, i) * (c ^ (2 - i)) * (d ^ i))
        Next i
        MainExpansion.x = MainExpansion.x + (nCr(2, r) * (a ^ (2 - r)) * (b ^ r)) * Expansion1.x
        MainExpansion.y = MainExpansion.y + (nCr(2, r) * (a ^ (2 - r)) * (b ^ r)) * Expansion1.y
    Next r
    
    NthDPatchTC = MainExpansion
End Function

Private Function NthDPatch(x As Single, y As Single, cp() As bsp_vertex) As vect
    Dim r As Long, i As Long
    Dim Expansion1 As vect
    Dim MainExpansion As vect
    Dim a As Single, b As Single, c As Single, d As Single
    
    a = x: b = 1 - x
    c = y: d = 1 - y
    
    For r = 0 To 2
        Expansion1.x = 0
        Expansion1.y = 0
        Expansion1.z = 0
        For i = 0 To 2
            'calculate the  expansion, and store the result
            Expansion1.x = Expansion1.x + (cp(r, i).position(0) * nCr(2, i) * (c ^ (2 - i)) * (d ^ i))
            Expansion1.y = Expansion1.y + (cp(r, i).position(1) * nCr(2, i) * (c ^ (2 - i)) * (d ^ i))
            Expansion1.z = Expansion1.z + (cp(r, i).position(2) * nCr(2, i) * (c ^ (2 - i)) * (d ^ i))
        Next i
        MainExpansion.x = MainExpansion.x + (nCr(2, r) * (a ^ (2 - r)) * (b ^ r)) * Expansion1.x
        MainExpansion.y = MainExpansion.y + (nCr(2, r) * (a ^ (2 - r)) * (b ^ r)) * Expansion1.y
        MainExpansion.z = MainExpansion.z + (nCr(2, r) * (a ^ (2 - r)) * (b ^ r)) * Expansion1.z
    Next r
    
    NthDPatch = MainExpansion
End Function

Private Sub buildpatch(face As Long, lod As Long)
    Dim Patch As bsp_face
    Dim cp() As bsp_vertex
    Dim i As Long, j As Long
    Dim result As bsp_patch
    Dim logstr As String
    
    Patch = faces(face)
    
    ReDim cp(Patch.numVertices - 1)
    For i = 0 To Patch.numVertices - 1
        cp(i) = vertices(i + Patch.vertexIndex)
    Next i
    
    Dim x As Long, y As Long
    Dim patchesX As Single, patchesY As Single
    Dim tmpVect As vect
    Dim p(2, 2) As bsp_vertex
    Dim step As Single
    Dim tx As Single, ty As Single
    Dim tmpTC As vect2
    Dim tmpLTC As vect2 'lightmap tex coords
    
    Dim qX As Single, qY As Single
    Dim gX As Long, gY As Long
    
    patchesX = (Patch.size(0) - 1) / 2 - 1
    patchesY = (Patch.size(1) - 1) / 2 - 1
    
    result.patchesX = patchesX
    result.patchesY = patchesY
    
    ReDim result.patches(patchesX, patchesY)
    
    For y = 0 To patchesY
        gY = y * 2
        For x = 0 To patchesX
            gX = x * 2
            
            With result.patches(x, y)
                .w = levelToWidth(lod)
                .l = levelToWidth(lod)
                
                ReDim .verts(.w - 1, .l - 1)
                ReDim .textures(.w - 1, .l - 1)
                ReDim .lightmaps(.w - 1, .l - 1)
                
                p(0, 0) = cp(gY * Patch.size(0) + gX)
                p(0, 1) = cp(gY * Patch.size(0) + gX + 1)
                p(0, 2) = cp(gY * Patch.size(0) + gX + 2)
                
                p(1, 0) = cp((gY + 1) * Patch.size(0) + gX)
                p(1, 1) = cp((gY + 1) * Patch.size(0) + gX + 1)
                p(1, 2) = cp((gY + 1) * Patch.size(0) + gX + 2)
                
                p(2, 0) = cp((gY + 2) * Patch.size(0) + gX)
                p(2, 1) = cp((gY + 2) * Patch.size(0) + gX + 1)
                p(2, 2) = cp((gY + 2) * Patch.size(0) + gX + 2)
                
                For qX = 0 To .w - 1 'x * patchesX To result.W + x * patchesX
                    For qY = 0 To .l - 1 'y * patchesY To result.l + y * patchesY
                        .verts(qX, qY) = NthDPatch(qX / (.w - 1), qY / (.l - 1), p)
                        
                        tmpTC = NthDPatchTC(qX / (.w - 1), qY / (.l - 1), p)
                        
                        .textures(qX, qY).x = tmpTC.x
                        .textures(qX, qY).y = tmpTC.y
                        
                        tmpLTC = NthDPatchLTC(qX / (.w - 1), qY / (.l - 1), p)
                        
                        .lightmaps(qX, qY).x = tmpLTC.x
                        .lightmaps(qX, qY).y = tmpLTC.y
                    Next qY
                Next qX
            End With
        Next x
    Next y
    
    numPatches = numPatches + 1
    ReDim Preserve patches(numPatches)
    patches(numPatches) = result
End Sub

'loading
Private Function getHeader(id As String, version As Long) As Boolean
    Dim header As bsp_header
    Get filenr, , header
    
    If header.id <> id Then
        Log.Log "CBSP", "getHeader", "Invalid bsp file"
        Exit Function
    End If
    If header.version <> version Then
        Log.Log "CBSP", "getheader", "invalid bsp file"
        Exit Function
    End If
    
    getHeader = True
End Function

Private Sub getLumpData()
    Dim i As Long
    For i = 0 To 16
        Get filenr, , lumpData(i)
    Next i
End Sub

Private Sub getEntities()
    Dim entitiesString As String
    Dim entitiesFilenr As Long
    
    If lumpData(lump_entities).offset + 1 < 0 Then Exit Sub
    
    entitiesFilenr = FreeFile
    
    Seek filenr, lumpData(lump_entities).offset + 1
    
    entitiesString = Space(lumpData(lump_entities).length)
    
    Get filenr, , entitiesString
    
    entitiesString = Replace(entitiesString, Chr(10), vbNewLine)
    
    Open App.path & "\entities.txt" For Output As entitiesFilenr
        Print #entitiesFilenr, entitiesString
    Close entitiesFilenr
    
    objEntityLib.init_entityLib
    objEntityLib.parseEntities App.path & "\entities.txt"
    
    Kill App.path & "\entities.txt"
End Sub

Private Sub getTextures()
    Dim tmpTexture As bsp_texture
    Dim texdata As EGLTEXTUREDATA
    Dim texPath As String
    Dim i As Long
    
    numTextures = lumpData(lump_textures).length / Len(tmpTexture) - 1
    
    Seek filenr, lumpData(lump_textures).offset + 1
    
    'clearAll
    
    ReDim textureData(numTextures)
    ReDim textureIDs(numTextures)
    For i = 0 To numTextures
        drawloadscreen
        
        Get filenr, , textureData(i)
        
        texPath = Replace(LEFT(textureData(i).path, InStr(1, textureData(i).path, Chr(0)) - 1), "/", "\")
        If texPath <> "noshader" And texPath <> "clip" And texPath <> "" Then
            textureIDs(i) = openQuakeTexture(texPath)
            If Not textureIDs(i) = -1 Then
                Log.Log "", "", "texture loaded: " & textureData(i).path
            Else: Log.Log "", "", "invalid texture: " & textureData(i).path
            End If
        Else: textureIDs(i) = 0
        End If
    Next i
End Sub

'this function makes a list of faces that are (partially) transparant and thus, should be drawed last
Public Sub parseSurfaceInfo()
    Dim i As Long, j As Long
    
    If Not mapLoaded Then Exit Sub
    
    For i = 0 To numFaces
        If Not faces(i).textureID > UBound(textureIDs) Then
            If textureIDs(faces(i).textureID) > -1 Then
                If shaderIndices(textureIDs(faces(i).textureID)) > -1 Then
                    Dim index As Long
                    index = shaderIndices(textureIDs(faces(i).textureID))
                    
                    If shaderShouldBeParsed(index) Then
                        surfaceFlags(i) = parsedShaders(index).surfaceparms
                        
                        facetransparant(i) = surfaceFlags(i) And surfaceparm_trans
                        If surfaceFlags(i) And surfaceparm_trans Then
                            facetransparant(i) = True
                            GoTo next_i
                        End If
                        
                        For j = 0 To parsedShaders(index).numStages
                            If parsedShaders(index).stages(j).alphafunc Or parsedShaders(index).stages(j).shouldBlend Then
                                facetransparant(i) = True
                                GoTo next_i
                            End If
                        Next j
                    End If
                End If
            End If
        End If
next_i:
    Next i
End Sub

Public Sub getPlanes()
    Dim tmpplane As bsp_plane
    Dim i As Long
    
    numPlanes = lumpData(lump_planes).length / Len(tmpplane) - 1
    
    Seek filenr, lumpData(lump_planes).offset + 1
    
    Dim m As matrix
    Dim r(2) As Single
    Dim v As vector
    r(0) = 270
    Math3d.mat_loadIdentity m
    Math3d.mat_setRotationDegrees m, r
    
    ReDim planes(numPlanes)
    For i = 0 To numPlanes
        Get filenr, , planes(i)
        
        v.pos(0) = planes(i).normals(0)
        v.pos(1) = planes(i).normals(1)
        v.pos(2) = planes(i).normals(2)
        Math3d.vect_transform v, m
        planes(i).normals(0) = v.pos(0)
        planes(i).normals(1) = v.pos(1)
        planes(i).normals(2) = v.pos(2)
    Next i
End Sub

Private Sub getNodes()
    Dim tmpNode As bsp_node
    Dim i As Long
    
    numNodes = lumpData(lump_nodes).length / Len(tmpNode) - 1
    
    Seek filenr, lumpData(lump_nodes).offset + 1
    
    Dim m As matrix
    Dim r(2) As Single
    Dim v As vector
    r(0) = 270
    Math3d.mat_loadIdentity m
    Math3d.mat_setRotationDegrees m, r
    
    ReDim nodes(numNodes)
    For i = 0 To numNodes
        Get filenr, , nodes(i)
        
        v.pos(0) = nodes(i).mins(0)
        v.pos(1) = nodes(i).mins(1)
        v.pos(2) = nodes(i).mins(2)
        Math3d.vect_transform v, m
        nodes(i).mins(0) = v.pos(0)
        nodes(i).mins(1) = v.pos(1)
        nodes(i).mins(2) = v.pos(2)
        
        v.pos(0) = nodes(i).maxs(0)
        v.pos(1) = nodes(i).maxs(1)
        v.pos(2) = nodes(i).maxs(2)
        Math3d.vect_transform v, m
        nodes(i).maxs(0) = v.pos(0)
        nodes(i).maxs(1) = v.pos(1)
        nodes(i).maxs(2) = v.pos(2)
    Next i
End Sub

Private Sub getLeafs()
    Dim tmpLeaf As bsp_leaf
    Dim i As Long
    
    numLeafs = lumpData(lump_leafs).length / Len(tmpLeaf)
    
    Seek filenr, lumpData(lump_leafs).offset + 1
    
    Dim m As matrix
    Dim r(2) As Single
    Dim v As vector
    r(0) = 270
    Math3d.mat_loadIdentity m
    Math3d.mat_setRotationDegrees m, r
    
    ReDim leafs(numLeafs)
    ReDim leafDistance(numLeafs)
    For i = 0 To numLeafs
        Get filenr, , leafs(i)
        
        v.pos(0) = leafs(i).mins(0)
        v.pos(1) = leafs(i).mins(1)
        v.pos(2) = leafs(i).mins(2)
        Math3d.vect_transform v, m
        leafs(i).mins(0) = v.pos(0)
        leafs(i).mins(1) = v.pos(1)
        leafs(i).mins(2) = v.pos(2)
        
        v.pos(0) = leafs(i).maxs(0)
        v.pos(1) = leafs(i).maxs(1)
        v.pos(2) = leafs(i).maxs(2)
        Math3d.vect_transform v, m
        leafs(i).maxs(0) = v.pos(0)
        leafs(i).maxs(1) = v.pos(1)
        leafs(i).maxs(2) = v.pos(2)
    Next i
End Sub

Private Sub getLeafsFaces()
    Dim tmpLeafface As Long
    Dim i As Long
    
    numLeafsfaces = lumpData(lump_leafsfaces).length / Len(tmpLeafface)
    
    Seek filenr, lumpData(lump_leafsfaces).offset + 1
    
    ReDim leafsfaces(numLeafsfaces)
    For i = 0 To numLeafsfaces
        Get filenr, , leafsfaces(i)
    Next i
End Sub

Private Sub getLeafsBrushes()
    Dim tmpLeafbrush As Long
    Dim i As Long
    
    numLeafsbrushes = lumpData(lump_LeafsBrushes).length / Len(tmpLeafbrush)
    
    Seek filenr, lumpData(lump_LeafsBrushes).offset + 1
    
    ReDim leafsbrushes(numLeafsbrushes)
    For i = 0 To numLeafsbrushes
        Get filenr, , leafsbrushes(i)
    Next i
End Sub

Public Sub getModels()
    Dim tmpmodel As bsp_model
    Dim i As Long
    
    numModels = lumpData(lump_models).length / Len(tmpmodel)
    
    Seek filenr, lumpData(lump_models).offset + 1
    
    ReDim models(numModels)
    For i = 0 To numModels
        Get filenr, , models(i)
    Next i
End Sub

Private Sub getBrushes()
    Dim tmpbrush As bsp_brush
    Dim i As Long
    
    numBrushes = lumpData(lump_brushes).length / Len(tmpbrush)
    
    Seek filenr, lumpData(lump_brushes).offset + 1
    
    ReDim brushes(numBrushes)
    For i = 0 To numBrushes
        Get filenr, , brushes(i)
    Next i
End Sub

Private Sub getBrushessides()
    Dim tmpbrushside As bsp_brushSide
    Dim i As Long
    
    numBrushsides = lumpData(lump_brushessides).length / Len(tmpbrushside)
    
    Seek filenr, lumpData(lump_brushessides).offset + 1
    
    ReDim brushside(numBrushsides)
    For i = 0 To numBrushsides
        Get filenr, , brushside(i)
    Next i
End Sub

Private Sub getVertices()
    Dim tmpvertex As bsp_vertex
    Dim i As Long
    
    numVertices = lumpData(lump_vertices).length / Len(tmpvertex)
    
    Seek filenr, lumpData(lump_vertices).offset + 1
    
    Dim m As matrix
    Dim r(2) As Single
    Dim v As vector
    r(0) = 270
    Math3d.mat_loadIdentity m
    Math3d.mat_setRotationDegrees m, r
    
    ReDim vertices(numVertices)
    For i = 0 To numVertices
        Get filenr, , vertices(i)
        vertices(i).texture(1) = 1 - vertices(i).texture(1)
        
        v.pos(0) = vertices(i).position(0)
        v.pos(1) = vertices(i).position(1)
        v.pos(2) = vertices(i).position(2)
        Math3d.vect_transform v, m
        vertices(i).position(0) = v.pos(0)
        vertices(i).position(1) = v.pos(1)
        vertices(i).position(2) = v.pos(2)
    Next i
End Sub

Private Sub getMeshVertices()
    Dim tmpvertex As Long
    Dim i As Long
    
    numMeshVertices = lumpData(lump_meshvertices).length / Len(tmpvertex)
    
    Seek filenr, lumpData(lump_meshvertices).offset + 1
    
    ReDim meshvertices(numMeshVertices)
    For i = 0 To numMeshVertices
        Get filenr, , meshvertices(i)
    Next i
End Sub

Private Sub getShaders()
    Dim tmpshader As bsp_shader
    Dim i As Long
    
    numShaders = lumpData(lump_shaders).length / Len(tmpshader)
    
    Seek filenr, lumpData(lump_shaders).offset + 1
    
    ReDim shaders(numShaders)
    For i = 0 To numShaders
        Get filenr, , shaders(i)
        'objShaderlib.precacheShader shaders(i).path
    Next i
    
    objShaderlib.calcShaders
End Sub

Private Sub getFaces()
    Dim tmpface As bsp_face
    Dim i As Long
    
    numFaces = lumpData(lump_faces).length / Len(tmpface)
    
    Seek filenr, lumpData(lump_faces).offset + 1
    
    Log.Log "", "", ""
    
    ReDim faces(numFaces)
    ReDim faceDrawn(numFaces)
    ReDim patchIndex(numFaces)
    ReDim surfaceFlags(numFaces)
    ReDim facetransparant(numFaces)
    For i = 0 To numFaces
        Get filenr, , faces(i)
        
        If faces(i).type = 2 Then
            buildpatch i, 2
            patchIndex(i) = numPatches
        Else: patchIndex(i) = -1
        End If
    Next i
    Log.Log "", "", ""
End Sub

Private Sub adjustGamma(color() As Byte)
    Dim i As Long
    Dim r As Single
    Dim g As Single
    Dim b As Single
    Dim Temp As Single
    Dim thescale As Single
    
    For i = 0 To UBound(color) Step 3
        thescale = 1
        Temp = 0
        
        r = CSng(color(i))
        g = CSng(color(i + 1))
        b = CSng(color(i + 2))
        
        r = r * factor / 255
        g = g * factor / 255
        b = b * factor / 255
        
        If r = 0 Then r = 1 / 255
        If g = 0 Then g = 1 / 255
        If b = 0 Then b = 1 / 255
        
        Temp = 1 / r
        If r > 1 And Temp < thescale Then thescale = Temp
        Temp = 1 / g
        If g > 1 And Temp < thescale Then thescale = Temp
        Temp = 1 / b
        If b > 1 And Temp < thescale Then thescale = Temp
        
        thescale = thescale * 255
        
        r = r * thescale
        g = g * thescale
        b = b * thescale
        
        color(i) = r
        color(i + 1) = g
        color(i + 2) = b
    Next i
End Sub

Private Function makeLightmap(color() As Byte) As Long
    Dim tID As Long
    Dim i As Long
    
    glGenTextures 1, tID
    glBindTexture GL_TEXTURE_2D, tID
    
    glPixelStorei GL_PACK_ALIGNMENT, 1
    
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR_MIPMAP_NEAREST
    
    glTexEnvi GL_TEXTURE_2D, GL_TEXTURE_ENV_MODE, GL_MODULATE
    
    adjustGamma color
    
    gluBuild2DMipmaps GL_TEXTURE_2D, 3, 128, 128, GL_RGB, GL_UNSIGNED_BYTE, color(0)
    
    'glTexImage2D GL_TEXTURE_2D, 0, GL_RGB, 128, 128, 0, GL_RGB, GL_UNSIGNED_BYTE, color(0)
    makeLightmap = tID
End Function

Private Sub getLightmaps()
    Dim tmpLight() As bsp_lightmap
    Dim texdata As EGLTEXTUREDATA
    Dim i As Long
    
    glDeleteTextures numLightmaps + 1, lightmaps(0)
    
    ReDim tmpLight(0)
    
    numLightmaps = lumpData(lump_lightmaps).length / Len(tmpLight(0))
    
    Seek filenr, lumpData(lump_lightmaps).offset + 1
    
    ReDim lightmaps(numLightmaps)
    For i = 0 To numLightmaps
        Get filenr, , tmpLight(0)
        lightmaps(i) = makeLightmap(tmpLight(0).color)
    Next i
End Sub

Private Sub getLightvolumes()
    Dim tmpLight As bsp_lightdata
    Dim i As Long
    
    numLightvolumes = lumpData(lump_lightvolumes).length / Len(tmpLight)
    
    Seek filenr, lumpData(lump_lightvolumes).offset + 1
    
    ReDim lightvolumes(numLightvolumes)
    For i = 0 To numLightvolumes
        Get filenr, , lightvolumes(i)
    Next i
End Sub

Private Sub getVisdata()
    Dim i As Long
    
    Seek filenr, lumpData(lump_visdata).offset + 1
    
    Get filenr, , visdata.numClusters
    Get filenr, , visdata.bytesPerCluster
    
    ReDim visdata.bytes(visdata.numClusters * visdata.bytesPerCluster - 1)
    For i = 0 To visdata.numClusters * visdata.bytesPerCluster - 1
        Get filenr, , visdata.bytes(i)
    Next i
End Sub

Private Sub drawloadscreen()
    useTexture loadscreen 'glBindTexture GL_TEXTURE_2D, loadscreen
    
    glEnable GL_TEXTURE_2D
    
    glTranslatef 0, 0, -10
    
    glBegin GL_QUADS
        glTexCoord2f 0, 0
        glVertex3f 0, 0, 0
        glTexCoord2f 0, 1
        glVertex3f 0, 10, 0
        glTexCoord2f 1, 1
        glVertex3f 10, 10, 0
        glTexCoord2f 1, 0
        glVertex3f 10, 0, 0
    glEnd
    glDisable GL_TEXTURE_2D
    
    Graphics.nextFrame
    DoEvents
End Sub

Public Sub loadBSP(path As String, Optional id As String = "IBSP", Optional version As Long = 46)
    Dim texdata As EGLTEXTUREDATA
    Dim consoleShotPath As String
    Dim oldSHOW_TEXT As Long
    
    On Error GoTo errhandler
    
    Log.Log "", "", ""
    Log.Log "", "", "loading bsp " & path
    
    mapLoaded = False
    
    If Not InStr(1, path, ":") Then
        If LEFT(path, 1) = "\" Then
            path = App.path & path
        Else: path = App.path & "\" & path
        End If
    End If
    
    path = Replace(path, "/", "\")
    
    If Not pathexists(path) Then GoTo errhandler
    
    consoleShotPath = LEFT(path, Len(path) - 4)
    consoleShotPath = RIGHT(consoleShotPath, Len(consoleShotPath) - InStrRev(consoleShotPath, "\"))
    
    loadscreen = openQuakeTexture("levelshots\" & consoleShotPath)
    
    oldSHOW_TEXT = SHOW_TEXT
    SHOW_TEXT = 0
    
    filenr = FreeFile
    
    glMatrixMode GL_PROJECTION       ' Select The Projection Matrix
    glLoadIdentity                  ' Reset The Projection Matrix
    glOrtho 0, 10, 0, 10, 0, 10     ' set to ortho mode to display the loadscreen properly
    glMatrixMode GL_MODELVIEW        ' Select The Modelview Matrix
    glLoadIdentity                  ' Reset The Modelview Matrix
    
    Open path For Binary As filenr
        If Not getHeader(id, version) Then GoTo errhandler
        drawloadscreen
        
        getLumpData
        drawloadscreen
        
        getEntities
        drawloadscreen
        
        getTextures
        drawloadscreen
        
        getPlanes
        drawloadscreen
        
        getNodes
        drawloadscreen
        
        getLeafs
        drawloadscreen
        
        getLeafsFaces
        drawloadscreen
        
        getLeafsBrushes
        drawloadscreen
        
        getModels
        drawloadscreen
        
        getBrushes
        drawloadscreen
        
        getBrushessides
        drawloadscreen
        
        getVertices
        drawloadscreen
        
        getMeshVertices
        drawloadscreen
        
        getShaders
        drawloadscreen
        
        getFaces
        drawloadscreen
        
        getLightmaps
        drawloadscreen
        
        getLightvolumes
        drawloadscreen
        
        getVisdata
        drawloadscreen
    Close filenr
    
    Graphics.resize
    
    SHOW_TEXT = oldSHOW_TEXT
    
    mapLoaded = True
    
    Dim mapname As String
    
    mapname = RIGHT(path, Len(path) - Len(App.path))
    mapname = LEFT(mapname, Len(mapname) - 4)
    mapname = RIGHT(mapname, Len(mapname) - InStrRev(mapname, "\"))
    
    setCVar "map", mapname
    
    Log.Log "", "", ""
    Log.Log "", "", mapname & " succesfully loaded"
    Log.Log "", "", ""
    
    Exit Sub
    
errhandler:
    Graphics.resize
    SHOW_TEXT = oldSHOW_TEXT
    Log.Log "", "", "Unable to load map: " & path
    Log.Log "", "", ""
End Sub

Public Function findleaf(pos As vect) As Long
    Dim distance As Single
    Dim i As Long
    
    Do While i >= 0
        If i > numNodes Then
            findleaf = -1
            Exit Function
        End If
        
        distance = planes(nodes(i).plane).normals(0) * pos.x _
        + planes(nodes(i).plane).normals(1) * pos.y _
        + planes(nodes(i).plane).normals(2) * pos.z _
        - planes(nodes(i).plane).d
        
        If distance >= 0 Then
            i = nodes(i).FRONT
        Else
            i = nodes(i).BACK
        End If
    Loop
    
    findleaf = -(i + 1)
End Function

Public Function clusterVisible(test As Long, current As Long) As Boolean
    Dim offset As Long
    Dim bit As Long
    Dim theByte As Byte
    
    If current < 0 Or test < 0 Then
        clusterVisible = True
        Exit Function
    End If
    
    bit = test Mod visdata.bytesPerCluster
    offset = (test - (test Mod 8)) / 8 + current * visdata.bytesPerCluster
    
    If offset > UBound(visdata.bytes) Then
        clusterVisible = True
        Exit Function
    End If
    
    theByte = visdata.bytes(offset)
    clusterVisible = (theByte And (2 ^ bit) = 2 ^ bit)
End Function

Private Sub drawPatch(face As Long)
    Dim Patch As bsp_patch
    
    'Exit Sub
    
    'glColor3f 1, 0, 0
    'glDisable GL_CULL_FACE
    'glDisable GL_TEXTURE_2D
    
    objShaderlib.resetStates
    
    Patch = patches(patchIndex(face))
    
    Dim x As Long, y As Long, i As Long, j As Long
    
    If ObjEngine.texturesOn Then
        If Not faces(face).textureID > UBound(textureIDs) Then
            useTexture textureIDs(faces(face).textureID) ', faces(face) 'glBindTexture GL_TEXTURE_2D, faces(face).textureID
            
            For x = 0 To IIf(Patch.patchesX = 0, 0, Patch.patchesX)
                For y = 0 To IIf(Patch.patchesY = 0, 0, Patch.patchesY)
                    With Patch.patches(x, y)
                        For i = 0 To .w - 2
                            glBegin GL_TRIANGLE_STRIP
                                For j = 0 To .l - 1
                                    glTexCoord2f .textures(i, j).x, .textures(i, j).y
                                    glVertex3fv .verts(i, j).x
                                    glTexCoord2f .textures(i + 1, j).x, .textures(i + 1, j).y
                                    glVertex3fv .verts(i + 1, j).x
                                Next j
                            glEnd
                        Next i
                    End With
                Next y
            Next x
        End If
    End If
    
    If ObjEngine.lighting Then
        If Not faces(face).lightmapID > UBound(lightmaps) Then
            If ObjEngine.texturesOn Then
                objShaderlib.resetStates
                glBlendFunc GL_ZERO, GL_SRC_COLOR
                glEnable GL_BLEND
                glDepthMask GL_FALSE
                glDepthFunc GL_EQUAL
            End If
            
            If faces(face).lightmapID > -1 Then
                glEnable GL_TEXTURE_2D
                glBindTexture GL_TEXTURE_2D, lightmaps(faces(face).lightmapID)
            Else: glBindTexture GL_TEXTURE_2D, 0
            End If
            
            For x = 0 To IIf(Patch.patchesX = 0, 0, Patch.patchesX)
                For y = 0 To IIf(Patch.patchesY = 0, 0, Patch.patchesY)
                    With Patch.patches(x, y)
                        For i = 0 To .w - 2
                            glBegin GL_TRIANGLE_STRIP
                                For j = 0 To .l - 1
                                    glTexCoord2f .lightmaps(i, j).x, .lightmaps(i, j).y
                                    glVertex3fv .verts(i, j).x
                                    glTexCoord2f .lightmaps(i + 1, j).x, .lightmaps(i + 1, j).y
                                    glVertex3fv .verts(i + 1, j).x
                                Next j
                            glEnd
                        Next i
                    End With
                Next y
            Next x
            
            If ObjEngine.texturesOn Then
                glDepthMask GL_TRUE
                glDepthFunc GL_LEQUAL
                glDisable GL_BLEND
            End If
        End If
    End If
    
    'glEnable GL_TEXTURE_2D
    'glEnable GL_CULL_FACE
    'glColor3f 1, 1, 1
End Sub

Private Sub drawFace(face As Long)
    Dim j As Long
    
    If face > UBound(faces) Then Exit Sub
    
    resetStates
    
    If Not faceDrawn(face) Then
        If faces(face).type = 1 Then
            'Exit Sub
            If ObjEngine.texturesOn Then
                If Not faces(face).textureID > UBound(textureIDs) Then
                    'If Not useTextureWithShader(textureIDs(faces(face).textureID), faces(face)) Then  'glBindTexture GL_TEXTURE_2D, textureIDs(faces(face).textureID)
                        glBegin GL_TRIANGLE_FAN
                            For j = faces(face).vertexIndex To faces(face).vertexIndex + faces(face).numVertices - 1
                                glTexCoord2f vertices(j).texture(0), vertices(j).texture(1)
                                glVertex3fv vertices(j).position(0)
                            Next j
                        glEnd
                    'End If
                End If
            End If
            
            If ObjEngine.lighting Then
                If ObjEngine.texturesOn Then
                    objShaderlib.resetStates
                    glBlendFunc GL_ZERO, GL_SRC_COLOR
                    glEnable GL_BLEND
                    glDepthMask GL_FALSE
                    glDepthFunc GL_EQUAL
                End If
                
                If faces(face).lightmapID > -1 Then
                    glEnable GL_TEXTURE_2D
                    glBindTexture GL_TEXTURE_2D, lightmaps(faces(face).lightmapID)
                Else: glBindTexture GL_TEXTURE_2D, 0
                End If
                
                glBegin GL_TRIANGLE_FAN
                    For j = faces(face).vertexIndex To faces(face).vertexIndex + faces(face).numVertices - 1
                        glTexCoord2f vertices(j).lightmap(0), vertices(j).lightmap(1)
                        glVertex3fv vertices(j).position(0)
                    Next j
                glEnd
                
                If ObjEngine.texturesOn Then
                    glDepthMask GL_TRUE
                    glDepthFunc GL_LEQUAL
                    glDisable GL_BLEND
                End If
            End If
        ElseIf faces(face).type = 2 Then
            drawPatch face
        ElseIf faces(face).type = 3 Then
            'Exit Sub
            'glDisable GL_CULL_FACE
            If ObjEngine.texturesOn Then
                If Not faces(face).textureID > UBound(textureIDs) Then
                    If Not useTextureWithShader(textureIDs(faces(face).textureID), faces(face)) Then  'glBindTexture GL_TEXTURE_2D, textureIDs(faces(face).textureID)
                        glBegin GL_TRIANGLES
                            For j = faces(face).meshVertexIndex To faces(face).meshVertexIndex + faces(face).meshNumVertices - 1
                                glTexCoord2f vertices(meshvertices(j) + faces(face).vertexIndex).texture(0), vertices(meshvertices(j) + faces(face).vertexIndex).texture(1)
                                glVertex3fv vertices(meshvertices(j) + faces(face).vertexIndex).position(0)
                            Next j
                        glEnd
                    End If
                End If
            End If
            
            If ObjEngine.lighting Then
                If ObjEngine.texturesOn Then
                    objShaderlib.resetStates
                    glEnable GL_BLEND
                    glBlendFunc GL_ZERO, GL_SRC_COLOR
                    glDepthMask GL_FALSE
                    glDepthFunc GL_EQUAL
                End If
                
                If faces(face).lightmapID > -1 Then
                    glBindTexture GL_TEXTURE_2D, lightmaps(faces(face).lightmapID)
                Else: glBindTexture GL_TEXTURE_2D, 0
                End If
                glBegin GL_TRIANGLES
                    For j = faces(face).meshVertexIndex To faces(face).meshVertexIndex + faces(face).meshNumVertices - 1
                        glTexCoord2f vertices(meshvertices(j) + faces(face).vertexIndex).lightmap(0), vertices(meshvertices(j) + faces(face).vertexIndex).lightmap(1)
                        glVertex3fv vertices(meshvertices(j) + faces(face).vertexIndex).position(0)
                    Next j
                glEnd
                
                If ObjEngine.texturesOn Then
                    glDepthMask GL_TRUE
                    glDepthFunc GL_LEQUAL
                    glDisable GL_BLEND
                End If
            End If
            'glEnable GL_CULL_FACE
        End If
        
        faceDrawn(face) = True
    End If
End Sub

'draw part
Public Sub drawMap()
    Dim i As Long, j As Long
    Dim drawSkybox As Boolean
    
    If Not mapLoaded Then Exit Sub
    
    drawSkybox = True
    
    glPushMatrix
    
    glEnable GL_CULL_FACE
    
    If drawSkybox Then
        glPushMatrix
        glLoadIdentity
        
        glRotatef angleY, 1, 0, 0
        glRotatef angleX, 0, 1, 0
        
        glDisable GL_DEPTH_TEST
        
        For i = 0 To numFaces
            If surfaceFlags(i) And surfaceparm_sky Then
                objShaderlib.drawSkybox shaderIndices(textureIDs(faces(i).textureID))
                GoTo exit_skybox
            End If
        Next i
        
exit_skybox:
        
        glEnable GL_DEPTH_TEST
        
        glPopMatrix
    End If
    
    'get the current cluster and draw it
    Dim currentleaf As bsp_leaf
    Dim currentleafNR As Long
    Dim distance As Single
    Dim campos As vect
    
    campos = ObjEngine.getpos
    currentleafNR = findleaf(campos)
        
    If currentleafNR <> -1 Then
        currentleaf = leafs(currentleafNR)
        currentcluster = currentleaf.cluster
    Else: currentcluster = -1
    End If
    
    getfrustrum
    
    drawInfoPlayerStart
    
    WalkBSP 0, currentcluster, campos
    
    'For i = 0 To numLeafs
    '    If clusterVisible(leafs(i).cluster, currentcluster) Then
    '        If Frustrum.boxInFrustrum(CSng(leafs(i).mins(0)), CSng(leafs(i).mins(1)), CSng(leafs(i).mins(2)), CSng(leafs(i).maxs(0)), CSng(leafs(i).maxs(1)), CSng(leafs(i).maxs(2))) Then
    '            For j = leafs(i).leafFace To leafs(i).leafFace + leafs(i).numLeaffaces
    '                If Not facetransparant(leafsfaces(j)) Then
    '                    drawFace leafsfaces(j)
    '                End If
    '            Next j
    '        End If
    '    End If
    'Next i
    '
    'For i = 0 To numLeafs
    '    If clusterVisible(leafs(i).cluster, currentcluster) Then
    '        If Frustrum.boxInFrustrum(CSng(leafs(i).mins(0)), CSng(leafs(i).mins(1)), CSng(leafs(i).mins(2)), CSng(leafs(i).maxs(0)), CSng(leafs(i).maxs(1)), CSng(leafs(i).maxs(2))) Then
    '            For j = leafs(i).leafFace To leafs(i).leafFace + leafs(i).numLeaffaces
    '                If facetransparant(leafsfaces(j)) Then
    '                    drawFace leafsfaces(j)
    '                End If
    '            Next j
    '        End If
    '    End If
    'Next i
    
    glPopMatrix
    
    For i = 0 To numFaces
        faceDrawn(i) = False
    Next i
    
    resetStates
End Sub

'initializing and terminating
Public Sub BSP_init()
    ReDim textureIDs(0)
    ReDim lightmaps(0)
    
    numPatches = -1
    
    factor = 2.5
    
    levelToWidth(0) = 2
    levelToWidth(1) = 4
    levelToWidth(2) = 6
    levelToWidth(3) = 8
    levelToWidth(4) = 10
    levelToWidth(5) = 12
    levelToWidth(6) = 14
End Sub

Public Sub BSP_quit()
    clearAll
    glDeleteTextures numLightmaps + 1, lightmaps(0)
End Sub

Private Function dist_to_plane(planenr As Long, pos As vect) As Single
    dist_to_plane = planes(planenr).normals(0) * pos.x _
    + planes(planenr).normals(1) * pos.y _
    + planes(planenr).normals(2) * pos.z _
    - planes(planenr).d
End Function

Public Function detect_collision(p As vect) As Boolean
    Dim pointleaf As Long
    Dim i As Long, j As Long
    
    pointleaf = findleaf(p)
    
    With leafs(pointleaf)
        For i = .leafBrush To .leafBrush + .numLeafbrushes
            For j = brushes(leafsbrushes(i)).brushside To brushes(leafsbrushes(i)).brushside + brushes(leafsbrushes(i)).numBrushsides
                detect_collision = detect_collision And (dist_to_plane(brushside(j).plane, p) >= 0)
            Next j
        Next i
    End With
End Function

Public Function detect_point_visible(p As vect) As Boolean
    Dim pointleaf As Long
    Dim camleaf As Long
    Dim campos As vect
    Dim i As Long
    
    campos = ObjEngine.getpos
    
    pointleaf = findleaf(p)
    camleaf = findleaf(campos)
    
    If clusterVisible(leafs(camleaf).cluster, leafs(pointleaf).cluster) Then
        If boxInFrustrum(CSng(leafs(pointleaf).mins(0)), CSng(leafs(pointleaf).mins(1)), CSng(leafs(pointleaf).mins(2)), CSng(leafs(pointleaf).maxs(0)), CSng(leafs(pointleaf).maxs(1)), CSng(leafs(pointleaf).maxs(2))) Then
            detect_point_visible = True
        End If
    End If
End Function

Private Function planedist(p As vect, plane As Long) As Single
    With planes(plane)
        planedist = .normals(0) * p.x + .normals(1) * p.y + .normals(2) * p.z - .d
    End With
End Function

'used for collision detection
Public Function trace(move As movement, leaf As Long) As Boolean
    Dim i As Long, j As Long
    Dim dist1 As Single, dist2 As Single
    Dim result As Boolean
    Dim tmpresult As Boolean
    
    For i = leafs(leaf).leafBrush To leafs(leaf).leafBrush + leafs(leaf).numLeafbrushes
        For j = brushes(leafsbrushes(i)).brushside To brushes(leafsbrushes(i)).brushside + brushes(leafsbrushes(i)).numBrushsides
            With brushside(j)
                dist1 = planedist(move.start, .plane)
                dist2 = planedist(move.end, .plane)
                
                If (dist1 > 0 And dist2 < 0) Or (dist1 < 0 And dist2 > 0) Then
                    tmpresult = True
                Else: tmpresult = False
                End If
            End With
        Next j
        
        result = result Or tmpresult
    Next i
    
    trace = result
End Function

Private Function WalkBSP(node As Long, cluster As Long, pos As vect)
    Dim n As Long
    Dim d As Single
    Dim i As Long
    If node > -1 Then
        If Not boxInFrustrum(CSng(nodes(node).mins(0)), CSng(nodes(node).mins(1)), CSng(nodes(node).mins(2)), CSng(nodes(node).maxs(0)), CSng(nodes(node).maxs(1)), CSng(nodes(node).maxs(2))) Then Exit Function
        
        'Front to back rendering
        d = planedist(pos, nodes(node).plane)
        If d < 0 Then
            'we are behind the plane so goto to the front child first
            WalkBSP nodes(node).FRONT, cluster, pos
            WalkBSP nodes(node).BACK, cluster, pos
        Else
            'we are in front of the plane so goto to the back child first
            WalkBSP nodes(node).BACK, cluster, pos
            WalkBSP nodes(node).FRONT, cluster, pos
        End If
    Else
        n = -(node + 1)
        If clusterVisible(leafs(n).cluster, cluster) Then
            If boxInFrustrum(CSng(leafs(n).mins(0)), CSng(leafs(n).mins(1)), CSng(leafs(n).mins(2)), CSng(leafs(n).maxs(0)), CSng(leafs(n).maxs(1)), CSng(leafs(n).maxs(2))) Then
                For i = leafs(n).leafFace To leafs(n).leafFace + leafs(n).numLeaffaces
                    drawFace leafsfaces(i)
                Next i
            End If
        End If
    End If
End Function
