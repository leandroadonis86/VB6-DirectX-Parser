VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import directx file"
   ClientHeight    =   1545
   ClientLeft      =   3045
   ClientTop       =   3195
   ClientWidth     =   3810
   Icon            =   "Form1.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   3810
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All coded by Leandro Silva 25/10/2021
' Importer for directx (may need improve for your needs) for VB6 ver.8176
' Essential for 3D objects and collusions with Mesh, VB6 ver.8176
' Essencial para criar objectos 3D ou definir colisões por Mesh
'
'*** Formatacao do ficheiro .x, File Format .x *** vvv
'Frame Mundo { 1
'  FrameTransformMatrix { 1.1 = 2 (loading...)
'  ..,..,..,..,..,..,..,..,..,..,..,..,..,..,..,..;;
'  }
'  Frame Nome { 1.2 = 3
'    FrameTransformMatrix { 1.2.1 = 4 (loading...)
'    ..,..,..,..,..,..,..,..,..,..,..,..,..,..,..,..;;
'    }
'    Mesh Nome { 1.2.2 = 5 (loading...)
'    nVertices;
'    ..;..;..;,
'    nFaces;
'    4;..,..,..,..;,
'    MeshNormals { 1.2.3 = 6 (loading...)
'      nNormais;
'      ..;..;..;,
'      nNormaisnasFaces;
'      4;..,..,..,..;,
'    } //fecha MeshNormals
'    MeshTextureCoords { 1.2.4 = 7
'      nTextureCoord;
'      ..;..;,
'    } //fecha uv coords
'    MeshMaterialList { 1.2.5 = 8
'      nMaterial;
'      ..;
'      ..,
'      ..;;
'      Mateiral Nome { 1.2.6 = 9
'      ..;..;.;..;;
'      TextureFilename { 1.2.7 = 10
'      "";
'      } //fecha TextureFile; end TextureFile
'    } //fecha Materiallist; end Materiallist
'    } //fecha Mesh; end Mesh
'  } //fecha Frame Nome; end Frame Nome
'} //fecha Frame Mundo; end Frame Mundo
' *** *** ^^^

Private Type Vertice
    x As Single
    y As Single
    z As Single
End Type

Private Type Triangle
    v1 As Vertice
    v2 As Vertice
    v3 As Vertice
End Type

'only if needed
Private Type Face
    t() As Triangle
End Type

Private Type Matrice4
    x(4) As Single
    y(4) As Single
    z(4) As Single
    d(4) As Single
End Type

Private Type Idtable
    nIds As Integer
    ids() As String
End Type

Private Type Mesh
    nVertices As Integer
    arrayVertices() As Vertice
    nFaces As Integer
    idtables() As Idtable
End Type

Private Type Normals
    nNormais As Integer
    arrayNormals() As Vertice
    nNormalinFaces As Integer
    idtables() As Idtable
End Type

Private Type Obje
    n As Integer
    arrayy() As Vertice
    nn As Integer
    idx() As Idtable
End Type

Private Type Model
    sceFrameTransformMatrix  As Matrice4
    objFrameTransformMatrix  As Matrice4
    objMesh As Mesh
    objNormals As Normals
End Type

Private Type Loading
    section As Integer
    line As Integer
End Type

Dim Itera1 As Integer 'cada componente; each component
Dim Itera2 As Integer 'linhas das componentes; line of component
Dim mymodel As Model  'container principal dos dados .x; main container of data .x
Dim myload As Loading 'loading pogress

Private Function xParser(sFileText As String)
'no inicio, at start
If Itera1 >= 0 Then
    'verificar arvore das componentes do objeto pela {; verify object componentes from the tree using {
    If InStr(1, LCase$(sFileText), LCase$(" {")) <> 0 Then
        Itera2 = 0
        myload.line = Itera2
        'encontrar primeiro onde começa o Frame Mundo, (1,3); find 1st start Frame Mundo at (1,3)
        If InStr(1, LCase$(sFileText), LCase$("Frame ")) <> 0 Then
                'MsgBox "Frame "
                If myload.section = 0 Then
                    myload.section = 1
                Else
                    myload.section = 3
                End If
'                Debug.Print "Frame " + Str(myload.section)
                fowardItera1
                Exit Function
        End If
        'verificar linha com FrameTransformMatrix, (2,4); verify line with FrameTransformMatrix at (2,4)
        If InStr(1, LCase$(sFileText), LCase$("FrameTransformMatrix ")) <> 0 Then
                'MsgBox "FrameTransformMatrix "
                If myload.section = 1 Then
                    myload.section = 2
                Else
                    myload.section = 4
                End If
'                Debug.Print "FrameTransformMatrix " + Str(myload.section)
                fowardItera1
                Exit Function
        End If
        'verificar linha com Mesh, 5; verify line with Mesh data, 5
        If InStr(1, LCase$(sFileText), LCase$("Mesh ")) <> 0 Then
                'MsgBox "Mesh "
                myload.section = 5
'                Debug.Print "Mesh " + Str(myload.section)
                fowardItera1
                Exit Function
        End If
        'verificar linha com MeshNormals, 6; verify line with MeshNormals, 6
        If InStr(1, LCase$(sFileText), LCase$("MeshNormals ")) <> 0 Then
                'MsgBox "MeshNormals "
                 myload.section = 6
'                Debug.Print "MeshNormals " + Str(myload.section)
                fowardItera1
                Exit Function
        End If
        'verificar linha com MeshTextureCoords; verify line with MeshTextureCoords
        If InStr(1, LCase$(sFileText), LCase$("MeshTextureCoords ")) <> 0 Then
                'MsgBox "MeshTextureCoords "
                fowardItera1
                Exit Function
        End If
        'verificar linha com MeshMaterialList; verify line with MeshMaterialList
        If InStr(1, LCase$(sFileText), LCase$("MeshMaterialList ")) <> 0 Then
                'MsgBox "MeshMaterialList "
                fowardItera1
                Exit Function
        End If
        'verificar linha com Material; verify line with Material
        If InStr(1, LCase$(sFileText), LCase$("Material ")) <> 0 Then
                'MsgBox "Material "
                fowardItera1
                Exit Function
        End If
        'verificar linha com TextureFilename; verify line with TextureFilename
        If InStr(1, LCase$(sFileText), LCase$("TextureFilename ")) <> 0 Then
                'MsgBox "TextureFilename "
                fowardItera1
                Exit Function
        End If
    End If
End If

If Itera1 > 0 And Itera1 <= 6 Then
    'estruturar dados para memoria, 0 nada aconteceu; struct data to memory, 0 nothing happened
    xParseStruct sFileText
End If

'fechar chavetas, itera2 fica negativo; close brace itera2 turn negative
If InStr(1, LCase$(sFileText), LCase$(" }")) <> 0 Then
    If Itera1 > 10 Then
        Itera1 = Itera1 - 11
    End If
    Itera2 = -1
    Exit Function
End If

'iterador das linhas avancar; line iterator to go foward
fowardItera2

End Function

Private Function xParseStruct(sFileText As String) As Integer
'funcao de formatacao crucial que pode ocorrer erro; crucial format function error may occur
xParseStruct = 1
    'se for uma matriz então estruturar dados; if its matrix struct data
    If isMatrice(sFileText) = True Then
        sFileText = formatingData(sFileText)
        'criar matriz ...
        Dim shwMx As Matrice4
        shwMx = xToMatrice(sFileText)
        If myload.section = 2 Then mymodel.sceFrameTransformMatrix = shwMx 'para sce
        If myload.section = 4 Then mymodel.objFrameTransformMatrix = shwMx 'para obj
'        Debug.Print "Matrice! " + Str$(shwMx.x(0)) + ", " + Str$(shwMx.y(0)) + ", " + Str$(shwMx.z(0)) + ", " + Str$(shwMx.d(0)) + ", " _
'                    + Str$(shwMx.x(1)) + ", " + Str$(shwMx.y(1)) + ", " + Str$(shwMx.z(1)) + ", " + Str$(shwMx.d(1)) + ", " _
'                    + Str$(shwMx.x(2)) + ", " + Str$(shwMx.y(2)) + ", " + Str$(shwMx.z(2)) + ", " + Str$(shwMx.d(2)) + ", " _
'                    + Str$(shwMx.x(3)) + ", " + Str$(shwMx.y(3)) + ", " + Str$(shwMx.z(3)) + ", " + Str$(shwMx.d(3))
    Exit Function
    End If
    
    'se for um total; if is a value
    If isTotal(sFileText) = True Then
        sFileText = formatingData(sFileText)
        'guardar total na matriz ...; safe to matrix
        Dim shwTotal As Integer
        shwTotal = xToTotal(sFileText)
        'Mesh section
        If myload.section = 5 Then
            If mymodel.objMesh.nVertices = 0 Then
                mymodel.objMesh.nVertices = shwTotal
                'e' necessario redefinir o tamanho da array para memoria; is necessary redefine array lenght for memory
                ReDim mymodel.objMesh.arrayVertices(shwTotal) As Vertice
            Else
                mymodel.objMesh.nFaces = shwTotal
                'e' necessario redefinir o tamanho da array para memoria
                ReDim mymodel.objMesh.idtables(shwTotal) As Idtable
            End If
        End If
        'Normal section
        If myload.section = 6 Then
            If mymodel.objNormals.nNormais = 0 Then
                mymodel.objNormals.nNormais = shwTotal
                ReDim mymodel.objNormals.arrayNormals(shwTotal) As Vertice
            Else
                mymodel.objNormals.nNormalinFaces = shwTotal
                ReDim mymodel.objNormals.idtables(shwTotal) As Idtable
            End If
        End If
'        Debug.Print "Totals! " + Str$(shwTotal) + ";"
    Exit Function
    End If
    
    'se for um vertice xyz; if its a vertice xyz
    If isVertice(sFileText) = True Then
        sFileText = formatingData(sFileText)
        'guardar vertice ...; save vertice ...
        Dim shwVce As Vertice
        shwVce = xToVertice(sFileText)
        'Mesh section
        If myload.section = 5 Then
            If Itera2 <= mymodel.objMesh.nVertices And mymodel.objMesh.nFaces = 0 Then
                mymodel.objMesh.arrayVertices(Itera2 - 1) = shwVce
            End If
        End If
        'Normal section
        If myload.section = 6 Then
            If Itera2 <= mymodel.objNormals.nNormais And mymodel.objNormals.nNormalinFaces = 0 Then
                mymodel.objNormals.arrayNormals(Itera2 - 1) = shwVce
            End If
        End If
'        Debug.Print "Vertice! " + Str$(shwVce.x) + ";" + Str$(shwVce.y) + ";" + Str$(shwVce.z) + ";"
    Exit Function
    End If
    
    'se for uma tabela de ids; if it's a table with ids
    If isTable(sFileText) = True Then
        sFileText = formatingData(sFileText)
        'porque é no seguimento da seccao e nfaces nunca é maior que os vertices
        If myload.section = 5 Then
            If Itera2 > mymodel.objMesh.nFaces Then Itera2 = 1
        End If
        If myload.section = 6 Then
            If Itera2 > mymodel.objNormals.nNormalinFaces Then Itera2 = 1
        End If
        'guardar tabela ...; save table
        Dim shwTbl As Idtable
        shwTbl = xToTable(sFileText)
        'Mesh section
        If myload.section = 5 Then
            If Itera2 <= mymodel.objMesh.nFaces And mymodel.objMesh.nFaces <> 0 Then
                mymodel.objMesh.idtables(Itera2 - 1) = shwTbl
            End If
        End If
        If myload.section = 6 Then
            If Itera2 <= mymodel.objNormals.nNormalinFaces And mymodel.objNormals.nNormalinFaces <> 0 Then
                mymodel.objNormals.idtables(Itera2 - 1) = shwTbl
            End If
        End If
'        Dim sidtables As String
'        For i = LBound(shwTbl.ids) To UBound(shwTbl.ids)
'            sidtables = sidtables + shwTbl.ids(i) + ","
'        Next
'        Debug.Print "Table! " + Str$(shwTbl.nIds) + ";" + sidtables + ";,"
    Exit Function
    End If

xParseStruct = 0
End Function

Private Sub Form_Activate()
 Me.Print "Click form to open .x file!"
End Sub

'importar e estruturar dados; import and struct data
Private Sub Form_Click()
    'Itera para a secção, outro iterado para as linhas; itera1 for section\component, intera2 for lines
    Itera1 = 0
    Itera2 = 0
    
    'loading calc...
    myload.section = 0
    myload.line = 0
    
    'abrir ficheiro .x para importar; open file .x to import
    With CommonDialog1
        .FileName = "" 'qualquer nome a selecionar
        .Filter = "All files (*.x) |*.x|" 'filtro para ficheiros directx; filter for .x files
        .ShowOpen 'mostrar; visible dialog
    End With
    
    Dim sFileText As String
    Dim iInputFile As Integer
    Dim i As Integer
    
    i = 0
    iInputFile = FreeFile
    If Len(CommonDialog1.FileName) > 0 Then
        Open CommonDialog1.FileName For Input As #iInputFile
        Do While Not EOF(iInputFile)
           'importar cada linha; import each line
           Line Input #iInputFile, sFileText
           'parse dos dados contidos na linha; parse data in line
           xParser sFileText
           'mostrar informação do ficheiro e iterar
'           Debug.Print Str(i) + " i1:" + Str(Itera1) + " i2:" + Str(Itera2) + " out: " + sFileText
           i = i + 1 'proxima linha; next line
        Loop
        Close #iInputFile
        toStringModel
    End If

End Sub

Private Function formatingData(sFileText As String) As String
    'remover tabulação; remove tabulation
    formatingData = Replace(sFileText, " ", "")
    'se tiver limitador final 2 ;; if theres a end limiter with two ;;
    If InStr(1, LCase$(sFileText), LCase$(";;")) <> 0 Then
        'remover 2 ;; remove two ;;
        formatingData = Replace(sFileText, ";;", "")
        'remover espaços para cada valores; remove spaces for each values
        formatingData = Replace(sFileText, " ", "")
    End If
End Function

Private Function countSpecificChar(txt As String, ch As String) As Integer
    'conta um caractere especifico; count specific character
    countSpecificChar = Len(txt) - Len(Replace(txt, ch, ""))
End Function

Private Function isMatrice(txt As String) As Boolean
    'é matriz com uma linha de 15 virgulas; is a matrix with 1 line and 15 comma
   If countSpecificChar(txt, ",") = 15 Then isMatrice = True
   If countSpecificChar(txt, ".") < 9 Then isMatrice = False
End Function

Private Function isTotal(n As String) As Boolean
    'é matriz com uma linha de 1 unico ponto e virgula; is a matrix with 1 line and 1 semicolon
   If countSpecificChar(n, ";") = 1 Then isTotal = True
   If countSpecificChar(n, ".") > 0 Then isTotal = False
End Function

Private Function isVertice(n As String) As Boolean
   If countSpecificChar(n, ";") >= 3 Then
        If countSpecificChar(n, ".") = 3 Then
             isVertice = True
        Else
             isVertice = False
        End If
   End If
End Function

Private Function isTable(n As String) As Boolean
    If countSpecificChar(n, ".") > 0 Then
        isTable = False
        Exit Function
    ElseIf countSpecificChar(n, ";") >= 2 Then
        If countSpecificChar(n, ",") >= 3 Then
            isTable = True
        End If
    End If
End Function
Private Function xToTotal(x As String) As Integer
    Dim v As Integer
    v = Val(Split(x, ";")(0))
    
    xToTotal = v
End Function
Private Function xToTable(x As String) As Idtable
    Dim t As Idtable
    Dim ttable() As String
    Dim tidstable() As String
    ttable = Split(x, ";")
    tidstable = Split(ttable(1), ",")
        
    With t
        .nIds = Val(ttable(0))
        .ids = tidstable
    End With
    
    xToTable = t
End Function
Private Function xToVertice(x As String) As Vertice
    Dim v As Vertice
    Dim vvert() As String
    vvert = Split(x, ";")
    
    'construir vertice; build vertice
    With v
        .x = Val(vvert(0))
        .y = Val(vvert(1))
        .z = Val(vvert(2))
    End With
    
    xToVertice = v
End Function
Private Function xToMatrice(x As String) As Matrice4
    Dim m As Matrice4
    'devolver cada valor entre virgulas para uma array; value in between comma to array
    Dim vmat() As String
    vmat = Split(x, ",")

    'array de string para matriz de valor; string array for matrix values
    For i = 0 To 3
        m.x(i) = Val(vmat(0 + i * 4))
        m.y(i) = Val(vmat(1 + i * 4))
        m.z(i) = Val(vmat(2 + i * 4))
        m.d(i) = Val(vmat(3 + i * 4))
    Next
    
    xToMatrice = m
End Function

Private Function toStringTable(tb() As Idtable) As String
''*** *** vvv
''Private Type Idtable
''    nIds As Integer
''    ids() As String
''End Type
''*** *** ^^^
Dim txt As String
For i = LBound(tb) To UBound(tb) - 1
    txt = txt + Str(tb(i).nIds) + ":"
    For ii = 0 To tb(i).nIds - 1
        txt = txt + tb(i).ids(ii) + ","
    Next
    'If i <= tb(i).nIds Then
    txt = txt + vbCrLf + vbTab + vbTab
Next
    toStringTable = txt

End Function

Private Function toStringVertice(arrv() As Vertice) As String
''*** *** vvv
''Private Type Vertice
''    x As Single
''    y As Single
''    z As Single
''End Type
''*** *** ^^^
Dim txt As String
For i = LBound(arrv) To UBound(arrv) - 1
    txt = txt + Str(arrv(i).x) + "," + Str(arrv(i).y) + "," + Str(arrv(i).z) + ";, " + vbCrLf + vbTab + vbTab
Next
    toStringVertice = txt

End Function

Private Function toStringObj(obj As Obje) As String
''*** *** vvv
''Private Type Mesh
''    nVertices As Integer
''    arrayVertices() As Vertice
''    nFaces As Integer
''    idtables() As Idtable
''End Type
''*** *** ^^^
''*** *** vvv
''Private Type Normals
''    nNormais As Integer
''    arrayNormals() As Vertice
''    nNormalinFaces As Integer
''    idtables() As Idtable
''End Type
''*** *** ^^^

toStringObj = "" + Str(obj.n) _
    + "; " + vbCrLf + vbTab + vbTab _
    + toStringVertice(obj.arrayy) _
    + "" + Str(obj.nn) _
    + "; " + vbCrLf + vbTab + vbTab _
    + toStringTable(obj.idx)
End Function

Private Function toStringMatrix4(m As Matrice4) As String
''*** *** vvv
''Private Type Matrice4
''    x(4) As Single
''    y(4) As Single
''    z(4) As Single
''    d(4) As Single
''End Type
''*** *** ^^^

';)
toStringMatrix4 = "["
For i = 0 To 3
    With m
        toStringMatrix4 = toStringMatrix4 + Str(.x(i)) + ", "
        toStringMatrix4 = toStringMatrix4 + Str(.y(i)) + ", "
        toStringMatrix4 = toStringMatrix4 + Str(.x(i)) + ", "
        toStringMatrix4 = toStringMatrix4 + Str(.d(i))
    End With
    If i < 3 Then toStringMatrix4 = toStringMatrix4 + ", "
Next
toStringMatrix4 = toStringMatrix4 + "]"
End Function


Public Function toStringModel() As String
''*** *** vvv
''Private Type Model
''    sceFrameTransformMatrix  As Matrice4
''    objFrameTransformMatrix  As Matrice4
''    objMesh As Mesh
''    objNormals As Normals
''End Type
''*** *** ^^^
'Debug.Print toStringMatrix4(mymodel.sceFrameTransformMatrix)

toStringModel = "Model { " + vbCrLf + vbTab _
                    + "sceFrameTransformMatrix: " _
                    + toStringMatrix4(mymodel.sceFrameTransformMatrix) _
                    + ", " + vbCrLf + vbTab _
                    + "objFrameTransformMatrix: " _
                    + toStringMatrix4(mymodel.objFrameTransformMatrix) _
                    + ", " + vbCrLf + vbTab _
                    + "objMesh {" + vbCrLf + vbTab + vbTab _
                    + toStringObj(meshtoObject(mymodel.objMesh)) _
                    + "} " + vbCrLf + vbTab _
                    + "objNormals {" + vbCrLf + vbTab + vbTab _
                    + toStringObj(normalstoObject(mymodel.objNormals)) _
                    + "} " + vbCrLf _
                 + "}"
Debug.Print toStringModel
End Function

Private Function meshtoObject(c As Mesh) As Obje
'converter tipo Mesh para objecto; convert type Mesh to object
    Dim o As Obje
    o.n = c.nVertices
    o.arrayy = c.arrayVertices
    o.nn = c.nFaces
    o.idx = c.idtables
    
    meshtoObject = o
End Function

Private Function normalstoObject(c As Normals) As Obje
'coverter tipo Normals para objecto; convert type Normals to object
    Dim o As Obje
    o.n = c.nNormais
    o.arrayy = c.arrayNormals
    o.nn = c.nNormalinFaces
    o.idx = c.idtables
    
    normalstoObject = o
End Function

Private Function createFace2Triangles(t() As Triangle) As Face
    'uma face é 1 ou mais tringulos, normalmente 2 triangulos; face is 1 ou more tris, standart 2 tris
    Dim f As Face
    f.t = t
    
    createFace2Triangles = f
End Function

Private Function createTriangles3Vertices(v() As Vertice) As Triangle
    'um triangulo são 3 vertices; one tris is 3 vertix
    Dim tp As Triangle
    tp.v1 = v(0)
    tp.v2 = v(1)
    tp.v3 = v(2)
    
    createTriangles3Vertices = tp
End Function

Private Function createVerticeXYZ(p() As Vertice) As Vertice
    'um vertice e' um ponto com x,y,z; one vertice is one point xyz
    Dim tp As Vertice
    tp.x = p(0)
    tp.y = p(1)
    tp.z = p(2)
    
    createVerticeXYZ = tp
End Function

Private Function fowardItera1(Optional i As Integer = 1)
    'sempre que houver ";;" avança ; on each ";;" go foward
     Itera1 = Itera1 + i
End Function

Private Function fowardItera2(Optional i As Integer = 1)
    'se houver mais linhas na secção ; if more lines in section
     Itera2 = Itera2 + i
End Function

Private Sub Form_Terminate()
'terminou
Debug.Print "** Ended **"
End Sub
