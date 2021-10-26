# VB6-DirectX-Parser
Application to parse DirectX files ".x" to data structures in VB6

All coded by Leandro Silva 25/10/2021
Importer for directx (may need improve for your needs) for VB6 ver.8176
Essencial para criar objectos 3D ou definir colisões por Mesh

```
*** Formatacao do ficheiro .x *** vvv
Frame Mundo { 1
  FrameTransformMatrix { 1.1 = 2 (loading...)
  ..,..,..,..,..,..,..,..,..,..,..,..,..,..,..,..;;
  }
  Frame Nome { 1.2 = 3
    FrameTransformMatrix { 1.2.1 = 4 (loading...)
    ..,..,..,..,..,..,..,..,..,..,..,..,..,..,..,..;;
    }
    Mesh Nome { 1.2.2 = 5 (loading...)
    nVertices;
    ..;..;..;,
    nFaces;
    4;..,..,..,..;,
    MeshNormals { 1.2.3 = 6 (loading...)
      nNormais;
      ..;..;..;,
      nNormaisnasFaces;
      4;..,..,..,..;,
    } //fecha MeshNormals
    MeshTextureCoords { 1.2.4 = 7
      nTextureCoord;
      ..;..;,
    } //fecha uv coords
    MeshMaterialList { 1.2.5 = 8
      nMaterial;
      ..;
      ..,
      ..;;
      Mateiral Nome { 1.2.6 = 9
      ..;..;.;..;;
      TextureFilename { 1.2.7 = 10
      "";
      } //fecha TextureFile
    } //fecha Materiallist
    } //fecha Mesh
  } //fecha Frame Nome
} //fecha Frame Mundo
```
