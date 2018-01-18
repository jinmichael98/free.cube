VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Surface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private tiles(3, 3) As Tile
Private side As Integer

Public Sub initTile(pic As Image, tileColor As Color, X As Integer, y As Integer)

    tiles(X, y).setImg pic
    tiles(X, y).setTColor tileColor
    'tiles(x, y).x = x
    'tiles(x, y).y = y
    
End Sub

Public Function getTile(X As Integer, y As Integer) As Tile

    Set getTile = tiles(X, y)

End Function

Public Sub setSide(s As Integer)

    side = s

End Sub

Public Sub changeTileImg(X As Integer, y As Integer, path As String)

    tiles(X, y).getImg().Picture = LoadPicture(path)
    tiles(X, y).setImgPath path
    
End Sub

Public Sub changeTileColor(X As Integer, y As Integer, c As Color)
    
    tiles(X, y).setTColor c
    tiles(X, y).setImgPath getColorPath(c, side)
    tiles(X, y).getImg().Picture = LoadPicture(tiles(X, y).getImgPath())
    
End Sub

Public Sub updateTileImg(X As Integer, y As Integer)
    
    tiles(X, y).setImgPath getColorPath(tiles(X, y).getTColor, side)
    tiles(X, y).getImg().Picture = LoadPicture(tiles(X, y).getImgPath())
    
End Sub
Public Sub rotateStructure(clockwise As Integer)

    Dim temps(3, 3) As Tile
    
    For y = 0 To 2
        For X = 0 To 2
            
            Dim newPos() As Integer
            newPos = turn(Int(X), Int(y), clockwise)
            
            Set temps(newPos(0), newPos(1)) = New Tile
            temps(newPos(0), newPos(1)).copy tiles(X, y)
            
        Next
    Next
    
    For y = 0 To 2
        For X = 0 To 2
            tiles(X, y).copy temps(X, y)
        Next
    Next

End Sub

Public Sub rotate(clockwise As Integer)
    
    rotateStructure clockwise
    
    For y = 0 To 2
        For X = 0 To 2
            changeTileColor Int(X), Int(y), tiles(X, y).getTColor()
        Next
    Next
    
End Sub

Private Sub Class_Initialize()

    For X = 0 To 2
        For y = 0 To 2
            Set tiles(X, y) = New Tile
        Next
    Next
            
End Sub
