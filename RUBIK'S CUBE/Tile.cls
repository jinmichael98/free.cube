VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private img As Image
Private imgPath As String
Private tColor As Color
'Private x As Integer
'Private y As Integer

' getters n setters
Public Function getImg() As Image
    Set getImg = img
End Function
Public Sub setImg(newImg As Image)
    Set img = newImg
End Sub

Public Function getImgPath() As String
    getImgPath = imgPath
End Function
Public Sub setImgPath(newImgPath As String)
    imgPath = newImgPath
End Sub

Public Function getTColor() As Color
    getTColor = tColor
End Function
Public Sub setTColor(newTColor As Color)
    tColor = newTColor
End Sub

Public Sub copy(t As Tile)

    tColor = t.getTColor()
    imgPath = t.getImgPath()

End Sub
