Attribute VB_Name = "modGlobals"


'Type RubikCube
'          ___
'         | 4 |
'  ___ ___|___|___
' | 0 | 2 | 3 | 5 |3
' |___|___|___|___|| side
'  top    | 1 |-3-
'         |___|
'         front
 '   surfaces(6) As Surface ' face in {surface, x, y}
'    colorsDb(6, 3) As Color ' color, surface
'End Type


Enum Color
    BLUE
    GREEN
    WHITE
    YELLOW
    ORANGE
    RED
End Enum

Global RESOURCES_ROOT As String
Global HORIZONTAL_FACES(4) As Surface
Global VERTICAL_FACES(4) As Surface
Global LATERAL_FACES(4) As Surface

Function getColorPath(tileColor As Color, side As Integer) As String
    Dim s As Integer
    If side < 3 Then
        s = side
    Else
        s = 6 Mod side
    End If
    getColorPath = RESOURCES_ROOT & "\" & (s) & "\" & Int(tileColor) & ".wmf"
End Function

Function turn(X As Integer, y As Integer, clockwise As Integer) As Integer() ' 'turns' a tile 90 degrees in its surface
    
    Dim newPos(2) As Integer
    
    If clockwise = 1 Then
        newPos(0) = 2 - y
        newPos(1) = X
    Else
        newPos(0) = y
        newPos(1) = 2 - X
    End If
    
    turn = newPos
    
End Function
