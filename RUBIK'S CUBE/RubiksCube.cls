VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RubiksCube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private surfaces(6) As Surface
'          ___
'         | 4 |
'  ___ ___|___|___
' | 0 | 3 | 5 | 2 |3
' |___|___|___|___|| side
'  top    | 1 |-3-
'         |___|
'         front

Private Sub Class_Initialize()
    For i = 0 To 5
        Set surfaces(i) = New Surface
        surfaces(i).setSide (i)
    Next
End Sub

Public Sub reset()

    Dim X As Integer, y As Integer
    For s = 0 To 5
        For X = 0 To 2
            For y = 0 To 2
                surfaces(Int(s)).changeTileColor X, y, Int(s)
            Next
        Next
    Next

End Sub

Public Function getSurface(faceIndex As Integer) As Surface
    Set getSurface = surfaces(faceIndex)
End Function

Public Sub initTileOnSurface(faceIndex As Integer, pic As Control, tileColor As Color, X As Integer, y As Integer)
    
    surfaces(faceIndex).initTile pic, tileColor, X, y
    surfaces(faceIndex).changeTileImg X, y, getColorPath(tileColor, faceIndex)
End Sub

Public Sub directShift(relevant_faces() As Surface, spinAxis As Integer, axisIndex As Integer, direction As Integer) ' coordinates of tiles stay the same, only surface changes
    
    Dim temp As Color
            Dim X As Integer ' not necessarily x-axis
            Dim a As Integer, b As Integer ' x and y are inverted if vertical shift
            
            For X = 0 To 2
                If spinAxis = 0 Then
                    a = X
                    b = axisIndex
                        
                ElseIf spinAxis = 1 Then
                    a = axisIndex
                    b = X
                
                End If
                
                If direction = 1 Then
                    temp = relevant_faces(3).getTile(a, b).getTColor()
                    
                    For s = 2 To 0 Step -1 ' max horizontal face index
                        'If (spinAxis = 1 And s = 1) Then
                            'Dim pos() As Integer
                            'pos = turn(a, b, 1) 'turn twice to get coordinates
                            'pos = turn(pos(0), pos(1), 1)
                            'relevant_faces(2).getTile(pos(0), pos(1)).setTColor ( _
                                'relevant_faces(1).getTile(a, b).getTColor() _
                            ')
                            
                        'Else
                            relevant_faces((s + 1) Mod 4).getTile(a, b).setTColor ( _
                                relevant_faces(s).getTile(a, b).getTColor() _
                            )
                       ' End If
                    Next
                    relevant_faces(0).getTile(a, b).setTColor temp
                Else
                    temp = relevant_faces(0).getTile(a, b).getTColor()
                    For s = 0 To 2 ' max horizontal face index
                        
                        relevant_faces(s).getTile(a, b).setTColor ( _
                            relevant_faces(s + 1).getTile(a, b).getTColor() _
                        )
                    Next
                    relevant_faces(3).getTile(a, b).setTColor temp
                End If
            Next
            
            For s = 0 To 3
                For X = 0 To 2
                    If spinAxis = 0 Then
                        a = X
                    Else
                        b = X
                    End If
                    
                    relevant_faces(s).updateTileImg a, b
                Next
            Next
    
End Sub

Public Sub spin(spinAxis As Integer, axisIndex As Integer, direction As Integer)
    
    Select Case spinAxis
        Case 0: ' x - horizontal shift
            directShift HORIZONTAL_FACES, 0, axisIndex, direction
        
        Case 1: ' y - vertical shift
            'temporarily rotate back surface 180 to 'fix' continuity
            surfaces(4).rotateStructure (1)
            surfaces(4).rotateStructure (1)
            directShift VERTICAL_FACES, 1, axisIndex, direction
            surfaces(4).rotate (1)
            surfaces(4).rotate (1)
        
        Case 2: ' z - lateral shift??
            'temporarily rotate each surface 90 * distance from face 2 to 'fix' continuity
            For i = 0 To 3
                For j = 0 To i - 1
                    LATERAL_FACES(i).rotateStructure (-1)
                Next
            Next
            directShift LATERAL_FACES, 1, axisIndex, direction
            For i = 0 To 3
                For j = 0 To i - 1
                    LATERAL_FACES(i).rotate (1)
                Next
            Next
        
    End Select
    
End Sub

Public Sub shift(spinAxis As Integer, axisIndex As Integer, direction As Integer)
    
    spin spinAxis, axisIndex, direction
    
    Select Case spinAxis
        Case 0:
            If axisIndex = 0 Then
                surfaces(0).rotate -direction
            ElseIf axisIndex = 2 Then
                surfaces(3).rotate direction
            End If
            
        Case 1:
            If axisIndex = 0 Then
                surfaces(5).rotate direction
            ElseIf axisIndex = 2 Then
                surfaces(2).rotate -direction
            End If
            
        Case 2:
            If axisIndex = 0 Then
                surfaces(1).rotate direction
            ElseIf axisIndex = 2 Then
                surfaces(4).rotate -direction
            End If
    End Select
    
    If isComplete Then
        MsgBox "RUBIK'S CUBE COMPLETE"
    End If
    
End Sub

Public Function isComplete() As Boolean
    
    Dim ret As Boolean
    ret = True
    
    For s = 0 To 5
        Dim compColor As Color
        compColor = surfaces(s).getTile(0, 0).getTColor()
        
        Dim X As Integer, y As Integer
        For X = 0 To 2
            For y = 0 To 2
                If compColor <> surfaces(s).getTile(X, y).getTColor Then
                    ret = False
                    Exit Function
                End If
            Next
        Next
    Next
    
    isComplete = ret
    
End Function
