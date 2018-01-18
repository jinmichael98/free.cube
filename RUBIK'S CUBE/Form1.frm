VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "free.cube"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdToggleDir 
      Caption         =   "invert direction"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "help"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Timer tmrGame 
      Interval        =   100
      Left            =   360
      Top             =   2040
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "reset"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdScramble 
      Caption         =   "start"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   6360
      Width           =   975
   End
   Begin VB.Timer tmrShuffle 
      Interval        =   100
      Left            =   240
      Top             =   1320
   End
   Begin VB.CommandButton cmdInvert 
      Caption         =   "flip"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblMoveC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "moves done: 0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Image imgBot 
      Height          =   495
      Index           =   8
      Left            =   1200
      Picture         =   "Form1.frx":0000
      Top             =   5640
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgBot 
      Height          =   495
      Index           =   7
      Left            =   1680
      Picture         =   "Form1.frx":01B4
      Top             =   5160
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgBot 
      Height          =   495
      Index           =   6
      Left            =   2160
      Picture         =   "Form1.frx":0368
      Top             =   4680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgBot 
      Height          =   495
      Index           =   5
      Left            =   2160
      Picture         =   "Form1.frx":051C
      Top             =   5640
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgBot 
      Height          =   495
      Index           =   4
      Left            =   2640
      Picture         =   "Form1.frx":06D0
      Top             =   5160
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgBot 
      Height          =   495
      Index           =   3
      Left            =   3120
      Picture         =   "Form1.frx":0884
      Top             =   4680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgBot 
      Height          =   495
      Index           =   2
      Left            =   3120
      Picture         =   "Form1.frx":0A38
      Top             =   5640
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgBot 
      Height          =   495
      Index           =   1
      Left            =   3600
      Picture         =   "Form1.frx":0BEC
      Top             =   5160
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgBot 
      Height          =   495
      Index           =   0
      Left            =   4080
      Picture         =   "Form1.frx":0DA0
      Top             =   4680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgFrontB 
      Height          =   1500
      Index           =   8
      Left            =   2160
      Picture         =   "Form1.frx":0F54
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFrontB 
      Height          =   1500
      Index           =   7
      Left            =   1680
      Picture         =   "Form1.frx":1108
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFrontB 
      Height          =   1500
      Index           =   6
      Left            =   1200
      Picture         =   "Form1.frx":12BC
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFrontB 
      Height          =   1500
      Index           =   5
      Left            =   2160
      Picture         =   "Form1.frx":1470
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFrontB 
      Height          =   1500
      Index           =   4
      Left            =   1680
      Picture         =   "Form1.frx":1624
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFrontB 
      Height          =   1500
      Index           =   3
      Left            =   1200
      Picture         =   "Form1.frx":17D8
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFrontB 
      Height          =   1500
      Index           =   2
      Left            =   2160
      Picture         =   "Form1.frx":198C
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFrontB 
      Height          =   1500
      Index           =   1
      Left            =   1680
      Picture         =   "Form1.frx":1B40
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgFrontB 
      Height          =   1500
      Index           =   0
      Left            =   1200
      Picture         =   "Form1.frx":1CF4
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgLSide 
      Height          =   1005
      Index           =   5
      Left            =   4560
      Picture         =   "Form1.frx":1EA8
      Top             =   2760
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image imgLSide 
      Height          =   1005
      Index           =   4
      Left            =   3600
      Picture         =   "Form1.frx":205C
      Top             =   2760
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image imgLSide 
      Height          =   1005
      Index           =   3
      Left            =   2640
      Picture         =   "Form1.frx":2210
      Top             =   2760
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image imgLSide 
      Height          =   1005
      Index           =   2
      Left            =   4560
      Picture         =   "Form1.frx":23C4
      Top             =   1800
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image imgLSide 
      Height          =   1005
      Index           =   1
      Left            =   3600
      Picture         =   "Form1.frx":2578
      Top             =   1800
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image imgLSide 
      Height          =   1005
      Index           =   0
      Left            =   2640
      Picture         =   "Form1.frx":272C
      Top             =   1800
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image imgLSide 
      Height          =   1005
      Index           =   6
      Left            =   2640
      Picture         =   "Form1.frx":28E0
      Top             =   3720
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image imgLSide 
      Height          =   1005
      Index           =   7
      Left            =   3600
      Picture         =   "Form1.frx":2A94
      Top             =   3720
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image imgLSide 
      Height          =   1005
      Index           =   8
      Left            =   4560
      Picture         =   "Form1.frx":2C48
      Top             =   3720
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image imgSide 
      Height          =   1500
      Index           =   8
      Left            =   5040
      Picture         =   "Form1.frx":2DFC
      Top             =   3720
      Width           =   495
   End
   Begin VB.Image imgSide 
      Height          =   1500
      Index           =   7
      Left            =   4560
      Picture         =   "Form1.frx":2FB0
      Top             =   4200
      Width           =   495
   End
   Begin VB.Image imgSide 
      Height          =   1500
      Index           =   6
      Left            =   4080
      Picture         =   "Form1.frx":3164
      Top             =   4680
      Width           =   495
   End
   Begin VB.Image imgSide 
      Height          =   1500
      Index           =   5
      Left            =   5040
      Picture         =   "Form1.frx":3318
      Top             =   2760
      Width           =   495
   End
   Begin VB.Image imgSide 
      Height          =   1500
      Index           =   4
      Left            =   4560
      Picture         =   "Form1.frx":34CC
      Top             =   3240
      Width           =   495
   End
   Begin VB.Image imgSide 
      Height          =   1500
      Index           =   3
      Left            =   4080
      Picture         =   "Form1.frx":3680
      Top             =   3720
      Width           =   495
   End
   Begin VB.Image imgSide 
      Height          =   1500
      Index           =   2
      Left            =   5040
      Picture         =   "Form1.frx":3834
      Top             =   1800
      Width           =   495
   End
   Begin VB.Image imgSide 
      Height          =   1500
      Index           =   1
      Left            =   4560
      Picture         =   "Form1.frx":39E8
      Top             =   2280
      Width           =   495
   End
   Begin VB.Image imgTop 
      Height          =   495
      Index           =   8
      Left            =   3120
      Picture         =   "Form1.frx":3B9C
      Top             =   2760
      Width           =   1500
   End
   Begin VB.Image imgTop 
      Height          =   495
      Index           =   7
      Left            =   2160
      Picture         =   "Form1.frx":3D50
      Top             =   2760
      Width           =   1500
   End
   Begin VB.Image imgTop 
      Height          =   495
      Index           =   6
      Left            =   1200
      Picture         =   "Form1.frx":3F04
      Top             =   2760
      Width           =   1500
   End
   Begin VB.Image imgTop 
      Height          =   495
      Index           =   5
      Left            =   3600
      Picture         =   "Form1.frx":40B8
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Image imgTop 
      Height          =   495
      Index           =   4
      Left            =   2640
      Picture         =   "Form1.frx":426C
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Image imgTop 
      Height          =   495
      Index           =   3
      Left            =   1680
      Picture         =   "Form1.frx":4420
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Image imgTop 
      Height          =   495
      Index           =   2
      Left            =   4080
      Picture         =   "Form1.frx":45D4
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Image imgTop 
      Height          =   495
      Index           =   1
      Left            =   3120
      Picture         =   "Form1.frx":4788
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Image imgFront 
      Height          =   1005
      Index           =   8
      Left            =   3120
      Picture         =   "Form1.frx":493C
      Top             =   5160
      Width           =   1005
   End
   Begin VB.Image imgFront 
      Height          =   1005
      Index           =   7
      Left            =   2160
      Picture         =   "Form1.frx":4AF0
      Top             =   5160
      Width           =   1005
   End
   Begin VB.Image imgFront 
      Height          =   1005
      Index           =   6
      Left            =   1200
      Picture         =   "Form1.frx":4CA4
      Top             =   5160
      Width           =   1005
   End
   Begin VB.Image imgFront 
      Height          =   1005
      Index           =   5
      Left            =   3120
      Picture         =   "Form1.frx":4E58
      Top             =   4200
      Width           =   1005
   End
   Begin VB.Image imgFront 
      Height          =   1005
      Index           =   4
      Left            =   2160
      Picture         =   "Form1.frx":500C
      Top             =   4200
      Width           =   1005
   End
   Begin VB.Image imgFront 
      Height          =   1005
      Index           =   3
      Left            =   1200
      Picture         =   "Form1.frx":51C0
      Top             =   4200
      Width           =   1005
   End
   Begin VB.Image imgFront 
      Height          =   1005
      Index           =   2
      Left            =   3120
      Picture         =   "Form1.frx":5374
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Image imgFront 
      Height          =   1005
      Index           =   1
      Left            =   2160
      Picture         =   "Form1.frx":5528
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Image imgTop 
      Height          =   495
      Index           =   0
      Left            =   2160
      Picture         =   "Form1.frx":56DC
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Image imgFront 
      Height          =   1005
      Index           =   0
      Left            =   1200
      Picture         =   "Form1.frx":5890
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Image imgSide 
      Height          =   1500
      Index           =   0
      Left            =   4080
      Picture         =   "Form1.frx":5A44
      Top             =   2760
      Width           =   495
   End
   Begin VB.Image imgBg 
      Height          =   7200
      Left            =   -120
      Picture         =   "Form1.frx":5BF8
      Top             =   -120
      Width           =   9600
   End
   Begin VB.Image imgBgReflection 
      Height          =   7200
      Left            =   -2500
      Picture         =   "Form1.frx":6FE3
      Top             =   0
      Visible         =   0   'False
      Width           =   9600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cube As RubiksCube
Dim moveModifier As Integer
Public moveCount As Integer

Private Sub cmdHelp_Click()

    MsgBox "Click on tiles to move them with their entire row or column." & vbNewLine & _
    vbNewLine & _
    "Default directions:" & vbNewLine & _
    vbNewLine & _
    "    Front view:" & vbNewLine & _
    "        Front-facing tiles to the right" & vbNewLine & _
    "        Side tiles to the bottom" & vbNewLine & _
    "        Top tiles to the front" & vbNewLine & _
    vbNewLine & _
    "    Back view:" & vbNewLine & _
    "        Front-facing (opposite of the front view right face) tiles to the top" & vbNewLine & _
    "        Side (opposite of front face) tiles to the right" & vbNewLine & _
    "        Bottom tiles to the side (opposite of front face)" & vbNewLine & _
    vbNewLine & _
    "You move in oppsite directions by pressing the Invert Direction button"
    

End Sub

Private Sub cmdToggleDir_Click()

    moveModifier = -moveModifier
    If moveModifier <> 1 Then
        cmdToggleDir.Caption = "revert direction"
    Else
        cmdToggleDir.Caption = "invert direction"
    End If

End Sub

Private Sub cmdInvert_Click()
    For i = 0 To 8
        imgFront(i).Visible = Not imgFront(i).Visible
        imgLSide(i).Visible = Not imgLSide(i).Visible
            
        imgSide(i).Visible = Not imgSide(i).Visible
        imgFrontB(i).Visible = Not imgFrontB(i).Visible
            
        imgTop(i).Visible = Not imgTop(i).Visible
        imgBot(i).Visible = Not imgBot(i).Visible
    Next
    
    imgBg.Visible = Not imgBg.Visible
    imgBgReflection.Visible = Not imgBg.Visible
    
End Sub

Private Sub cmdReset_Click()
    
    tmrShuffle.Enabled = False
    cube.reset
    cmdScramble.Caption = "shuffle"
    moveCount = 0
    lblMoveC.Caption = "moves done: 0"
        
End Sub

Private Sub cmdScramble_Click()
    
    tmrShuffle.Enabled = Not tmrShuffle.Enabled
    
    If tmrShuffle.Enabled Then
        cmdScramble.Caption = "start"
    Else
        moveCount = 0
        lblMoveC.Caption = "moves done: 0"
        cmdScramble.Caption = "shuffle"
    End If
    
End Sub

Private Sub Form_Click()

    tmrShuffle.Enabled = Not tmrShuffle.Enabled

End Sub

Private Sub Form_Load()

    RESOURCES_ROOT = App.path & "\resources"
    
    moveCount = 0
    moveModifier = 1
    
    Randomize
    
    Set cube = New RubiksCube
    
    ' VB6 One line array definitions plz
    Set HORIZONTAL_FACES(0) = cube.getSurface(1)
    Set HORIZONTAL_FACES(1) = cube.getSurface(2)
    Set HORIZONTAL_FACES(2) = cube.getSurface(4)
    Set HORIZONTAL_FACES(3) = cube.getSurface(5)
    
    Set VERTICAL_FACES(0) = cube.getSurface(1)
    Set VERTICAL_FACES(1) = cube.getSurface(3)
    Set VERTICAL_FACES(2) = cube.getSurface(4)
    Set VERTICAL_FACES(3) = cube.getSurface(0)
    
    Set LATERAL_FACES(0) = cube.getSurface(2)
    Set LATERAL_FACES(1) = cube.getSurface(3)
    Set LATERAL_FACES(2) = cube.getSurface(5)
    Set LATERAL_FACES(3) = cube.getSurface(0)
        
    For i = 0 To 5 ' initialize, one color each side
        For X = 0 To 2
            For y = 0 To 2
                Dim t As Integer
                t = 3 * y + X
                Select Case i
                '          ___
                '         | 4 |
                '  ___ ___|___|___
                ' | 0 | 5 | 3 | 2 |3
                ' |___|___|___|___|| side
                '  top    | 1 |-3-
                '         |___|
                '         front
                    Case 0:
                        cube.initTileOnSurface Int(i), imgTop(t), Int(i), Int(X), Int(y)
                    Case 1:
                        cube.initTileOnSurface Int(i), imgFront(t), Int(i), Int(X), Int(y)
                    Case 2:
                        cube.initTileOnSurface Int(i), imgSide(t), Int(i), Int(X), Int(y)
                    Case 3:
                        cube.initTileOnSurface Int(i), imgBot(t), Int(i), Int(X), Int(y)
                    Case 4:
                        cube.initTileOnSurface Int(i), imgFrontB(t), Int(i), Int(X), Int(y)
                    Case 5:
                        cube.initTileOnSurface Int(i), imgLSide(t), Int(i), Int(X), Int(y)
                        
                End Select
            Next
        Next
    Next i
    
    cmdScramble.Left = Me.ScaleWidth / 2 - cmdScramble.Width / 2
    lblMoveC.Left = Me.ScaleWidth / 2 - lblMoveC.Width / 2

End Sub

Private Sub addMove()

    moveCount = moveCount + 1
    lblMoveC.Caption = "moves done: " & moveCount

End Sub

Private Sub imgBot_Click(Index As Integer)
    
    cube.shift 1, Index Mod 3, moveModifier
    addMove
    
End Sub

Private Sub imgFront_Click(Index As Integer)
    
    cube.shift 0, Int(Index / 3), moveModifier
    addMove
    
End Sub

Private Sub imgFrontB_Click(Index As Integer)
    
    cube.shift 0, Int(Index / 3), moveModifier
    addMove
    
End Sub

Private Sub imgLSide_Click(Index As Integer)
        
    cube.shift 2, Abs((Index Mod 3) - 2), moveModifier
    addMove
    
End Sub

Private Sub imgSide_Click(Index As Integer)
    
    cube.shift 2, Index Mod 3, moveModifier
    addMove
    
End Sub

Private Sub imgTop_Click(Index As Integer)
    
    cube.shift 1, Index Mod 3, moveModifier
    addMove
    
End Sub

Private Sub tmrShuffle_Timer()
    
    cube.shift Int(Rnd * 3), Int(Rnd * 3), 1
    
End Sub
