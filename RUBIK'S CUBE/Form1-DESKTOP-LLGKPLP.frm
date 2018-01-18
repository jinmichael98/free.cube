VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "PLAY CUBE EVERYDAY"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInvert 
      Caption         =   "Invert"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imgBot 
      Height          =   495
      Index           =   8
      Left            =   3120
      Picture         =   "Form1.frx":0000
      Top             =   5640
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgBot 
      Height          =   495
      Index           =   7
      Left            =   2160
      Picture         =   "Form1.frx":01B4
      Top             =   5640
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgBot 
      Height          =   495
      Index           =   6
      Left            =   1200
      Picture         =   "Form1.frx":0368
      Top             =   5640
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgBot 
      Height          =   495
      Index           =   5
      Left            =   3600
      Picture         =   "Form1.frx":051C
      Top             =   5160
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
      Left            =   1680
      Picture         =   "Form1.frx":0884
      Top             =   5160
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgBot 
      Height          =   495
      Index           =   2
      Left            =   4080
      Picture         =   "Form1.frx":0A38
      Top             =   4680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgBot 
      Height          =   495
      Index           =   1
      Left            =   3120
      Picture         =   "Form1.frx":0BEC
      Top             =   4680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgBot 
      Height          =   495
      Index           =   0
      Left            =   2160
      Picture         =   "Form1.frx":0DA0
      Top             =   4680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgLSide 
      Height          =   1500
      Index           =   8
      Left            =   1200
      Picture         =   "Form1.frx":0F54
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgLSide 
      Height          =   1500
      Index           =   7
      Left            =   1200
      Picture         =   "Form1.frx":1108
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgLSide 
      Height          =   1500
      Index           =   6
      Left            =   1200
      Picture         =   "Form1.frx":12BC
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgLSide 
      Height          =   1500
      Index           =   5
      Left            =   1680
      Picture         =   "Form1.frx":1470
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgLSide 
      Height          =   1500
      Index           =   4
      Left            =   1680
      Picture         =   "Form1.frx":1624
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgLSide 
      Height          =   1500
      Index           =   3
      Left            =   1680
      Picture         =   "Form1.frx":17D8
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgLSide 
      Height          =   1500
      Index           =   2
      Left            =   2160
      Picture         =   "Form1.frx":198C
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgLSide 
      Height          =   1500
      Index           =   1
      Left            =   2160
      Picture         =   "Form1.frx":1B40
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgLSide 
      Height          =   1500
      Index           =   0
      Left            =   2160
      Picture         =   "Form1.frx":1CF4
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgBack 
      Height          =   1005
      Index           =   5
      Left            =   4560
      Picture         =   "Form1.frx":1EA8
      Top             =   2760
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image imgBack 
      Height          =   1005
      Index           =   4
      Left            =   3600
      Picture         =   "Form1.frx":205C
      Top             =   2760
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image imgBack 
      Height          =   1005
      Index           =   3
      Left            =   2640
      Picture         =   "Form1.frx":2210
      Top             =   2760
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image imgBack 
      Height          =   1005
      Index           =   2
      Left            =   4560
      Picture         =   "Form1.frx":23C4
      Top             =   3720
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image imgBack 
      Height          =   1005
      Index           =   1
      Left            =   3600
      Picture         =   "Form1.frx":2578
      Top             =   3720
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image imgBack 
      Height          =   1005
      Index           =   0
      Left            =   2640
      Picture         =   "Form1.frx":272C
      Top             =   3720
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image imgBack 
      Height          =   1005
      Index           =   6
      Left            =   2640
      Picture         =   "Form1.frx":28E0
      Top             =   1800
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image imgBack 
      Height          =   1005
      Index           =   7
      Left            =   3600
      Picture         =   "Form1.frx":2A94
      Top             =   1800
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image imgBack 
      Height          =   1005
      Index           =   8
      Left            =   4560
      Picture         =   "Form1.frx":2C48
      Top             =   1800
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
      Left            =   4080
      Picture         =   "Form1.frx":3B9C
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Image imgTop 
      Height          =   495
      Index           =   7
      Left            =   3120
      Picture         =   "Form1.frx":3D50
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Image imgTop 
      Height          =   495
      Index           =   6
      Left            =   2160
      Picture         =   "Form1.frx":3F04
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Image imgTop 
      Height          =   495
      Index           =   5
      Left            =   1680
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
      Left            =   3600
      Picture         =   "Form1.frx":4420
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Image imgTop 
      Height          =   495
      Index           =   2
      Left            =   1200
      Picture         =   "Form1.frx":45D4
      Top             =   2760
      Width           =   1500
   End
   Begin VB.Image imgTop 
      Height          =   495
      Index           =   1
      Left            =   2160
      Picture         =   "Form1.frx":4788
      Top             =   2760
      Width           =   1500
   End
   Begin VB.Image imgFront 
      Height          =   1005
      Index           =   8
      Left            =   1200
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
      Left            =   3120
      Picture         =   "Form1.frx":4CA4
      Top             =   5160
      Width           =   1005
   End
   Begin VB.Image imgFront 
      Height          =   1005
      Index           =   5
      Left            =   1200
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
      Left            =   3120
      Picture         =   "Form1.frx":51C0
      Top             =   4200
      Width           =   1005
   End
   Begin VB.Image imgFront 
      Height          =   1005
      Index           =   2
      Left            =   1200
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
      Left            =   3120
      Picture         =   "Form1.frx":56DC
      Top             =   2760
      Width           =   1500
   End
   Begin VB.Image imgFront 
      Height          =   1005
      Index           =   0
      Left            =   3120
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RESOURCES_ROOT As String

Private Sub cmdInvert_Click()
    For i = 0 To 8
        imgFront(i).Visible = Not imgFront(i).Visible
        imgBack(i).Visible = Not imgBack(i).Visible
            
        imgSide(i).Visible = Not imgSide(i).Visible
        imgLSide(i).Visible = Not imgLSide(i).Visible
            
        imgTop(i).Visible = Not imgTop(i).Visible
        imgBot(i).Visible = Not imgBot(i).Visible
    Next
    
End Sub

Private Sub Form_Load()
    RESOURCES_ROOT = App.Path & "\resources"
    Dim cube As RubiksCube
    
    For i = Color.BLUE To Color.RED ' initialize, one color each side
        For x = 0 To 2
            For y = 0 To 2
                cube.surfaces(i).tiles(x, y) = i
                
                Select Case i
                End Select
            Next
        Next
    Next i

End Sub
