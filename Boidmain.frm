VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "vbBoid : Example of emergent behaviour  © RL 1999"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11190
   ForeColor       =   &H8000000F&
   Icon            =   "Boidmain.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   8055
      Left            =   0
      TabIndex        =   69
      Top             =   0
      Width           =   1875
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Index           =   3
         Left            =   120
         Max             =   50
         TabIndex        =   96
         Top             =   360
         Value           =   30
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Clear"
         Height          =   375
         Left            =   960
         TabIndex        =   95
         Top             =   7080
         Width           =   855
      End
      Begin VB.CheckBox chkEnclosed 
         Caption         =   "Enclosed"
         Height          =   375
         Left            =   180
         TabIndex        =   94
         Top             =   6660
         Width           =   1335
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   60
         TabIndex        =   87
         Top             =   7500
         Width           =   855
      End
      Begin VB.CheckBox chkShowCentre 
         Caption         =   "Show"
         Height          =   255
         Left            =   840
         TabIndex        =   86
         Top             =   720
         Width           =   735
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Index           =   0
         Left            =   120
         Max             =   10
         TabIndex        =   85
         Top             =   960
         Value           =   3
         Width           =   1575
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Index           =   1
         Left            =   120
         Max             =   10
         TabIndex        =   84
         Top             =   1560
         Value           =   3
         Width           =   1575
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Index           =   2
         Left            =   120
         Max             =   10
         TabIndex        =   83
         Top             =   2160
         Value           =   6
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1200
         TabIndex        =   82
         Text            =   ".5"
         Top             =   3060
         Width           =   495
      End
      Begin VB.CommandButton cmdExecute 
         Caption         =   "Execute"
         Height          =   375
         Left            =   60
         TabIndex        =   81
         Top             =   7080
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1200
         TabIndex        =   80
         Text            =   "80"
         Top             =   2580
         Width           =   495
      End
      Begin VB.CheckBox chkGrid 
         Caption         =   "Show Grid"
         Height          =   255
         Left            =   180
         TabIndex        =   79
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CheckBox chkTrails 
         Caption         =   "Show Trails"
         Height          =   255
         Left            =   180
         TabIndex        =   78
         Top             =   4800
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1200
         TabIndex        =   77
         Text            =   "4"
         Top             =   3540
         Width           =   495
      End
      Begin VB.CheckBox chkShowSep 
         Caption         =   "Show "
         Height          =   255
         Left            =   840
         TabIndex        =   76
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox chkShowAlign 
         Caption         =   "Show "
         Height          =   255
         Left            =   840
         TabIndex        =   75
         Top             =   1920
         Width           =   855
      End
      Begin VB.CheckBox chkShowSensor 
         Caption         =   "Show Vision"
         Height          =   255
         Left            =   180
         TabIndex        =   74
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CheckBox chkShowBox 
         Caption         =   "Show Obstacle Box"
         Height          =   375
         Left            =   180
         TabIndex        =   73
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CheckBox chkShowColours 
         Caption         =   "Show Colours"
         Height          =   255
         Left            =   180
         TabIndex        =   72
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CheckBox chkShowArrow 
         Caption         =   "Show Arrow"
         Height          =   375
         Left            =   180
         TabIndex        =   71
         Top             =   5940
         Width           =   1335
      End
      Begin VB.CheckBox chkShowCircle 
         Caption         =   "Show Circle"
         Height          =   375
         Left            =   180
         TabIndex        =   70
         Top             =   6300
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Boid Count"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   97
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Centre"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   93
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Separate"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   92
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Align"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   91
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Sensor Range (Pix)"
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   90
         Top             =   2580
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Max Turn (Rad)"
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   89
         Top             =   3060
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Field of View (Rad)"
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   88
         Top             =   3540
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   7215
      Left            =   11760
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton cmdSet 
         Caption         =   "Set Boids"
         Height          =   375
         Left            =   1080
         TabIndex        =   68
         Top             =   6120
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   2160
         TabIndex        =   67
         Top             =   6120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCalcForces 
         Caption         =   "Calc Forces"
         Height          =   375
         Left            =   0
         TabIndex        =   66
         Top             =   6480
         Width           =   1095
      End
      Begin VB.CommandButton cmdDraw 
         Caption         =   "Draw"
         Height          =   375
         Left            =   1080
         TabIndex        =   65
         Top             =   6480
         Width           =   1095
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "Move"
         Height          =   375
         Left            =   2160
         TabIndex        =   64
         Top             =   6480
         Width           =   1095
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run"
         Height          =   375
         Left            =   600
         TabIndex        =   63
         Top             =   6840
         Width           =   1095
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset Boids"
         Height          =   375
         Left            =   0
         TabIndex        =   62
         Top             =   6120
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Height          =   6135
         Index           =   3
         Left            =   2520
         TabIndex        =   49
         Top             =   0
         Width           =   675
         Begin VB.TextBox Text4 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   61
            Text            =   "10"
            Top             =   1680
            Width           =   500
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   60
            Text            =   "2"
            Top             =   1200
            Width           =   500
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   59
            Text            =   "132"
            Top             =   720
            Width           =   500
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   58
            Text            =   "143"
            Top             =   240
            Width           =   500
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   57
            Text            =   "10"
            Top             =   2160
            Width           =   500
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   56
            Text            =   "10"
            Top             =   2640
            Width           =   500
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   55
            Text            =   "10"
            Top             =   3120
            Width           =   500
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   54
            Text            =   "10"
            Top             =   3600
            Width           =   500
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   53
            Text            =   "10"
            Top             =   4080
            Width           =   500
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   52
            Text            =   "10"
            Top             =   4560
            Width           =   500
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   51
            Text            =   "10"
            Top             =   5040
            Width           =   500
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   50
            Text            =   "10"
            Top             =   5520
            Width           =   500
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6135
         Index           =   2
         Left            =   1920
         TabIndex        =   36
         Top             =   120
         Width           =   675
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   48
            Text            =   "10"
            Top             =   1680
            Width           =   500
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   47
            Text            =   "3"
            Top             =   1200
            Width           =   500
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   46
            Text            =   "110"
            Top             =   720
            Width           =   500
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   45
            Text            =   "118"
            Top             =   240
            Width           =   500
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   44
            Text            =   "10"
            Top             =   2160
            Width           =   500
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   43
            Text            =   "10"
            Top             =   2640
            Width           =   500
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   42
            Text            =   "10"
            Top             =   3120
            Width           =   500
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   41
            Text            =   "10"
            Top             =   3600
            Width           =   500
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   40
            Text            =   "10"
            Top             =   4080
            Width           =   500
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   39
            Text            =   "10"
            Top             =   4560
            Width           =   500
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   38
            Text            =   "10"
            Top             =   5040
            Width           =   500
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   37
            Text            =   "10"
            Top             =   5520
            Width           =   500
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6135
         Index           =   1
         Left            =   1320
         TabIndex        =   23
         Top             =   120
         Width           =   675
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   35
            Text            =   "10"
            Top             =   1680
            Width           =   500
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   34
            Text            =   "1"
            Top             =   1200
            Width           =   500
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   33
            Text            =   "102"
            Top             =   720
            Width           =   500
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Text            =   "153"
            Top             =   240
            Width           =   500
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   31
            Text            =   "10"
            Top             =   2160
            Width           =   500
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   30
            Text            =   "10"
            Top             =   2640
            Width           =   500
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   29
            Text            =   "10"
            Top             =   3120
            Width           =   500
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   28
            Text            =   "10"
            Top             =   3600
            Width           =   500
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   27
            Text            =   "10"
            Top             =   4080
            Width           =   500
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   26
            Text            =   "10"
            Top             =   4560
            Width           =   500
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   25
            Text            =   "10"
            Top             =   5040
            Width           =   500
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   24
            Text            =   "10"
            Top             =   5520
            Width           =   500
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6135
         Index           =   0
         Left            =   720
         TabIndex        =   10
         Top             =   0
         Width           =   675
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Text            =   "120"
            Top             =   240
            Width           =   500
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Text            =   "160"
            Top             =   720
            Width           =   500
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   20
            Text            =   "0"
            Top             =   1200
            Width           =   500
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   19
            Text            =   "10"
            Top             =   1680
            Width           =   500
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   18
            Text            =   "10"
            Top             =   2160
            Width           =   500
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   17
            Text            =   "10"
            Top             =   2640
            Width           =   500
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   16
            Text            =   "10"
            Top             =   3120
            Width           =   500
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   15
            Text            =   "10"
            Top             =   3600
            Width           =   500
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   14
            Text            =   "10"
            Top             =   4080
            Width           =   500
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   13
            Text            =   "10"
            Top             =   4560
            Width           =   500
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   12
            Text            =   "10"
            Top             =   5040
            Width           =   500
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   11
            Text            =   "10"
            Top             =   5520
            Width           =   500
         End
      End
      Begin VB.Frame Frame2 
         Height          =   6135
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   735
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "X"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Y"
            Height          =   195
            Index           =   1
            Left            =   100
            TabIndex        =   8
            Top             =   840
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "DIR"
            Height          =   195
            Index           =   2
            Left            =   100
            TabIndex        =   7
            Top             =   1320
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Len"
            Height          =   195
            Index           =   3
            Left            =   100
            TabIndex        =   6
            Top             =   1800
            Width           =   600
         End
         Begin VB.Label label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Align"
            ForeColor       =   &H00FF00FF&
            Height          =   195
            Index           =   1
            Left            =   100
            TabIndex        =   5
            Top             =   2400
            Width           =   600
         End
         Begin VB.Label label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Centre"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   0
            Left            =   100
            TabIndex        =   4
            Top             =   3360
            Width           =   600
         End
         Begin VB.Label label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Sep"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   2
            Left            =   100
            TabIndex        =   3
            Top             =   4440
            Width           =   600
         End
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   7875
      Left            =   1860
      ScaleHeight     =   521
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   613
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 0) As SAFEARRAYBOUND
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type point
    r As Byte
    c As Byte
End Type

Dim dwn%
Dim bkg(255, 255) As Byte
Dim blnRun As Boolean
Dim iXval As Integer
Dim iYval As Integer

Private Sub cmdCalcForces_Click()
Dim boid As BoidClass

    CalcForces flock, HScroll1(0), HScroll1(1), HScroll1(2), Val(Text5), Val(Text7)

    For Each boid In flock

        Select Case boid.id
        Case 0
            Text1(4) = boid.DesireAlignTurn
            Text1(5) = boid.DesireAlignWeight
            Text1(6) = boid.DesireCentreTurn
            Text1(7) = boid.DesireCentreWeight
            Text1(8) = boid.DesireSeparateTurn
            Text1(9) = boid.DesireSeparateWeight
            Text1(10) = boid.AveX
            Text1(11) = boid.AveY
        Case 1
            Text2(4) = boid.DesireAlignTurn
            Text2(5) = boid.DesireAlignWeight
            Text2(6) = boid.DesireCentreTurn
            Text2(7) = boid.DesireCentreWeight
            Text2(8) = boid.DesireSeparateTurn
            Text2(9) = boid.DesireSeparateWeight
            Text2(10) = boid.AveX
            Text2(11) = boid.AveY
        Case 2
            Text3(4) = boid.DesireAlignTurn
            Text3(5) = boid.DesireAlignWeight
            Text3(6) = boid.DesireCentreTurn
            Text3(7) = boid.DesireCentreWeight
            Text3(8) = boid.DesireSeparateTurn
            Text3(9) = boid.DesireSeparateWeight
            Text3(10) = boid.AveX
            Text3(11) = boid.AveY
        Case 3
            Text4(4) = boid.DesireAlignTurn
            Text4(5) = boid.DesireAlignWeight
            Text4(6) = boid.DesireCentreTurn
            Text4(7) = boid.DesireCentreWeight
            Text4(8) = boid.DesireSeparateTurn
            Text4(9) = boid.DesireSeparateWeight
            Text4(10) = boid.AveX
            Text4(11) = boid.AveY
        End Select
        
    Next
    Set boid = Nothing

End Sub

Private Sub cmdClear_Click()
    Picture1.Cls
End Sub

Private Sub cmdDraw_Click()
Dim i%, j%

    For i% = 0 To Picture1.ScaleWidth Step 10
        Picture1.Line (i%, 0)-(i%, Picture1.ScaleHeight), vbButtonFace
    Next
    For i% = 0 To Picture1.ScaleHeight Step 10
        Picture1.Line (0, i%)-(Picture1.ScaleWidth, i%), vbButtonFace
    Next
    
    DrawBoid flock, Me.Picture1, Me.chkShowColours, Me.chkShowArrow, Me.chkShowCircle
    
    DrawObjects objects, Me.Picture1

    DrawForces flock, Me.Picture1, Val(Text5), Val(Text7), Me.chkShowCentre, Me.chkShowSep, Me.chkShowAlign, Me.chkShowSensor, Me.chkShowBox
    
End Sub

Private Sub cmdExecute_Click()
Dim i%, j%
Dim boid As BoidClass

Select Case cmdExecute.Caption
Case "Execute"
    cmdExecute.Caption = "Stop"
    blnRun = True
    
    Do
    
    
        CalcForces flock, HScroll1(0), HScroll1(1), HScroll1(2), Val(Text5), Val(Text7)
    
        If Me.chkTrails = False Then
            Picture1.Cls
        End If
        
        If Me.chkGrid <> False Then
            For i% = 0 To Picture1.ScaleWidth Step 10
                Picture1.Line (i%, 0)-(i%, Picture1.ScaleHeight), vbButtonFace
            Next
            For i% = 0 To Picture1.ScaleHeight Step 10
                Picture1.Line (0, i%)-(Picture1.ScaleWidth, i%), vbButtonFace
            Next
        End If
        
        DrawBoid flock, Me.Picture1, Me.chkShowColours, Me.chkShowArrow, Me.chkShowCircle
        DrawObjects objects, Me.Picture1
        
        DrawForces flock, Me.Picture1, Val(Text5), Val(Text7), Me.chkShowCentre, Me.chkShowSep, Me.chkShowAlign, Me.chkShowSensor, Me.chkShowBox
    
        DoEvents
        
        MoveBoid flock, Val(Text6), Picture1.ScaleHeight, Picture1.ScaleWidth, Val(Text5), Me.chkEnclosed
        
    
    Loop While blnRun = True
    
    
Case "Stop"
    cmdExecute.Caption = "Execute"
    blnRun = False
    
End Select
    
 End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdMove_Click()
    MoveBoid flock, Val(Text6), Picture1.ScaleHeight, Picture1.ScaleWidth, Val(Text5), Me.chkEnclosed
End Sub

Private Sub cmdReset_Click()
Dim boid As BoidClass

    For Each boid In flock

        Select Case boid.id
        Case 0
            Text1(0) = 100
            Text1(1) = 150
            Text1(2) = 0
        Case 1
            Text2(0) = 100
            Text2(1) = 160
            Text2(2) = 0
        Case 2
            Text3(0) = 400
            Text3(1) = 430
            Text3(2) = 2
        Case 3
            Text4(0) = 400
            Text4(1) = 440
            Text4(2) = 2
        End Select
        
    Next
    Set boid = Nothing
End Sub

Private Sub cmdRun_Click()

    Picture1.Cls

Dim boid As BoidClass

    CalcForces flock, HScroll1(0), HScroll1(1), HScroll1(2), Val(Text5), Val(Text7)

    For Each boid In flock

        Select Case boid.id
            Case 0
                Text1(0) = boid.X
                Text1(1) = boid.Y
                Text1(2) = boid.direction
                Text1(3) = boid.ClosestDist
                
                Text1(4) = boid.DesireAlignTurn
                Text1(5) = boid.DesireAlignWeight
                Text1(6) = boid.DesireCentreTurn
                Text1(7) = boid.DesireCentreWeight
                Text1(8) = boid.DesireSeparateTurn
                Text1(9) = boid.DesireSeparateWeight
                Text1(10) = boid.AveX
                Text1(11) = boid.AveY
            Case 1
                Text2(0) = boid.X
                Text2(1) = boid.Y
                Text2(2) = boid.direction
                Text2(3) = boid.ClosestDist
                
                Text2(4) = boid.DesireAlignTurn
                Text2(5) = boid.DesireAlignWeight
                Text2(6) = boid.DesireCentreTurn
                Text2(7) = boid.DesireCentreWeight
                Text2(8) = boid.DesireSeparateTurn
                Text2(9) = boid.DesireSeparateWeight
                Text2(10) = boid.AveX
                Text2(11) = boid.AveY
            Case 2
                Text3(0) = boid.X
                Text3(1) = boid.Y
                Text3(2) = boid.direction
                Text3(3) = boid.ClosestDist
    
                Text3(3) = boid.AveSpeed
                Text3(4) = boid.DesireAlignTurn
                Text3(5) = boid.DesireAlignWeight
                Text3(6) = boid.DesireCentreTurn
                Text3(7) = boid.DesireCentreWeight
                Text3(8) = boid.DesireSeparateTurn
                Text3(9) = boid.DesireSeparateWeight
                Text3(10) = boid.AveX
                Text3(11) = boid.AveY
            Case 3
                Text4(0) = boid.X
                Text4(1) = boid.Y
                Text4(2) = boid.direction
                Text4(3) = boid.ClosestDist
    
                Text4(3) = boid.AveSpeed
                Text4(4) = boid.DesireAlignTurn
                Text4(5) = boid.DesireAlignWeight
                Text4(6) = boid.DesireCentreTurn
                Text4(7) = boid.DesireCentreWeight
                Text4(8) = boid.DesireSeparateTurn
                Text4(9) = boid.DesireSeparateWeight
                Text4(10) = boid.AveX
                Text4(11) = boid.AveY
        End Select
        
    Next
    Set boid = Nothing

Dim i%, j%

    For i% = 0 To Picture1.ScaleWidth Step 10
        Picture1.Line (i%, 0)-(i%, Picture1.ScaleHeight), vbButtonFace
    Next
    For i% = 0 To Picture1.ScaleHeight Step 10
        Picture1.Line (0, i%)-(Picture1.ScaleWidth, i%), vbButtonFace
    Next
    
    DrawBoid flock, Me.Picture1, Me.chkShowColours, Me.chkShowArrow, Me.chkShowCircle
    DrawObjects objects, Me.Picture1

    
    DrawForces flock, Me.Picture1, Val(Text5), Val(Text7), Me.chkShowCentre, Me.chkShowSep, Me.chkShowAlign, Me.chkShowSensor, Me.chkShowBox

    MoveBoid flock, Val(Text6), Picture1.ScaleHeight, Picture1.ScaleWidth, Val(Text5), Me.chkEnclosed

End Sub

Private Sub cmdSet_Click()
Dim i%
Dim H As Integer
Dim Theta As Double

Dim X As Integer
Dim Y As Integer
Dim NewX As Integer
Dim NewY As Integer


    Do While flock.Count > 0
        flock.Remove 1
    Loop
    
    

'Boid0
    
    X = Val(Text1(0))
    Y = Val(Text1(1))
    Theta = Val(Text1(2))
    H = Val(Text1(3))
    
    AddBoid flock, X, Y, Theta, vbBlue
    
    Sin (1)
'Boid1
    
    X = Val(Text2(0))
    Y = Val(Text2(1))
    Theta = Val(Text2(2))
    H = Val(Text2(3))
    
    AddBoid flock, X, Y, Theta, vbBlack
    
'Boid2

    X = Val(Text3(0))
    Y = Val(Text3(1))
    Theta = Val(Text3(2))
    H = Val(Text3(3))

    AddBoid flock, X, Y, Theta, vbBlack
'Boid2

    X = Val(Text4(0))
    Y = Val(Text4(1))
    Theta = Val(Text4(2))
    H = Val(Text4(3))


    AddBoid flock, X, Y, Theta, vbBlack
        
End Sub

Private Sub Command1_Click()
    Do While objects.Count > 0
        objects.Remove 1
    Loop
    
End Sub

Private Sub Form_Load()
Dim i%
    
    Randomize
    
    Label3(6) = "Boid Connt : " & HScroll1(3).Value
    For i% = 0 To Val(Mid$(Label3(6), 13))
        AddBoid flock, Int(Rnd(1) * 500), Int(Rnd(1) * 500), Rnd(1) * 6, vbBlack
    Next

'    For i% = 0 To 4
'        AddObstacle objects, Int(Rnd(1) * 300) + 100, Int(Rnd(1) * 300) + 100, 25
'    Next

    AddStartupObs
    
End Sub

Sub AddStartupObs()
        
        Do While objects.Count > 0
            objects.Remove 1
        Loop
        
        Picture1.Cls
        AddObstacle objects, Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2, 20
        AddObstacle objects, (Picture1.ScaleWidth / 2) + 100, Picture1.ScaleHeight / 2, 20
        AddObstacle objects, Picture1.ScaleWidth / 2, (Picture1.ScaleHeight / 2) + 100, 20
        AddObstacle objects, (Picture1.ScaleWidth / 2) - 100, Picture1.ScaleHeight / 2, 20
        AddObstacle objects, Picture1.ScaleWidth / 2, (Picture1.ScaleHeight / 2) - 100, 20
        
        AddObstacle objects, 0, 0, 50
        AddObstacle objects, Picture1.ScaleWidth, 0, 50
        AddObstacle objects, 0, Picture1.ScaleHeight, 50
        AddObstacle objects, Picture1.ScaleWidth, Picture1.ScaleHeight, 50

End Sub
Private Sub Form_Resize()
    
    Picture1.Width = Me.Width - Frame4.Width - 200
    Picture1.Height = Me.Height - 900
    AddStartupObs
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdExecute.Caption = "Stop" Then
        cmdExecute.Value = True
    End If
    End
End Sub

Private Sub HScroll1_Change(Index As Integer)

    Select Case Index
    Case 3
        Label3(6) = "Boid Connt : " & HScroll1(3).Value
        Do While flock.Count > Val(Mid$(Label3(6), 13))
            flock.Remove 1
        Loop
        Do While flock.Count < Val(Mid$(Label3(6), 13))
            AddBoid flock, Int(Rnd(1) * 500), Int(Rnd(1) * 500), Rnd(1) * 6, vbBlack
        Loop
       
    End Select
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'    iXval = X
'    iYval = Y

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    iXval = X
    iYval = Y
    AddObstacle objects, iXval, iYval, 25

End Sub

