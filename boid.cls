VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoidClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvardirection As Single 'local copy
Private mvarspeed As Single 'local copy
Private mvarid As Integer 'local copy
Private mvarX As Integer 'local copy
Private mvarY As Integer 'local copy
Private mvarNewX As Integer 'local copy
Private mvarNewY As Integer 'local copy
Private mvarColour As Long 'local copy

Private mvarAveX As Integer 'local copy
Private mvarAveY As Integer 'local copy
Private mvarAveDir As Single 'local copy
Private mvarAveSpeed As Single 'local copy
Private mvarCentreDist As Single 'local copy

Private mvarClosestX As Integer 'local copy
Private mvarClosestY As Integer 'local copy
Private mvarClosestDist As Single 'local copy

Private mvarDesireCentreTurn As Single 'local copy
Private mvarDesireCentreWeight As Double 'local copy
Private mvarDesireCentreX As Integer 'local copy
Private mvarDesireCentreY As Integer 'local copy

Private mvarDesireAlignTurn As Single 'local copy
Private mvarDesireAlignWeight As Double 'local copy
Private mvarDesireAlignX As Integer 'local copy
Private mvarDesireAlignY As Integer 'local copy

Private mvarDesireSeparateTurn As Single 'local copy
Private mvarDesireSeparateWeight As Double 'local copy
Private mvarDesireSeparateX As Integer 'local copy
Private mvarDesireSeparateY As Integer 'local copy

Private mvarDesireAvoidTurn As Single 'local copy
Private mvarDesireAvoidWeight As Double 'local copy
Private mvarDesireAvoidX As Integer 'local copy
Private mvarDesireAvoidY As Integer 'local copy

Private mvarDesireAvoidLeft As Boolean 'local copy
Private mvarDesireAvoidRight As Boolean 'local copy


Public Property Let id(ByVal vData As Integer)
    mvarid = vData
End Property

Public Property Get id() As Integer
    id = mvarid
End Property

Public Property Let Colour(ByVal vData As Long)
    mvarColour = vData
End Property

Public Property Get Colour() As Long
    Colour = mvarColour
End Property

Public Property Let Y(ByVal vData As Integer)
    mvarY = vData
End Property

Public Property Get Y() As Integer
    Y = mvarY
End Property

Public Property Let X(ByVal vData As Integer)
    mvarX = vData
End Property

Public Property Get X() As Integer
    X = mvarX
End Property


Public Property Let NewY(ByVal vData As Integer)
    mvarNewY = vData
End Property

Public Property Get NewY() As Integer
    NewY = mvarNewY
End Property

Public Property Let NewX(ByVal vData As Integer)
    mvarNewX = vData
End Property

Public Property Get NewX() As Integer
    NewX = mvarNewX
End Property


Public Property Let AveX(ByVal vData As Integer)
    mvarAveX = vData
End Property

Public Property Get AveX() As Integer
    AveX = mvarAveX
End Property

Public Property Let AveY(ByVal vData As Integer)
    mvarAveY = vData
End Property

Public Property Get AveY() As Integer
    AveY = mvarAveY
End Property

Public Property Let CentreDist(ByVal vData As Integer)
    mvarCentreDist = vData
End Property

Public Property Get CentreDist() As Integer
    CentreDist = mvarCentreDist
End Property

Public Property Let direction(ByVal vData As Single)
    mvardirection = vData
End Property

Public Property Get direction() As Single
    direction = mvardirection
End Property

Public Property Let speed(ByVal vData As Single)
    mvarspeed = vData
End Property

Public Property Get speed() As Single
    speed = mvarspeed
End Property

Public Property Let AveDir(ByVal vData As Single)
    mvarAveDir = vData
End Property

Public Property Get AveDir() As Single
    AveDir = mvarAveDir
End Property

Public Property Let AveSpeed(ByVal vData As Single)
    mvarAveSpeed = vData
End Property

Public Property Get AveSpeed() As Single
    AveSpeed = mvarAveSpeed
End Property

Public Property Let ClosestX(ByVal vData As Integer)
    mvarClosestX = vData
End Property

Public Property Get ClosestX() As Integer
    ClosestX = mvarClosestX
End Property

Public Property Let ClosestY(ByVal vData As Integer)
    mvarClosestY = vData
End Property

Public Property Get ClosestY() As Integer
    ClosestY = mvarClosestY
End Property

Public Property Let ClosestDist(ByVal vData As Integer)
    mvarClosestDist = vData
End Property

Public Property Get ClosestDist() As Integer
    ClosestDist = mvarClosestDist
End Property



Public Property Let DesireCentreTurn(ByVal vData As Single)
    mvarDesireCentreTurn = vData
End Property

Public Property Get DesireCentreTurn() As Single
    DesireCentreTurn = mvarDesireCentreTurn
End Property

Public Property Let DesireCentreWeight(ByVal vData As Single)
    mvarDesireCentreWeight = vData
End Property

Public Property Get DesireCentreWeight() As Single
    DesireCentreWeight = mvarDesireCentreWeight
End Property

Public Property Let DesireCentreX(ByVal vData As Integer)
    mvarDesireCentreX = vData
End Property

Public Property Get DesireCentreX() As Integer
    DesireCentreX = mvarDesireCentreX
End Property

Public Property Let DesireCentreY(ByVal vData As Integer)
    mvarDesireCentreY = vData
End Property

Public Property Get DesireCentreY() As Integer
    DesireCentreY = mvarDesireCentreY
End Property



Public Property Let DesireAlignTurn(ByVal vData As Single)
    mvarDesireAlignTurn = vData
End Property

Public Property Get DesireAlignTurn() As Single
    DesireAlignTurn = mvarDesireAlignTurn
End Property

Public Property Let DesireAlignWeight(ByVal vData As Single)
    mvarDesireAlignWeight = vData
End Property

Public Property Get DesireAlignWeight() As Single
    DesireAlignWeight = mvarDesireAlignWeight
End Property

Public Property Let DesireAlignX(ByVal vData As Integer)
    mvarDesireAlignX = vData
End Property

Public Property Get DesireAlignX() As Integer
    DesireAlignX = mvarDesireAlignX
End Property

Public Property Let DesireAlignY(ByVal vData As Integer)
    mvarDesireAlignY = vData
End Property

Public Property Get DesireAlignY() As Integer
    DesireAlignY = mvarDesireAlignY
End Property




Public Property Let DesireSeparateTurn(ByVal vData As Single)
    mvarDesireSeparateTurn = vData
End Property

Public Property Get DesireSeparateTurn() As Single
    DesireSeparateTurn = mvarDesireSeparateTurn
End Property

Public Property Let DesireSeparateWeight(ByVal vData As Single)
    mvarDesireSeparateWeight = vData
End Property

Public Property Get DesireSeparateWeight() As Single
    DesireSeparateWeight = mvarDesireSeparateWeight
End Property

Public Property Let DesireSeparateX(ByVal vData As Integer)
    mvarDesireSeparateX = vData
End Property

Public Property Get DesireSeparateX() As Integer
    DesireSeparateX = mvarDesireSeparateX
End Property

Public Property Let DesireSeparateY(ByVal vData As Integer)
    mvarDesireSeparateY = vData
End Property

Public Property Get DesireSeparateY() As Integer
    DesireSeparateY = mvarDesireSeparateY
End Property





Public Property Let DesireAvoidTurn(ByVal vData As Single)
    mvarDesireAvoidTurn = vData
End Property

Public Property Get DesireAvoidTurn() As Single
    DesireAvoidTurn = mvarDesireAvoidTurn
End Property

Public Property Let DesireAvoidWeight(ByVal vData As Single)
    mvarDesireAvoidWeight = vData
End Property

Public Property Get DesireAvoidWeight() As Single
    DesireAvoidWeight = mvarDesireAvoidWeight
End Property

Public Property Let DesireAvoidX(ByVal vData As Integer)
    mvarDesireAvoidX = vData
End Property

Public Property Get DesireAvoidX() As Integer
    DesireAvoidX = mvarDesireAvoidX
End Property

Public Property Let DesireAvoidY(ByVal vData As Integer)
    mvarDesireAvoidY = vData
End Property

Public Property Get DesireAvoidY() As Integer
    DesireAvoidY = mvarDesireAvoidY
End Property

Public Property Let DesireAvoidLeft(ByVal vData As Boolean)
    mvarDesireAvoidLeft = vData
End Property

Public Property Get DesireAvoidLeft() As Boolean
    DesireAvoidLeft = mvarDesireAvoidLeft
End Property

Public Property Let DesireAvoidRight(ByVal vData As Boolean)
    mvarDesireAvoidRight = vData
End Property

Public Property Get DesireAvoidRight() As Boolean
    DesireAvoidRight = mvarDesireAvoidRight
End Property





