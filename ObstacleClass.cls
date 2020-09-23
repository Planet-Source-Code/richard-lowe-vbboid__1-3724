VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObstacleClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mvarid As Integer 'local copy
Private mvarX As Integer 'local copy
Private mvarY As Integer 'local copy
Private mvarRadius As Integer 'local copy

Public Property Let id(ByVal vData As Integer)
    mvarid = vData
End Property

Public Property Get id() As Integer
    id = mvarid
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

Public Property Let Radius(ByVal vData As Integer)
    mvarRadius = vData
End Property

Public Property Get Radius() As Integer
    Radius = mvarRadius
End Property


