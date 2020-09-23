Attribute VB_Name = "Module1"
Option Explicit
Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpPoint As Long) As Long

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020         ' (DWORD) dest = source
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const SRCAND = &H8800C6          ' (DWORD) dest = source AND dest

Public Const PI = 3.1415926
Public Const PI2 = 3.1415926 * 2

Public flock As New Collection
Public objects As New Collection

Public Sub AddBoid(flock As Collection, X As Integer, Y As Integer, ByVal Dir As Integer, Bcol As Long)
'helper function to add Boid to the specified collection (flock)
Dim Colour As Integer
Dim boid As BoidClass
Set boid = New BoidClass

    boid.X = X
    boid.Y = Y
    
    boid.Colour = Bcol
    boid.id = flock.Count
    
    boid.direction = Dir
    boid.speed = 10
    
    flock.Add boid
    Set boid = Nothing
End Sub

Public Sub AddObstacle(objects As Collection, X As Integer, Y As Integer, ByVal Radius As Integer)
Dim obs As ObstacleClass
Set obs = New ObstacleClass

    obs.X = X
    obs.Y = Y
    
    obs.id = objects.Count
    
    obs.Radius = Radius
    
    objects.Add obs
    Set obs = Nothing
    
End Sub

Sub DrawBoid(flock As Collection, Canvas As PictureBox, ShowColours As Boolean, ShowArrow As Boolean, ShowCircle As Boolean)
Dim boid As BoidClass
Dim d As Integer
Dim u%
Dim NewX As Integer
Dim NewY As Integer

Dim XDist As Integer
Dim YDist As Integer

Dim AHx As Integer
Dim AHy As Integer
Dim Theta As Integer
Dim Bcol As Long

    For Each boid In flock

        Theta = boid.direction
        
        If ShowColours = True Then
            Bcol = boid.Colour
        Else
            Bcol = vbBlack
        End If
        
        boid.NewY = boid.Y + (10 * Sin(boid.direction))
        boid.NewX = boid.X + (10 * Cos(boid.direction))
        
        Canvas.Line (boid.X, boid.Y)-(boid.NewX, boid.NewY), Bcol
        
        If ShowCircle Then
            Canvas.Circle (boid.X, boid.Y), 5, Bcol
        End If
        
    'arrow head
        If ShowArrow Then
            AHx = 5 * Cos((Theta + 45))
            AHy = 5 * Sin((Theta + 45))
            Canvas.Line (boid.NewX, boid.NewY)-(boid.NewX - AHx, boid.NewY - AHy), Bcol
            AHx = 5 * Cos((Theta - 45))
            AHy = 5 * Sin((Theta - 45))
            Canvas.Line (boid.NewX, boid.NewY)-(boid.NewX - AHx, boid.NewY - AHy), Bcol
        End If
        
    Next
    Set boid = Nothing


End Sub

Sub DrawObjects(objects As Collection, Canvas As PictureBox)
Dim obs As ObstacleClass
    
    For Each obs In objects

        Canvas.Circle (obs.X, obs.Y), obs.Radius
        
    Next
    Set obs = Nothing


End Sub

Sub DrawForces(flock As Collection, Canvas As PictureBox, SensorDist As Integer, ViewTheta As Single, ShowCentre As Boolean, ShowSep As Boolean, ShowAlign As Boolean, ShowSensor As Boolean, ShowBox As Boolean)
Dim boid As BoidClass
Dim d As Integer
Dim u%

Dim tmpX1 As Integer
Dim tmpY1 As Integer
Dim tmpX2 As Integer
Dim tmpY2 As Integer
Dim tmpX3 As Integer
Dim tmpY3 As Integer
Dim tmpX4 As Integer
Dim tmpY4 As Integer

Dim tmpStart As Single
Dim tmpEnd As Single
Dim HalfTheta As Single

    HalfTheta = ViewTheta / 2
    For Each boid In flock

            If ShowSensor Then
            
                tmpX1 = boid.X + (SensorDist * Cos(boid.direction + HalfTheta))
                tmpY1 = boid.Y + (SensorDist * Sin(boid.direction + HalfTheta))
                tmpX2 = boid.X + (SensorDist * Cos(boid.direction - HalfTheta))
                tmpY2 = boid.Y + (SensorDist * Sin(boid.direction - HalfTheta))
                
                tmpStart = PI2 - (boid.direction + HalfTheta)
                tmpEnd = PI2 - (boid.direction - HalfTheta)
                   
                'Debug.Print tmpStart, tmpEnd
                
                If tmpStart > PI2 Then
                    tmpStart = tmpStart - PI2
                End If
                If tmpStart < 0 Then
                    tmpStart = tmpStart + PI2
                End If
                
                If tmpEnd > PI2 Then
                    tmpEnd = tmpEnd - PI2
                End If
                If tmpEnd < 0 Then
                    tmpEnd = tmpEnd + PI2
                End If
            
                Canvas.Circle (boid.X, boid.Y), SensorDist, vbBlack, tmpStart, tmpEnd
                Canvas.Line (boid.X, boid.Y)-(tmpX1, tmpY1), vbBlack
                Canvas.Line (boid.X, boid.Y)-(tmpX2, tmpY2), vbBlack
                
            End If

            If ShowCentre Then
                Canvas.Line (boid.X, boid.Y)-(boid.X + boid.DesireCentreX * 10, boid.Y + boid.DesireCentreY * 10), vbGreen
            End If
            
            If ShowAlign Then
                Canvas.Line (boid.X, boid.Y)-(boid.X + boid.DesireAlignX * 10, boid.Y + boid.DesireAlignY * 10), vbMagenta
            End If
            
            If ShowSep Then
                Canvas.Line (boid.X, boid.Y)-(boid.X + boid.DesireSeparateX * 10, boid.Y + boid.DesireSeparateY * 10), vbBlue
            End If


'show box used for collision detection

            If ShowBox Then
            'box to the right
                tmpX1 = boid.X + (5 * Cos(boid.direction + PI / 2))
                tmpY1 = boid.Y + (5 * Sin(boid.direction + PI / 2))
                tmpX2 = tmpX1 + (SensorDist * Cos(boid.direction))
                tmpY2 = tmpY1 + (SensorDist * Sin(boid.direction))
    
                If boid.DesireAvoidRight = False Then
                    Canvas.Line (boid.X, boid.Y)-(tmpX1, tmpY1), vbScrollBars
                    Canvas.Line (tmpX1, tmpY1)-(tmpX2, tmpY2), vbScrollBars
                Else
                    Canvas.Line (boid.X, boid.Y)-(tmpX1, tmpY1), vbRed
                    Canvas.Line (tmpX1, tmpY1)-(tmpX2, tmpY2), vbRed
                End If
                
            'box to the left
                tmpX3 = boid.X - (5 * Cos(boid.direction + PI / 2))
                tmpY3 = boid.Y - (5 * Sin(boid.direction + PI / 2))
                tmpX4 = tmpX3 + (SensorDist * Cos(boid.direction))
                tmpY4 = tmpY3 + (SensorDist * Sin(boid.direction))

                If boid.DesireAvoidLeft = False Then
                    Canvas.Line (boid.X, boid.Y)-(tmpX3, tmpY3), vbScrollBars
                    Canvas.Line (tmpX3, tmpY3)-(tmpX4, tmpY4), vbScrollBars
                Else
                    Canvas.Line (boid.X, boid.Y)-(tmpX3, tmpY3), vbRed
                    Canvas.Line (tmpX3, tmpY3)-(tmpX4, tmpY4), vbRed
                End If
                
            'complete box
                Canvas.Line (tmpX2, tmpY2)-(tmpX4, tmpY4), vbScrollBars
            End If
            
    Next
    Set boid = Nothing

End Sub

Public Sub CalcForces(flock As Collection, CentMult As Integer, SepMult As Integer, AliMult As Integer, SensorDist As Integer, ViewTheta As Single)
    
    Dim distance As Integer
    Dim i%
    Dim AveDir As Single
    Dim AveX As Integer
    Dim AveY As Integer
    Dim AveSpeed As Single
    
    Dim boid As BoidClass
    Dim obs As ObstacleClass
    
    Dim otherBoid As BoidClass
        
    Dim ClosestBoid As BoidClass
    
    Dim iLeaderX As Integer
    Dim iLeaderY As Integer
    
    Dim CloseBoidCount As Integer
    Dim GroupCount As Integer
    'Dim CloseBoidCount As Integer
    
    Dim AllDirChange As Single
'    Dim SensorDist As Integer
    
    Dim ClosestDist As Single
    Dim TmpDist As Single
    Dim TooClose As Boolean
    
    Dim TmpWeight As Double
    
    Dim AngDiff As Single
    Dim Angle As Single
    Dim HalfTheta As Single
    
    Dim blnResult As Boolean
    
    Dim X1 As Integer
    Dim Y1 As Integer
    Dim X2 As Integer
    Dim Y2 As Integer
    
    Dim LeftDist As Single
    Dim RightDist As Single
    
'================================================================================
'================================================================================
'================================================================================
    
    'SensorDist = 50
    
    HalfTheta = ViewTheta / 2
    
    AllDirChange = 0
    For Each boid In flock

        ClosestDist = SensorDist * 2
        boid.ClosestDist = 0
        TmpDist = 1000

        boid.AveX = boid.X
        boid.AveY = boid.Y
        boid.AveDir = boid.direction
        boid.AveSpeed = boid.speed
        boid.CentreDist = ClosestDist
        boid.DesireAvoidX = 0
        boid.DesireAvoidY = 0
        boid.DesireAvoidWeight = 0
        boid.DesireAvoidRight = False
        boid.DesireAvoidLeft = False
        CloseBoidCount = 1

        'Get Average information from flockmates in sensor range
        For Each otherBoid In flock

            If boid.id <> otherBoid.id Then 'as long as it's not itself

                distance = Abs(1 + Sqr((boid.X - otherBoid.X) ^ 2 + (boid.Y - otherBoid.Y) ^ 2))
                
                If (boid.X - otherBoid.X) <> 0 Then
                    Angle = Abs(Atn((boid.Y - otherBoid.Y) / (boid.X - otherBoid.X)))
                Else
                    Angle = Abs(Atn(90))
                End If
                
                If boid.X > otherBoid.X Then
                    If boid.Y > otherBoid.Y Then
                        Angle = Angle + PI
                    Else
                        Angle = PI - Angle
                    End If
                Else
                    If boid.Y > otherBoid.Y Then
                        Angle = (PI2) - Angle
                    Else
                        Angle = Angle
                    End If
                End If
                
                If (distance < SensorDist) Then
                
                    AngDiff = Abs(boid.direction - Angle)
                    
                    If AngDiff > PI Then
                        AngDiff = AngDiff - PI
                    End If
                    
                    If AngDiff < HalfTheta Then

                        'see if it is the closest
                        If distance < TmpDist Then
                            ClosestDist = distance
                            TmpDist = distance
                            Set ClosestBoid = otherBoid
                        End If
    
                        CloseBoidCount = CloseBoidCount + 1
    
                        boid.AveDir = boid.AveDir + otherBoid.direction
                        boid.AveX = boid.AveX + otherBoid.X
                        boid.AveY = boid.AveY + otherBoid.Y
                        boid.AveSpeed = boid.AveSpeed + otherBoid.speed
                        
                    End If
    
                End If

            End If
            TmpDist = SensorDist
        Next

        'Debug.Print

'averages
        boid.AveDir = boid.AveDir / CloseBoidCount
        boid.AveX = boid.AveX / CloseBoidCount
        boid.AveY = boid.AveY / CloseBoidCount
        boid.AveSpeed = boid.AveSpeed / CloseBoidCount

        If CloseBoidCount > 1 Then
'align
                boid.CentreDist = Abs(Sqr((boid.X - boid.AveX) ^ 2 + (boid.Y - boid.AveY) ^ 2))
                If boid.CentreDist = 0 Then boid.CentreDist = 1
                boid.DesireAlignTurn = boid.AveDir
                boid.DesireAlignWeight = 5
                
'centre
                boid.CentreDist = (Sqr((boid.X - boid.AveX) ^ 2 + (boid.Y - boid.AveY) ^ 2))
                If boid.CentreDist = 0 Then boid.CentreDist = 1
                
                If boid.X <> boid.AveX Then
                    boid.DesireCentreTurn = Abs(Atn((boid.Y - boid.AveY) / (boid.X - boid.AveX)))
                Else
                    boid.DesireCentreTurn = Atn(boid.Y - boid.AveY)
                End If
                
                'boid.DesireCentreWeight = (CSng(boid.CentreDist) * CSng(boid.CentreDist) * CSng(boid.CentreDist)) / 200
                boid.DesireCentreWeight = Log((CSng(boid.CentreDist) * CSng(boid.CentreDist) * CSng(boid.CentreDist)) / 200) + 5
                
                
                If boid.X > boid.AveX Then
                    If boid.Y > boid.AveY Then
                        boid.DesireCentreTurn = boid.DesireCentreTurn + PI
                    Else
                        boid.DesireCentreTurn = PI - boid.DesireCentreTurn
                    End If
                Else
                    If boid.Y > boid.AveY Then
                        boid.DesireCentreTurn = (PI2) - boid.DesireCentreTurn
                    Else
                        boid.DesireCentreTurn = boid.DesireCentreTurn
                    End If
                End If
'Separate
                boid.ClosestX = ClosestBoid.X
                boid.ClosestY = ClosestBoid.Y
                boid.ClosestDist = Abs(1 + Sqr((boid.X - ClosestBoid.X) ^ 2 + (boid.Y - ClosestBoid.Y) ^ 2))
    
                If boid.X <> ClosestBoid.X Then
                    boid.DesireSeparateTurn = Abs(Atn((boid.Y - ClosestBoid.Y) / (boid.X - ClosestBoid.X)))
                Else
                    boid.DesireSeparateTurn = Abs(Atn((boid.Y - ClosestBoid.Y)))
                End If
    
                If boid.X > boid.ClosestX Then
                    If boid.Y > boid.ClosestY Then
                        boid.DesireSeparateTurn = boid.DesireSeparateTurn + PI
                    Else
                        boid.DesireSeparateTurn = PI - boid.DesireSeparateTurn
                    End If
                Else
                    If boid.Y > boid.ClosestY Then
                        boid.DesireSeparateTurn = (PI2) - boid.DesireSeparateTurn
                    Else
                        boid.DesireSeparateTurn = boid.DesireSeparateTurn
                    End If
                End If
                boid.DesireSeparateTurn = boid.DesireSeparateTurn + PI
    
                If boid.DesireSeparateTurn > (PI2) Then boid.DesireSeparateTurn = boid.DesireSeparateTurn - (PI2)
    
                'boid.DesireSeparateWeight = (2000 / (CSng(boid.ClosestDist) * CSng(boid.ClosestDist) * CSng(boid.ClosestDist)))
                boid.DesireSeparateWeight = Log((2000 / (CSng(boid.ClosestDist) * CSng(boid.ClosestDist) * CSng(boid.ClosestDist)))) + 10
                
                                
            
'normalise

            TmpWeight = Abs(boid.DesireAlignWeight) + Abs(boid.DesireCentreWeight) + Abs(boid.DesireSeparateWeight)

            If TmpWeight > 0 Then

                boid.DesireCentreWeight = Abs(boid.DesireCentreWeight / TmpWeight) * CentMult
                boid.DesireSeparateWeight = Abs(boid.DesireSeparateWeight / TmpWeight) * SepMult
                boid.DesireAlignWeight = Abs(boid.DesireAlignWeight / TmpWeight) * AliMult


            End If

            boid.DesireCentreX = (boid.DesireCentreWeight * Cos(boid.DesireCentreTurn) * CentMult)
            boid.DesireCentreY = (boid.DesireCentreWeight * Sin(boid.DesireCentreTurn) * CentMult)

            boid.DesireSeparateX = (boid.DesireSeparateWeight * Cos(boid.DesireSeparateTurn) * SepMult)
            boid.DesireSeparateY = (boid.DesireSeparateWeight * Sin(boid.DesireSeparateTurn) * SepMult)

            boid.DesireAlignX = (boid.DesireAlignWeight * Cos(boid.DesireAlignTurn) * AliMult)
            boid.DesireAlignY = (boid.DesireAlignWeight * Sin(boid.DesireAlignTurn) * AliMult)

            
        Else

            boid.DesireAlignTurn = 0
            boid.DesireAlignWeight = 0
            boid.DesireAlignX = 0
            boid.DesireAlignY = 0

            boid.DesireCentreTurn = 0
            boid.DesireCentreWeight = 0
            boid.DesireCentreX = 0
            boid.DesireCentreY = 0

            boid.ClosestX = 0
            boid.ClosestY = 0
            boid.ClosestDist = 0

            boid.DesireSeparateTurn = 0
            boid.DesireSeparateWeight = 0
            boid.DesireSeparateX = 0
            boid.DesireSeparateY = 0

            boid.ClosestDist = SensorDist * 2
        End If

'================================================================================
'================================================================================
'================================================================================
 
 'obstacle avoidance
 
        distance = 0
        TmpDist = 2 * SensorDist
        
        For Each obs In objects
        
            distance = Abs(1 + Sqr((boid.X - obs.X) ^ 2 + (boid.Y - obs.Y) ^ 2))
            
            If distance < obs.Radius Then
                'Debug.Print boid.id, distance
                'Stop
            End If
            
            If distance < SensorDist + obs.Radius Then
                
                'box to the right
                X1 = boid.X + (5 * Cos(boid.direction + PI / 2))
                Y1 = boid.Y + (5 * Sin(boid.direction + PI / 2))
                X2 = X1 + (SensorDist * Cos(boid.direction))
                Y2 = Y1 + (SensorDist * Sin(boid.direction))
                
                RightDist = SegCirInt(X1, Y1, X2, Y2, obs.X, obs.Y, obs.Radius)
                'Debug.Print X1, Y1, X2, Y2, obs.X, obs.Y, obs.Radius, blnResult
                
                'box to the left
                X1 = boid.X - (5 * Cos(boid.direction + PI / 2))
                Y1 = boid.Y - (5 * Sin(boid.direction + PI / 2))
                X2 = X1 + (SensorDist * Cos(boid.direction))
                Y2 = Y1 + (SensorDist * Sin(boid.direction))
            
                LeftDist = SegCirInt(X1, Y1, X2, Y2, obs.X, obs.Y, obs.Radius)
                'Debug.Print X1, Y1, X2, Y2, obs.X, obs.Y, obs.Radius, blnResult
            
                If LeftDist <> 0 Or RightDist <> 0 Then
                    
                    If LeftDist <> 0 And RightDist <> 0 Then
                        Select Case LeftDist - RightDist
                        Case Is < 0
                            'Debug.Print LeftDist, RightDist, "TURN RIGHT"
                            'Need to move right
                            boid.DesireAvoidTurn = boid.direction + PI / 4
                            boid.DesireAvoidLeft = True
                            
                            boid.Colour = vbRed
                        Case Is > 0
                            'Debug.Print LeftDist, RightDist, "TURN LEFT"
                            'Need to move left
                            boid.DesireAvoidTurn = boid.direction - PI / 4
                            boid.DesireAvoidRight = True
    
                            boid.Colour = vbBlue
                        Case Else
                            'Debug.Print LeftDist, RightDist, "RANDOM TURN"
                            i% = Int(2 * Rnd)
                            If i% = 0 Then
                                boid.DesireAvoidTurn = boid.direction - PI / 4
                                boid.Colour = vbYellow
                            Else
                                boid.DesireAvoidTurn = boid.direction + PI / 4
                                boid.Colour = vbRed
                            End If
                            
                            boid.DesireAvoidRight = True
                            boid.DesireAvoidLeft = True
                        End Select
                        
                    Else
                        If LeftDist <> 0 Then
                            boid.DesireAvoidTurn = boid.direction + PI / 4
                            boid.DesireAvoidLeft = True
                            
                            boid.Colour = vbMagenta
                        End If
                        If RightDist <> 0 Then
                            boid.DesireAvoidTurn = boid.direction - PI / 4
                            boid.DesireAvoidRight = True
    
                            boid.Colour = vbCyan
                        End If
                    End If
                    
                If boid.DesireAvoidTurn > PI2 Then
                    boid.DesireAvoidTurn = boid.DesireAvoidTurn - PI2
                End If
                If boid.DesireAvoidTurn < 0 Then
                    boid.DesireAvoidTurn = boid.DesireAvoidTurn + PI2
                End If
                
                boid.DesireAvoidX = 10 * Cos(boid.DesireAvoidTurn)
                boid.DesireAvoidY = 10 * Sin(boid.DesireAvoidTurn)
                boid.DesireAvoidWeight = 5

                End If
            End If
        
        Next
    Next

'================================================================================
'================================================================================
'================================================================================

    Set boid = Nothing
    Set obs = Nothing
    Set ClosestBoid = Nothing
    Set otherBoid = Nothing
    
End Sub

Sub MoveBoid(flock As Collection, MaxTurn As Single, iHeight As Integer, iWidth As Integer, SensorDist As Integer, Enclosed As Boolean)

    Dim boid As BoidClass
    Dim NewX As Integer
    Dim NewY As Integer
    
'================================================================================
'================================================================================
'================================================================================
Dim NewDir As Single
 Dim Diff As Single
 
    For Each boid In flock

'store starting pos
            boid.Y = boid.Y + (boid.speed * Sin(boid.direction))
            boid.X = boid.X + (boid.speed * Cos(boid.direction))
            
'select action
            If boid.DesireAvoidWeight = 0 Then
                Select Case boid.ClosestDist
                Case 0 To 5
                    boid.DesireSeparateY = boid.DesireSeparateY + Int(Rnd(1) * 10)
                    boid.DesireSeparateX = boid.DesireSeparateX + Int(Rnd(1) * 10)
                Case 6 To 20
                    boid.DesireAlignWeight = 0
                    boid.DesireCentreWeight = 0
                    boid.DesireSeparateY = boid.DesireSeparateY * 5
                    boid.DesireSeparateX = boid.DesireSeparateX * 5
                Case Is > SensorDist
    '                Stop
                Case Else
                End Select
    
    'start from last finish point
                NewY = boid.Y
                NewX = boid.X
    'add forces
                NewY = NewY + boid.DesireAlignY
                NewX = NewX + boid.DesireAlignX
                
                NewY = NewY + boid.DesireCentreY
                NewX = NewX + boid.DesireCentreX
                
                NewY = NewY + boid.DesireSeparateY
                NewX = NewX + boid.DesireSeparateX
            Else
            
                NewY = boid.Y
                NewX = boid.X
                
                NewY = NewY + boid.DesireAvoidY * 10
                NewX = NewX + boid.DesireAvoidX * 10
              
            End If

            If NewX = boid.X And NewY = boid.Y Then
                'Stop
            Else
            
'add direction
                NewY = NewY + (Sin((boid.direction)) * boid.speed)
                NewX = NewX + (Cos((boid.direction)) * boid.speed)
    
'get new dir
                'If boid.Y - NewY = 0 Then NewY = NewY + 1
                
                If (boid.X - NewX) <> 0 Then
                    NewDir = Abs(Atn((boid.Y - NewY) / (boid.X - NewX)))
                Else
                    NewDir = Atn(90)
                End If
    
'add static turn
                If boid.X > NewX Then
                    If boid.Y > NewY Then
                        NewDir = Abs(NewDir) + PI
                    Else
                        NewDir = PI - Abs(NewDir)
                    End If
                Else
                    If boid.Y > NewY Then
                        NewDir = (PI2) - Abs(NewDir)
                    Else
                        NewDir = Abs(NewDir)
                    End If
                End If
     
'check for max allowed turn
                If boid.direction > NewDir Then
                
                    If (boid.direction - NewDir) > PI Then
                        Diff = ((PI2) - (boid.direction - NewDir))
                        If Diff > MaxTurn Then
                            Diff = MaxTurn
                        End If
                        boid.direction = boid.direction + Diff
                    Else
                        Diff = (boid.direction - NewDir)
                        If Diff > MaxTurn Then
                            Diff = MaxTurn
                        End If
                        boid.direction = boid.direction - Diff
                    End If
                    
                Else
                
                    If (NewDir - boid.direction) > PI Then
                    
                        Diff = ((PI2) - (NewDir - boid.direction))
                        If Diff > MaxTurn Then
                            Diff = MaxTurn
                        End If
                        boid.direction = boid.direction - Diff
                        
                    Else
                        Diff = (NewDir - boid.direction)
                        If Diff > MaxTurn Then
                            Diff = MaxTurn
                        End If
                        boid.direction = boid.direction + Diff
                    End If
                End If
                
                
                If boid.direction > (PI2) Then
                    boid.direction = boid.direction - (PI2)
                End If
                If boid.direction < 0 Then
                    boid.direction = boid.direction + (PI2)
                End If
                            
            End If
            
            boid.NewY = boid.Y + boid.speed * Sin(boid.direction)
            boid.NewX = boid.X + boid.speed * Cos(boid.direction)




'================================================================================
'================================================================================
'================================================================================
'    If boid.CentreDist = 0 Then Stop

        Select Case boid.CentreDist
        Case Is < SensorDist
            If Abs(boid.DesireCentreTurn - boid.direction) > PI Then
                boid.speed = boid.speed - 0.1
                boid.Colour = vbRed
            Else
                boid.speed = boid.speed + 0.1
                boid.Colour = vbGreen
            End If
        Case Is > SensorDist
                boid.speed = boid.speed + 0.1
                boid.Colour = vbGreen
        End Select

'================================================================================
'================================================================================
'================================================================================

        Do
            If boid.direction > (PI2) Then boid.direction = boid.direction - (PI2)
            If boid.direction < 0 Then boid.direction = boid.direction + (PI2)
            
            If boid.speed > 10 Then
                boid.speed = 10
                boid.Colour = &HC000&
            End If
            
            If boid.speed < 5 Then
                boid.speed = 5
                boid.Colour = &H40C0&
            End If
            
        Loop While (boid.direction > (PI2)) Or (boid.direction < 0)

'================================================================================
'================================================================================
'================================================================================

        If Enclosed = False Then
        
            If boid.X < 0 Then boid.X = iWidth
            If boid.X > iWidth Then boid.X = 0
            
            If boid.Y < 0 Then boid.Y = iHeight
            If boid.Y > iHeight Then boid.Y = 0
        
        Else
        
            If boid.X > iWidth Then
                boid.X = iWidth
                Select Case boid.direction
                Case 0 To PI / 2
                    boid.direction = PI - boid.direction
                Case PI + PI / 2 To PI2
                    boid.direction = PI + (PI2 - boid.direction)
                Case Else
                    'Stop
                End Select
            End If
                
            If boid.X < 0 Then
                boid.X = 0
                Select Case boid.direction
                Case PI / 2 To PI
                    boid.direction = PI - boid.direction
                Case PI To PI + PI / 2
                    boid.direction = PI2 - (boid.direction - PI)
                Case Else
                    'Stop
                End Select
            End If
                
            If boid.Y > iHeight Then
                boid.Y = iHeight
                Select Case boid.direction
                Case PI / 2 To PI
                    boid.direction = PI + (PI - boid.direction)
                Case 0 To PI
                    boid.direction = PI2 - boid.direction
                Case Else
                    'Stop
                End Select
            End If
                
            If boid.Y < 0 Then
                boid.Y = 0
                Select Case boid.direction
                Case PI To PI + PI / 2
                    boid.direction = PI - (boid.direction - PI)
                Case PI + PI / 2 To PI2
                    boid.direction = PI2 - boid.direction
                Case Else
                    'Stop
                End Select
            End If
        End If
        
'================================================================================
'================================================================================
'================================================================================

Next
    
End Sub

Sub RandDir(flock As Collection)
Dim boid As BoidClass

    For Each boid In flock
        boid.direction = Rnd * 360
    Next
    
End Sub

Function SegCirInt(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, X3 As Integer, Y3 As Integer, r As Integer) As Single
Dim a As Double
Dim b As Double
Dim intClosestPointX As Double
Dim intClosestPointY As Double
Dim cpDist As Single
Dim result As Double
Dim lngX1 As Long
Dim lngX2 As Long
Dim lngX3 As Long
Dim lngY1 As Long
Dim lngY2 As Long
Dim lngY3 As Long


    lngX1 = X1
    lngX2 = X2
    lngX3 = X3
    lngY1 = Y1
    lngY2 = Y2
    lngY3 = Y3
    
'check if closest point of line to centre of the object is less than the radius
'of the object

    a = (lngX3 - lngX1) * (lngX2 - lngX1) + (lngY3 - lngY1) * (lngY2 - lngY1)
    b = (lngX2 - lngX1) * (lngX2 - lngX1) + (lngY2 - lngY1) * (lngY2 - lngY1)
    
    result = a / b
    
    If (result > 0) And (result < 1) Then
    'closest point is between the start and end points of the line
    'so now check the distance to the centre of the object
    
        intClosestPointX = X1 + (X2 - X1) * result
        intClosestPointY = Y1 + (Y2 - Y1) * result
        
        cpDist = Sqr(((intClosestPointX - X3) * (intClosestPointX - X3)) + ((intClosestPointY - Y3) * (intClosestPointY - Y3)))
        If cpDist < r Then
            SegCirInt = cpDist
        Else
            SegCirInt = 0
        End If
    Else
    'closest point is outside the start and end points
    'Now check if either end of the line is inside the circle
    
        'check end of line
        cpDist = Sqr((lngX2 - lngX3) * (lngX2 - lngX3) + (lngY2 - lngY3) * (lngY2 - lngY3))
        If cpDist < r Then
            SegCirInt = cpDist
        Else
            'check  start of line
            cpDist = Sqr((lngX1 - lngX3) * (lngX1 - lngX3) + (lngY1 - lngY3) * (lngY1 - lngY3))
            If cpDist < r Then
                SegCirInt = cpDist
            Else
                SegCirInt = 0
            End If
        End If
        
    End If
    
    
End Function

Function LineCirInt(X1, Y1, X2, Y2, X3, Y3, r) As Boolean
Dim a As Long
Dim b As Long
Dim c As Double
Dim result As Long

    a = ((X2 - X1) * (X2 - X1)) + ((Y2 - Y1) * (Y2 - Y1))
    b = 2 * ((X2 - X1) * (X1 - X3) + (Y2 - Y1) * (Y1 - Y3))
    c = (X3 * X3) + (Y3 * Y3) + (X1 * X1) + (Y1 * Y1) - 2 * (X3 * X1 + Y3 * Y1) - (r * r)
    
    result = (b * b) - 4 * a * c
    
    If result > 0 Then
        LineCirInt = True
    Else
        LineCirInt = False
    End If
    
    
End Function


