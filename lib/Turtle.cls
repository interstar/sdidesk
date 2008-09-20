VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Turtle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const twoPi = 3.14159265 * 2

Public Enum penState
  up = 0
  down = 1
End Enum

Public Enum collideStyle
  ignore = 0
  die = 1
  rotate = 2
End Enum

Public x As Single  ' x co-ord
Public y As Single  ' y co-ord

Public angle As Single ' angle of movement
Public velocity As Single ' velocity

Public dx As Single  ' velocity in x direction
Public dy As Single  ' velocity in y direction

Public vx As Single  ' distance to look ahead x
Public vy As Single  ' distance to look ahead y

Public r As Single  ' brushRadius

Public maxAge As Integer  ' how many time steps before expires
Public currentAge As Integer ' current time steps

Public state As penState ' is pen up or down
Public colour As Long ' colour of pen

Public collide As collideStyle ' what to do when something in way
Public turnAngle As Single ' turn angle
Public turnState As Boolean ' am I currently turning?
Public turnStart As Single ' remember which angle the turn started

Public Canvas As PictureBox  ' canvas

Public Sub setUp(c As PictureBox)
  Set Canvas = c
  turnState = False
End Sub

Public Sub setPos(ex As Single, wy As Single)
  x = ex
  y = wy
End Sub

Public Sub setAngle(a As Single)
  angle = a
  Call setVel(angleToDx(a, velocity), angleToDy(a, velocity))
  Dim s As String
  s = CStr(dx) + ", " + CStr(dy)
'  MsgBox (s)
End Sub

Public Sub turn(a As Single)
 ' MsgBox (angle)
  Call setAngle(angle + a)
End Sub

Public Sub setVel(ex As Single, wy As Single)
  dx = ex
  dy = wy
  vx = ex * 2.3
  vy = wy * 2.3
End Sub

' trig stuff

Public Function angleToDx(a As Single, v As Single) As Single
  angleToDx = Cos(a) * v
End Function

Public Function angleToDy(a As Single, v As Single) As Single
  angleToDy = Sin(a) * v
End Function

Public Function oneTouch() As Boolean
  Canvas.FillStyle = 0
  Canvas.FillColor = colour
  Canvas.Circle (x, y), r, colour
End Function

Public Function testCollide() As Boolean
  ' see if a collision
  Dim j As Integer
  Dim sx As Single
  Dim sy As Single
  Dim hit As Boolean
  Dim PCol As Long
  
  hit = False
  For j = 1 To 10
'     sx = r / 2 * vx + j * vx
'     sy = r / 2 * vy + j * vy
     PCol = Canvas.Point(x + vx, y + vy)
     If PCol <> RGB(255, 255, 255) Then
        hit = True
     End If
  Next j
  testCollide = hit
End Function

Public Function onCollide() As Boolean
  ' what to do if collision occurs
  Select Case collide
    Case die:
      currentAge = 0
    Case rotate:
      If turnState = True Then
        If turnAngle > 0 And angle > turnStart + twoPi Then
          ' turned over 360 degrees, no way out so die
          MsgBox (CStr(turnAngle) + ", " + CStr(angle))
          currentAge = 0
        End If
        If turnAngle < 0 And angle < turnStart - twoPi Then
        ' turned over 360 degrees, no way out so die
          currentAge = 0
        End If
      Else
        ' not currently turning, so start us off
        turnState = True
        turnStart = angle
      End If
      
      Call turn(turnAngle)
          
    Case Else:
      turnState = False ' we ain't turning no more, reset flag
      Call normalForward
   End Select
End Function

Public Sub normalForward()
     ' mark and move on
     If state = down Then
        Call oneTouch
     End If
     ' move on
     x = x + dx
     y = y + dy
End Sub

Public Sub moveForward(dist As Integer)
  Canvas.FillStyle = 0
  Canvas.FillColor = colour
  Canvas.Line (x, y)-(x + (dx * dist), y + dy * dist)

End Sub

Public Function nextStep(finger As Boolean) As Boolean
  ' finger is whether mouse button is down or not, true = yes
  ' return false if dies
   
  ' collide test
  If testCollide() Then
     Call onCollide
  Else
     Call normalForward
  End If
    
  If finger = False Then currentAge = currentAge - 1
  
  If currentAge <= 0 Then
      nextStep = False
  Else
      nextStep = True
  End If

End Function

Public Function pi() As Double
  pi = twoPi / 2
End Function
