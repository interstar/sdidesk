VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VseCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This is the code which manages a Visual Structure Editor
' Canvas. Doesn't actually create the canvas (that's given by
' the main form, but does know about and manipulate it

' has all the handlers for things you click on etc.

Public Canvas As PictureBox
Public RawText As RichTextBox
Public chef As PageCooker
Public mainForm As Form
Public nim As NetworkInfoManager

Public Enum vseMode
   View
   Edit
End Enum

Public mode As vseMode ' is the canvas in view or edit mode

Private Type vsePoint
   x As Long
   y As Long
End Type

Private dragFlag As Boolean ' true when drawing an arc
Private dragStartNode As Integer ' which node started from
Private dragStartPoint As vsePoint ' used when drawing
Private dragEndPoint As vsePoint ' used when drawing


Public Sub init(p As PictureBox, rt As RichTextBox, pc As PageCooker, mf As Form)
  Set Canvas = p
  
  Set RawText = rt
  Set chef = pc
  Set mainForm = mf
  Set nim = New NetworkInfoManager
  Call nim.init(NodeInfo, ArcInfo, Me)
  mode = View
  dragFlag = False
End Sub

Public Sub setMode(m As vseMode)
  mode = m
  If m = View Then
    Canvas.BackColor = RGB(255, 255, 255)
  Else
    Canvas.BackColor = RGB(200, 200, 255)
  End If
End Sub

Public Sub changed(n As Network)
     n.innerPage.raw = n.spitAsPrettyPersist
     n.parseFromPrettyPersist (n.innerPage.raw)
     Call draw(n, mode)
     RawText.text = n.spitAsPrettyPersist
     mainForm.MagicNotebook.getSingleUserState.changesSaved = False
End Sub

Public Sub MouseDownOnCanvas(n As Network, Button As Integer, shift As Integer, x As Single, y As Single)
   Dim p As Page
   Set p = n
   
   Dim pn As String ' page name if we're going anywhere else
   Dim newName As String
   
   ' detect hitting nodes
   Dim nodeId As Integer
   nodeId = n.hitNodeDetect(x, y)
   Dim arcId As Point
   Set arcId = n.hitArcDetect(x, y)
   
   Dim theNode As Node
   
   If mode = View Then
        ' in view mode
     
        If Button = 1 Then
        
            If nodeId <> -1 Then
                ' clicked on a node
                Set theNode = n.getNode(nodeId)
                pn = theNode.name
                Call mainForm.controller.actionLoad(pn, False)
            End If
    
            If arcId.x <> -1 Then
                ' clicked on an arc
        
                Dim theArc As Arc
                Set theArc = n.getArc(arcId.x, arcId.y)
        
                pn = theArc.label
                If pn <> "" Then
                    Call mainForm.controller.actionLoad(pn, False)
                End If
            End If
    
        Else
            ' switch into edit mode if right click.
            ' this is counter-intuitive but convenient. Which is better?
            Call mainForm.controller.actionEdit( _
                mainForm.MagicNotebook.getControllableModel.getCurrentPage.pageName, False)
        End If
          
   Else
     ' edit mode
     
     If nodeId = -1 Then
       ' haven't hit an existing node,
       
       ' have we hit an arc?
       If arcId.x = -1 Then
         ' didn't hit an arc either,
         'must want to add a new node
         Call n.addNode(CLng(x), CLng(y), "new node " + CStr(n.nextId))
         Call changed(n)
       Else
         ' hit an arc
         Set theArc = n.getArc(arcId.x, arcId.y)
         If Button = 2 Then
           Call nim.editAnArc(theArc, n, p)
         End If
      End If
     
     Else
       ' hit an existing node
       Set theNode = n.getNode(nodeId)
       
       If Button = 1 Then
         Call startDrag(nodeId)
       Else
         Call nim.editANode(theNode, n, p)
       End If
       
     End If
   End If
End Sub

Public Sub startDrag(nodeId)
  dragStartNode = nodeId
  dragFlag = True
End Sub

Public Sub endDrag(n As Network, Button As Integer, shift As Integer, x As Long, y As Long)
  
  If dragFlag = True Then
    dragFlag = False
    Dim p As Page
    Set p = n
    
    Dim theNode As Node
    Dim nodeIndex As Integer
    
    nodeIndex = n.hitNodeDetect(CLng(x), CLng(y))
    If nodeIndex > -1 Then
      ' dragged to another node, so connect them
     If nodeIndex <> dragStartNode Then
        ' only connect nodes if they aren't the same
        Call n.connectNodes(dragStartNode, nodeIndex, "", noDirection)
        Call changed(n)
      End If
    Else
     ' dragged to nowhere, let's see if we're moving the node
       Set theNode = n.getNode(dragStartNode)
       theNode.x = x
       theNode.y = y
       Call changed(n)
    End If
  End If
End Sub

Public Sub drawArc(n1 As Node, n2 As Node, a As Arc, m As vseMode)
  Canvas.Line (n1.x, n1.y)-(n2.x, n2.y)
  If m = Edit Or a.label <> "" Then
     ' draw label
     Canvas.FillColor = RGB(200, 255, 255)
     Canvas.FillStyle = 0
     Canvas.Circle (a.x, a.y), 100, RGB(255, 255, 255), , , 1
     Canvas.FillStyle = 1
     Canvas.Circle (a.x, a.y), 100, RGB(0, 0, 0), , , 1
     
     Canvas.Line (a.x - 180, a.y + 50)-(a.x - 180, a.y + 50)
     Canvas.Print a.label
  End If
  
  If a.direction = one Then
     ' draw one arrow
     Dim turt As New Turtle
     Call turt.setUp(Canvas)
     turt.state = down
     turt.collide = Ignore
     turt.velocity = 5
     
     Dim i As Integer, apx As Single, apy As Single
     apx = (n2.x - a.x) / 2 + a.x
     apy = (n2.y - a.y) / 2 + a.y
     
     If n2.x > n1.x Then
     
       Call turt.setPos(apx, apy)
       Call turt.setAngle(a.angle + turt.pi - turt.pi / 8)
       Call turt.moveForward(40)
     
       Call turt.setPos(apx, apy)
       Call turt.setAngle(a.angle + turt.pi + turt.pi / 8)
       Call turt.moveForward(40)
     
     Else
    
       Call turt.setPos(apx, apy)
       Call turt.setAngle(a.angle - turt.pi / 8)
       Call turt.moveForward(40)
     
       Call turt.setPos(apx, apy)
       Call turt.setAngle(a.angle + turt.pi / 8)
       Call turt.moveForward(40)
    
    
    End If
    
  End If
  
  
End Sub

Public Sub drawNode(n As Node, nt As Network)
  Dim w As Long, h As Long
  w = Canvas.TextWidth(n.name) + nt.drawSize
  h = Canvas.TextHeight(n.name) + nt.drawSize * nt.drawAspect
  Call n.setHitBox(n.x - w / 2, n.y, n.x + w / 2, n.y + h)
  Canvas.FillStyle = 1
  Canvas.FillColor = RGB(255, 200, 255)
  Canvas.ForeColor = RGB(0, 0, 0)
  Canvas.FillStyle = 0
'  Canvas.Circle (n.x, n.y), nt.drawSize, RGB(0, 0, 0), , , nt.drawAspect
  Canvas.Line (n.x - w / 2, n.y)-(n.x + w / 2, n.y + h), RGB(255, 255, 255), BF
  Canvas.FillColor = RGB(255, 255, 255)
  Canvas.Line (n.x - w / 2, n.y)-(n.x + w / 2, n.y + h), RGB(0, 0, 0), B
  Canvas.ForeColor = RGB(255, 255, 255)
  Canvas.Line (n.x - w / 2 + nt.drawSize / 2, n.y + nt.drawSize * nt.drawAspect * 0.5)-(n.x - w / 2 + nt.drawSize / 2, n.y + nt.drawSize * nt.drawAspect * 0.5), RGB(255, 255, 255), B
  Canvas.ForeColor = RGB(0, 0, 0)
  Canvas.Print n.name
  
End Sub

Public Sub draw(n As Network, m As vseMode)
  Me.Canvas.Cls
  Dim i As Integer, j As Integer
  For i = 0 To n.noNodes - 1
    For j = 0 To n.noNodes - 1
       If Not n.getArc(i, j) Is Nothing And n.getArc(i, j).exists = True Then
         Call drawArc(n.getNode(i), n.getNode(j), n.getArc(i, j), m)
       End If
    Next j
  Next i
  For i = 0 To n.noNodes - 1
    If Not n.getNode(i) Is Nothing Then
      Call drawNode(n.getNode(i), n)
    End If
  Next i
End Sub


