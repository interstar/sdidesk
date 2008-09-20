VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Network"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' It's now a *decorator* of a page,
' ie. it implements the Page interface
' and contains a page.

Implements Page

' The idea is that Page itself shouldn't have to know about
' networks (or tables)
' So when we turn into one, we just create the tabe or network
' as a wrapper.

' this manages a network (containing nodes and arcs)
' including knowing how to draw it on a canvas (though this should be changed)
' and knowing how to parse one from our plain-text format
' and write it back

Private nodes() As Node
Private arcs() As Arc
Public noNodes As Integer      ' number of nodes
Private arraySize As Integer    ' actual dimension of array

Public drawSize As Long        ' how big to draw boxes
Public drawAspect As Single    ' aspect ratio for drawing

Public innerPage As Page      ' to do the pagey stuff

Public Sub init(nn As Integer, ds As Long, da As Single)
  arraySize = nn
  ReDim nodes(arraySize)
  ReDim arcs(arraySize, arraySize)
  Dim i As Integer, j As Integer
  For i = 0 To arraySize - 1
    For j = 0 To arraySize - 1
      Set arcs(i, j) = New Arc
    Next j
  Next i
  noNodes = 0
  drawSize = ds
  drawAspect = da
End Sub

Public Function asPage() As Page
    Set asPage = Me
End Function

Public Function nextId() As Integer
  nextId = noNodes
End Function

Public Sub addNode(ex As Long, wy As Long, name As String)
  If noNodes > arraySize Then
    MsgBox ("too many nodes")
  Else
    Set nodes(noNodes) = New Node
    Call nodes(noNodes).init(ex, wy, name)
    noNodes = noNodes + 1
  End If
End Sub

Public Function getNode(id As Integer) As Node
  Set getNode = nodes(id)
End Function

Public Function getArc(i As Integer, j As Integer) As Arc
  Set getArc = arcs(i, j)
End Function

Public Function getIdByName(n As String) As Integer
  Dim i As Integer
  For i = 0 To noNodes - 1
    If Not nodes(i) Is Nothing Then
        If nodes(i).name = n Then
          getIdByName = i
          Exit Function
        End If
    End If
  Next i
  getIdByName = -1
End Function


Public Function addArc(n1 As Integer, n2 As Integer, l As String, d As ArcDirectionality, ex As Long, wy As Long, a As Single) As Boolean
  Set arcs(n1, n2) = New Arc
  arcs(n1, n2).label = l
  arcs(n1, n2).exists = True
  arcs(n1, n2).direction = d
  arcs(n1, n2).x = ex
  arcs(n1, n2).y = wy
  arcs(n1, n2).angle = a
  arcs(n1, n2).n1 = n1
  arcs(n1, n2).n2 = n2
End Function

Public Sub removeArc(n1 As Integer, n2 As Integer)
  Set arcs(n1, n2) = Nothing
End Sub

Public Sub removeArcByArc(a As Arc)
  Call removeArc(a.n1, a.n2)
End Sub

Public Sub removeNode(n As Integer)
    Dim i As Integer
    For i = 0 To noNodes - 1
        Call removeArc(n, i)
        Call removeArc(i, n)
    Next i
    Set nodes(n) = Nothing
End Sub

Public Sub removeNodeByNode(n As Node)
    Dim i As Integer
    For i = 0 To noNodes - 1
        If nodes(i) Is n Then
            Call removeNode(i)
            Exit Sub
        End If
    Next i
End Sub

Public Function hitNodeDetect(ex As Single, wy As Single) As Integer
  Dim whichHit As Integer
  Dim i As Integer
  whichHit = -1
  For i = 0 To noNodes - 1
    If Not nodes(i) Is Nothing Then
        If nodes(i).hit(CLng(ex), CLng(wy)) Then
            whichHit = i
        End If
    End If
  Next i
  hitNodeDetect = whichHit
End Function


Public Function hitArcDetect(ex As Single, wy As Single) As Point
  Dim p As New Point
  Dim i As Integer, j As Integer
  p.x = -1 ' defaults if no arcs hit
  p.y = -1 ' defaults if no arcs hit
  For i = 0 To noNodes - 1
    For j = 0 To noNodes - 1
      If arcs(i, j).hit(CLng(ex), CLng(wy)) Then
        p.x = i
        p.y = j
      End If
    Next j
  Next i
  Set hitArcDetect = p
End Function


Public Function half(x1 As Long, x2 As Long) As Long
   half = x1 + ((x2 - x1) / 2)
End Function


Public Sub connectNodes(fromNode As Integer, toNode As Integer, l As String, d As ArcDirectionality)
  If fromNode <= noNodes And toNode <= noNodes And fromNode >= 0 And toNode >= 0 Then
    Dim hx As Long, hy As Long, a As Single
    hx = half(getNode(fromNode).x, getNode(toNode).x)
    hy = half(getNode(fromNode).y, getNode(toNode).y)
    
    ' angle is arcTan of dY/dX
    Dim xDif As Integer
    xDif = getNode(fromNode).x - getNode(toNode).x
    If xDif = 0 Then xDif = 1
      
    a = Atn((getNode(fromNode).y - getNode(toNode).y) / xDif)
    
    Call addArc(fromNode, toNode, l, d, hx, hy, a)
  End If
End Sub

Public Sub parseFromPrettyPersist(s As String)
' pretty persist form of a network is
' #Network,, version
' nodeName1,, x,, y
' nodeName2,, x,, y
' ----
' node1,, node2,, type,, directionality
' ----
   Dim version As Integer
   Dim lines() As String
   Dim parts() As String
   lines = Split(s, vbCrLf)
   Dim first As String
   first = lines(0)

   If Left(first, 8) <> "#Network" Then
     MsgBox ("error in network::parseFromPrettyPersist")
   Else
     parts = Split(first, ",, ")
     version = parts(1)
     If CInt(version) <> 1 Then
       MsgBox ("only know how to use version 1")
     Else
       ' count the number of nodes
       noNodes = 0
       Do While lines(noNodes + 1) <> "----"
          noNodes = noNodes + 1
       Loop
       ' OK, now create the nodes
       Call init(noNodes, drawSize, drawAspect)
       Dim i As Integer
       i = 1
       Do While lines(i) <> "----"
          parts = Split(lines(i), ",, ")
          Call addNode(CLng(parts(1)), CLng(parts(2)), parts(0))
          i = i + 1
       Loop
       i = i + 1
       Do While lines(i) <> "----"
          parts = Split(lines(i), ",, ")
          Dim n1 As Integer, n2 As Integer
          n1 = getIdByName(parts(0))
          n2 = getIdByName(parts(1))
          Dim label As String
          label = ""
          Dim direct As ArcDirectionality
                    
          If UBound(parts) > 1 Then
             direct = CInt(parts(2))
          End If
          
          If UBound(parts) > 2 Then
            label = parts(3)
          End If
          
          Call connectNodes(n1, n2, label, direct)
          i = i + 1
       Loop
     End If
   End If
End Sub

Public Function spitAsPrettyPersist() As String
  Dim s As String
  Dim i As Integer
  
  s = "#Network,, 1" + vbCrLf
  For i = 0 To noNodes - 1
    If Not nodes(i) Is Nothing Then
        s = s + nodes(i).name & ",, " & nodes(i).x & ",, " & nodes(i).y & vbCrLf
    End If
  Next i
  s = s + "----" & vbCrLf
  Dim j As Integer
  For i = 0 To noNodes - 1
    For j = 0 To noNodes - 1
      If Not (arcs(i, j) Is Nothing) Then
        If arcs(i, j).exists = True Then
         s = s + nodes(i).name & ",, " & nodes(j).name & ",, " & CInt(arcs(i, j).direction) & ",, " & arcs(i, j).label & vbCrLf
        End If
      End If
    Next j
  Next i
  s = s + "----" & vbCrLf
  spitAsPrettyPersist = s
End Function

Private Property Let Page_categories(ByVal RHS As String)
    innerPage.categories = RHS
End Property

Private Property Get Page_categories() As String
    Page_categories = innerPage.categories
End Property

Private Sub Page_cook(prep As PagePreparer, chef As PageCooker, backlinks As Boolean)
    Call innerPage.cook(prep, chef, backlinks)
End Sub

Private Property Let Page_cooked(ByVal RHS As String)
    innerPage.cooked = RHS
End Property

Private Property Get Page_cooked() As String
    Page_cooked = innerPage.cooked
End Property

Private Property Let Page_createdDate(ByVal RHS As Date)
    innerPage.createdDate = RHS
End Property

Private Property Get Page_createdDate() As Date
    Page_createdDate = innerPage.createdDate
End Property

Private Function Page_getDataDictionary() As VCollection
    Set Page_getDataDictionary = innerPage.getDataDictionary
End Function

Private Function Page_getFirstLine() As String
    Page_getFirstLine = innerPage.getFirstLine
End Function

Private Function Page_getMyType() As String
    Page_getMyType = innerPage.getMyType
End Function

Private Function Page_getRedirectPage() As String
    Page_getRedirectPage = innerPage.getRedirectPage
End Function

Private Function Page_getTable() As table
    Set Page_getTable = innerPage.getTable
End Function

Private Function Page_getVal(key As String) As String
    Page_getVal = innerPage.getVal(key)
End Function

Private Function Page_hasVar(key As String) As Boolean
    Page_hasVar = innerPage.hasVar(key)
End Function

Private Function Page_isNetwork() As Boolean
    Page_isNetwork = True
End Function

Private Function Page_isNew() As Boolean
    Page_isNew = False
End Function

Private Function Page_isRedirect() As Boolean
    Page_isRedirect = False
End Function

Private Function Page_isTable() As Boolean
    Page_isTable = False
End Function

Private Property Let Page_lastEdited(ByVal RHS As Date)
    innerPage.lastEdited = RHS
End Property

Private Property Get Page_lastEdited() As Date
    Page_lastEdited = innerPage.lastEdited
End Property

Private Property Let Page_pageName(ByVal RHS As String)
    innerPage.pageName = RHS
End Property

Private Property Get Page_pageName() As String
    Page_pageName = innerPage.pageName
End Property

Private Property Let Page_pageType(ByVal RHS As String)
End Property

Private Property Get Page_pageType() As String
    Page_pageType = "network"
End Property

Private Sub Page_prepare(prep As PagePreparer, backlinks As Boolean)
    Call innerPage.prepare(prep, backlinks)
End Sub

Private Property Let Page_prepared(ByVal RHS As String)
    innerPage.prepared = RHS
End Property

Private Property Get Page_prepared() As String
    Page_prepared = innerPage.prepared
End Property

Private Property Let Page_raw(ByVal RHS As String)
    innerPage.raw = RHS
End Property

Private Property Get Page_raw() As String
    Page_raw = innerPage.raw
End Property

Private Sub Page_setVal(aKey As String, aVal As String)
    Call innerPage.setVal(aKey, aVal)
End Sub

Private Function Page_spawnCopy() As Page
    Set Page_spawnCopy = innerPage.spawnCopy()
End Function

Private Function Page_varsToString() As String
    Page_varsToString = innerPage.varsToString
End Function

Private Function Page_wordCount() As Integer
    Page_wordCount = innerPage.wordCount
End Function
