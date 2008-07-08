Option Strict Off
Option Explicit On
Friend Class Network
	Implements _Page
	
	' It's now a *decorator* of a page,
	' ie. it implements the Page interface
	' and contains a page.
	
	
	' The idea is that Page itself shouldn't have to know about
	' networks (or tables)
	' So when we turn into one, we just create the tabe or network
	' as a wrapper.
	
	' this manages a network (containing nodes and arcs)
	' including knowing how to draw it on a canvas (though this should be changed)
	' and knowing how to parse one from our plain-text format
	' and write it back
	
	Private nodes() As Node
    Private arcs(0, 0) As Arc
	Public noNodes As Short ' number of nodes
	Private arraySize As Short ' actual dimension of array
	
	Public drawSize As Integer ' how big to draw boxes
	Public drawAspect As Single ' aspect ratio for drawing
	
	Public innerPage As _Page ' to do the pagey stuff
	
	Public Sub init(ByRef nn As Short, ByRef ds As Integer, ByRef da As Single)
		arraySize = nn
		ReDim nodes(arraySize)
		ReDim arcs(arraySize, arraySize)
		Dim i, j As Short
		For i = 0 To arraySize - 1
			For j = 0 To arraySize - 1
				arcs(i, j) = New Arc
			Next j
		Next i
		noNodes = 0
		drawSize = ds
		drawAspect = da
	End Sub
	
	Public Function asPage() As _Page
		asPage = Me
	End Function
	
	Public Function nextId() As Short
		nextId = noNodes
	End Function
	
	Public Sub addNode(ByRef ex As Integer, ByRef wy As Integer, ByRef name As String)
		If noNodes > arraySize Then
			MsgBox("too many nodes")
		Else
			nodes(noNodes) = New Node
			Call nodes(noNodes).init(ex, wy, name)
			noNodes = noNodes + 1
		End If
	End Sub
	
	Public Function getNode(ByRef id As Short) As Node
		getNode = nodes(id)
	End Function
	
	Public Function getArc(ByRef i As Short, ByRef j As Short) As Arc
		getArc = arcs(i, j)
	End Function
	
	Public Function getIdByName(ByRef n As String) As Short
		Dim i As Short
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
	
	
	Public Function addArc(ByRef n1 As Short, ByRef n2 As Short, ByRef l As String, ByRef d As Arc.ArcDirectionality, ByRef ex As Integer, ByRef wy As Integer, ByRef a As Single) As Boolean
		arcs(n1, n2) = New Arc
		arcs(n1, n2).label = l
		arcs(n1, n2).exists = True
		arcs(n1, n2).direction = d
		arcs(n1, n2).x = ex
		arcs(n1, n2).y = wy
		arcs(n1, n2).angle = a
		arcs(n1, n2).n1 = n1
		arcs(n1, n2).n2 = n2
	End Function
	
	Public Sub removeArc(ByRef n1 As Short, ByRef n2 As Short)
		'UPGRADE_NOTE: Object arcs() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		arcs(n1, n2) = Nothing
	End Sub
	
	Public Sub removeArcByArc(ByRef a As Arc)
		Call removeArc((a.n1), (a.n2))
	End Sub
	
	Public Sub removeNode(ByRef n As Short)
		Dim i As Short
		For i = 0 To noNodes - 1
			Call removeArc(n, i)
			Call removeArc(i, n)
		Next i
		'UPGRADE_NOTE: Object nodes() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		nodes(n) = Nothing
	End Sub
	
	Public Sub removeNodeByNode(ByRef n As Node)
		Dim i As Short
		For i = 0 To noNodes - 1
			If nodes(i) Is n Then
				Call removeNode(i)
				Exit Sub
			End If
		Next i
	End Sub
	
	Public Function hitNodeDetect(ByRef ex As Single, ByRef wy As Single) As Short
		Dim whichHit As Short
		Dim i As Short
		whichHit = -1
		For i = 0 To noNodes - 1
			If Not nodes(i) Is Nothing Then
				'UPGRADE_WARNING: Couldn't resolve default property of object nodes(i).hit(CLng(ex), CLng(wy)). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If nodes(i).hit(CInt(ex), CInt(wy)) Then
					whichHit = i
				End If
			End If
		Next i
		hitNodeDetect = whichHit
	End Function
	
	
	Public Function hitArcDetect(ByRef ex As Single, ByRef wy As Single) As Point
		Dim p As New Point
		Dim i, j As Short
		p.x = -1 ' defaults if no arcs hit
		p.y = -1 ' defaults if no arcs hit
		For i = 0 To noNodes - 1
			For j = 0 To noNodes - 1
				If arcs(i, j).hit(CInt(ex), CInt(wy)) Then
					p.x = i
					p.y = j
				End If
			Next j
		Next i
		hitArcDetect = p
	End Function
	
	
	Public Function half(ByRef x1 As Integer, ByRef x2 As Integer) As Integer
		half = x1 + ((x2 - x1) / 2)
	End Function
	
	
	Public Sub connectNodes(ByRef fromNode As Short, ByRef toNode As Short, ByRef l As String, ByRef d As Arc.ArcDirectionality)
		Dim xDif As Short
		Dim hx, hy As Integer
		Dim a As Single
		If fromNode <= noNodes And toNode <= noNodes And fromNode >= 0 And toNode >= 0 Then
			hx = half((getNode(fromNode).x), (getNode(toNode).x))
			hy = half((getNode(fromNode).y), (getNode(toNode).y))
			
			' angle is arcTan of dY/dX
			xDif = getNode(fromNode).x - getNode(toNode).x
			If xDif = 0 Then xDif = 1
			
			a = System.Math.Atan((getNode(fromNode).y - getNode(toNode).y) / xDif)
			
			Call addArc(fromNode, toNode, l, d, hx, hy, a)
		End If
	End Sub
	
	Public Sub parseFromPrettyPersist(ByRef s As String)
		' pretty persist form of a network is
		' #Network,, version
		' nodeName1,, x,, y
		' nodeName2,, x,, y
		' ----
		' node1,, node2,, type,, directionality
		' ----
		Dim version As Short
		Dim lines() As String
		Dim parts() As String
		lines = Split(s, vbCrLf)
		Dim first As String
		first = lines(0)
		
		Dim i As Short
		Dim n1, n2 As Short
		Dim label As String
		Dim direct As Arc.ArcDirectionality
		If Left(first, 8) <> "#Network" Then
			MsgBox("error in network::parseFromPrettyPersist")
		Else
			parts = Split(first, ",, ")
			version = CShort(parts(1))
			If CShort(version) <> 1 Then
				MsgBox("only know how to use version 1")
			Else
				' count the number of nodes
				noNodes = 0
				Do While lines(noNodes + 1) <> "----"
					noNodes = noNodes + 1
				Loop 
				' OK, now create the nodes
				Call init(noNodes, drawSize, drawAspect)
				i = 1
				Do While lines(i) <> "----"
					parts = Split(lines(i), ",, ")
					Call addNode(CInt(parts(1)), CInt(parts(2)), parts(0))
					i = i + 1
				Loop 
				i = i + 1
				Do While lines(i) <> "----"
					parts = Split(lines(i), ",, ")
					n1 = getIdByName(parts(0))
					n2 = getIdByName(parts(1))
					label = ""
					
					If UBound(parts) > 1 Then
						direct = CShort(parts(2))
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
		Dim i As Short
		
		s = "#Network,, 1" & vbCrLf
		For i = 0 To noNodes - 1
			If Not nodes(i) Is Nothing Then
				s = s & nodes(i).name & ",, " & nodes(i).x & ",, " & nodes(i).y & vbCrLf
			End If
		Next i
		s = s & "----" & vbCrLf
		Dim j As Short
		For i = 0 To noNodes - 1
			For j = 0 To noNodes - 1
				If Not (arcs(i, j) Is Nothing) Then
					If arcs(i, j).exists = True Then
						s = s & nodes(i).name & ",, " & nodes(j).name & ",, " & CShort(arcs(i, j).direction) & ",, " & arcs(i, j).label & vbCrLf
					End If
				End If
			Next j
		Next i
		s = s & "----" & vbCrLf
		spitAsPrettyPersist = s
	End Function
	
	
	Private Property Page_categories() As String Implements _Page.categories
		Get
			Page_categories = innerPage.categories
		End Get
		Set(ByVal Value As String)
			innerPage.categories = Value
		End Set
	End Property
	
	
	Private Property Page_cooked() As String Implements _Page.cooked
		Get
			Page_cooked = innerPage.cooked
		End Get
		Set(ByVal Value As String)
			innerPage.cooked = Value
		End Set
	End Property
	
	
	Private Property Page_createdDate() As Date Implements _Page.createdDate
		Get
			Page_createdDate = innerPage.createdDate
		End Get
		Set(ByVal Value As Date)
			innerPage.createdDate = Value
		End Set
	End Property
	
	
	Private Property Page_lastEdited() As Date Implements _Page.lastEdited
		Get
			Page_lastEdited = innerPage.lastEdited
		End Get
		Set(ByVal Value As Date)
			innerPage.lastEdited = Value
		End Set
	End Property
	
	
	Private Property Page_pageName() As String Implements _Page.pageName
		Get
			Page_pageName = innerPage.pageName
		End Get
		Set(ByVal Value As String)
			innerPage.pageName = Value
		End Set
	End Property
	
	
	Private Property Page_pageType() As String Implements _Page.pageType
		Get
			Page_pageType = "network"
		End Get
		Set(ByVal Value As String)
		End Set
	End Property
	
	
	Private Property Page_prepared() As String Implements _Page.prepared
		Get
			Page_prepared = innerPage.prepared
		End Get
		Set(ByVal Value As String)
			innerPage.prepared = Value
		End Set
	End Property
	
	
	Private Property Page_raw() As String Implements _Page.raw
		Get
			Page_raw = innerPage.raw
		End Get
		Set(ByVal Value As String)
			innerPage.raw = Value
		End Set
	End Property
	
	Private Sub Page_cook(ByRef prep As PagePreparer, ByRef chef As _PageCooker, ByRef backlinks As Boolean) Implements _Page.cook
		Call innerPage.cook(prep, chef, backlinks)
	End Sub
	
	Private Function Page_getDataDictionary() As VCollection Implements _Page.getDataDictionary
		Page_getDataDictionary = innerPage.getDataDictionary
	End Function
	
	Private Function Page_getFirstLine() As String Implements _Page.getFirstLine
		Page_getFirstLine = innerPage.getFirstLine
	End Function
	
	Private Function Page_getMyType() As String Implements _Page.getMyType
		Page_getMyType = innerPage.getMyType
	End Function
	
	Private Function Page_getRedirectPage() As String Implements _Page.getRedirectPage
		Page_getRedirectPage = innerPage.getRedirectPage
	End Function
	
	Private Function Page_getTable() As Table Implements _Page.getTable
		Page_getTable = innerPage.getTable
	End Function
	
	Private Function Page_getVal(ByRef key As String) As String Implements _Page.getVal
		Page_getVal = innerPage.getVal(key)
	End Function
	
	Private Function Page_hasVar(ByRef key As String) As Boolean Implements _Page.hasVar
		Page_hasVar = innerPage.hasVar(key)
	End Function
	
	Private Function Page_isNetwork() As Boolean Implements _Page.isNetwork
		Page_isNetwork = True
	End Function
	
	Private Function Page_isNew() As Boolean Implements _Page.isNew
		Page_isNew = False
	End Function
	
	Private Function Page_isRedirect() As Boolean Implements _Page.isRedirect
		Page_isRedirect = False
	End Function
	
	Private Function Page_isTable() As Boolean Implements _Page.isTable
		Page_isTable = False
	End Function
	
	Private Sub Page_prepare(ByRef prep As PagePreparer, ByRef backlinks As Boolean) Implements _Page.prepare
		Call innerPage.prepare(prep, backlinks)
	End Sub
	
	Private Sub Page_setVal(ByRef aKey As String, ByRef aVal As String) Implements _Page.setVal
		Call innerPage.setVal(aKey, aVal)
	End Sub
	
	Private Function Page_spawnCopy() As _Page Implements _Page.spawnCopy
		Page_spawnCopy = innerPage.spawnCopy()
	End Function
	
	Private Function Page_varsToString() As String Implements _Page.varsToString
		Page_varsToString = innerPage.varsToString
	End Function
	
	Private Function Page_wordCount() As Short Implements _Page.wordCount
		Page_wordCount = innerPage.wordCount
	End Function
End Class