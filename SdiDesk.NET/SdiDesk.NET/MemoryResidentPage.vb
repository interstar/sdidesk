Option Strict Off
Option Explicit On
Friend Class MemoryResidentPage
	Implements _Page
	
	
	' this is the basic page object which holds data about a page
	' As far as possible, EVERYTHING is a page in sdiDesk
	
	' implement properties
	Private myRaw As String ' the raw text of the page
	Private myPrepared As String ' done includes and inlines, but not prettification
	Private myCooked As String ' the presentation view of the page.
	
	Private myPageName As String ' name of the page
	Private myCategories As String ' the categories box
	Private myCreatedDate As Date ' date this was created
	Private myLastEdited As Date ' date last edited
	
	' instance vars for this page
	Private dataDictionary As VCollection
	
	' alternative versions
	Private myTable As Table
	
	Private myType As String
	
	' helpers
	Private st As StringTool
	
	Public Function asPage() As _Page
		asPage = Me
	End Function
	
	Public Function getFirstLine() As String
		Dim lines() As String
		If InStr(myRaw, vbCrLf) Then
			lines = Split(myRaw, vbCrLf)
			getFirstLine = lines(0)
		Else
			getFirstLine = myRaw
		End If
	End Function
	
	Public Sub trimSpacesFromEnd()
		Dim flag As Boolean
		flag = False
		While flag = False
			If Len(myRaw) > 1 Then
				If Right(myRaw, 1) = vbCrLf Or Right(myRaw, 1) = " " Or Asc(Right(myRaw, 1)) = 10 Then
					myRaw = Left(myRaw, Len(myRaw) - 2)
				Else
					flag = True
				End If
			Else
				flag = True
			End If
		End While
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		dataDictionary = New VCollection
		st = New StringTool
		myType = "new page"
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object dataDictionary may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		dataDictionary = Nothing
		'UPGRADE_NOTE: Object st may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		st = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	Private Property Page_categories() As String Implements _Page.categories
		Get
			Page_categories = myCategories
		End Get
		Set(ByVal Value As String)
			myCategories = Value
		End Set
	End Property
	
	
	Private Property Page_cooked() As String Implements _Page.cooked
		Get
			Page_cooked = myCooked
		End Get
		Set(ByVal Value As String)
			myCooked = Value
		End Set
	End Property
	
	
	Private Property Page_createdDate() As Date Implements _Page.createdDate
		Get
			Page_createdDate = myCreatedDate
		End Get
		Set(ByVal Value As Date)
			myCreatedDate = Value
		End Set
	End Property
	
	
	Private Property Page_lastEdited() As Date Implements _Page.lastEdited
		Get
			Page_lastEdited = myLastEdited
		End Get
		Set(ByVal Value As Date)
			myLastEdited = Value
		End Set
	End Property
	
	
	Private Property Page_pageName() As String Implements _Page.pageName
		Get
			Page_pageName = myPageName
		End Get
		Set(ByVal Value As String)
			myPageName = Value
		End Set
	End Property
	
	
	Private Property Page_pageType() As String Implements _Page.pageType
		Get
			Page_pageType = myType
		End Get
		Set(ByVal Value As String)
			myType = Value
		End Set
	End Property
	
	
	Private Property Page_prepared() As String Implements _Page.prepared
		Get
			Page_prepared = myPrepared
		End Get
		Set(ByVal Value As String)
			myPrepared = Value
		End Set
	End Property
	
	
	Private Property Page_raw() As String Implements _Page.raw
		Get
			Page_raw = myRaw
		End Get
		Set(ByVal Value As String)
			myRaw = Value
		End Set
	End Property
	
	Private Sub Page_cook(ByRef prep As PagePreparer, ByRef chef As _PageCooker, ByRef backlinks As Boolean) Implements _Page.cook
		Dim t As New Table
		If t.isValidTable(myRaw) And InStr(myRaw, "____") Then
			myType = "table"
		Else
			myType = "normal"
		End If
		'UPGRADE_NOTE: Object t may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		t = Nothing
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Me. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call prep.prepare(Me, backlinks)
		'UPGRADE_WARNING: Couldn't resolve default property of object Me. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		myCooked = chef.cook(Me)
	End Sub
	
	Private Function Page_getDataDictionary() As VCollection Implements _Page.getDataDictionary
		Page_getDataDictionary = dataDictionary
	End Function
	
	Private Function Page_getFirstLine() As String Implements _Page.getFirstLine
		Page_getFirstLine = getFirstLine()
	End Function
	
	Private Function Page_getMyType() As String Implements _Page.getMyType
		Dim t As String
		t = "normal"
		If Page_isNetwork Then t = "network"
		If Page_isTable Then t = "table"
		If Page_isRedirect Then t = "redirect"
		If Page_isNew Then t = "new page"
		Page_getMyType = t
	End Function
	
	
	
	Private Function Page_getRedirectPage() As String Implements _Page.getRedirectPage
		Dim s As String
		s = getFirstLine()
		Dim parts() As String
		parts = Split(s, " ")
		Page_getRedirectPage = parts(1)
	End Function
	
	Private Function Page_getTable() As Table Implements _Page.getTable
		If Page_isTable Then
			Page_getTable = myTable
		Else
			MsgBox("Error trying to get table")
			End
		End If
	End Function
	
	Private Function Page_getVal(ByRef key As String) As String Implements _Page.getVal
		Dim k As String
		k = st.strip(key)
		If Page_hasVar(k) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object dataDictionary.Item(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Page_getVal = dataDictionary.Item(k)
		Else
			Page_getVal = "<font color='#990000'> undefined variable : " & k & " on page " & myPageName & "</font>"
		End If
	End Function
	
	Private Function Page_hasVar(ByRef key As String) As Boolean Implements _Page.hasVar
		Dim k As String
		k = st.strip(key)
		Page_hasVar = dataDictionary.hasKey(k)
	End Function
	
	Private Function Page_isNetwork() As Boolean Implements _Page.isNetwork
		Dim f As String
		f = getFirstLine()
		If Len(f) > 8 And Left(f, 8) = "#Network" Then
			Page_isNetwork = True
		Else
			Page_isNetwork = False
		End If
	End Function
	
	Private Function Page_isNew() As Boolean Implements _Page.isNew
		If myType = "new page" Then
			Page_isNew = True
		Else
			Page_isNew = False
		End If
	End Function
	
	Private Function Page_isRedirect() As Boolean Implements _Page.isRedirect
		Dim f As String
		f = getFirstLine()
		If InStr(f, "#REDIRECT ") > 0 Then
			Page_isRedirect = True
		Else
			Page_isRedirect = False
		End If
	End Function
	
	Private Function Page_isTable() As Boolean Implements _Page.isTable
		If myType = "table" Then
			Page_isTable = True
		Else
			Page_isTable = False
		End If
	End Function
	
	Private Sub Page_prepare(ByRef prep As PagePreparer, ByRef backlinks As Boolean) Implements _Page.prepare
		Dim t As New Table
		If Not Page_isNetwork() Then
			If t.isValidTable(myRaw) And InStr(myRaw, "____") Then
				myType = "table"
			Else
				myType = "normal"
			End If
			'UPGRADE_NOTE: Object t may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			t = Nothing
			
			'UPGRADE_WARNING: Couldn't resolve default property of object Me. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call prep.prepare(Me, backlinks)
		End If
	End Sub
	
	
	Private Sub Page_setVal(ByRef aKey As String, ByRef aVal As String) Implements _Page.setVal
		Dim v, k As String
		v = st.strip(aVal)
		k = st.strip(aKey)
		If dataDictionary.hasKey(k) Then
			Call dataDictionary.Remove(k)
		End If
		
		Call dataDictionary.Add(v, k)
	End Sub
	
	Private Function Page_spawnCopy() As _Page Implements _Page.spawnCopy
		' creates a new object of class MemoryResidentPage, and populates it with copies
		' of all data except cooked and network
		Dim p2 As _Page
		p2 = POLICY_getFactory().getNewPageInstance()
		p2.raw = myRaw
		p2.pageName = myPageName
		p2.categories = myCategories
		p2.createdDate = myCreatedDate
		p2.lastEdited = myLastEdited
		Page_spawnCopy = p2
		'UPGRADE_NOTE: Object p2 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		p2 = Nothing
	End Function
	
	Private Function Page_varsToString() As String Implements _Page.varsToString
		Page_varsToString = dataDictionary.toString_Renamed()
	End Function
	
	Private Function Page_wordCount() As Short Implements _Page.wordCount
		' nb : a rough word-count
		Dim lines() As String
		Dim words() As String
		Dim v, w As Object
		Dim s As String
		Dim wmg As New WikiMarkupGopher
		Dim Count As Short
		Count = 0
		lines = Split(myRaw, vbCrLf)
		For	Each w In lines
			'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			words = Split(CStr(w), " ")
			For	Each v In words
				'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				s = CStr(v)
				If s <> "" And wmg.isAlpha(Left(s, 1)) Then
					Count = Count + 1
				End If
			Next v
		Next w
		Page_wordCount = Count
		'UPGRADE_NOTE: Object wmg may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		wmg = Nothing
	End Function
End Class