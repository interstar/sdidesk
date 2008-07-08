Option Strict Off
Option Explicit On
Interface _Page
	 Property raw As String
	 Property prepared As String
	 Property cooked As String
	 Property pageName As String
	 Property categories As String
	 Property createdDate As Date
	 Property lastEdited As Date
	 Property pageType As String
	Function isNetwork() As Boolean
	Function isTable() As Boolean
	Function isRedirect() As Boolean
	Function isNew() As Boolean
	Function getMyType() As String
	Function getFirstLine() As String
	Function getTable() As Table
	Function getRedirectPage() As String
	Sub prepare(ByRef prep As PagePreparer, ByRef backlinks As Boolean)
	Sub cook(ByRef prep As PagePreparer, ByRef chef As _PageCooker, ByRef backlinks As Boolean)
    Function spawnCopy() As _Page
	Sub setVal(ByRef aKey As String, ByRef aVal As String)
	Function hasVar(ByRef key As String) As Boolean
	Function getVal(ByRef key As String) As String
	Function varsToString() As String
	Function getDataDictionary() As VCollection
	Function wordCount() As Short
End Interface
Friend Class Page
	Implements _Page
	
	' this is the basic page object which holds data about a page
	' As far as possible, EVERYTHING is a page in sdiDesk
	' Page is now an interface, so we can have
	' memory resident, external and remote pages
	
	Dim raw_MemberVariable As String
	Public Property raw() As String Implements _Page.raw
		Get
			raw = raw_MemberVariable
		End Get
		Set(ByVal Value As String)
			raw_MemberVariable = Value
		End Set
	End Property ' the raw text of the page
	Dim prepared_MemberVariable As String
	Public Property prepared() As String Implements _Page.prepared
		Get
			prepared = prepared_MemberVariable
		End Get
		Set(ByVal Value As String)
			prepared_MemberVariable = Value
		End Set
	End Property ' done includes and inlines, but not pretification
	Dim cooked_MemberVariable As String
	Public Property cooked() As String Implements _Page.cooked
		Get
			cooked = cooked_MemberVariable
		End Get
		Set(ByVal Value As String)
			cooked_MemberVariable = Value
		End Set
	End Property ' the presentation view of the page.
	' Could be HTML or something more exotic
	
	Dim pageName_MemberVariable As String
	Public Property pageName() As String Implements _Page.pageName
		Get
			pageName = pageName_MemberVariable
		End Get
		Set(ByVal Value As String)
			pageName_MemberVariable = Value
		End Set
	End Property ' name of the page
	Dim categories_MemberVariable As String
	Public Property categories() As String Implements _Page.categories
		Get
			categories = categories_MemberVariable
		End Get
		Set(ByVal Value As String)
			categories_MemberVariable = Value
		End Set
	End Property ' the categories
	Dim createdDate_MemberVariable As Date
	Public Property createdDate() As Date Implements _Page.createdDate
		Get
			createdDate = createdDate_MemberVariable
		End Get
		Set(ByVal Value As Date)
			createdDate_MemberVariable = Value
		End Set
	End Property ' date this was created
	Dim lastEdited_MemberVariable As Date
	Public Property lastEdited() As Date Implements _Page.lastEdited
		Get
			lastEdited = lastEdited_MemberVariable
		End Get
		Set(ByVal Value As Date)
			lastEdited_MemberVariable = Value
		End Set
	End Property ' date last edited
	
	Dim pageType_MemberVariable As String
	Public Property pageType() As String Implements _Page.pageType
		Get
			pageType = pageType_MemberVariable
		End Get
		Set(ByVal Value As String)
			pageType_MemberVariable = Value
		End Set
	End Property ' the type
	
	' types
	Public Function isNetwork() As Boolean Implements _Page.isNetwork ' is it a network
	End Function
	
	Public Function isTable() As Boolean Implements _Page.isTable ' is it a table
	End Function
	
	Public Function isRedirect() As Boolean Implements _Page.isRedirect ' is it a redirect
	End Function
	
	Public Function isNew() As Boolean Implements _Page.isNew ' is it a new page
	End Function
	
	Public Function getMyType() As String Implements _Page.getMyType ' depends on type
	End Function
	
	Public Function getFirstLine() As String Implements _Page.getFirstLine ' gets first line
	End Function
	
	' transforms
	
	
	
	' if table
	Public Function getTable() As Table Implements _Page.getTable
	End Function
	
	' if redirect
	Public Function getRedirectPage() As String Implements _Page.getRedirectPage
	End Function
	
	Public Sub prepare(ByRef prep As PagePreparer, ByRef backlinks As Boolean) Implements _Page.prepare
	End Sub
	
	Public Sub cook(ByRef prep As PagePreparer, ByRef chef As _PageCooker, ByRef backlinks As Boolean) Implements _Page.cook
	End Sub
	
	Public Function spawnCopy() As _Page Implements _Page.spawnCopy
	End Function
	
	
	' for handling instance variables
	Public Sub setVal(ByRef aKey As String, ByRef aVal As String) Implements _Page.setVal
	End Sub
	
	Public Function hasVar(ByRef key As String) As Boolean Implements _Page.hasVar
	End Function
	
	Public Function getVal(ByRef key As String) As String Implements _Page.getVal
	End Function
	
	Public Function varsToString() As String Implements _Page.varsToString
	End Function
	
	Public Function getDataDictionary() As VCollection Implements _Page.getDataDictionary
	End Function
	
	Public Function wordCount() As Short Implements _Page.wordCount
	End Function
End Class