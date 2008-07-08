Option Strict Off
Option Explicit On
Interface _PageCooker
	 Property LinkProcessor As _LinkProcessor
	 Property LinkWrapper As _LinkWrapper
	Function cook(ByRef aPage As _page) As String
	Function cookObject(ByRef aPage As _page) As Object
End Interface
Friend Class PageCooker
	Implements _PageCooker
	
	' turns raw pages into cooked ones
	' ie. processes the raw text of pages to
	' produce nice, HTML formatted one
	' Also does clever things like turning WikiWords into links
	' and http://blah into HTML links etc.
	
	
	Dim LinkProcessor_MemberVariable As LinkProcessor
    Public Property LinkProcessor() As _LinkProcessor Implements _PageCooker.LinkProcessor
        Get
            LinkProcessor = LinkProcessor_MemberVariable
        End Get
        Set(ByVal Value As _LinkProcessor)
            LinkProcessor_MemberVariable = Value
        End Set
    End Property ' to parse links
	Dim LinkWrapper_MemberVariable As LinkWrapper
    Public Property LinkWrapper() As _LinkWrapper Implements _PageCooker.LinkWrapper
        Get
            LinkWrapper = LinkWrapper_MemberVariable
        End Get
        Set(ByVal Value As _LinkWrapper)
            LinkWrapper_MemberVariable = Value
        End Set
    End Property ' to wrap links
	
	' returns the cooked version of the page
	' expects the page's raw and prepared to be filled
	
	Public Function cook(ByRef aPage As _page) As String Implements _PageCooker.cook
	End Function
	
	Public Function cookObject(ByRef aPage As _page) As Object Implements _PageCooker.cookObject
		' this version of cook can return *any* object,
		' however wild the objects become in future
	End Function
End Class