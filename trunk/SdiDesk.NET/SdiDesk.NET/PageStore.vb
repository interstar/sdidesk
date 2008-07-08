Option Strict Off
Option Explicit On
Interface _PageStore
	 Property pictureLocality As String
	Function getPageStoreIdentifier() As String
	Function pageExists(ByRef pageName As String) As Boolean
	Function safeDate(ByRef s As String) As Date
    Function loadRaw(ByRef pageName As String) As _Page
    Function loadUntilNotRedirectRaw(ByRef pageName As String) As _Page
    Function loadOldPage(ByRef pageName As String, ByRef version As Short) As _Page
	Sub savePage(ByRef p As _Page)
	Function deletePage(ByRef pageName As String) As Object
	Function loadMonth(ByRef month As Short, ByRef year As Short) As String
	Sub saveMonth(ByRef month As Short, ByRef year As Short, ByRef body As String)
	Function timeIndexAsWikiFormat(ByRef month As Short, ByRef year As Short, ByRef order As Boolean) As Object
	Function pageContains(ByRef pageName As String, ByRef searchText As String) As Boolean
	Function getPageSetOfAllPagesStartingWith(ByRef s As String) As PageSet
	Function AllPages() As PageSet
	Function getPageSetContaining(ByRef searchText As Object) As PageSet
End Interface
Friend Class PageStore
	Implements _PageStore
	
	' The PageStore is an interface to an object which hides the storage and
	' searching of pages.
	
	' It knows how to save a page somewhere and fetch it
	
	' Also to bring list of all pages or search for pages with some criterion
	
	Dim pictureLocality_MemberVariable As String
	Public Property pictureLocality() As String Implements _PageStore.pictureLocality
		Get
			pictureLocality = pictureLocality_MemberVariable
		End Get
		Set(ByVal Value As String)
			pictureLocality_MemberVariable = Value
		End Set
	End Property ' the protocol and location for pictures
	
	Public Function getPageStoreIdentifier() As String Implements _PageStore.getPageStoreIdentifier
		' string which represents the "address" of this page-store
		' eg. a directory or URL
	End Function
	
	Public Function pageExists(ByRef pageName As String) As Boolean Implements _PageStore.pageExists
	End Function
	
	Public Function safeDate(ByRef s As String) As Date Implements _PageStore.safeDate
	End Function
	
	Public Function loadRaw(ByRef pageName As String) As _Page Implements _PageStore.loadRaw
		' load a page but don't prepare or cook it, hence page is still raw
	End Function
	
	Public Function loadUntilNotRedirectRaw(ByRef pageName As String) As _Page Implements _PageStore.loadUntilNotRedirectRaw
	End Function
	
	Public Function loadOldPage(ByRef pageName As String, ByRef version As Short) As _Page Implements _PageStore.loadOldPage
	End Function
	
	Public Sub savePage(ByRef p As _Page) Implements _PageStore.savePage
	End Sub
	
	Public Function deletePage(ByRef pageName As String) As Object Implements _PageStore.deletePage
	End Function
	
	'UPGRADE_NOTE: year was upgraded to year_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: month was upgraded to month_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function loadMonth(ByRef month_Renamed As Short, ByRef year_Renamed As Short) As String Implements _PageStore.loadMonth
		' loads a page content containing time data
	End Function
	
	'UPGRADE_NOTE: year was upgraded to year_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: month was upgraded to month_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub saveMonth(ByRef month_Renamed As Short, ByRef year_Renamed As Short, ByRef body As String) Implements _PageStore.saveMonth
	End Sub
	
	'UPGRADE_NOTE: year was upgraded to year_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_NOTE: month was upgraded to month_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function timeIndexAsWikiFormat(ByRef month_Renamed As Short, ByRef year_Renamed As Short, ByRef order As Boolean) As Object Implements _PageStore.timeIndexAsWikiFormat
	End Function
	
	Public Function pageContains(ByRef pageName As String, ByRef searchText As String) As Boolean Implements _PageStore.pageContains
	End Function
	
	Public Function getPageSetOfAllPagesStartingWith(ByRef s As String) As PageSet Implements _PageStore.getPageSetOfAllPagesStartingWith
	End Function
	
	Public Function AllPages() As PageSet Implements _PageStore.AllPages
	End Function
	
	Public Function getPageSetContaining(ByRef searchText As Object) As PageSet Implements _PageStore.getPageSetContaining
	End Function
End Class