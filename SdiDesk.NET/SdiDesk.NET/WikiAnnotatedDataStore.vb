Option Strict Off
Option Explicit On
Interface _WikiAnnotatedDataStore
	 Property store As _PageStore
	Function pageExists(ByRef pName As String) As Boolean
	Function getPageVar(ByRef pageName As String, ByRef varName As String) As String
	Function getPageSetContaining(ByRef s As String) As PageSet
	Function getRawPageData(ByRef pName As String) As String
End Interface
Friend Class WikiAnnotatedDataStore
	Implements _WikiAnnotatedDataStore
	
	' Conceptually the WADS encapsulates all the responsibilities
	' for managing a PageStore, producing native and export versions of
	' wiki pages
	
	' this is an interface which encapsulates the management
	' of a PageStore, processing pages for native display
	' or export etc.
	
	' we are refactoring to this interface gradually
	
	Dim store_MemberVariable As PageStore
    Public Property store() As _PageStore Implements _WikiAnnotatedDataStore.store
        Get
            store = store_MemberVariable
        End Get
        Set(ByVal Value As _PageStore)
            store_MemberVariable = Value
        End Set
    End Property
	
	Public Function pageExists(ByRef pName As String) As Boolean Implements _WikiAnnotatedDataStore.pageExists
		
	End Function
	
	Public Function getPageVar(ByRef pageName As String, ByRef varName As String) As String Implements _WikiAnnotatedDataStore.getPageVar
		
	End Function
	
	Public Function getPageSetContaining(ByRef s As String) As PageSet Implements _WikiAnnotatedDataStore.getPageSetContaining
		
	End Function
	
	Public Function getRawPageData(ByRef pName As String) As String Implements _WikiAnnotatedDataStore.getRawPageData
		
	End Function
End Class