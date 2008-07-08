Option Strict Off
Option Explicit On

Public Enum PageEditState
    LoadedState
    RawState
    EditedState
    PreviewState
    SavedState
End Enum

Interface _SingleUserState
	 Property currentPageName As String
	 Property oldPageName As String
	 Property currentPage As _Page
	 Property backlinks As Boolean
	 Property isLoading As Boolean
	 Property changesSaved As Boolean
	 Property editState As PageEditState
	 Property history As NavigationHistory
End Interface
Friend Class SingleUserState
	Implements _SingleUserState
	
	' This interface class represents that part of the ModelLevel
	' that has to keep track of the interaction state with a user
	
	' Includes access to a NavigationHistory, CurrentPage etc.
	
	' implemented by ModelLevel
	
	
	
	Dim currentPageName_MemberVariable As String
	Public Property currentPageName() As String Implements _SingleUserState.currentPageName
		Get
			currentPageName = currentPageName_MemberVariable
		End Get
		Set(ByVal Value As String)
			currentPageName_MemberVariable = Value
		End Set
	End Property ' the name of the current page
	Dim oldPageName_MemberVariable As String
	Public Property oldPageName() As String Implements _SingleUserState.oldPageName
		Get
			oldPageName = oldPageName_MemberVariable
		End Get
		Set(ByVal Value As String)
			oldPageName_MemberVariable = Value
		End Set
	End Property ' the name of the previous page
	
	Dim currentPage_MemberVariable As Page
    Public Property currentPage() As _Page Implements _SingleUserState.currentPage
        Get
            currentPage = currentPage_MemberVariable
        End Get
        Set(ByVal Value As _Page)
            currentPage_MemberVariable = Value
        End Set
    End Property ' current page
	
	Dim backlinks_MemberVariable As Boolean
	Public Property backlinks() As Boolean Implements _SingleUserState.backlinks
		Get
			backlinks = backlinks_MemberVariable
		End Get
		Set(ByVal Value As Boolean)
			backlinks_MemberVariable = Value
		End Set
	End Property ' do we automatically show backlinks?
	
	Dim isLoading_MemberVariable As Boolean
	Public Property isLoading() As Boolean Implements _SingleUserState.isLoading
		Get
			isLoading = isLoading_MemberVariable
		End Get
		Set(ByVal Value As Boolean)
			isLoading_MemberVariable = Value
		End Set
	End Property ' is the page loading (so ignore onChange)
	Dim changesSaved_MemberVariable As Boolean
	Public Property changesSaved() As Boolean Implements _SingleUserState.changesSaved
		Get
			changesSaved = changesSaved_MemberVariable
		End Get
		Set(ByVal Value As Boolean)
			changesSaved_MemberVariable = Value
		End Set
	End Property ' record if the changes were saved
	Dim editState_MemberVariable As PageEditState
	Public Property editState() As PageEditState Implements _SingleUserState.editState
		Get
			editState = editState_MemberVariable
		End Get
		Set(ByVal Value As PageEditState)
			editState_MemberVariable = Value
		End Set
	End Property ' the state of editing of this page
	
	Dim history_MemberVariable As NavigationHistory
	Public Property history() As NavigationHistory Implements _SingleUserState.history
		Get
			history = history_MemberVariable
		End Get
		Set(ByVal Value As NavigationHistory)
			history_MemberVariable = Value
		End Set
	End Property ' user's nav history
End Class