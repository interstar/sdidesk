Option Strict Off
Option Explicit On
Interface _SystemConfigurations
	 Property startPage As String
	 Property configPage As String
	 Property helpIndexPage As String
	 Property allPage As String
	 Property recentChangesPage As String
	 Property interMap As InterWikiMap
	Sub setLinkTypeManager(ByRef l As LinkTypeManager)
	Function getTypeColour(ByRef typeName_Renamed As String) As String
End Interface
Friend Class SystemConfigurations
	Implements _SystemConfigurations
	
	' This interface, implemented by the model-level
	' is what takes care of the configuration of this
	' copy of SdiDesk
	
	' eg. linkTypes, intermap etc.
	
	
	Dim startPage_MemberVariable As String
	Public Property startPage() As String Implements _SystemConfigurations.startPage
		Get
			startPage = startPage_MemberVariable
		End Get
		Set(ByVal Value As String)
			startPage_MemberVariable = Value
		End Set
	End Property ' page to start on
	Dim configPage_MemberVariable As String
	Public Property configPage() As String Implements _SystemConfigurations.configPage
		Get
			configPage = configPage_MemberVariable
		End Get
		Set(ByVal Value As String)
			configPage_MemberVariable = Value
		End Set
	End Property ' page for configs
	Dim helpIndexPage_MemberVariable As String
	Public Property helpIndexPage() As String Implements _SystemConfigurations.helpIndexPage
		Get
			helpIndexPage = helpIndexPage_MemberVariable
		End Get
		Set(ByVal Value As String)
			helpIndexPage_MemberVariable = Value
		End Set
	End Property ' help index
	Dim allPage_MemberVariable As String
	Public Property allPage() As String Implements _SystemConfigurations.allPage
		Get
			allPage = allPage_MemberVariable
		End Get
		Set(ByVal Value As String)
			allPage_MemberVariable = Value
		End Set
	End Property ' where all pages are
	Dim recentChangesPage_MemberVariable As String
	Public Property recentChangesPage() As String Implements _SystemConfigurations.recentChangesPage
		Get
			recentChangesPage = recentChangesPage_MemberVariable
		End Get
		Set(ByVal Value As String)
			recentChangesPage_MemberVariable = Value
		End Set
	End Property ' where recent changes are listed
	
	Dim interMap_MemberVariable As InterWikiMap
	Public Property interMap() As InterWikiMap Implements _SystemConfigurations.interMap
		Get
			interMap = interMap_MemberVariable
		End Get
		Set(ByVal Value As InterWikiMap)
			interMap_MemberVariable = Value
		End Set
	End Property
	
	' typed links
	Public Sub setLinkTypeManager(ByRef l As LinkTypeManager) Implements _SystemConfigurations.setLinkTypeManager
	End Sub
	
	'UPGRADE_NOTE: typeName was upgraded to typeName_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function getTypeColour(ByRef typeName_Renamed As String) As String Implements _SystemConfigurations.getTypeColour
	End Function
End Class