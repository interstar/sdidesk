Option Strict Off
Option Explicit On
Interface _LinkWrapper
	 Property remoteWads As _WikiAnnotatedDataStore
	 Property remoteSysConf As _SystemConfigurations
	 Property remoteInterMap As InterWikiMap
	Function wrap(ByRef l As Link) As String
End Interface
Friend Class LinkWrapper
	Implements _LinkWrapper
	' Interface for LinkWrappers
	
	Dim remoteWads_MemberVariable As WikiAnnotatedDataStore
    Public Property remoteWads() As _WikiAnnotatedDataStore Implements _LinkWrapper.remoteWads
        Get
            remoteWads = remoteWads_MemberVariable
        End Get
        Set(ByVal Value As _WikiAnnotatedDataStore)
            remoteWads_MemberVariable = Value
        End Set
    End Property
	Dim remoteSysConf_MemberVariable As SystemConfigurations
    Public Property remoteSysConf() As _SystemConfigurations Implements _LinkWrapper.remoteSysConf
        Get
            remoteSysConf = remoteSysConf_MemberVariable
        End Get
        Set(ByVal Value As _SystemConfigurations)
            remoteSysConf_MemberVariable = Value
        End Set
    End Property
	Dim remoteInterMap_MemberVariable As InterWikiMap
	Public Property remoteInterMap() As InterWikiMap Implements _LinkWrapper.remoteInterMap
		Get
			remoteInterMap = remoteInterMap_MemberVariable
		End Get
		Set(ByVal Value As InterWikiMap)
			remoteInterMap_MemberVariable = Value
		End Set
	End Property
	
	Public Function wrap(ByRef l As Link) As String Implements _LinkWrapper.wrap
		
	End Function
End Class