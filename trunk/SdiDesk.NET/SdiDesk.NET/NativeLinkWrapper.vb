Option Strict Off
Option Explicit On
Friend Class NativeLinkWrapper
	Implements _LinkWrapper
	
	
	Private myWads As _WikiAnnotatedDataStore
	Private mySysConf As _SystemConfigurations
	Private myMap As InterWikiMap
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object myWads may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		myWads = Nothing
		'UPGRADE_NOTE: Object mySysConf may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mySysConf = Nothing
		'UPGRADE_NOTE: Object myMap may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		myMap = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Function asLinkWrapper() As _LinkWrapper
		asLinkWrapper = Me
	End Function
	
	
	Private Property LinkWrapper_remoteInterMap() As InterWikiMap Implements _LinkWrapper.remoteInterMap
		Get
			LinkWrapper_remoteInterMap = myMap
		End Get
		Set(ByVal Value As InterWikiMap)
			myMap = Value
		End Set
	End Property
	
	
	Private Property LinkWrapper_remoteSysConf() As _SystemConfigurations Implements _LinkWrapper.remoteSysConf
		Get
			LinkWrapper_remoteSysConf = mySysConf
		End Get
		Set(ByVal Value As _SystemConfigurations)
			mySysConf = Value
		End Set
	End Property
	
	
	Private Property LinkWrapper_remoteWads() As _WikiAnnotatedDataStore Implements _LinkWrapper.remoteWads
		Get
			LinkWrapper_remoteWads = myWads
		End Get
		Set(ByVal Value As _WikiAnnotatedDataStore)
			myWads = Value
		End Set
	End Property
	
	Private Function LinkWrapper_wrap(ByRef l As Link) As String Implements _LinkWrapper.wrap
		Dim s, iMapUrl As String
		
		If l.interMap = True Then
			'MsgBox (l.nameSpace)
			'MsgBox (l.toString)
			If myMap.getUrl((l.nameSpace_Renamed)) <> "ERROR" Then
				'MsgBox (l.target)
				'MsgBox (l.nameSpace)
				l.target = myMap.getUrl((l.nameSpace_Renamed)) & l.target
				LinkWrapper_wrap = "<a href='" & l.target & "'>" & l.text & "</a>"
				Exit Function
			Else
				'MsgBox (l.nameSpace)
				LinkWrapper_wrap = "<font color=''>Warning. '" & l.nameSpace_Renamed & "' not defined</font>                    "
				Exit Function
			End If
		End If
		
		If l.external = True Then
			s = "<a href='" & l.target & "' id='external' target='new'>" & l.text & "</a>"
		Else
			If myWads.pageExists((l.target)) Or Left(l.target, 1) = "#" Then
				s = "<a href='about:blank' class='" & l.linkType & "' id='" & l.target & "'><font color='" & mySysConf.getTypeColour((l.linkType)) & "'>" & l.text & "</font></a>"
			Else
				s = "<a href='about:blank' class='newPage' id='" & l.target & "'><font color='#ff6666'>" & l.text & "</font></a>"
			End If
		End If
		LinkWrapper_wrap = s
	End Function
End Class