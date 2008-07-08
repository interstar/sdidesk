Option Strict Off
Option Explicit On
Friend Class NativePageCooker
	Implements _PageCooker
	
	
	' This class is the concrete implementation of the abstract / interface
	' class : PageCooker
	
	' This is the page cooker which turns the pages into native HTML
	
	Private myWikiToHtml As WikiToHtml ' to do all the HTML work,
	
	Public myLinkProcessor As _LinkProcessor ' to parse links
	Public myLinkWrapper As _LinkWrapper ' to wrap links
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		myWikiToHtml = New WikiToHtml
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object myWikiToHtml may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		myWikiToHtml = Nothing
		'UPGRADE_NOTE: Object myLinkProcessor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		myLinkProcessor = Nothing
		'UPGRADE_NOTE: Object myLinkWrapper may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		myLinkWrapper = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Function asPageCooker() As _PageCooker
		asPageCooker = Me
	End Function
	
	Private Function PageCooker_cook(ByRef aPage As _Page) As String Implements _PageCooker.cook
		PageCooker_cook = myWikiToHtml.mainTransform((aPage.prepared), myLinkProcessor, myLinkWrapper)
	End Function
	
	Private Function PageCooker_cookObject(ByRef aPage As _Page) As Object Implements _PageCooker.cookObject
		Dim s As New StringTool ' dummy object
		PageCooker_cookObject = s
	End Function
	
	
	
	Private Property PageCooker_LinkProcessor() As _LinkProcessor Implements _PageCooker.LinkProcessor
		Get
			PageCooker_LinkProcessor = myLinkProcessor
		End Get
		Set(ByVal Value As _LinkProcessor)
			myLinkProcessor = Value
		End Set
	End Property
	
	
	Private Property PageCooker_LinkWrapper() As _LinkWrapper Implements _PageCooker.LinkWrapper
		Get
			PageCooker_LinkWrapper = myLinkWrapper
		End Get
		Set(ByVal Value As _LinkWrapper)
			myLinkWrapper = Value
		End Set
	End Property
End Class