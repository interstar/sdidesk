Option Strict Off
Option Explicit On
Interface _ExportSubsystem
	 Property pageStoreIdentifier As String
	Sub refreshExportManager(ByRef wads As _WikiAnnotatedDataStore)
    Function makeExportsPage() As _Page
    Function makeExportersPage() As _Page
    Function makeChooseExporterPage(ByRef currentPageName As String) As _Page
	Sub scanForExports()
	Sub doExport(ByRef name As String)
	Sub doInstantExport(ByRef exporterName As String, ByRef pageName As String)
End Interface
Friend Class ExportSubsystem
	Implements _ExportSubsystem
	
	' interface for the export subsystem
	Dim pageStoreIdentifier_MemberVariable As String
	Public Property pageStoreIdentifier() As String Implements _ExportSubsystem.pageStoreIdentifier
		Get
			pageStoreIdentifier = pageStoreIdentifier_MemberVariable
		End Get
		Set(ByVal Value As String)
			pageStoreIdentifier_MemberVariable = Value
		End Set
	End Property
	' this is the string the ExportManager will pass to any
	' export programs so they can find the PageStore
	' Currently the main data directory
	' though may later be a URL
	
	Public Sub refreshExportManager(ByRef wads As _WikiAnnotatedDataStore) Implements _ExportSubsystem.refreshExportManager
		' reload the details from the PageStore
		' if the definitions of exports have been updated.
	End Sub
	
	Public Function makeExportsPage() As _Page Implements _ExportSubsystem.makeExportsPage
		' create a page that's a list of currently available exports
	End Function
	
	Public Function makeExportersPage() As _Page Implements _ExportSubsystem.makeExportersPage
		' create a page that details the export programs available
	End Function
	
	Public Function makeChooseExporterPage(ByRef currentPageName As String) As _Page Implements _ExportSubsystem.makeChooseExporterPage
		' when the user wants to export the current page, a list of
		' exporter programs to choose from
	End Function
	
	Public Sub scanForExports() Implements _ExportSubsystem.scanForExports
		' scans the local drive for export plug-ins
	End Sub
	
	Public Sub doExport(ByRef name As String) Implements _ExportSubsystem.doExport
		' fire off the export
	End Sub
	
	Public Sub doInstantExport(ByRef exporterName As String, ByRef pageName As String) Implements _ExportSubsystem.doInstantExport
		' fire off an instant export
	End Sub
End Class