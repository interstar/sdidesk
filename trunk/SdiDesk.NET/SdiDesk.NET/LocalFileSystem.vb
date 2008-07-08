Option Strict Off
Option Explicit On
Interface _LocalFileSystem
	Function hasLocalFileSystem() As Boolean
    Function makeDirectoryPage(ByRef path As String) As _Page
	Sub changeDirectory(ByRef path As String)
	Function getMainDataDirectory() As String
	Function getExporterDirectory() As String
	Function getDirectoryListingAsVCollection(ByRef d As String) As VCollection
End Interface
Friend Class LocalFileSystem
	Implements _LocalFileSystem
	
	' an interface for the model level to deal with the local file
	' system, eg. make directory page
	
	'still a bit of a mess because ...
	
	' makeDirectoryPage is used for looking at a local directory
	' whereas changeDirectory is really about where the PageStore is going
	' to put things, doesn't make sense if there's a remote PageStore
	' (as I hope there will be one day)
	
	Public Function hasLocalFileSystem() As Boolean Implements _LocalFileSystem.hasLocalFileSystem
	End Function
	
	Public Function makeDirectoryPage(ByRef path As String) As _Page Implements _LocalFileSystem.makeDirectoryPage
	End Function
	
	Public Sub changeDirectory(ByRef path As String) Implements _LocalFileSystem.changeDirectory
	End Sub
	
	Public Function getMainDataDirectory() As String Implements _LocalFileSystem.getMainDataDirectory
	End Function
	
	Public Function getExporterDirectory() As String Implements _LocalFileSystem.getExporterDirectory
		' gets the exporter directory
	End Function
	
	Public Function getDirectoryListingAsVCollection(ByRef d As String) As VCollection Implements _LocalFileSystem.getDirectoryListingAsVCollection
		' returns a vcollection of directory d
	End Function
End Class