Option Strict Off
Option Explicit On
Interface _LinkProcessor
	Function wrapAllLinks(ByRef l As String, ByRef wrapper As _LinkWrapper) As String
	Function getAllLinks(ByRef l As String) As OCollection
	Function getAllLinksInBigDocument(ByRef doc As String) As OCollection
End Interface
Friend Class LinkProcessor
	Implements _LinkProcessor
	
	' This interface analyses a line to extract links
	' and wraps them all with the LinkWrapper
	
	' these operate on the level of the individual line
	Public Function wrapAllLinks(ByRef l As String, ByRef wrapper As _LinkWrapper) As String Implements _LinkProcessor.wrapAllLinks
	End Function
	
	Public Function getAllLinks(ByRef l As String) As OCollection Implements _LinkProcessor.getAllLinks
	End Function
	
	
	' but sometimes the processor should be able to tackle a whole
	' document. We'll leave it up to the processer how it does it
	Public Function getAllLinksInBigDocument(ByRef doc As String) As OCollection Implements _LinkProcessor.getAllLinksInBigDocument
	End Function
End Class