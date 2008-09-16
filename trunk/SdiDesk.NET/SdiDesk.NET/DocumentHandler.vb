Option Strict Off
Option Explicit On
Imports mshtml

Friend Class DocumentHandler
	
	' this class handles the HTML document in the document viewer
	' it collects the HTML links and accepts the
	
	Public HtmlView As System.Windows.Forms.WebBrowser
    Private links As HtmlElementCollection  'Object '  these will be the link elements of the document
	Public noLinks As Short
    Private linkObjects() As Object ' HtmlElementCollection
	
    Public Sub recalc()
        Dim doc As IHTMLDocument2 = DirectCast(HtmlView.Document.DomDocument, IHTMLDocument2)
        Dim allLinks As IHTMLElementCollection = DirectCast(doc.links, IHTMLElementCollection)
        For Each link As HTMLAnchorElement In allLinks
            Dim anEvent As New HtmlEvent 
            anEvent.Event_Details(Me, "HTML_Click", link.id)
            AddHandler DirectCast(link, HTMLAnchorEvents2_Event).onclick, AddressOf anEvent.HTML_Event
        Next 
    End Sub 
    Public Sub HTML_Click(ByVal a As String)
        Call WADSMainForm.waitPageLoad()
        If a = "external" Then
        Else
            Call WADSMainForm.controller.processCommand(a, False)
        End If
    End Sub
End Class