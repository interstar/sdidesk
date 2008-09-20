VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NativeWikiToHtmlDecorator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements WikiToHtml

' Now this is a decorator which delegates to WikiToHtml which
' adds the appropriate link wrapping

' We need to produce an HTML for export to stand-alone
' files. This HTML is *almost* the same, but has some differences.
' Links have to refer to html pages, not to nowhere, with the
' destination in the "id" field.

Private w2h As WikiToHtml
Public sizeChange As String

Public Function qq(s As String) As String
  ' puts string in quotes
  qq = Chr(34) + s + Chr(34)
End Function


Private Sub Class_Initialize()
  Set w2h = New WikiToHtml
  Me.sizeChange = "+0"
End Sub

Private Function WikiToHtml_allHeaders(s As String) As String
    WikiToHtml_allHeaders = w2h.allHeaders(s)
End Function

Private Function WikiToHtml_changeURIs(l As String, linkProtocol As String) As String
    WikiToHtml_changeURIs = w2h.changeURIs(l, linkProtocol)
End Function

Private Function WikiToHtml_getBodyHeader() As String
    WikiToHtml_getBodyHeader = w2h.getBodyHeader()
End Function

Private Function WikiToHtml_getFooter() As String
    WikiToHtml_getFooter = w2h.getFooter()
End Function

Private Function WikiToHtml_getHtmlHead() As String
    WikiToHtml_getHtmlHead = w2h.getHtmlHead()
End Function

Private Sub WikiToHtml_init(ml As ModelLevel)
   Call w2h.init(ml)
End Sub

Private Function WikiToHtml_isImage(url As String) As Boolean
    WikiToHtml_isImage = w2h.isImage(url)
End Function

Private Function WikiToHtml_lineOfTable(l2 As String) As String
    WikiToHtml_lineOfTable = w2h.lineOfTable(l2)
End Function

Private Function WikiToHtml_mainTransform(raw As String, this As WikiToHtml) As String
    WikiToHtml_mainTransform = w2h.mainTransform(raw, Me)
End Function

Private Property Set WikiToHtml_model(ByVal RHS As ModelLevel)
    Set w2h.model = RHS
End Property

Private Property Get WikiToHtml_model() As ModelLevel
    Set WikiToHtml_model = w2h.model
End Property

Private Function WikiToHtml_processBracketContents(s As String, this As WikiToHtml) As String
    WikiToHtml_processBracketContents = w2h.processBracketContents(s, this)
End Function

Private Function WikiToHtml_processFootnotes(raw As String, dummy As Boolean) As Variant
   WikiToHtml_processFootnotes = w2h.processFootnotes(raw, True)
End Function

Private Function WikiToHtml_singleSquare(l As String, linkProtocol As String) As String
    WikiToHtml_singleSquare = w2h.singleSquare(l, linkProtocol)
End Function

Private Function WikiToHtml_squareBrackets(l As String, this As WikiToHtml) As String
    WikiToHtml_squareBrackets = w2h.squareBrackets(l, this)
End Function

Private Property Set WikiToHtml_template(RHS As HtmlTemplate)
    Set w2h.template = RHS
End Property

Private Property Get WikiToHtml_template() As Variant
    Set WikiToHtml_template = w2h.template
End Property

Private Function WikiToHtml_wikiWord(l As String, this As WikiToHtml) As String
    WikiToHtml_wikiWord = w2h.wikiWord(l, this)
End Function

Private Function WikiToHtml_wrapLink(s2 As String, aType As String) As String
   Dim s As String, link As String, colour As String, com As String
   Dim st As New StringTool
     
   s = s2
   com = ""

   colour = w2h.model.linkTypeMan.getColour(aType)
   
   ' test if this is actually a command
   If Left(s, 1) = "#" Then
      com = st.stripRightSpaces(st.leftsa(s, " ", 1)) & "+"
      s = st.stripHead(s, " ", 1)
   End If
   
   s = st.strip(s)
   
   ' Does page exist?
   If w2h.model.pageExists(s) Or Left(com, 1) = "#" Then
     link = "<a href=" + qq("about:blank")
     link = link + " class=" + qq(aType)
     link = link + " id=" + qq(com + s) + "><font size="
     link = link + qq(Me.sizeChange) + " Color = "
     link = link + qq(colour) + ">" + com + s + "</font></a>"
   Else
     link = s + "<a href=" + qq("about:blank")
     link = link + " class=" + qq(aType)
     link = link + " id=" + qq(com + s) + "> <font size=+1 color="
     link = link + qq(colour) + "><b>(?)</b></font></a>"
   End If
  
   WikiToHtml_wrapLink = link
End Function

