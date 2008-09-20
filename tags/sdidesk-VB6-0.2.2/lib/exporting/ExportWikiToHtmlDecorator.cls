VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExportWikiToHtmlDecorator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements WikiToHtml

' This is an alternative decorator to NativeWikiToHtmlDecorator.

' It wraps links to look for flat HTML files in a certain directory

Private w2h As WikiToHtml ' we're going to delegate most of the work to this

Public path As String
Public exportedPages As PageSet

Public Sub init(aPath As String, ps As PageSet, md As ModelLevel)
  path = aPath ' set up a path for the files to be put into into
  Set exportedPages = ps ' we need a pageset of all pages we're
  ' exporting, so we know which links are real or not
  Set w2h.model = md
End Sub

Public Function qq(s As String) As String
  ' quick quote
  qq = Chr(34) + s + Chr(34)
End Function


Public Function pageTitle(pt As String) As String
   If pageTitle <> "" Then
     pageTitle = "<h2>" + pt + "</h2>" + vbCrLf
   Else
     pageTitle = ""
   End If
End Function


Public Function getFooter(parent As WikiToHtml) As String
  getFooter = parent.template.footer + vbCrLf + "</body></html>"
End Function

Private Sub Class_Initialize()
  Set w2h = New WikiToHtml
End Sub

Private Sub Class_Terminate()
  Set w2h = Nothing
End Sub

Private Function WikiToHtml_allHeaders(s As String) As String
    WikiToHtml_allHeaders = w2h.allHeaders(s)
End Function

Private Function WikiToHtml_changeURIs(l As String, linkProtocol As String) As String
    WikiToHtml_changeURIs = w2h.changeURIs(l, linkProtocol)
End Function

Private Function WikiToHtml_getBodyHeader() As String
    WikiToHtml_getBodyHeader = w2h.template.header
End Function

Private Function WikiToHtml_getFooter() As String
    WikiToHtml_getFooter = w2h.template.footer
End Function

Private Function WikiToHtml_getHtmlHead() As String
   WikiToHtml_getHtmlHead = "<html>" + vbCrLf + "<head>" + vbCrLf + w2h.template.styleSheet + "</head>" + vbCrLf
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
   WikiToHtml_processFootnotes = w2h.processFootnotes(raw, False)
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

Private Function WikiToHtml_wrapLink(s2 As String, linkType As String) As String
    Dim s As String
    Dim link As String, url As String, sizeChange As String, colour As String
    
    s = s2
    sizeChange = "+1"
    colour = w2h.model.linkTypeMan.getColour(linkType)
    url = ".\" + s + ".html"
   
    If exportedPages.hasPage(s) Then
        link = "<a href=" + qq(url)
        link = link + " class=" + qq(linkType)
        link = link + " id=" + qq(s) + "> <font size="
        link = link + qq(sizeChange) + " Color = "
        link = link + qq(colour) + ">" + s + "</font></a>"
   Else
        ' nothing to link to
        link = s
   End If

   WikiToHtml_wrapLink = link
End Function


