VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NativePageCooker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements PageCooker

' This class is the concrete implementation of the abstract / interface
' class : PageCooker

' This is the page cooker which turns the pages into native HTML

Private myWikiToHtml As WikiToHtml ' to do all the HTML work,

Public myLinkProcessor As LinkProcessor ' to parse links
Public myLinkWrapper As LinkWrapper ' to wrap links


Private Sub Class_Initialize()
    Set myWikiToHtml = New WikiToHtml
End Sub

Private Sub Class_Terminate()
    Set myWikiToHtml = Nothing
    Set myLinkProcessor = Nothing
    Set myLinkWrapper = Nothing
End Sub

Public Function asPageCooker() As PageCooker
    Set asPageCooker = Me
End Function

Private Function PageCooker_cook(aPage As Page) As String
   PageCooker_cook = myWikiToHtml.mainTransform(aPage.prepared, myLinkProcessor, myLinkWrapper)
End Function

Private Function PageCooker_cookObject(aPage As Page) As Object
    Dim s As New StringTool ' dummy object
    Set PageCooker_cookObject = s
End Function


Private Property Set PageCooker_LinkProcessor(ByVal RHS As LinkProcessor)
    Set myLinkProcessor = RHS
End Property

Private Property Get PageCooker_LinkProcessor() As LinkProcessor
    Set PageCooker_LinkProcessor = myLinkProcessor
End Property

Private Property Set PageCooker_LinkWrapper(ByVal RHS As LinkWrapper)
    Set myLinkWrapper = RHS
End Property

Private Property Get PageCooker_LinkWrapper() As LinkWrapper
    Set PageCooker_LinkWrapper = myLinkWrapper
End Property

