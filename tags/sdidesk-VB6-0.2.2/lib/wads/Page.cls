VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Page"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' this is the basic page object which holds data about a page
' As far as possible, EVERYTHING is a page in sdiDesk
' Page is now an interface, so we can have
' memory resident, external and remote pages
 
Public raw As String ' the raw text of the page
Public prepared As String ' done includes and inlines, but not pretification
Public cooked As String ' the presentation view of the page.
' Could be HTML or something more exotic

Public pageName As String ' name of the page
Public categories As String ' the categories
Public createdDate As Date ' date this was created
Public lastEdited As Date ' date last edited

Public pageType As String ' the type

' types
Public Function isNetwork() As Boolean  ' is it a network
End Function

Public Function isTable() As Boolean ' is it a table
End Function

Public Function isRedirect() As Boolean ' is it a redirect
End Function

Public Function isNew() As Boolean ' is it a new page
End Function

Public Function getMyType() As String ' depends on type
End Function

Public Function getFirstLine() As String ' gets first line
End Function

' transforms



' if table
Public Function getTable() As table
End Function

' if redirect
Public Function getRedirectPage() As String
End Function

Public Sub prepare(prep As PagePreparer, backlinks As Boolean)
End Sub

Public Sub cook(prep As PagePreparer, chef As PageCooker, backlinks As Boolean)
End Sub

Public Function spawnCopy() As Page
End Function


' for handling instance variables
Public Sub setVal(aKey As String, aVal As String)
End Sub

Public Function hasVar(key As String) As Boolean
End Function

Public Function getVal(key As String) As String
End Function

Public Function varsToString() As String
End Function

Public Function getDataDictionary() As VCollection
End Function

Public Function wordCount() As Integer
End Function

