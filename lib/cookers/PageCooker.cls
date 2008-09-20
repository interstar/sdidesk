VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageCooker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' turns raw pages into cooked ones
' ie. processes the raw text of pages to
' produce nice, HTML formatted one
' Also does clever things like turning WikiWords into links
' and http://blah into HTML links etc.


Public LinkProcessor As LinkProcessor ' to parse links
Public LinkWrapper As LinkWrapper ' to wrap links

' returns the cooked version of the page
' expects the page's raw and prepared to be filled

Public Function cook(aPage As page) As String
End Function

Public Function cookObject(aPage As page) As Object
' this version of cook can return *any* object,
' however wild the objects become in future
End Function

