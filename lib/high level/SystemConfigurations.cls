VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemConfigurations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This interface, implemented by the model-level
' is what takes care of the configuration of this
' copy of SdiDesk

' eg. linkTypes, intermap etc.


Public startPage As String ' page to start on
Public configPage As String ' page for configs
Public helpIndexPage As String ' help index
Public allPage As String ' where all pages are
Public recentChangesPage As String ' where recent changes are listed

Public interMap As InterWikiMap

' typed links
Public Sub setLinkTypeManager(l As LinkTypeManager)
End Sub

Public Function getTypeColour(typeName As String) As String
End Function



