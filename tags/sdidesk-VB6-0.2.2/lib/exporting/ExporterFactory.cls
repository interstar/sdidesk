VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExporterFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Interface for all things that should be given to exporters

' This is an Abstract Factory, right?

' Makes all the things necessary for an exporter
' The user can assume all parts are wired up correctly

Public Sub init(wads As WikiAnnotatedDataStore, sysConf As SystemConfigurations, iMap As interWikiMap)
' sets these objects which will be injected into the various objects
' made by this factory
End Sub

Public Function getLinkWrapper() As LinkWrapper
End Function

Public Function getLinkProcessor() As LinkProcessor
End Function

Public Function getPageCooker() As PageCooker
End Function

Public Function getPagePreparer() As PagePreparer
End Function

Public Function getExporter() As Exporter
End Function
