VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModelLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Now the ModelLevel is the interface to that part of
' the program that implements the model part (in the MVC sense)
' of SdiDesk

' now (as of March 2005) the model level contains several sub-components

' Currently
' WikiAnnotatedDataStore
' SingleUserState
' SystemConfiguration
' CrawlerSubsystem
' ExportSubsystem
' ControllableModel

' how are we going to refactor this?

' we are going to try to break out the separate modules
' for testing in programs *without* ModelLayer,
' and we are going to add whatever functionality we need
' to test them, to the WikiAnnotatedDataStore etc and we will have
' fake WADS, SysConfs etc. for testing


Public Function getWikiAnnotatedDataStore() As WikiAnnotatedDataStore
End Function

Public Function getSingleUserState() As SingleUserState
End Function

Public Function getSystemConfigurations() As SystemConfigurations
End Function

Public Function getCrawlerSubsystem() As CrawlerSubsystem
End Function

Public Function getExportSubsystem() As ExportSubsystem
End Function

Public Function getControllableModel() As ControllableModel
End Function

Public Function getLocalFileSystem() As LocalFileSystem
End Function

Public Sub setCallBackForm(f As Form)
End Sub

Public Sub setPagePreparer(pp As PagePreparer)
End Sub

Public Sub setPageCooker(pc As PageCooker)
End Sub

Public Sub setForm(f As Form)
End Sub
