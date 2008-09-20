VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LinkWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Interface for LinkWrappers

Public remoteWads As WikiAnnotatedDataStore
Public remoteSysConf As SystemConfigurations
Public remoteInterMap As InterWikiMap

Public Function wrap(l As Link) As String

End Function





