VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SysConfStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements SystemConfigurations


Public Function asSystemConfigurations() As SystemConfigurations
    Set asSystemConfigurations = Me
End Function

Private Property Let SystemConfigurations_allPage(ByVal RHS As String)

End Property

Private Property Get SystemConfigurations_allPage() As String
    SystemConfigurations_allPage = "AllPage"
End Property

Private Property Let SystemConfigurations_configPage(ByVal RHS As String)

End Property

Private Property Get SystemConfigurations_configPage() As String
    SystemConfigurations_configPage = "ConfigPage"
End Property

Private Function SystemConfigurations_getTypeColour(typeName As String) As String
    SystemConfigurations_getTypeColour = "#ffeedd"
End Function

Private Property Let SystemConfigurations_helpIndexPage(ByVal RHS As String)

End Property

Private Property Get SystemConfigurations_helpIndexPage() As String
    SystemConfigurations_helpIndexPage = "HelpIndex"
End Property

Private Property Set SystemConfigurations_interMap(ByVal RHS As InterWikiMap)

End Property

Private Property Get SystemConfigurations_interMap() As InterWikiMap
    Set SystemConfigurations_interMap = New InterWikiMap
End Property

Private Property Let SystemConfigurations_recentChangesPage(ByVal RHS As String)

End Property

Private Property Get SystemConfigurations_recentChangesPage() As String
    SystemConfigurations_recentChangesPage = "RecentChanges"
End Property

Private Sub SystemConfigurations_setLinkTypeManager(l As LinkTypeManager)

End Sub

Private Property Let SystemConfigurations_startPage(ByVal RHS As String)
End Property

Private Property Get SystemConfigurations_startPage() As String
    SystemConfigurations_startPage = "StartPage"
End Property
