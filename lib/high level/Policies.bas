Attribute VB_Name = "Policies"
Option Explicit

' this for global functions which are "policies" of the
' program. How to get access to things outside the object

' factory is the root of all "dependency injection"
' though not sure if this is the way it should be done

Private factory As New SdiDeskConfigurationFactory
Private timerText As String

Public Function POLICY_getFactory() As SdiDeskConfigurationFactory
    Set POLICY_getFactory = factory
End Function

Public Sub POLICY_recordEvent(s As String)
    timerText = timerText & Time & " : " & s & vbCrLf
End Sub

