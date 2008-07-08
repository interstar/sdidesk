Option Strict Off
Option Explicit On
Module Policies
	
	' this for global functions which are "policies" of the
	' program. How to get access to things outside the object
	
	' factory is the root of all "dependency injection"
	' though not sure if this is the way it should be done
	
	Private factory As New SdiDeskConfigurationFactory
    Private timerText As String
    Public Sub CreateFactoryObj()
        Try
            factory = New SdiDeskConfigurationFactory

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
	
	Public Function POLICY_getFactory() As SdiDeskConfigurationFactory
		POLICY_getFactory = factory
	End Function
	
	Public Sub POLICY_recordEvent(ByRef s As String)
		timerText = timerText & TimeOfDay & " : " & s & vbCrLf
	End Sub
End Module