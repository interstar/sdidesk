Option Strict Off
Option Explicit On
Imports mshtml
Friend Class HtmlEvent
	
	'needed when we trap the "clicked on a link" event in the HTML viewer
	
	Private handler As Object ' what handles the event
	Private action As String ' the action to call
	Private id As String
	
	Public Sub Event_Details(ByRef aHandler As Object, ByRef anAction As String, ByRef anId As String)
		handler = aHandler
		action = anAction
		id = anId
	End Sub
	
    'Public Sub HTML_Event()
    '	' HTML EVENT TRIGGERED
    '	' This MUST be the Default Procedure for the Class Module
    '	' The routine will be processed when the event occurs
    '       ' MsgBox ("hi " + id)
    '	Dim a(0) As Object
    '	'UPGRADE_WARNING: Couldn't resolve default property of object a(0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '	a(0) = id
    '	CallByName(handler, action, CallType.Method, a)
    '   End Sub

    Public Function HTML_Event(ByVal e As IHTMLEventObj) As Boolean
        Try
            'Dim a(0) As Object
            'UPGRADE_WARNING: Couldn't resolve default property of object a(0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'a(0) = DirectCast(id, Object)
            'CallByName(handler, action, CallType.Method, a)
            CallByName(handler, action, CallType.Method, id)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Function
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		On Error Resume Next
		'UPGRADE_NOTE: Object handler may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		handler = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class