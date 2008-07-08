Option Strict Off
Option Explicit On
Friend Class Link
	
	' Object that represents a link and all it's potential attributes
	
	Public target As String ' where the link is going
	Public text As String ' what you see
	Public linkType As String ' the type of this link
	'UPGRADE_NOTE: nameSpace was upgraded to nameSpace_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public nameSpace_Renamed As String ' is the link at a remote site, if so namespace
	Public external As Boolean ' is this a link to the outside world?
	Public interMap As Boolean ' is this an intermap link?
	
	' used for WikiMap
	
	Public Sub init(ByRef txt As String, ByRef targ As String, ByRef lTyp As String, ByRef ns As String, ByRef ext As Boolean, ByRef imap As Boolean)
		target = targ
		text = txt
		linkType = lTyp
		nameSpace_Renamed = ns
		external = ext
		interMap = imap
	End Sub
	
	Public Function deepCopy() As Link
		Dim l As New Link
		l.external = external
		l.interMap = interMap
		l.linkType = linkType
		l.nameSpace_Renamed = nameSpace_Renamed
		l.target = target
		l.text = text
		deepCopy = l
	End Function
	
	'UPGRADE_NOTE: toString was upgraded to toString_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function toString_Renamed() As String
		toString_Renamed = "(" & target & ", " & text & ", " & linkType & ", " & nameSpace_Renamed & ", " & external & ")"
	End Function
	
	Public Function isCommand() As Boolean
		If Left(target, 1) = "#" Then
			isCommand = True
		Else
			isCommand = False
		End If
	End Function
	
	Public Function matches(ByRef l2 As Link) As Boolean
		matches = True
		If l2.target <> target Then
			matches = False
			Exit Function
		End If
		
		If l2.text <> text Then
			matches = False
			Exit Function
		End If
		
		If l2.external <> external Then
			matches = False
			Exit Function
		End If
		
		If l2.interMap <> interMap Then
			matches = False
			Exit Function
		End If
		
		If l2.linkType <> linkType Then
			matches = False
			Exit Function
		End If
		
		If l2.nameSpace_Renamed <> nameSpace_Renamed Then
			matches = False
			Exit Function
		End If
		
	End Function
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		target = ""
		text = ""
		linkType = ""
		nameSpace_Renamed = ""
		external = False
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class