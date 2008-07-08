Option Strict Off
Option Explicit On
Friend Class NavCommand
	
	' don't entirely know what this is for yet,
	' but intuitively it's going to help me sort out my
	' problems with history and stuff
	
	' a command can be one of several things
	' 1) PageName
	' 2) #raw PageName
	' 3) #history PageName
	' 4) #delete PageName etc.
	
	
	Public full As String ' the full text of the command
	'UPGRADE_NOTE: command was upgraded to command_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public command_Renamed As String ' the command word
	Public pageName As String ' the argument
	Public tail As String ' everything after the command
	Private parts() As String ' all parts
	Private st As StringTool ' useful
	
	Public Sub init(ByRef s As String)
		st = New StringTool
		full = s
		
		If Left(full, 1) = "#" Then
			' it's a command
			' this allows us to use + instead of space as separator, when
			' useful
			
			full = Replace(full, "+", " ", 1, 1) ' only the first
			
			If InStr(full, " ") > 0 Then
				command_Renamed = st.strip(st.leftsa(full, " ", 1))
				'UPGRADE_WARNING: Couldn't resolve default property of object st.star(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pageName = st.strip(st.star(full, " ", 1, 1))
				tail = st.strip(st.stripHead(full, " ", 1))
			Else
				command_Renamed = full
			End If
		Else
			command_Renamed = "#load"
			pageName = full
		End If
		
	End Sub
	
	Public Function getPageName() As String
		getPageName = pageName
	End Function
	
	Public Function getCommand() As String
		getCommand = command_Renamed
	End Function
	
	Public Function getFull() As String
		getFull = full
	End Function
	
	Public Function getTail() As String
		getTail = tail
	End Function
	
	
	Public Function afterFirstSpace() As String
		' when searching for things with spaces, need
		' to return *everything* after first space
		afterFirstSpace = st.stripHead(full, " ", 1)
	End Function
End Class