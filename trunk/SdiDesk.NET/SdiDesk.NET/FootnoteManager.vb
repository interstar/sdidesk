Option Strict Off
Option Explicit On
Friend Class FootnoteManager
	
	' this object manages the extraction of footnotes embedded in pages and
	' places them at the bottom of the page
	
	Public footnotes As Collection
	Private counter As Short
	
	Public Sub init()
		footnotes = New Collection
		counter = 1
	End Sub
	
	Public Sub addFootnote(ByRef index As Short, ByRef note As String)
		Call footnotes.Add(note, CStr(index))
	End Sub
	
	Public Function extractFootnotes(ByRef line As String, ByRef native As Boolean) As String
		Dim s, note As String
		Dim i As Short
		Dim v As Object
		Dim parts() As String
		If InStr(line, "{{") Then
			parts = Split(line, "{{")
			s = parts(0)
			For i = 1 To UBound(parts)
				If InStr(parts(i), "}}") Then
					note = Left(parts(i), InStr(parts(i), "}}") - 1)
					Call addFootnote(counter, note)
					If Not native Then
						'internal links not working in the native html
						MsgBox(1)
						s = s & "(<b><a href='#" & counter & "'>" & counter & "</a></b>)"
					Else
						s = s & "(<b>" & counter & "</b>)"
					End If
					s = s & Mid(parts(i), InStr(parts(i), "}}") + 2)
					counter = counter + 1
				Else
					s = s & parts(i)
				End If
			Next i
			
			extractFootnotes = s
		Else
			extractFootnotes = line
		End If
	End Function
	
	
	Public Function getFootnoteCollection() As Collection
		getFootnoteCollection = footnotes
	End Function
	
	Public Function getFootnotesAsHtmlString() As String
		Dim s As String
		Dim i As Short
		s = ""
		For i = 1 To counter - 1
			' nasty, shouldn't be here.
			s = s & "* (<a name='" & i & "'>" & i & "</a>) "
			'UPGRADE_WARNING: Couldn't resolve default property of object footnotes.item(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			s = s + footnotes.Item(i)
			s = s & vbCrLf
		Next i
		getFootnotesAsHtmlString = s
	End Function
End Class