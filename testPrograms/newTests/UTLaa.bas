Attribute VB_Name = "UTLaa"
Option Explicit

' tests for lineArgAnalyser
Public Sub clearText()
    utlaaForm.RichTextBox1.Text = ""
End Sub

Public Sub pt(s As String)
    utlaaForm.RichTextBox1.Text = utlaaForm.RichTextBox1.Text + s + vbCrLf
End Sub

Public Function testVC() As String
    testVC = "AOK"
    
    Dim vc As New VCollection
    
    Call vc.add("val1", "key1")
    Call vc.add("val2", "key2")
    
    If vc.Item("key1") <> "val1" Then
        testVC = "testVC Err1"
        Exit Function
    End If
    
    If vc.Item("key2") <> "val2" Then
        testVC = "testVC Err1"
        Exit Function
    End If
    
    If vc.Count <> 2 Then
        testVC = "testVC Err3"
        Exit Function
    End If
    
    Dim c As Collection
    Set c = vc.toCollection
    If c.Item("key1") <> "val1" Then
        testVC = "testVC Err4"
        Exit Function
    End If
    
    Set c = vc.keyCollection
    If c.Item("key1") <> "key1" Then
        testVC = "testVC Err5"
        Exit Function
    End If
    
    If c.Item("key2") <> "key2" Then
        testVC = "testVC Err6"
        Exit Function
    End If
    
    
End Function

Public Function testPL() As String
    testPL = "AOK"
    
    Dim laa As New LineArgAnaliser
    Dim x As VCollection
    
    Call laa.analise("-name 'Hello'")
    Set x = laa.asVCollection()
    
    If Not x.hasKey("name") Then
        testPL = "testPL : Err 1"
        Exit Function
    End If
    
    If Not x.Item("name") = "Hello" Then
        testPL = "TestPL Err 2"
        Exit Function
    End If
    Set laa = New LineArgAnaliser
    Call laa.analise("-name 'Hello' -pageName 'Somewhere else' -psi 'C:\Program Files'")
    Set x = laa.asVCollection()
        
    If Not x.hasKey("pageName") Then
        testPL = "testPL : Err 3"
        Exit Function
    End If
    
    If Not x.Item("psi") = "C:\Program Files" Then
        testPL = "testPL : Err 4"
        Exit Function
    End If
    
    pt "== " & laa.toString & vbCrLf

End Function
