Attribute VB_Name = "LinkProcessingUnitTests"
Option Explicit

' Unit tests for various link processing and wrapping

Public line As String

Private Sub pt(s As String)
    LTForm.RichText1.text = LTForm.RichText1.text + vbCrLf + s
End Sub

Public Function testWmg(s As String) As String
    Dim wmg As New WikiMarkupGopher
    testWmg = "AOK"
    
    If Not wmg.qq("Hello World") = Chr(34) & "Hello World" & Chr(34) Then
        testWmg = "Error in WikiMarkupGopher.qq"
        Exit Function
    End If
    
    If Not wmg.isAlpha("a") Then
        testWmg = "Error in WikiMarkupGopher.isAlpha"
        Exit Function
    End If
    
    If Not wmg.isAlpha("z") Then
        testWmg = "Error in WikiMarkupGopher.isAlpha"
        Exit Function
    End If
    
    If Not wmg.isAlpha("A") Then
        testWmg = "Error in WikiMarkupGopher.isAlpha"
        Exit Function
    End If
    
    If Not wmg.isAlpha("Z") Then
        testWmg = "Error in WikiMarkupGopher.isAlpha"
        Exit Function
    End If
    
    If Not wmg.isAlpha("Z") Then
        testWmg = "Error in WikiMarkupGopher.isAlpha"
        Exit Function
    End If

    If Not wmg.isAlphaOrSlash("a") Then
        testWmg = "Error 1 in WikiMarkupGopher.isAlphaOrSlash"
        Exit Function
    End If

    If Not wmg.isAlphaOrSlash("A") Then
        testWmg = "Error 2 in WikiMarkupGopher.isAlphaOrSlash"
        Exit Function
    End If

    If Not wmg.isAlphaOrSlash("/") Then
        testWmg = "Error 3 in WikiMarkupGopher.isAlphaOrSlash"
        Exit Function
    End If
    
    If Not wmg.isAlphaOrSlash("3") Then
        testWmg = "Error 4 in WikiMarkupGopher.isAlphaOrSlash"
        Exit Function
    End If
    
    If wmg.isAlphaOrSlash(".") Then
        testWmg = "Error 5 in WikiMarkupGopher.isAlphaOrSlash"
        Exit Function
    End If
    
    If wmg.isAlphaOrSlash(" ") Then
        testWmg = "Error 6 in WikiMarkupGopher.isAlphaOrSlash"
        Exit Function
    End If

    If wmg.hasCapital("abcdefg") Then
        testWmg = "Error 1 in WikiMarkupGopher.hasCapital"
        Exit Function
    End If

    If Not wmg.hasCapital("abcdF    fg ") Then
        testWmg = "Error 2 in WikiMarkupGopher.hasCapital"
        Exit Function
    End If
    
    If Not wmg.hasCapital("ABC") Then
        testWmg = "Error 3 in WikiMarkupGopher.hasCapital"
        Exit Function
    End If

    If Not wmg.hasCapital("A") Then
        testWmg = "Error 4 in WikiMarkupGopher.hasCapital"
        Exit Function
    End If

    If Not wmg.hasCapital("Z") Then
        testWmg = "Error 3 in WikiMarkupGopher.hasCapital"
        Exit Function
    End If
    
    If wmg.hasCapital("a") Then
        testWmg = "Error 4 in WikiMarkupGopher.hasCapital"
        Exit Function
    End If

    If wmg.isWikiWord("abc") Then
        testWmg = "Error 1 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If
    
    If wmg.isWikiWord("aFbcEre") Then
        testWmg = "Error 2 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If
    
    If wmg.isWikiWord("abc ier kjh rw ErjhErehj hjk") Then
        testWmg = "Error 3 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If

    If Not wmg.isWikiWord("AabcVer") Then
        testWmg = "Error 4 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If

    If wmg.isWikiWord("Eabc") Then
        testWmg = "Error 5 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If

    If Not wmg.isWikiWord("AFooBar") Then
        testWmg = "Error 6 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If
    
    If Not wmg.isWikiWord("EekAMouse") Then
        testWmg = "Error 7 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If

    If Not wmg.isWikiWord("MainPage/Subpage") Then
        testWmg = "Error 8 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If

    If Not wmg.isWikiWord("MainPage/fred") Then
        testWmg = "Error 9 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If
    
    If Not wmg.isWikiWord("MainPage/SubPage") Then
        testWmg = "Error 10 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If

    If wmg.isWikiWord("") Then
        testWmg = "Error 11 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If
    
    If wmg.isWikiWord("Mai nPage") Then
        testWmg = "Error 12 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If


    If Not wmg.isWikiWord("M4inPage") Then
        testWmg = "Error 13 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If
    
    If wmg.isWikiWord("Mai.nPage") Then
        testWmg = "Error 14 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If

    If wmg.isWikiWord("Mai_nPage") Then
        testWmg = "Error 15 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If
    
    If wmg.isWikiWord("Main*Page") Then
        testWmg = "Error 16 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If

    If wmg.isWikiWord("(MainPage") Then
        testWmg = "Error 17 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If
    
    If wmg.isWikiWord("MainPage)") Then
        testWmg = "Error 18 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If

    If wmg.isWikiWord("MainPage]") Then
        testWmg = "Error 19 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If

    If wmg.isWikiWord("MainXYPage") Then
        testWmg = "Error 20 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If

    If Not wmg.isWikiWord("MainAPageAThing") Then
        testWmg = "Error 21 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If

    If wmg.isWikiWord("MainA/PageAThing") Then
        testWmg = "Error 22 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If

    If wmg.isWikiWord("MainAPage//AThing") Then
        testWmg = "Error 23 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If
    
    If Not wmg.isWikiWord("MainAPage/AThing") Then
        testWmg = "Error 24 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If

    If wmg.isWikiWord("Main/APage/AThing") Then
        testWmg = "Error 25 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If

    If wmg.isWikiWord("AThing") Then
        testWmg = "Error 26 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If
    
    If wmg.isWikiWord("Some more /bits and pieces") Then
        testWmg = "Error 27 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If

    If wmg.isWikiWord("SomeThing/") Then
        testWmg = "Error 28 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If

    If Not wmg.isWikiWord("Wiki:PhilJones") Then
        testWmg = "Error 29 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If
    
    If wmg.isWikiWord(":PhilJones") Then
        testWmg = "Error 30 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If
    
    If wmg.isWikiWord("PhilJones:") Then
        testWmg = "Error 31 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If
    
    If wmg.isWikiWord("Wiki:PhilJones:AnotherColon") Then
        testWmg = "Error 32 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If
    
    If wmg.isWikiWord("TS") Then
        testWmg = "Error 33 in WikiMarkupGopher.isWikiWord"
        Exit Function
    End If

    If wmg.measureWikiWordAtFront("HelloWorld") <> 10 Then
        testWmg = "Error 1 in WikiMarkupGopher.measureWikiWordAtFront"
        Exit Function
    End If

    If wmg.measureWikiWordAtFront("HelloWorld more stuff") <> 10 Then
        testWmg = "Error 2 in WikiMarkupGopher.measureWikiWordAtFront"
        Exit Function
    End If
    
    If wmg.measureWikiWordAtFront("HelloWorld and SomethingElse more") <> 10 Then
        testWmg = "Error 3 in WikiMarkupGopher.measureWikiWordAtFront"
        Exit Function
    End If
    

End Function

Public Function testLink() As String
    testLink = "AOK"
    
    Dim l As New Link
    Call l.init("blah", "HelloWorld", "normal", "ns", False, False)
    
    Dim l2 As Link
    Set l2 = l.deepCopy()
    
    If l.toString <> "(HelloWorld, blah, normal, ns, False)" Then
        pt l.toString()
        testLink = "Error 1 in Link.toString"
        Exit Function
    End If
    
    If l.toString() <> l2.toString Then
        testLink = "Error 1 in Link.deepCopy"
        Exit Function
    End If
    
    If Not l.matches(l2) Then
        testLink = "Error 1 in matches"
        Exit Function
    End If
    
    l2.text = "Blooo"
    l2.target = "GoodbyeCruelWorld"
    
    If l.matches(l2) Then
        testLink = "Error 2 in matches"
        Exit Function
    End If
    
End Function

Public Function testSLP(s As String) As String
    Dim x As String, y As String
    testSLP = testLink()
    
    Dim slp As New StandardLinkProcessor
    
    Dim l As Link
    Set l = slp.wikiWordToLink("WikiWord")
    
    If l.target <> "WikiWord" Then
        testSLP = "Error 1 in slp.wikiWordToLink"
        Exit Function
    End If
    
    If l.text <> "WikiWord" Then
        testSLP = "Error 2 in slp.wikiWordToLink"
        Exit Function
    End If
    
    If l.nameSpace <> "" Then
        testSLP = "Error 3 in slp.wikiWordToLink"
        Exit Function
    End If
    
    If l.linkType <> "normal" Then
        testSLP = "Error 4 in slp.wikiWordToLink"
        Exit Function
    End If
    
    If l.external <> False Then
        testSLP = "Error 5 in slp.wikiWordToLink"
        Exit Function
    End If
    
    If l.interMap = True Then
        testSLP = "Error 5a in slp.wikiWordToLink"
        Exit Function
    End If
    
    Set l = slp.wikiWordToLink("Wiki:WikiWord")
    
    If l.target <> "WikiWord" Then
        testSLP = "Error 6 in slp.wikiWordToLink"
        Exit Function
    End If
    
    If l.text <> "Wiki:WikiWord" Then
        testSLP = "Error 7 in slp.wikiWordToLink"
        Exit Function
    End If
    
    If l.nameSpace <> "Wiki" Then
        testSLP = "Error 8 in slp.wikiWordToLink"
        Exit Function
    End If
    
    If l.linkType <> "normal" Then
        testSLP = "Error 9 in slp.wikiWordToLink"
        Exit Function
    End If
    
    If l.external <> True Then
        testSLP = "Error 10 in slp.wikiWordToLink"
        Exit Function
    End If
    
    If l.interMap = False Then
        testSLP = "Error 11 in slp.wikiWordToLink"
        Exit Function
    End If
    
    
    Set l = slp.wikiWordToLink("TS:WikiWord")
    
    If l.text <> "TS:WikiWord" Then
        testSLP = "Error 12 in slp.wikiWordToLink"
        pt l.text
        Exit Function
    End If
    
    If l.nameSpace <> "TS" Then
        testSLP = "Error 13 in slp.wikiWordToLink"
        Exit Function
    End If
    
    If l.linkType <> "normal" Then
        testSLP = "Error 14 in slp.wikiWordToLink"
        Exit Function
    End If
    
    If l.external <> True Then
        testSLP = "Error 15 in slp.wikiWordToLink"
        Exit Function
    End If
    
    If l.interMap = False Then
        testSLP = "Error 16 in slp.wikiWordToLink"
        Exit Function
    End If
    
    
    If slp.wikiWords("Hello teenage America") <> "Hello teenage America" Then
        testSLP = "Error 1 in slp.wikiWords"
        Exit Function
    End If
    
    If slp.wikiWords("Hello TeenageAmerica") <> "Hello LINK0" Then
        testSLP = "Error 2 in slp.wikiWords"
        Exit Function
    End If

    
    'pt slp.wikiWords("Hello TeenageAmerica goodbye CruelWorld")
    
    If slp.wikiWords("Hello TeenageAmerica goodbye CruelWorld") <> "Hello LINK1 goodbye LINK2" Then
        testSLP = "Error 3 in slp.wikiWords"
        Exit Function
    End If
    
    If slp.wikiWords("Hello TeenageAmerica/allstars goodbye MissAmericanPy") <> "Hello LINK3 goodbye LINK4" Then
        testSLP = "Error 4 in slp.wikiWords"
        Exit Function
    End If
    
    x = "(TeenageAmerica, TeenageAmerica, normal, , False)" & vbCrLf & _
"(TeenageAmerica, TeenageAmerica, normal, , False)" & vbCrLf & _
"(CruelWorld, CruelWorld, normal, , False)" & vbCrLf & _
"(TeenageAmerica/allstars, TeenageAmerica/allstars, normal, , False)" & vbCrLf & _
"(MissAmericanPy, MissAmericanPy, normal, , False)" & vbCrLf
'    pt "*" & slp.linksToString & "*"
'    pt "*" & x & "*"
    
    
    If slp.linksToString() <> x Then
        testSLP = "Error 5 in slp.wikiWords"
        Exit Function
    End If
    
    x = slp.wikiWords("HelloWorld")
    If x <> "LINK5" Then
        testSLP = "Error 6 in slp.wikiWords"
        pt x
        Exit Function
    End If
    
    Set slp = New StandardLinkProcessor

    x = slp.looseURL("hello world", "http://")
    If x <> "hello world" Then
        pt x
        testSLP = "Error 1 in WikiToHtml.changeURIs"
        Exit Function
    End If
    
    x = slp.looseURL("my site is at http://www.synaesmedia.com Check it out", "http://")
    y = "my site is at LINK0 Check it out"
    If x <> y Then
        testSLP = "Error 2 in WikiToHtml.looseURL"
        pt "**" & x & "**"
        pt "*" & y & "*"
        
        Exit Function
    End If
    
    x = slp.looseURL("my site is at https://www.synaesmedia.com Check it out and there's more at https://www.somewhere-else.net/myScript.cgi?act=go&data=3 moooooooo", "https://")
    y = "my site is at LINK1 Check it out and there's more at LINK2 moooooooo"
    If x <> y Then
        testSLP = "Error 3 in WikiToHtml.changeURIs"
        pt "**" & x & "**"
        pt "*" & y & "*"
        
        Exit Function
    End If
    
    

End Function

Public Function testDB() As String
    Dim slp As New StandardLinkProcessor
    Dim x As String
    
    testDB = "AOK"
    
    If slp.wikiWordsAmongBrackets("hello world") <> "hello world" Then
        testDB = "Error 1 in slp.amongBrackets"
        Exit Function
    End If
    
    'pt "*" & slp.wikiWordsAmongBrackets("hello [[teenage america]]") & "*"
    
    If slp.wikiWordsAmongBrackets("hello [[teenage america]]") <> "hello LINK0" Then
        testDB = "Error 2 in slp.amongBrackets"
        Exit Function
    End If
    
    If slp.wikiWordsAmongBrackets("hello [[teenage america]] and [[type>more stuff | alt text]]") <> "hello LINK1 and LINK2" Then
        testDB = "Error 3 in slp.amongBrackets"
        Exit Function
    End If
    
    If slp.wikiWordsAmongBrackets("hello CruelWorld [[teenage america]] and [[counter>against phil | U R Stoopid]]") <> _
    "hello LINK3 LINK4 and LINK5" Then
        testDB = "Error 4 in slp.amongBrackets"
        Exit Function
    End If
    
    x = slp.wikiWordsAmongBrackets("hello Wiki:PhilJones more [[a new link]]")
    If x <> "hello LINK6 more LINK7" Then
        testDB = "Error 5 in slp.amongBrackets"
        pt x
        Exit Function
    End If
    
    
    If slp.doubleBracketContent("explanation>some text | alternative") <> "LINK8" Then
        testDB = "Error 1 in slp.doubleBracketContent"
        Exit Function
    End If
    
    Call slp.singleBracketContent("http://www.synaesmedia.net Synaesmedia")
    Call slp.singleBracketContent("https://www.oldskool.com Old Skool Jungle")

    x = slp.singleBrackets("there is [http://www.synaesmedia.net something interesting] about SdiDesk", "http://")
    If x <> "there is LINK11 about SdiDesk" Then
        pt x
        testDB = "Error 1 in singleBrackets"
        Exit Function
    End If

    x = slp.singleBrackets("there is [https://www.synaesmedia.net something interesting] about SdiDesk", "https://")
    If x <> "there is LINK12 about SdiDesk" Then
        pt x
        testDB = "Error 2 in singleBrackets"
        Exit Function
    End If
    
    x = slp.singleBrackets("there is [https://www.synaesmedia.net something interesting] about SdiDesk", "http://")
    If x = "there is LINK13 about SdiDesk" Then
        pt x
        testDB = "Error 3 in singleBrackets"
        Exit Function
    End If
    
    x = "(teenage_america, teenage america, , , False)" & vbCrLf & _
"(teenage_america, teenage america, , , False)" & vbCrLf & _
"(more_stuff, alt text, type, , False)" & vbCrLf & _
"(CruelWorld, CruelWorld, normal, , False)" & vbCrLf & _
"(teenage_america, teenage america, , , False)" & vbCrLf & _
"(against_phil, U R Stoopid, counter, , False)" & vbCrLf & _
"(PhilJones, Wiki:PhilJones, normal, Wiki, True)" & vbCrLf & _
"(a_new_link, a new link, , , False)" & vbCrLf & _
"(some_text, alternative, explanation, , False)" & vbCrLf & _
"(http://www.synaesmedia.net, Synaesmedia, , , True)" & vbCrLf & _
"(https://www.oldskool.com, Old Skool Jungle, , , True)" & vbCrLf & _
"(http://www.synaesmedia.net, something interesting, , , True)" & vbCrLf & _
"(https://www.synaesmedia.net, something interesting, , , True)" & vbCrLf

    If slp.linksToString <> x Then
        testDB = "Error 4 in slp.linksToString"
        pt slp.linksToString
        Exit Function
    End If
    pt slp.linksToString

    Set slp = New StandardLinkProcessor
    Dim y As String
    y = "first WikiWord and [[in brackets]] and SecondWiki and "
    x = slp.wikiWordsAmongBrackets(y)
    If x <> "first LINK0 and LINK1 and LINK2 and " Then
        testDB = "Error 6 in slp.wikiWordsAmongBrackets"
        pt x
        Exit Function
    End If

End Function

Public Function testW2H() As String
    testW2H = "AOK"
    Dim w2h As New WikiToHtml
    Dim x As String, y As String
        
    testW2H = testLine
    
    
End Function

Public Function testSt() As String
    ' testing string tool
    testSt = "AOK"
    
    Dim st As New StringTool
    
    Dim s As String
    Dim p() As String
    
    p = st.mySplit("hello world", ",,", "")
    If UBound(p) > 0 Then
        testSt = "Error 1 in StringTool.split"
        Exit Function
    End If
    
    If p(0) <> "hello world" Then
        testSt = "Error 2 in StringTool.split"
        Exit Function
    End If
    
    p = st.mySplit("hello world", " ", "")
    If UBound(p) <> 1 Then
        testSt = "Error 3 in StringTool.split"
        Exit Function
    End If
    
    If p(0) <> "hello" Then
        testSt = "Error 4 in StringTool.split"
        Exit Function
    End If
    
    If p(1) <> "world" Then
        testSt = "Error 5 in StringTool.split"
        Exit Function
    End If
    
    p = st.mySplit("hello teenage america", " ", "")
    If UBound(p) <> 2 Then
        testSt = "Error 6 in StringTool.split"
        Exit Function
    End If
    
    If p(2) <> "america" Then
        testSt = "Error 6 in StringTool.split"
        Exit Function
    End If
    
    p = st.mySplit("hello 'teenage america'", " ", "'")
    
    If UBound(p) <> 1 Then
        testSt = "Error 7 in StringTool.split"
        Exit Function
    End If
    
    If p(1) <> "'teenage america'" Then
        pt p(1)
        testSt = "Error 8 in StringTool.split"
        Exit Function
    End If
    
    p = st.mySplit("'hello teenage' america 'the great' and", " ", "'")
    If UBound(p) <> 3 Then
        Dim x As Variant
        For Each x In p
            pt CStr(x)
        Next x
        testSt = "Error 9 in StringTool.split"
        Exit Function
    End If
    
    If p(0) <> "'hello teenage'" Or p(1) <> "america" Or p(2) <> "'the great'" Or p(3) <> "and" Then
        pt p(0) & p(1) & p(2) & p(3)
        testSt = "Error 10 in StringTool.split"
        Exit Function
    End If
    
    
End Function

Public Function testNlw() As String
    testNlw = "AOK"
    
    Dim nlw As New NativeLinkWrapper
    Dim sysco As New SysConfStub
    Dim map As InterWikiMap
    Set map = sysco.asSystemConfigurations.interMap
    Dim slp As New StandardLinkProcessor
    Set nlw.asLinkWrapper.remoteInterMap = map
    Call map.add("http://www.synaesmedia.net/wiki/wiki.cgi?", "TiSo")
    Call map.add("www.google.com/", "XY")
    Set nlw.asLinkWrapper.remoteWads = New WADSStub
    Set nlw.asLinkWrapper.remoteSysConf = sysco
    
    Dim s As String
    s = slp.asLinkProcessor.wrapAllLinks("HelloWorld I'm TiSo:ComingHome to [http://www.nooranch.com NooRanch] bloo Wiki:wikiWay kjh", nlw)
    pt s
    pt slp.linksToString

    s = slp.asLinkProcessor.wrapAllLinks("as XY:boo lkjs ' iu XY:Yeah kj", nlw)
    pt s
    pt slp.linksToString

End Function

Public Function ocContains(oc As OCollection, l As Link) As Boolean
' does oc contain a Link that matches l
    Dim l2 As Link
    For Each l2 In oc.toCollection
        If l2.matches(l) Then
            ocContains = True
            Exit Function
        End If
    Next l2
End Function

Public Function testLine() As String
    testLine = "AOK"
    
    Dim raw As String, cooked As String, x As String
    Dim w2h As New WikiToHtml
    Dim lp As New StandardLinkProcessor
    Dim lw As New NativeLinkWrapper
    Dim dw As New DummyWads
    Dim scs As New SysConfStub
    
    Set lw.asLinkWrapper.remoteWads = dw
    Set lw.asLinkWrapper.remoteSysConf = scs
    
    raw = "hello world"
    
    cooked = w2h.mainTransform(raw, lp, lw)
    If cooked <> "hello world" Then
        testLine = "Error 1 in mainTransform"
        Exit Function
    End If
    
    raw = "hello world and WikiWord and some"
    
    cooked = w2h.mainTransform(raw, lp, lw)
    If cooked <> "hello world and <a href='about:blank' class='normal' id='WikiWord'><font color='#ffeedd'>WikiWord</font></a> and some" Then
        pt "==" & cooked & "=="
        testLine = "Error 2 in mainTransform"
        Exit Function
    End If
    
    raw = "This is in ''italics''"
    
    cooked = w2h.mainTransform(raw, lp, lw)
    If cooked <> "This is in <i>italics</i>" Then
        pt cooked
        testLine = "Error 3 in mainTransform"
        Exit Function
    End If
    
    raw = "[http://www.synaesmedia.net Syn] blah"
    cooked = w2h.mainTransform(raw, lp, lw)
    x = "<a href='http://www.synaesmedia.net' id='external' target='new'>Syn</a> blah"
    If cooked <> x Then
        pt raw
        pt "==" & cooked & "=="
        pt "--" & x & "--"
        testLine = "Error 4 in mainTransform"
   '     Exit Function
    End If
    
    raw = "* hello world"
    cooked = w2h.mainTransform(raw, lp, lw)
    x = "<ul>" & vbCrLf & "<li> hello world</li>"
    If cooked <> x Then
        pt "-"
        pt "==" & cooked & "=="
        pt "--" & x & "--"
        testLine = "Error 4a in mainTransform"
   '     Exit Function
    End If
    
    
    raw = "* http://www.synaesmedia.net blah"
    x = "<ul>" & vbCrLf & "<li> <a href='http://www.synaesmedia.net' id='external' target='new'>http://www.synaesmedia.net</a> blah</li>"
        
    cooked = w2h.mainTransform(raw, lp, lw)
    If cooked <> x Then
        pt "-"
        pt raw
        pt "==" & cooked & "=="
        pt "--" & x & "--"
        
        testLine = "Error 5 in mainTransform"
 '       Exit Function
    End If
    
    'pt vbCrLf & "6"
    raw = "* http://www.synaesmedia.net"
    cooked = w2h.mainTransform(raw, lp, lw)
    x = "<ul>" & vbCrLf & "<li> <a href='http://www.synaesmedia.net' id='external' target='new'>http://www.synaesmedia.net</a></li>"
    If cooked <> x Then
        pt "-"
        pt raw
        pt "==" & cooked & "=="
        pt "--" & x & "--"
        
        testLine = "Error 6 in mainTransform"
   '     Exit Function
    End If
    
    'pt vbCrLf & "7"
    raw = "* http://www.synaesmedia.net" + vbCrLf
    x = "<ul>" & vbCrLf & "<li> <a href='http://www.synaesmedia.net' id='external' target='new'>http://www.synaesmedia.net</a></li>" + vbCrLf + "</ul>" + vbCrLf + "<p />" + vbCrLf
    cooked = w2h.mainTransform(raw, lp, lw)
    If cooked <> x Then
        pt "-"
        pt raw
        pt cooked
        pt x
        testLine = "Error 7 in mainTransform"
        Exit Function
    End If
    
    
End Function
