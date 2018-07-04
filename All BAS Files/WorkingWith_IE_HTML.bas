Attribute VB_Name = "WorkingWith_IE_HTML"
'"Microsoft Internet Controls" library in your VBA references.
Sub ExplorerTest()
Const myPageTitle As String = "Wikipedia"
Const myPageURL As String = "http://en.wikipedia.org/wiki/Main_Page"
Const mySearchForm As String = "searchform"
Const mySearchInput As String = "searchInput"
Const mySearchTerm As String = "Document Object Model"
Const myButton As String = "Go"

Dim myIE As SHDocVw.InternetExplorer

  'check if page is already open
  Set myIE = GetOpenIEByTitle(myPageTitle, False)
  
  If myIE Is Nothing Then
    'page isn't open yet
    'create new IE instance
    Set myIE = GetNewIE
    'make IE window visible
    myIE.Visible = True
    'load page
    If LoadWebPage(myIE, myPageURL) = False Then
      'page wasn't loaded
      MsgBox "Couldn't open page"
      Exit Sub
    End If
  End If
  
  With myIE.document.forms(mySearchForm)
    'enter search term in text field
    .elements(mySearchInput).Value = mySearchTerm
    'press button "Go"
    .elements(myButton).Click
  End With
       
End Sub
'returns new instance of Internet Explorer
Function GetNewIE() As SHDocVw.InternetExplorer
  'create new IE instance
  Set GetNewIE = New SHDocVw.InternetExplorer
  'start with a blank page
  GetNewIE.Navigate2 "about:Blank"
End Function
'loads a web page and returns True or False depending on
'whether the page could be loaded or not
Function LoadWebPage(i_IE As SHDocVw.InternetExplorer, _
                     i_URL As String) As Boolean
  With i_IE
    'open page
    .navigate i_URL
    'wait until IE finished loading the page
    Do While .readyState <> READYSTATE_COMPLETE
      Application.Wait Now + TimeValue("0:00:01")
    Loop
    'check if page could be loaded
    If .document.URL = i_URL Then
      LoadWebPage = True
    End If
  End With
End Function
'finds an open IE site by checking the URL
Function GetOpenIEByURL(ByVal i_URL As String) As SHDocVw.InternetExplorer
Dim objShellWindows As New SHDocVw.ShellWindows

  'ignore errors when accessing the document property
  On Error Resume Next
  'loop over all Shell-Windows
  For Each GetOpenIEByURL In objShellWindows
    'if the document is of type HTMLDocument, it is an IE window
    If TypeName(GetOpenIEByURL.document) = "HTMLDocument" Then
      'check the URL
      If GetOpenIEByURL.document.URL = i_URL Then
        'leave, we found the right window
        Exit Function
      End If
    End If
  Next
End Function
'finds an open IE site by checking the title
Function GetOpenIEByTitle(i_Title As String, _
                          Optional ByVal i_ExactMatch As Boolean = True) As SHDocVw.InternetExplorer
Dim objShellWindows As New SHDocVw.ShellWindows

  If i_ExactMatch = False Then i_Title = "*" & i_Title & "*"
  'ignore errors when accessing the document property
  On Error Resume Next
  'loop over all Shell-Windows
  For Each GetOpenIEByTitle In objShellWindows
    'if the document is of type HTMLDocument, it is an IE window
    If TypeName(GetOpenIEByTitle.document) = "HTMLDocument" Then
      'check the title
      If GetOpenIEByTitle.document.Title Like i_Title Then
        'leave, we found the right window
        Exit Function
      End If
    End If
  Next
End Function

