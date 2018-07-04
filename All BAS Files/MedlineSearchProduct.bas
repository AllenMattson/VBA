Attribute VB_Name = "MedlineSearchProduct"
'"Microsoft Internet Controls" library in your VBA references.
Sub SearchProduct()
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .Calculation = xlAutomatic
End With
     'we define the essential variables
    Dim ie As Object
    Dim pwd, SearchStr
    Dim button
    Dim Xobj As HTMLDivElement
    Dim i As Long
    Dim trs, tds ', username, pwd, button
    Dim tr As Long, td As Long, r As Long, C As Long
    'add the "Microsoft Internet Controls" reference in your VBA Project indirectly
    Set ie = CreateObject("InternetExplorer.Application")
    With ie
        .Visible = True
        ''''''''''''''''''''''
        .navigate ("https://www.medline.com/account/login.jsp?fromst=true#")
        While ie.readyState <> 4
            DoEvents
        Wend
        Set username = .document.getElementById("ext-gen1004") 'id of the username control (HTML Control)'ext-gen1004
        Set pwd = .document.getElementById("ext-gen1005") 'id of the password control (HTML Control) 'ext-gen1005
        Set button = .document.getElementById("submitbutton") 'id of the button control (HTML Control) 'submitbutton
        username.Value = "danya"
        pwd.Value = "Geritom10"
        button.Click
        While ie.readyState <> 4
            DoEvents
        Wend
        Application.Wait Now + TimeValue("0:00:02")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Set SearchStr = .document.getElementById("searchQuestion") 'id of the username control (HTML Control)'ext-gen1004
        Set button = .document.getElementById("searchSubmit") '("button-blue search right") 'id of the button control (HTML Control) 'submitbutton
        SearchStr.Value = "COI21053"
        button.Click
        While ie.readyState <> 4
            DoEvents
        Wend
     Application.Wait Now + TimeValue("0:00:02")
    Set Xobj = ie.document.getElementById("galleryView") ''''"myDiv"'''''
    Set Xobj = Xobj.getElementsByClassName("medGridViewSkuListWrapper").Item(0)
    Set Xobj = Xobj.getElementsByClassName("medGridViewSkuList persist-area")(0)
        Set trs = Xobj.getElementsByClassName("skuRow    ")
        For r = 0 To trs.Length - 1
            Set tds = trs(r).getElementsByTagName("td")
                For C = 0 To tds.Length - 1
                    Debug.Print tds(C).innerText 'SKUsh.Cells(100000, C + 1).End(xlUp).Offset(1, 0).Value = tds(C).innerText
                Next C
        Next r
    End With
    Set ie = Nothing
    
End Sub
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

