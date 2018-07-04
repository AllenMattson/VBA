Attribute VB_Name = "MedlineImageFiles"
'Required References (Tools -> References in Visual Basic editor):
'1. Microsoft HTML Object Library
'2. Microsoft Internet Controls


Public Sub Test()
    Find_Matching_Images "https://www.medline.com/search/product?Ntt=Gloves&product=%2Fsearch%2Fproduct%3FNtt%3DGloves&xref=%2Fsearch%2Fxref%3Fquestion%3DGloves&contentsearch=%2Fsearch%2Fcontent%3FNtt%3DGloves", _
    "images", Worksheets("Sheet1").Range("A1")
    MsgBox "Finished"
End Sub


Private Sub Find_Matching_Images(sWebSiteURL As String, sImageSearchString As String, destinationStartCell As Range)
Dim ImgSTR As String
    Dim IE As InternetExplorer
    Dim HTMLdoc As HTMLDocument
    Dim imgElements As IHTMLElementCollection
    Dim imgElement As HTMLImg
    Dim aElement As HTMLAnchorElement
    Dim n As Integer
   
    Set IE = New InternetExplorer
   
    IE.Visible = True
    IE.navigate sWebSiteURL
    Do Until IE.readyState = READYSTATE_COMPLETE: DoEvents: Loop
    Do Until IE.document.readyState = "complete": DoEvents: Loop
   
    Set HTMLdoc = IE.document
       
    Set imgElements = HTMLdoc.getElementsByTagName("IMG")
    n = 0
    For Each imgElement In imgElements
    ImgSTR = Trim(imgElement.src)
        Debug.Print imgElement.src & ":" & imgElement.tagName
        Debug.Print imgElement.ParentNode.nodeName
        If Right(ImgSTR, 4) = ".JPG" Then
        'If InStr(imgElement.src, sImageSearchString) Then
            If imgElement.ParentNode.nodeName = "A" Then
                Set aElement = imgElement.ParentNode
                With destinationStartCell
                    .Offset(n, 0).Value = imgElement.src
                    .Offset(n, 1).Value = aElement.href
                End With
                n = n + 1
                Debug.Print imgElement.src
                Debug.Print aElement.href
            End If
        End If
    Next

    IE.Quit
       
End Sub
