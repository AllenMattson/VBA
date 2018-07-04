Attribute VB_Name = "GetInnerTextofHTML"
Option Explicit
'set reference to Microsoft HTML Object Library
Sub GetInnerTextofHTML()
    Dim objIe As Object, xobj As HTMLDivElement

    Set objIe = CreateObject("InternetExplorer.Application")
    objIe.Visible = True

    objIe.navigate "C:\a.htm"

    While (objIe.Busy Or objIe.READYSTATE <> 4): DoEvents: Wend

    Set xobj = objIe.document.getElementById("myDiv")
    Set xobj = xobj.getElementsByClassName("myTable").Item(0)
    Set xobj = xobj.getElementsByClassName("data")(0)

    Debug.Print xobj.innerText

    Set xobj = Nothing

    objIe.Quit
    Set objIe = Nothing
End Sub
