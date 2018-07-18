Attribute VB_Name = "Module1"
Option Explicit

' The procedure below will publish a worksheet
' with an embedded chart as static HTML

Sub PublishOnWeb(strSheetName As String, _
                 strFileName As String)

    Dim objPub As Excel.PublishObject
    Set objPub = ThisWorkbook.PublishObjects.Add( _
       SourceType:=xlSourceSheet, _
       Filename:=strFileName, Sheet:=strSheetName, _
       HtmlType:=xlHtmlStatic, Title:="Calls Analysis")
    objPub.Publish True
End Sub

Sub CreateHTMLFile()
    Call PublishOnWeb("Help Desk", _
         "C:\Excel2013_ByExample\WorksheetWithChart.htm")
End Sub



