VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_AfterXMLExport _
    (ByVal Map As XmlMap, _
    ByVal Url As String, _
    ByVal Result As XlXmlExportResult)
    If Result = xlXmlExportSuccess Then
       MsgBox ("XML export succeeded.")
    Else
       MsgBox ("XML export failed.")
    End If
End Sub


Private Sub Workbook_AfterXMLImport _
    (ByVal Map As XmlMap, _
    ByVal IsRefresh As Boolean, _
    ByVal Result As XlXmlImportResult)
    If Result = xlXmlImportSuccess Then
        MsgBox ("XML import succeeded.")
        ActiveSheet.ListObjects(1).Range.Select
     Selection.Interior.ColorIndex = 35
     ActiveCell.Select
    Else
        MsgBox ("XML import failed.")
    End If
End Sub

Private Sub Workbook_BeforeXMLExport _
    (ByVal Map As XmlMap, _
    ByVal Url As String, _
    Cancel As Boolean)

    If (Map.IsExportable) Then
        If MsgBox("Excel is about" & _
        " to export XML from the" & _
        Map.Name & "." & vbCrLf & "Do" & _
        " you want to continue?", _
        vbYesNo + vbQuestion, _
        "XML Export Process") = 7 Then
        Cancel = True
        End If
    End If
End Sub

Private Sub Workbook_BeforeXMLImport _
     (ByVal Map As XmlMap, _
     ByVal Url As String, _
     ByVal IsRefresh As Boolean, _
     Cancel As Boolean)
    
     If MsgBox("Excel is about " & _
        " to import XML into the" & _
        " workbook. Continue with" & _
        " importing?", _
        vbYesNo + vbQuestion, _
        "XML Import Process") = 7 Then
        Cancel = True
     End If
End Sub


