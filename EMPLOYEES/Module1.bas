Attribute VB_Name = "Module1"
Option Explicit

Sub ExportToString()
  Dim strEmpData As String

  ActiveWorkbook.XmlMaps("dataroot_Map").ExportXml _
    Data:=strEmpData
  Debug.Print strEmpData
End Sub


