Attribute VB_Name = "OLE"
Option Explicit

Sub InsertLetter()
    Workbooks.Add
    ActiveSheet.Shapes.AddOLEObject _
        Filename:="C:\Excel2013_HandsOn\Hello.docx"
End Sub

