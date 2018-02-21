Attribute VB_Name = "Module2"
Option Explicit

Sub AddNew_XMLMap()
    Dim lstCourses As ListObject
    Dim lstCol As ListColumn
    Dim objMap As XmlMap
    Dim mapName As String
    Dim strXPath As String

    On Error GoTo ErrorHandler

    ' Create a new XML map
    ActiveWorkbook.XmlMaps.Add _
        ("C:\Excel2013_XML\Courses1.xml ", _
        "Courses").Name = "Courses_Map"

    'location for the new XML table
    Set objMap = ActiveWorkbook.XmlMaps("Courses_Map")
    Range("B20").Select
    
    ' Create a new List object
    Set lstCourses = ActiveSheet.ListObjects.Add

    ' Bind the first XML element to the first table column
    strXPath = "/Courses/Course/@ID"
    With lstCourses.ListColumns(1)
        .XPath.SetValue objMap, strXPath
        .Name = "ID"
    End With

    ' Add a column to the table
    ' and bind it to an XML node
    Set lstCol = lstCourses.ListColumns.Add
    strXPath = "/Courses/Course/Title"
    With lstCol
        .XPath.SetValue objMap, strXPath
        .Name = "Title"
    End With

    ' Add a column to the table
    ' and bind it to an XML node
    Set lstCol = lstCourses.ListColumns.Add

    strXPath = "/Courses/Course/Startdate"
    With lstCol
        .XPath.SetValue objMap, strXPath
        .Name = "Start Date"
    End With

    ' Add a column to the table
    ' and bind it to an XML node
    Set lstCol = lstCourses.ListColumns.Add

    strXPath = "/Courses/Course/Sessions"
    With lstCol
        .XPath.SetValue objMap, strXPath
        .Name = "Sessions"
    End With

    ' Set some XML properties
    With ActiveWorkbook.XmlMaps("Courses_Map")
        .ShowImportExportValidationErrors = False
        .AdjustColumnWidth = True
        .PreserveColumnFilter = True
        .PreserveNumberFormatting = True
        .AppendOnImport = False
    End With

    ' Refresh the XML table in the worksheet
    ActiveWorkbook.XmlMaps("Courses_Map").DataBinding.Refresh
Exit Sub

ErrorHandler:
    MsgBox "The following error has occurred: " & vbCrLf _
    & Err.Description
End Sub


