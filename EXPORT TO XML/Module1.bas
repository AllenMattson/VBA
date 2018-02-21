Attribute VB_Name = "Module1"
Option Explicit

Sub ExportToXML()
    Dim Filename As Variant
    Dim Rng As Range
    Dim r As Long, c As Long
    
'   Set the range
    Set Rng = Range("Table1[#All]")

'   Get a file name
    Filename = Application.GetSaveAsFilename( _
        InitialFileName:="myrange.xml", _
        fileFilter:="XML Files(*.xml), *.xml")
    If Filename = False Then Exit Sub
    
'   Open the text file
    Open Filename For Output As #1
    
'   Write the <xml> tags
    Print #1, "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
    Print #1, "<EmployeeList xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">"
    
'   Loop through the cells
    For r = 2 To Rng.Rows.Count
        Print #1, "<Employee>"
        For c = 1 To Rng.Columns.Count
            Print #1, "<" & Rng.Cells(1, c) & ">";
            If IsDate(Rng.Cells(r, c)) Then
                Print #1, Format(Rng.Cells(r, c), "yyyy-mm-dd");
            Else
                Print #1, Rng.Cells(r, c).Text;
            End If
            Print #1, "</" & Rng.Cells(1, c) & ">"
        Next c
        Print #1, "</Employee>"
    Next r
'   Close the table
    Print #1, "</EmployeeList>"

'   Close the file
    Close #1
    
'   Tell the user
    MsgBox Rng.Rows.Count - 1 & " records were exported to " & Filename
End Sub

