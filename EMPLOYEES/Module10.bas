Attribute VB_Name = "Module10"
Option Explicit

Sub OpenAdoFile()
    Dim rst As ADODB.Recordset
    Dim StartRange As Range
    Dim h As Integer

    ' Create a recordset and fill it with
    ' the data from the XML file
    Set rst = New ADODB.Recordset
    rst.Open "C:\Excel2013_XML\Products.xml", _
    "Provider=MSPersist"

    ' Display the number of records
    MsgBox rst.RecordCount

    ' Open a new workbook
    Workbooks.Add

    ' Copy field names as headings to the first row
    ' of the worksheet
    For h = 1 To rst.Fields.Count
        ActiveSheet.Cells(1, h).Value = rst.Fields(h - 1).Name
    Next

    ' Specify the cell range to receive the data (A2)
    Set StartRange = ActiveSheet.Cells(2, 1)

    ' Copy the records from the recordset
    ' beginning in cell A2
    StartRange.CopyFromRecordset rst

    ' Autofit the columns to make the data fit
    Range("A1").CurrentRegion.Select
    Columns.AutoFit

    ' Close the workbook and save the file
    ActiveWorkbook.Close SaveChanges:=True, _
    Filename:="C:\Excel2013_ByExample\Products.xlsx"
End Sub


