Attribute VB_Name = "Module1"
Option Explicit

Sub GetData_From_Excel_Sheet()

    Dim MyConnect As String
    Dim MyRecordset As ADODB.Recordset
    Dim MySQL As String
    
    MyConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
               "Data Source=" & ThisWorkbook.FullName & ";" & _
               "Extended Properties=Excel 12.0"

    MySQL = " SELECT * FROM [SampleData$]" & _
            " WHERE Region ='NORTH'"

    Set MyRecordset = New ADODB.Recordset
    MyRecordset.Open MySQL, MyConnect, adOpenStatic, adLockReadOnly

     ThisWorkbook.Sheets.Add
     ActiveSheet.Range("A2").CopyFromRecordset MyRecordset

    With ActiveSheet.Range("A1:F1")
        .Value = Array("Region", "Market", "Branch_Number", _
        "Invoice_Number", "Sales_Amount", "Contracted Hours")
        .EntireColumn.AutoFit
    End With

End Sub

