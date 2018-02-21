Attribute VB_Name = "Module1"
Sub Macro92()

'Step 1:  Declare your variables
    Dim MyDatabase As DAO.Database
    Dim MyQueryDef As DAO.QueryDef
    Dim MyRecordset As DAO.Recordset
    Dim i As Integer
        

'Step 2:  Identify the database and query
    Set MyDatabase = DBEngine.OpenDatabase _
    ("C:\Temp\YourAccessDatabse.accdb")
    
    Set MyQueryDef = MyDatabase.QueryDefs("Your Query Name")
   

'Step 3:  Open the query
    Set MyRecordset = MyQueryDef.OpenRecordset
   

'Step 4:  Clear previous contents
     Sheets("Sheet1").Select
     ActiveSheet.Range("A6:K10000").ClearContents
     

'Step 5:  Copy the recordset to Excel
     ActiveSheet.Range("A7").CopyFromRecordset MyRecordset


'Step 6: Add column heading names to the spreadsheet
    For i = 1 To MyRecordset.Fields.Count
    ActiveSheet.Cells(6, i).Value = MyRecordset.Fields(i - 1).Name
    Next i
   
End Sub


