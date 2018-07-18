Attribute VB_Name = "Examples_DAO"
Option Explicit

Sub DAO_OpenDatabase(strDbPathName As String)
  Dim db As DAO.Database
  Dim tbl As Variant

  Set db = DBEngine.OpenDatabase(strDbPathName)

  MsgBox "There are " & db.TableDefs.Count & _
      " tables in " & strDbPathName & "." & vbCrLf & _
      " View the names in the Immediate window."

  For Each tbl In db.TableDefs
      Debug.Print tbl.Name
  Next

  db.Close
  Set db = Nothing
  MsgBox "The database has been closed."
End Sub

Sub NewDB_DAO()
  Dim db As DAO.Database
  Dim tbl As DAO.TableDef
  Dim strDb As String
  Dim strTbl As String

  On Error GoTo Error_CreateDb_DAO
  strDb = "C:\Excel2013_ByExample\ExcelDump.mdb"
  strTbl = "tblStates"
  ' Create a new database named ExcelDump
  Set db = CreateDatabase(strDb, dbLangGeneral)

  ' Create a new table named tblStates
  Set tbl = db.CreateTableDef(strTbl)

  ' Create fields and append them to the Fields collection
  With tbl
    .Fields.Append .CreateField("StateID", dbText, 2)
    .Fields.Append .CreateField("StateName", dbText, 25)
    .Fields.Append .CreateField("StateCapital", dbText, 25)
  End With

  ' Append the new tbl object to the TableDefs
  db.TableDefs.Append tbl
  ' Close the database
  db.Close
  Set db = Nothing
  MsgBox "There is a new database on your hard disk. " _
      & Chr(13) & "This database file contains a table " _
      & "named " & strTbl & "."
Exit_CreateDb_DAO:
  Exit Sub
Error_CreateDb_DAO:
  If Err.Number = 3204 Then
      ' Delete the database file if it
      ' already exists
      Kill strDb
      Resume
  Else
      MsgBox Err.Number & ": " & Err.Description
      Resume Exit_CreateDb_DAO
  End If
End Sub



