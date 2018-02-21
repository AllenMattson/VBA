Attribute VB_Name = "Module1"
Option Explicit

Sub ADO_Demo()
'   This demo requires a reference to
'   the Microsoft ActiveX Data Objects 2.x Library
    
    Dim DBFullName As String
    Dim Cnct As String, Src As String
    Dim Connection As ADODB.Connection
    Dim Recordset As ADODB.Recordset
    Dim Col As Integer

    Cells.Clear

'   Database information
    DBFullName = ThisWorkbook.Path & "\budget data.accdb"
    
'   Open the connection
    Set Connection = New ADODB.Connection
    Cnct = "Provider=Microsoft.ACE.OLEDB.12.0;"
    Cnct = Cnct & "Data Source=" & DBFullName & ";"
    Connection.Open ConnectionString:=Cnct
    
'   Create RecordSet
    Set Recordset = New ADODB.Recordset
    With Recordset
'       Filter
        Src = "SELECT * FROM Budget WHERE Item = 'Lease' "
        Src = Src & "and Division = 'N. America' "
        Src = Src & "and Year = '2008'"
        .Open Source:=Src, ActiveConnection:=Connection

        MsgBox "The Query:" & vbNewLine & vbNewLine & Src

'       Write the field names
        For Col = 0 To Recordset.Fields.Count - 1
           Range("A1").Offset(0, Col).Value = Recordset.Fields(Col).Name
        Next

'       Write the recordset
        Range("A1").Offset(1, 0).CopyFromRecordset Recordset
    End With
    Set Recordset = Nothing
    Connection.Close
    Set Connection = Nothing
End Sub
