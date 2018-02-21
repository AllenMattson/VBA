Attribute VB_Name = "Module9"
Option Explicit

Sub SaveRst_ADO()
    Dim rst As ADODB.Recordset
    Dim conn As New ADODB.Connection
    Const strConn = "Provider = Microsoft.Jet.OLEDB.4.0;" _
    & "Data Source = C:\Excel2013_HandsOn\Northwind.mdb"

    ' Open a connection to the database
    conn.Open strConn

    ' Execute a select SQL statement against the database
    Set rst = conn.Execute("SELECT * FROM Products")

    ' Delete the file if it exists
    On Error Resume Next
    Kill "C:\Excel2013_XML\Products.xml"

    ' Save the recordset as an XML file
    rst.Save "C:\Excel2013_XML\Products.xml", adPersistXML
    
    rst.Close
    conn.Close
End Sub


