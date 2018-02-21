Attribute VB_Name = "Module11"
Option Explicit

Sub SaveToDOM()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim xmlDoc As MSXML2.DOMDocument60
    Dim myNode As IXMLDOMNode
    Dim strCurValue As String

    ' Declare constant used as database connection string
    Const strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" _
    & "Data Source=C:\Excel2013_ByExample\Northwind.mdb"

    ' Open a connection to the database
    Set conn = New ADODB.Connection
    conn.Open strConn

    ' Open the Shippers table
    Set rst = New ADODB.Recordset
    rst.Open "Shippers", conn, adOpenStatic, adLockOptimistic

    ' Create a new XML DOMDocument object
    Set xmlDoc = New MSXML2.DOMDocument60

    ' Add the default namespace declaration
    ' to the Namespace names of the DOMDocument object
    ' using the setProperty method of the DOMDocument object

    xmlDoc.setProperty "SelectionNamespaces", _
    "xmlns:rs='urn:schemas-microsoft-com:rowset'" & _
    " xmlns:z='#RowsetSchema'"


    ' Save the recordset directly into
    ' the XML DOMDocument object
    rst.Save xmlDoc, adPersistXML
    Debug.Print xmlDoc.XML

    ' Modify shipper's phone
    Set myNode = xmlDoc.SelectSingleNode( _
    "//z:row[@CompanyName='Speedy Express']/@Phone")
    strCurValue = myNode.Text
    Debug.Print strCurValue
    myNode.Text = "(508)" & Right(strCurValue, 9)
    Debug.Print myNode.Text

    xmlDoc.Save "C:\Excel2013_XML\Shippers_Modified.xml"

    ' Cleanup
    Set xmlDoc = Nothing
    Set conn = Nothing
    Set rst = Nothing
    Set myNode = Nothing
End Sub


