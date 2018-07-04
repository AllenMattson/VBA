Attribute VB_Name = "ExtractHyperLinks"
Sub ExtractHL()
Dim HL As Hyperlink
For Each HL In ActiveSheet.Hyperlinks
HL.Range.Offset(0, 1).Value = HL.Address
Next
End Sub
Function GetURL(cell As Range, _
Optional default_value As Variant)
'Lists the Hyperlink Address for a Given Cell
'If cell does not contain a hyperlink, return default_value
If (cell.Range("A1").Hyperlinks.Count <> 1) Then
GetURL = default_value
Else
GetURL = cell.Range("A1").Hyperlinks(1).Address & "#" & cell.Range("A1").Hyperlinks(1).SubAddress
End If
End Function



