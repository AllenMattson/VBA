Attribute VB_Name = "Module1"
Option Explicit

Private Function FileExists(fname) As Boolean
'   Returns TRUE if the file exists
    Dim x As String
    x = Dir(fname)
    If x <> "" Then FileExists = True _
        Else FileExists = False
End Function

Private Function FileNameOnly(pname) As String
'   Returns the filename from a path/filename string
    Dim temp As Variant
    length = Len(pname)
    temp = Split(pname, Application.PathSeparator)
    FileNameOnly = temp(UBound(temp))
End Function

Private Function PathExists(pname) As Boolean
'   Returns TRUE if the path exists
  If Dir(pname, vbDirectory) = "" Then
    PathExists = False
 Else
    PathExists = (GetAttr(pname) And vbDirectory) = vbDirectory
 End If
End Function


Private Function RangeNameExists(nname) As Boolean
'   Returns TRUE if the range name exists
    Dim n As Name
    RangeNameExists = False
    For Each n In ActiveWorkbook.Names
        If UCase(n.Name) = UCase(nname) Then
            RangeNameExists = True
            Exit Function
        End If
    Next n
End Function

Private Function SheetExists(sname) As Boolean
'   Returns TRUE if sheet exists in the active workbook
    Dim x As Object
    On Error Resume Next
    Set x = ActiveWorkbook.Sheets(sname)
    If Err = 0 Then SheetExists = True _
        Else SheetExists = False
End Function

Private Function WorkbookIsOpen(wbname) As Boolean
'   Returns TRUE if the workbook is open
    Dim x As Workbook
    On Error Resume Next
    Set x = Workbooks(wbname)
    If Err = 0 Then WorkbookIsOpen = True _
        Else WorkbookIsOpen = False
End Function


Private Function IsInCollection(Coln As Object, Item As String) As Boolean
    Dim Obj As Object
    On Error Resume Next
    Set Obj = Coln(Item)
    IsInCollection = Not Obj Is Nothing
End Function


