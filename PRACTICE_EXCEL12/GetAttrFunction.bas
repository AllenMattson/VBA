Attribute VB_Name = "GetAttrFunction"
Sub GetAttributes()
    Dim attr As Integer
    Dim msg As String
    Dim strFileName As String
    
    strFileName = InputBox("Enter the complete file name:", _
        "Drive\Folder\Filename")
    If strFileName = "" Then Exit Sub
    attr = GetAttr(strFileName)
    
    msg = ""

    If attr And vbReadOnly Then msg = msg & "Read-Only (R)"
    If attr And vbHidden Then msg = msg & Chr(10) & "Hidden (H)"
    If attr And vbSystem Then msg = msg & Chr(10) & "System (S)"
    If attr And vbArchive Then msg = msg & Chr(10) & "Archive (A)"
    MsgBox msg, , strFileName

End Sub

