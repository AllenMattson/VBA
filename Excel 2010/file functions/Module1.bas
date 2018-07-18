Attribute VB_Name = "Module1"
Option Explicit

Function FileExists(fname) As Boolean
    FileExists = Dir(fname) <> ""
End Function

Function PathExists(pname) As Boolean
'   Returns TRUE if the path exists
    On Error Resume Next
    PathExists = (GetAttr(pname) And vbDirectory) = vbDirectory
End Function

