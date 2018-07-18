Attribute VB_Name = "Module1"
Function USER()
'   Returns the name of the current user
    USER = Application.UserName
End Function

Function EXCELDIR() As String
'   Returns the directory in which Excel is installed
    EXCELDIR = Application.Path
End Function

Function SHEETCOUNT()
'   Returns the number of workhseets
    Application.Volatile True
    SHEETCOUNT = Application.Caller.Parent.Parent.Sheets.Count
End Function

Function SHEETNAME(Optional sheetnum)
'   Returns the name of the worksheet
    Application.Volatile True
    SHEETNAME = Application.Caller.Parent.Name
End Function

Function STATICRAND()
'   Returns a random number that doesn't
'   change when recalculated
    STATICRAND = Rnd()
End Function

