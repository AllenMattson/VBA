Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveSheet.ListObjects("1st Qtr. 2013 Student Scores"). _
        Range.AutoFilter Field:=1, Criteria1:="Barbara O'Connor"
    'ActiveSheet.ListObjects("1st Qtr. 2013 Student Scores"). _
        Range.AutoFilter Field:=1
End Sub

Sub Macro2()
    '
    ' Macro1 Macro
    '
    
    '
    Dim strInput As String
    
    strInput = InputBox("Enter the search string:", "Find What")
    
    ActiveSheet.ListObjects("1st Qtr. 2013 Student Scores"). _
      Range.AutoFilter Field:=1, Criteria1:="=*" & strInput & "*"

End Sub



