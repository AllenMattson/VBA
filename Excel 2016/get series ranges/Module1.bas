Attribute VB_Name = "Module1"
Option Explicit
Option Private Module

Function SERIESNAME_FROM_SERIES(s As Series) As Variant
'   Returns a 2-element variant array
'   1st Element: Data type of the 1st SERIES formula argument(Range, Empty, String)
'   2nd Element: A range address, an empty string, or a string
'   Requres the SERIESFUNC function!
    Dim ResultArray As Variant
    Dim Func As String
    Dim ReturnArray(1 To 2) As Variant

'   Func = Replace(s.Formula, "=SERIES", "'" & ThisWorkbook.Name & "'!SERIESFUNC")
    Func = Application.Substitute(s.Formula, "=SERIES", "'" & ThisWorkbook.Name & "'!SERIESFUNC")
    ResultArray = Evaluate(Func)
    ReturnArray(1) = ResultArray(1, 1)
    ReturnArray(2) = ResultArray(2, 1)
    SERIESNAME_FROM_SERIES = ReturnArray
End Function

Function XVALUES_FROM_SERIES(s As Series) As Variant
'   Returns a 2-element variant array
'   1st Element: Data type of the 2nd SERIES formula argument(Range, Array, Empty, String)
'   2nd Element: A range address, an array, and empty string, or a string
'   Requres the SERIESFUNC function!
    Dim ResultArray As Variant
    Dim Func As String
    Dim ReturnArray(1 To 2) As Variant

'   Func = Replace(s.Formula, "=SERIES", "'" & ThisWorkbook.Name & "'!SERIESFUNC")
    Func = Application.Substitute(s.Formula, "=SERIES", "'" & ThisWorkbook.Name & "'!SERIESFUNC") 'Excel 97 does not support VBA's Replace function
    ResultArray = Evaluate(Func)
    ReturnArray(1) = ResultArray(1, 2)
    ReturnArray(2) = ResultArray(2, 2)
    XVALUES_FROM_SERIES = ReturnArray
End Function


Function VALUES_FROM_SERIES(s As Series) As Variant
'   Returns a 2-element variant array
'   1st Element: Data type of the 3rd SERIES formula argument (Range or Array)
'   2nd Element: A range address, or an array
'   Requres the SERIESFUNC function!
    Dim ResultArray As Variant
    Dim Func As String
    Dim ReturnArray(1 To 2) As Variant

'   Func = Replace(s.Formula, "=SERIES", "'" & ThisWorkbook.Name & "'!SERIESFUNC")
    Func = Application.Substitute(s.Formula, "=SERIES", "'" & ThisWorkbook.Name & "'!SERIESFUNC")
    ResultArray = Evaluate(Func)
    ReturnArray(1) = ResultArray(1, 3)
    ReturnArray(2) = ResultArray(2, 3)
    VALUES_FROM_SERIES = ReturnArray
End Function

Function BUBBLESIZE_FROM_SERIES(s As Series) As Variant
'   Returns a 2-element variant array
'   1st Element: Data type of the 5th SERIES formula argument (Range, Array, or Empty)
'   2nd Element: A range address, an array, or an empty string
'   This is relevant only for Bubble Charts.
'   Requres the SERIESFUNC function!
    Dim ResultArray As Variant
    Dim Func As String
    Dim ReturnArray(1 To 2) As Variant

'   Func = Replace(s.Formula, "=SERIES", "'" & ThisWorkbook.Name & "'!SERIESFUNC")
    Func = Application.Substitute(s.Formula, "=SERIES", "'" & ThisWorkbook.Name & "'!SERIESFUNC")
    ResultArray = Evaluate(Func)
    ReturnArray(1) = ResultArray(1, 5)
    ReturnArray(2) = ResultArray(2, 5)
    BUBBLESIZE_FROM_SERIES = ReturnArray
End Function

Function SERIESFUNC(Optional n, Optional cat, Optional Vals, Optional order, Optional BubSize) As Variant
'   Returns a 2x5 variant array
    Dim i As Long
    Dim Result(1 To 2, 1 To 5) As String

'   Series Name
    Select Case True
        Case IsMissing(n)
            Result(1, 1) = "Empty"
            Result(2, 1) = ""
        Case TypeName(n) = "Range"
            Result(1, 1) = "Range"
            Result(2, 1) = n.Address(, , , True)
        Case TypeName(n) = "String"
            Result(1, 1) = "String"
            Result(2, 1) = n
    End Select
        
'   Categories
    Select Case True
        Case IsMissing(cat)
            Result(1, 2) = "Empty"
            Result(2, 2) = ""
        Case TypeName(cat) = "Range"
            Result(1, 2) = "Range"
            For i = 1 To cat.Areas.Count
                Result(2, 2) = Result(2, 2) & cat.Areas(i).Address(, , , True)
                If i <> cat.Areas.Count Then Result(2, 2) = Result(2, 2) & ","
            Next i
        Case Else
            Result(1, 2) = "Array"
            Result(2, 2) = Result(2, 2) & "{"
            For i = LBound(cat) To UBound(cat)
                Result(2, 2) = Result(2, 2) & cat(i)
                If i <> UBound(cat) Then Result(2, 2) = Result(2, 2) & ","
            Next i
            Result(2, 2) = Result(2, 2) & "}"
    End Select
    
'   Values
    Select Case True
        Case TypeName(Vals) = "Range"
            Result(1, 3) = "Range"
            'Result(2, 3) = vals.Address(, , , True)
            For i = 1 To Vals.Areas.Count
                Result(2, 3) = Result(2, 3) & Vals.Areas(i).Address(, , , True)
                If i <> Vals.Areas.Count Then Result(2, 3) = Result(2, 3) & ","
            Next i
        Case Else
            Result(1, 3) = "Array"
            Result(2, 3) = Result(2, 3) & "{"
            For i = LBound(Vals) To UBound(Vals)
                Result(2, 3) = Result(2, 3) & Vals(i)
                If i <> UBound(Vals) Then Result(2, 3) = Result(2, 3) & ","
            Next i
            Result(2, 3) = Result(2, 3) & "}"
    End Select
    
'   Plot order
    Result(1, 4) = "Integer"
    Result(2, 4) = order
    
'   Bubble size
    Select Case True
        Case IsMissing(BubSize)
            Result(1, 5) = "Empty"
            Result(2, 5) = ""
        Case TypeName(BubSize) = "Range"
            Result(1, 5) = "Range"
            For i = 1 To BubSize.Areas.Count
                Result(2, 5) = Result(2, 5) & BubSize.Areas(i).Address(, , , True)
                If i <> BubSize.Areas.Count Then Result(2, 5) = Result(2, 5) & ","
            Next i
        Case Else
            Result(1, 5) = "Array"
            Result(2, 5) = Result(2, 5) & "{"
            For i = LBound(BubSize) To UBound(BubSize)
                Result(2, 5) = Result(2, 5) & BubSize(i)
                If i <> UBound(BubSize) Then Result(2, 5) = Result(2, 5) & ","
            Next i
            Result(2, 5) = Result(2, 5) & "}"
    End Select
    SERIESFUNC = Result
End Function

