Attribute VB_Name = "UsefulStuff"
Option Explicit
Public Function getLikelyColumnRange() As Range
    ' figure out the likely default value for the refedit.
    Dim rstart As Range, r As Range
    Set rstart = ActiveSheet.Cells(1, 1)

    While IsEmpty(rstart) And Not Intersect(rstart, rstart.Worksheet.UsedRange) Is Nothing
        Set rstart = rstart.Offset(, 1)
    Wend
    Set r = rstart
    While Not IsEmpty(r)
       Set r = r.Offset(, 1)
    Wend
    If r.Column > rstart.Column Then
        Set getLikelyColumnRange = rstart.Resize(1, r.Column - rstart.Column)
    Else
        Set getLikelyColumnRange = rstart
    End If
End Function
Function firstCell(inrange As Range) As Range
    Set firstCell = inrange.Cells(1, 1)
End Function
Function cleanFind(x As Variant, r As Range, Optional complain As Boolean = False, _
        Optional singlecell As Boolean = False) As Range
    ' does a normal .find, but catches where range is nothing
    Dim u As Range, ch As Boolean
    '
    ' we have a cache going
    '
    Set u = Nothing
    ch = False


    If Not ch Then
        If r Is Nothing Then
            Set u = Nothing
        Else
            Set u = r.Find(x, , xlValues, xlWhole)
        End If
    
        If singlecell And Not u Is Nothing Then
            Set u = firstCell(u)
        End If
        

    End If

    If complain And u Is Nothing Then
            Call msglost(x, r)
    End If
    
    Set cleanFind = u
    
End Function
Sub msglost(x As Variant, r As Range, Optional extra As String = "")

    MsgBox ("Couldnt find " & CStr(x) & " in " & SAd(r) & " " & extra)

End Sub
Function SAd(rngIn As Range, Optional target As Range = Nothing, Optional singlecell As Boolean = False, _
        Optional removeRowDollar As Boolean = False, Optional removeColDollar As Boolean = False) As String
    Dim strA As String
    Dim r As Range
    Dim u As Range
    
    ' creates an address including the worksheet name
    strA = ""
    For Each r In rngIn.Areas
        Set u = r
        If singlecell Then
            Set u = firstCell(u)
        End If
        strA = strA + SAdOneRange(u, target, singlecell, removeRowDollar, removeColDollar) & ","
    Next r
    SAd = Left(strA, Len(strA) - 1)
End Function


Function SAdOneRange(rngIn As Range, Optional target As Range = Nothing, Optional singlecell As Boolean = False, _
                        Optional removeRowDollar As Boolean = False, Optional removeColDollar As Boolean = False) As String
    Dim strA As String
    
    ' creates an address including the worksheet name
    
    strA = AddressNoDollars(rngIn, removeRowDollar, removeColDollar)
    
    ' dont bother with worksheet name if its on the same sheet, and its been asked to do that
    
    If Not target Is Nothing Then
        If target.Worksheet Is rngIn.Worksheet Then
            SAdOneRange = strA
            Exit Function
        End If
    End If

    ' otherwise add the sheet name
    
    SAdOneRange = "'" & rngIn.Worksheet.Name & "'!" & strA
        
End Function
Function AddressNoDollars(a As Range, Optional doRow As Boolean = True, Optional doColumn As Boolean = True) As String
' return address minus the dollars
    Dim st As String
    Dim p1 As Long, p2 As Long
    AddressNoDollars = a.Address
    
    If doRow And doColumn Then
        AddressNoDollars = Replace(a.Address, "$", "")
    Else
        p1 = InStr(1, a.Address, "$")
        p2 = 0
        If p1 > 0 Then
            p2 = InStr(p1 + 1, a.Address, "$")
        End If
        ' turn $A$1 into A$1
        If doColumn And p1 > 0 Then
            AddressNoDollars = Left(a.Address, p1 - 1) & Mid(a.Address, p1 + 1)
        
        ' turn $a$1 into $a1
        ElseIf doRow And p2 > 0 Then
            AddressNoDollars = Left(a.Address, p2 - 1) & Mid(a.Address, p2 + 1, p2 - p1)
    
        End If
    End If
    
    
End Function

Function toEmptyCol(r As Range) As Range
    Dim o As Range
    Set o = r
    
    While Not IsEmpty(firstCell(o))
        Set o = o.Offset(, 1)
    Wend
    If o.Column > r.Column Then
        Set toEmptyCol = r.Resize(r.Rows.Count, o.Column - r.Column)
    End If
    
End Function
Sub deleteAllShapes(r As Range, startingwith As String)
   
    Dim l As Long
    With r.Worksheet
        For l = .Shapes.Count To 1 Step -1
            If Left(.Shapes(l).Name, Len(startingwith)) = startingwith Then
                .Shapes(l).Delete
            End If
        Next l
    End With
    
End Sub
Function makearangeofShapes(r As Range, startingwith As String) As ShapeRange
   
    Dim s As Shape
    
    Dim n() As Variant, sz As Long
    With r.Worksheet
        For Each s In .Shapes
            If Left(s.Name, Len(startingwith)) = startingwith Then
                sz = sz + 1
                ReDim Preserve n(1 To sz) As Variant
                n(sz) = s.Name

            End If
        Next s
        Set makearangeofShapes = .Shapes.Range(n)
    End With
    
End Function
