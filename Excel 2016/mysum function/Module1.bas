Attribute VB_Name = "Module1"
Option Explicit

Function MYSUM(ParamArray args() As Variant) As Variant
' Emulates Excel's SUM function
  
' Variable declarations
  Dim i As Variant
  Dim TempRange As Range, cell As Range
  Dim ECode As String
  Dim m, n
  MYSUM = 0

' Process each argument
  For i = 0 To UBound(args)
'   Skip missing arguments
    If Not IsMissing(args(i)) Then
'     What type of argument is it?
      Select Case TypeName(args(i))
        Case "Range"
'         Create temp range to handle full row or column ranges
          Set TempRange = Intersect(args(i).Parent.UsedRange, args(i))
          For Each cell In TempRange
            If IsError(cell) Then
              MYSUM = cell ' return the error
              Exit Function
            End If
            If cell = True Or cell = False Then
              MYSUM = MYSUM + 0
            Else
              If IsNumeric(cell) Or IsDate(cell) Then _
                 MYSUM = MYSUM + cell
              End If
          Next cell
        Case "Variant()"
            n = args(i)
            For m = LBound(n) To UBound(n)
               MYSUM = MYSUM(MYSUM, n(m)) 'recursive call
            Next m
        Case "Null"  'ignore it
        Case "Error" 'return the error
          MYSUM = args(i)
          Exit Function
        Case "Boolean"
'         Check for literal TRUE and compensate
          If args(i) = "True" Then MYSUM = MYSUM + 1
        Case "Date"
          MYSUM = MYSUM + args(i)
        Case Else
          MYSUM = MYSUM + args(i)
      End Select
    End If
  Next i
End Function

