Attribute VB_Name = "Module1"
Option Explicit


Function ISBOLD(Cell) As Boolean
'   Returns TRUE if cell is bold
    ISBOLD = Cell.Range("A1").Font.Bold
End Function

Function ISITALIC(Cell) As Boolean
'   Returns TRUE if cell is italic
    ISITALIC = Cell.Range("A1").Font.Italic
End Function


Function ALLBOLD(Cell) As Boolean
'   Returns TRUE if all characters in cell are bold
    ALLBOLD = Not IsNull(Cell.Font.Bold)
End Function

Function FILLCOLOR(Cell) As Integer
'   Returns an integer corresponding to
'   cell's interior color
    Application.Volatile
    FILLCOLOR = Cell.Range("A1").Interior.ColorIndex
End Function

Function SAYIT(txt)
    Application.Speech.Speak (txt)
    SAYIT = txt
End Function

Function LASTSAVED()
    Application.Volatile
    LASTSAVED = ThisWorkbook. _
      BuiltinDocumentProperties("Last Save Time")
End Function

Function LASTPRINTED()
    Application.Volatile
    LASTPRINTED = ThisWorkbook. _
      BuiltinDocumentProperties("Last Print Date")
End Function


Function LastSaved2()
    Application.Volatile
    LastSaved2 = Application.Caller.Parent.Parent. _
      BuiltinDocumentProperties("Last Save Time")
End Function

Function LASTPRINTED2()
    Application.Volatile
    LASTPRINTED2 = Application.Caller.Parent.Parent. _
      BuiltinDocumentProperties("Last Print Date")
End Function


Function SheetName(ref) As String
    SheetName = ref.Parent.Name
End Function

Function WorkbookName(ref) As String
    WorkbookName = ref.Parent.Parent.Name
End Function

Function AppName(ref) As String
    AppName = ref.Parent.Parent.Parent.Name
End Function

Function COUNTBETWEEN(InRange, num1, num2) As Long
'   Counts number of values between num1 and num2
    With Application.WorksheetFunction
        If num1 <= num2 Then
            COUNTBETWEEN = .CountIfs(InRange, ">=" & num1, _
                InRange, "<=" & num2)
        Else
            COUNTBETWEEN = .CountIfs(InRange, ">=" & num2, _
                InRange, "<=" & num1)
        End If
    End With
End Function


Function LASTINCOLUMN(Rng As Range)
'   Returns the contents of the last non-empty cell in a column
    Dim LastCell As Range
    Application.Volatile
    With Rng.Parent
        With .Cells(.Rows.Count, Rng.Column)
            If Not IsEmpty(.Value) Then
                LASTINCOLUMN = .Value
            ElseIf IsEmpty(.End(xlUp)) Then
                LASTINCOLUMN = ""
            Else
                LASTINCOLUMN = .End(xlUp).Value
            End If
         End With
    End With
End Function


Function LASTINROW(Rng As Range)
'   Returns the contents of the last non-empty cell in a row
    Application.Volatile
    With Rng.Parent
        With .Cells(Rng.Row, .Columns.Count)
            If Not IsEmpty(.Value) Then
                LASTINROW = .Value
            ElseIf IsEmpty(.End(xlToLeft)) Then
                LASTINROW = ""
            Else
                LASTINROW = .End(xlToLeft).Value
            End If
         End With
    End With
End Function


Function ISLIKE(text As String, pattern As String) As Boolean
'   Returns true if the first argument is like the second
    ISLIKE = text Like pattern
End Function

Function EXTRACTELEMENT(txt, n, Separator) As String
'   Returns the nth element of a text string, where the
'   elements are separated by a specified separator character
    Dim AllElements As Variant
    AllElements = Split(txt, Separator)
    EXTRACTELEMENT = AllElements(n - 1)
End Function

Function EXTRACTELEMENT2(txt, n, Separator) As String
'   Returns the nth element of a text string, where the
'   elements are separated by a specified separator character

    Dim Txt1 As String, TempElement As String
    Dim ElementCount As Integer, i As Integer
    
    Txt1 = txt
'   If space separator, remove excess spaces
    If Separator = Chr(32) Then Txt1 = Application.Trim(Txt1)
    
'   Add a separator to the end of the string
    If Right(Txt1, Len(Txt1)) <> Separator Then _
        Txt1 = Txt1 & Separator
    
'   Initialize
    ElementCount = 0
    TempElement = ""
    
'   Extract each element
    For i = 1 To Len(Txt1)
        If Mid(Txt1, i, 1) = Separator Then
            ElementCount = ElementCount + 1
            If ElementCount = n Then
'               Found it, so exit
                EXTRACTELEMENT2 = TempElement
                Exit Function
            Else
                TempElement = ""
            End If
        Else
            TempElement = TempElement & Mid(Txt1, i, 1)
        End If
    Next i
    EXTRACTELEMENT2 = ""
End Function


Function STATFUNCTION(Rng, op)
    Select Case UCase(op)
        Case "SUM"
            STATFUNCTION = WorksheetFunction.Sum(Rng)
        Case "AVERAGE"
            STATFUNCTION = WorksheetFunction.Average(Rng)
        Case "MEDIAN"
            STATFUNCTION = WorksheetFunction.Median(Rng)
        Case "MODE"
            STATFUNCTION = WorksheetFunction.Mode(Rng)
        Case "COUNT"
            STATFUNCTION = WorksheetFunction.Count(Rng)
        Case "MAX"
            STATFUNCTION = WorksheetFunction.Max(Rng)
        Case "MIN"
            STATFUNCTION = WorksheetFunction.Min(Rng)
        Case "VAR"
            STATFUNCTION = WorksheetFunction.Var(Rng)
        Case "STDEV"
            STATFUNCTION = WorksheetFunction.StDev(Rng)
        Case Else
            STATFUNCTION = CVErr(xlErrNA)
    End Select
End Function

Function SHEETOFFSET(Offset As Long, Optional Cell As Variant)
'   Returns cell contents at Ref, in sheet offset
    Dim WksIndex As Long, WksNum As Long
    Dim wks As Worksheet
    Application.Volatile
    If IsMissing(Cell) Then Set Cell = Application.Caller
    WksNum = 1
    For Each wks In Application.Caller.Parent.Parent.Worksheets
        If Application.Caller.Parent.Name = wks.Name Then
            SHEETOFFSET = Worksheets(WksNum + Offset).Range(Cell(1).Address)
            Exit Function
        Else
            WksNum = WksNum + 1
        End If
    Next wks
End Function

Function MAXALLSHEETS(Cell)
    Dim MaxVal As Double
    Dim Addr As String
    Dim Wksht As Object
    Application.Volatile
    Addr = Cell.Range("A1").Address
    MaxVal = -9.9E+307
    For Each Wksht In Cell.Parent.Parent.Worksheets
        If Wksht.Name = Cell.Parent.Name And _
          Addr = Application.Caller.Address Then
        ' avoid circular reference
        Else
            If IsNumeric(Wksht.Range(Addr)) Then
                If Wksht.Range(Addr) > MaxVal Then _
                  MaxVal = Wksht.Range(Addr).Value
            End If
        End If
    Next Wksht
    If MaxVal = -9.9E+307 Then MaxVal = 0
    MAXALLSHEETS = MaxVal
End Function


Function RANDOMINTEGERS()
    Dim FuncRange As Range
    Dim V() As Variant, ValArray() As Variant
    Dim CellCount As Double
    Dim i As Integer, j As Integer
    Dim r As Integer, c As Integer
    Dim Temp1 As Variant, Temp2 As Variant
    Dim RCount As Integer, CCount As Integer
    
'   Create Range object
    Set FuncRange = Application.Caller

'   Return an error if FuncRange is too large
    CellCount = FuncRange.Count
    If CellCount > 1000 Then
        RANDOMINTEGERS = CVErr(xlErrNA)
        Exit Function
    End If
    
'   Assign variables
    RCount = FuncRange.Rows.Count
    CCount = FuncRange.Columns.Count
    ReDim V(1 To RCount, 1 To CCount)
    ReDim ValArray(1 To 2, 1 To CellCount)

'   Fill array with random numbers
'   and consecutive integers
    For i = 1 To CellCount
        ValArray(1, i) = Rnd
        ValArray(2, i) = i
    Next i

'   Sort ValArray by the random number dimension
    For i = 1 To CellCount
        For j = i + 1 To CellCount
            If ValArray(1, i) > ValArray(1, j) Then
                Temp1 = ValArray(1, j)
                Temp2 = ValArray(2, j)
                ValArray(1, j) = ValArray(1, i)
                ValArray(2, j) = ValArray(2, i)
                ValArray(1, i) = Temp1
                ValArray(2, i) = Temp2
            End If
        Next j
    Next i
    
'   Put the randomized values into the V array
    i = 0
    For r = 1 To RCount
        For c = 1 To CCount
            i = i + 1
            V(r, c) = ValArray(2, i)
        Next c
    Next r
    RANDOMINTEGERS = V
End Function

Function RANGERANDOMIZE(Rng)
    Dim V() As Variant, ValArray() As Variant
    Dim CellCount As Double
    Dim i As Integer, j As Integer
    Dim r As Integer, c As Integer
    Dim Temp1 As Variant, Temp2 As Variant
    Dim RCount As Integer, CCount As Integer
    
'   Return an error if rng is too large
    CellCount = Rng.Count
    If CellCount > 1000 Then
        RANGERANDOMIZE = CVErr(xlErrNA)
        Exit Function
    End If
    
'   Assign variables
    RCount = Rng.Rows.Count
    CCount = Rng.Columns.Count
    ReDim V(1 To RCount, 1 To CCount)
    ReDim ValArray(1 To 2, 1 To CellCount)

'   Fill ValArray with random numbers
'   and values from rng
    For i = 1 To CellCount
        ValArray(1, i) = Rnd
        ValArray(2, i) = Rng(i)
    Next i

'   Sort ValArray by the random number dimension
    For i = 1 To CellCount
        For j = i + 1 To CellCount
            If ValArray(1, i) > ValArray(1, j) Then
                Temp1 = ValArray(1, j)
                Temp2 = ValArray(2, j)
                ValArray(1, j) = ValArray(1, i)
                ValArray(2, j) = ValArray(2, i)
                ValArray(1, i) = Temp1
                ValArray(2, i) = Temp2
            End If
        Next j
    Next i
    
'   Put the randomized values into the V array
    i = 0
    For r = 1 To RCount
        For c = 1 To CCount
            i = i + 1
            V(r, c) = ValArray(2, i)
        Next c
    Next r
    RANGERANDOMIZE = V
End Function

Function SPELLDOLLARS(Cell) As Variant

    Dim Dollars As String
    Dim Cents As String
    Dim TextLen As Integer
    Dim Temp As String
    Dim Pos As Integer
    Dim iHundreds As Integer
    Dim iTens As Integer
    Dim iOnes As Integer
    Dim Units(2 To 5) As String
    Dim bHit As Boolean
    Dim Ones As Variant
    Dim Teens As Variant
    Dim Tens As Variant
    Dim NegFlag As Boolean

'   Is it a non-number?
    If Not IsNumeric(Cell) Then
        SPELLDOLLARS = CVErr(xlErrValue)
        Exit Function
    End If

'   Is it negative?
    If Cell < 0 Then
        NegFlag = True
        Cell = Abs(Cell)
    End If
    
    Dollars = Format(Cell, "###0.00")
    TextLen = Len(Dollars) - 3

'   Is it too large?
    If TextLen > 15 Then
        SPELLDOLLARS = CVErr(xlErrNum)
        Exit Function
    End If

'   Do the cents part
    Cents = Right(Dollars, 2) & "/100 Dollars"
    If Cell < 1 Then
        SPELLDOLLARS = Cents
        Exit Function
    End If

    Dollars = Left(Dollars, TextLen)

    Ones = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine")
    Teens = Array("Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen")
    Tens = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")

    Units(2) = "Thousand"
    Units(3) = "Million"
    Units(4) = "Billion"
    Units(5) = "Trillion"

    Temp = ""

    For Pos = 15 To 3 Step -3
        If TextLen >= Pos - 2 Then
            bHit = False
            If TextLen >= Pos Then
                iHundreds = Asc(Mid$(Dollars, TextLen - Pos + 1, 1)) - 48
                If iHundreds > 0 Then
                    Temp = Temp & " " & Ones(iHundreds) & " Hundred"
                    bHit = True
                End If
        End If
        iTens = 0
        iOnes = 0

        If TextLen >= Pos - 1 Then
            iTens = Asc(Mid$(Dollars, TextLen - Pos + 2, 1)) - 48
        End If

        If TextLen >= Pos - 2 Then
            iOnes = Asc(Mid$(Dollars, TextLen - Pos + 3, 1)) - 48
        End If

        If iTens = 1 Then
            Temp = Temp & " " & Teens(iOnes)
            bHit = True
        Else
            If iTens >= 2 Then
                Temp = Temp & " " & Tens(iTens)
                bHit = True
            End If
            If iOnes > 0 Then
                If iTens >= 2 Then
                    Temp = Temp & "-"
                Else
                    Temp = Temp & " "
                End If
                Temp = Temp & Ones(iOnes)
                bHit = True
            End If
        End If
        If bHit And Pos > 3 Then
            Temp = Temp & " " & Units(Pos \ 3)
        End If
    End If
    Next Pos

  SPELLDOLLARS = Trim(Temp) & " and " & Cents
  If NegFlag Then SPELLDOLLARS = "(" & SPELLDOLLARS & ")"
End Function


Function SORTED(Rng)
    Dim SortedData() As Variant
    Dim Cell As Range
    Dim Temp As Variant, i As Long, j As Long
    Dim NonEmpty As Long
   
'   Transfer data to SortedData
    For Each Cell In Rng
        If Not IsEmpty(Cell) Then
            NonEmpty = NonEmpty + 1
            ReDim Preserve SortedData(1 To NonEmpty)
            SortedData(NonEmpty) = Cell.Value
        End If
    Next Cell

'   Sort the array
    For i = 1 To NonEmpty
        For j = i + 1 To NonEmpty
            If SortedData(i) > SortedData(j) Then
                Temp = SortedData(j)
                SortedData(j) = SortedData(i)
                SortedData(i) = Temp
            End If
        Next j
    Next i
        
'   Transpose the array and return it
    SORTED = Application.Transpose(SortedData)
End Function

