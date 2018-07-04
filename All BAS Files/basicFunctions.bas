Attribute VB_Name = "basicFunctions"
Option Private Module
Option Base 1
Option Explicit

Public Function fieldNameIsOk(fieldName) As Boolean
    fieldNameIsOk = False
    If fieldName <> 0 And fieldName <> vbNullString And LCase(fieldName) <> "none" And Left$(fieldName, 2) <> "--" And Left$(fieldName, 1) <> "_" And LCase(fieldName) <> "(add...)" Then fieldNameIsOk = True
End Function
Public Function capitalizeFirstLetter(str As String) As String
    capitalizeFirstLetter = UCase(Left(str, 1)) & Right(str, Len(str) - 1)
End Function
Public Function parsePhotoID(photoNameAndId As String) As String
    Dim temp As String
    temp = Left(photoNameAndId, Len(photoNameAndId) - 1)
    temp = Right(temp, Len(temp) - InStrRev(temp, "("))
    parsePhotoID = temp
End Function
Public Function parseUserName(str As String) As String
    On Error GoTo errhandler
    Dim resultStr As String
    If InStr(1, str, "(id: ") = 0 Then
        parseUserName = str
    Else
        resultStr = Left(str, InStr(1, str, "(id: ") - 1)
        parseUserName = resultStr
    End If
    Exit Function
errhandler:
    parseUserName = str
End Function
Public Function parseUserID(str As String) As String
    On Error GoTo errhandler
    Dim resultStr As String
    If InStr(1, str, "(id: ") = 0 Then
        parseUserID = str
    Else
        resultStr = Right(str, Len(str) - InStr(1, str, "(id: ") - 4)
        resultStr = Left(resultStr, Len(resultStr) - 1)
        parseUserID = resultStr
    End If
    Exit Function
errhandler:
    parseUserID = str
End Function
Function fontIsInstalled(sFont) As Boolean
    On Error Resume Next
    Dim fontList As Object
    Dim TempBar As Object
    Dim i As Integer

    '   Returns True if sFont is installed
    fontIsInstalled = False
    Set fontList = Application.CommandBars("Formatting").FindControl(ID:=1728)

    '   If Font control is missing, create a temp CommandBar
    If fontList Is Nothing Then
        Set TempBar = Application.CommandBars.Add
        Set fontList = TempBar.Controls.Add(ID:=1728)
    End If

    For i = 0 To fontList.ListCount - 1
        If fontList.List(i + 1) = sFont Then
            fontIsInstalled = True
            On Error Resume Next
            TempBar.Delete
            Exit Function
        End If
    Next i

    '   Delete temp CommandBar if it exists
    On Error Resume Next
    TempBar.Delete
End Function
Public Function valueIsInArray(val As Variant, arr As Variant) As Boolean
    Dim rivi As Long
    valueIsInArray = False
    If Not IsArray(arr) Then Exit Function
    For rivi = LBound(arr) To UBound(arr)
        If arr(rivi) = val Then
            valueIsInArray = True
            Exit Function
        End If
    Next rivi
    valueIsInArray = False
End Function
Sub storeValue(settingName As String, settingValue, ws As Worksheet, Optional rangeName As String = "")
    Dim rivi As Long
    rivi = 0
    With ws
        On Error Resume Next
        rivi = Application.Match(settingName, .Columns("A"), 0)

        If debugMode = True Then On Error GoTo 0
        If rivi = 0 Then rivi = vikarivi(.Cells(1, 1)) + 1

        .Cells(rivi, 1).value = settingName
        .Cells(rivi, 2).value = settingValue
        If rangeName <> vbNullString Then .Cells(rivi, 2).Name = rangeName
    End With
End Sub

Public Function fetchValue(settingName As String, ws As Worksheet) As Variant
    Dim rivi As Long
    rivi = 0

    With ws
        On Error Resume Next
        rivi = Application.Match(settingName, .Columns("A"), 0)

        If debugMode = True Then On Error GoTo 0
        If rivi = 0 Then
            fetchValue = ""
        Else
            fetchValue = .Cells(rivi, 2).value
        End If
    End With
End Function

Public Function fetchSettingAddress(settingName As String, ws As Worksheet) As String
    Dim rivi As Long
    rivi = 0

    With ws
        On Error Resume Next
        rivi = Application.Match(settingName, .Columns("A"), 0)

        If debugMode = True Then On Error GoTo 0
        If rivi = 0 Then
            fetchSettingAddress = ""
        Else
            fetchSettingAddress = .Cells(rivi, 2).Address
        End If
    End With
End Function

Sub fetchValueToRange(settingName As String, ws As Worksheet, Optional inputValueTo As Range)
    Dim rivi As Long
    Dim settingValue As Variant
    rivi = 0

    With ws
        On Error Resume Next
        rivi = Application.Match(settingName, .Columns("A"), 0)

        If debugMode = True Then On Error GoTo 0
        If rivi = 0 Then
            settingValue = ""
        Else
            settingValue = .Cells(rivi, 2).value
            If Not IsMissing(inputValueTo) Then
                If Not inputValueTo Is Nothing Then inputValueTo.value = settingValue
            End If
        End If
    End With
End Sub


Function shapeExists(ByRef shapeName As String, Optional sheetName As String) As Boolean

    Dim ws As Worksheet
    If sheetName = "" Or IsMissing(sheetName) = True Then
        Set ws = ActiveSheet
    Else
        Set ws = Sheets(sheetName)
    End If

    shapeExists = False
    Dim sh As Shape
    For Each sh In ws.Shapes
        If sh.Name = shapeName Then
            shapeExists = True
            Exit Function
        End If
    Next sh
End Function
Function ChartExists(strChartName As String, wsTest As Worksheet) As Boolean
    Dim chTest As ChartObject

    On Error Resume Next
    Set chTest = wsTest.ChartObjects(strChartName)
    On Error GoTo 0

    If chTest Is Nothing Then
        ChartExists = False
    Else
        ChartExists = True
    End If

End Function
Sub laskealue()
    Selection.Calculate
End Sub
Sub breakLinks()
    Dim i As Long
    Dim astrLinks As Variant
    On Error Resume Next

    ' Define variable as an Excel link type.
    astrLinks = ThisWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)

    If Not astrLinks = "" Then
        ' Break the first link in the active workbook.
        For i = 1 To UBound(astrLinks)
            ThisWorkbook.BreakLink _
                    Name:=astrLinks(i), _
                    Type:=xlLinkTypeExcelLinks
        Next i
    End If
End Sub

Public Function parseVarFromName(nameStr As String, Var As String) As Variant
    Dim startC As Long
    Dim endC As Long
    'nameStr = LCase(nameStr)
    'Var = LCase(Var)
    startC = InStr(1, nameStr, "_" & Var)
    If startC = 0 Then Exit Function
    endC = InStr(startC + 1, nameStr, "_")
    If endC = 0 Then
        endC = Len(nameStr) + 5
    End If
    parseVarFromName = Mid(nameStr, startC + 1 + Len(Var), endC - startC - Len(Var) - 1)
End Function


Public Function ColumnLetter(l) As String
    Dim s0 As String, s1 As String, S2 As String, s3 As String
    If l > 18278 Then s0 = Chr$((Int((l - 18279) / 17576) Mod 26) + 65)
    If l > 702 Then s1 = Chr$((Int((l - 703) / 676) Mod 26) + 65)
    If l > 26 Then S2 = Chr$((Int((l - 27) / 26)) Mod 26 + 65)
    s3 = Chr$(((l - 1) Mod 26) + 65)
    ColumnLetter = s0 & s1 & S2 & s3
End Function

Public Function vikarivi(solu As Range) As Long
    vikarivi = solu.Worksheet.Cells(solu.Worksheet.Rows.Count, solu.Column).End(xlUp).row
End Function

Public Function vikasar(solu As Range) As Long
    vikasar = solu.Worksheet.Cells(solu.row, solu.Worksheet.Columns.Count).End(xlToLeft).Column
End Function


Public Function parseVarFromStr(ByVal str, Var, Optional separatorChar = "%") As String

    On Error GoTo errhandler

    Dim varStart As Long

    varStart = InStr(1, str, separatorChar & Var & "->")
    If varStart = 0 Then
        parseVarFromStr = ""
        'If debugMode = True Then Debug.Print "Variable " & var & " missing from string " & Left(str, 5000)
    Else
        parseVarFromStr = Mid(str, varStart + Len(separatorChar & Var & "->"), InStr(varStart + Len(separatorChar & Var & "->"), str, separatorChar) - varStart - Len(separatorChar & Var & "->"))
    End If
    Exit Function
errhandler:
    parseVarFromStr = ""
End Function
Function convertRSCL(ByVal str As Variant) As String
    str = Replace(str, rscL0, "%rscL0%")
    str = Replace(str, rscL1, "%rscL1%")
    str = Replace(str, rscL2, "%rscL2%")
    str = Replace(str, rscL3, "%rscL3%")
    str = Replace(str, rscL4, "%rscL4%")
    convertRSCL = str
End Function

Public Function findLastCell(sh As Worksheet) As Range

    On Error GoTo errhandler

    Dim LastColumn As Long
    Dim lastRow As Long
    Dim lastCell As Range
    If Application.CountA(sh.Cells) > 0 Then
        'Search for any entry, by searching backwards by Rows.
        lastRow = sh.Cells.Find(What:="*", after:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
        'Search for any entry, by searching backwards by Columns.
        LastColumn = sh.Cells.Find(What:="*", after:=[A1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

        Set findLastCell = sh.Cells(lastRow, LastColumn)
    Else
        Set findLastCell = Cells(1, 1)
    End If

    Exit Function

errhandler:

    Set findLastCell = Cells(1, 1)

End Function








Public Function arrMatch(ByVal arrname As Variant, ByVal value As Variant, Optional col As Long = 1)

    Dim rivi As Long

    For rivi = 1 To UBound(arrname)

        If arrname(rivi, col) = value Then
            arrMatch = rivi
            Exit Function
        End If

    Next rivi

    arrMatch = -1

End Function

Public Function weekDayMonSun(date1 As Date) As Integer


'MON-SUN
' weekDayMonSun = WeekDay(date1) - 1

'SUN-SAT
    weekDayMonSun = WeekDay(date1)

    If weekDayMonSun = 0 Then weekDayMonSun = 7
End Function

Public Function monthNameToNumber(ByVal monthName As String) As Integer
    monthName = Replace(monthName, ",", "")
    monthName = Trim(LCase(monthName))
    Select Case monthName
    Case "january"
        monthNameToNumber = 1
    Case "february"
        monthNameToNumber = 2
    Case "march"
        monthNameToNumber = 3
    Case "april"
        monthNameToNumber = 4
    Case "may"
        monthNameToNumber = 5
    Case "june"
        monthNameToNumber = 6
    Case "july"
        monthNameToNumber = 7
    Case "august"
        monthNameToNumber = 8
    Case "september"
        monthNameToNumber = 9
    Case "october"
        monthNameToNumber = 10
    Case "november"
        monthNameToNumber = 11
    Case "december"
        monthNameToNumber = 12
    End Select
End Function



Public Function getMonthName(monthNum As Integer) As String
    Select Case monthNum
    Case 1
        getMonthName = "January"
    Case 2
        getMonthName = "February"
    Case 3
        getMonthName = "March"
    Case 4
        getMonthName = "April"
    Case 5
        getMonthName = "May"
    Case 6
        getMonthName = "June"
    Case 7
        getMonthName = "July"
    Case 8
        getMonthName = "August"
    Case 9
        getMonthName = "September"
    Case 10
        getMonthName = "October"
    Case 11
        getMonthName = "November"
    Case 12
        getMonthName = "December"
    End Select
End Function


Public Function isTime(ByVal fieldName As String, Optional granularity As String = "", Optional includeMisc As Boolean = False) As Boolean

    fieldName = LCase(fieldName)
    Select Case fieldName
    Case "hour"
        If granularity = "hour" Or granularity = "" Then isTime = True
    Case "date", "day"
        If granularity = "date" Or granularity = "" Then isTime = True
    Case "dayofmonth", "dayofweek", "weekday"
        If (granularity = "date" Or granularity = "") And includeMisc Then isTime = True
    Case "week", "weekiso", "yearweek", "yearweekiso"
        If granularity = "week" Or granularity = "" Then isTime = True
    Case "month", "yearmonth"
        If granularity = "month" Or granularity = "" Then isTime = True
    Case "quarter"
        If (granularity = "quarter" Or granularity = "") And includeMisc Then isTime = True
    Case "year", "yearofisoweek"
        If granularity = "year" Or granularity = "" Then isTime = True
    End Select

End Function







Public Function CharCount(OrigString As String, _
                          Chars As String, Optional CaseSensitive As Boolean = False) _
                          As Long

'**********************************************
'PURPOSE: Returns Number of occurrences of a character or
'or a character sequencence within a string

'PARAMETERS:
'OrigString: String to Search in
'Chars: Character(s) to search for
'CaseSensitive (Optional): Do a case sensitive search
'Defaults to false

'RETURNS:
'Number of Occurrences of Chars in OrigString

'EXAMPLES:
'Debug.Print CharCount("FreeVBCode.com", "E") -- returns 3
'Debug.Print CharCount("FreeVBCode.com", "E", True) -- returns 0
'Debug.Print CharCount("FreeVBCode.com", "co") -- returns 2
''**********************************************

    Dim lLen As Long
    Dim lCharLen As Long
    Dim lAns As Long
    Dim sInput As String
    Dim sChar As String
    Dim lCtr As Long
    Dim lEndOfLoop As Long
    Dim bytCompareType As Byte

    sInput = OrigString
    If sInput = vbNullString Then Exit Function
    lLen = Len(sInput)
    lCharLen = Len(Chars)
    lEndOfLoop = (lLen - lCharLen) + 1
    bytCompareType = IIf(CaseSensitive, vbBinaryCompare, _
                         vbTextCompare)

    For lCtr = 1 To lEndOfLoop
        sChar = Mid$(sInput, lCtr, lCharLen)
        If StrComp(sChar, Chars, bytCompareType) = 0 Then _
           lAns = lAns + 1
    Next

    CharCount = lAns

End Function




Sub copyRangeNames()

    Dim n As Name
    For Each n In Names
        Debug.Print n.Name
        Debug.Print n.RefersTo
        Debug.Print Range(n).Address
        If InStr(1, n.RefersTo, "Analytics!") > 0 Then
            AdWords.Range(Range(n).Address).Name = n.Name & "AW"
        ElseIf InStr(1, n.RefersTo, "vars!") > 0 Then
            Sheets("varsAW").Range(Range(n).Address).Name = n.Name & "AW"
        End If
    Next
End Sub
Public Function uriEncode(ByVal str As String) As String
    On Error Resume Next
    If str = "" Then
        uriEncode = vbNullString
        Exit Function
    End If

    Dim resultStr As String
    resultStr = vbNullString


    str = UTF8_Encode(str)
    str = URLEncode2(str, True)

    str = Replace(str, "+", "%20")
    uriEncode = str


End Function



Public Function URLEncode2( _
       StringVal As String, _
       Optional SpaceAsPlus As Boolean = False _
     ) As String

    Dim StringLen As Long: StringLen = Len(StringVal)
    Dim result As Variant
    If StringLen > 0 Then
        ReDim result(StringLen) As String
        Dim i As Long, CharCode As Integer
        Dim char As String, Space As String

        If SpaceAsPlus Then Space = "+" Else Space = "%20"

        For i = 1 To StringLen
            char = Mid$(StringVal, i, 1)
            CharCode = Asc(char)
            Select Case CharCode
            Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                result(i) = char
            Case 32
                result(i) = Space
            Case 0 To 15
                result(i) = "%0" & Hex(CharCode)
            Case Else
                result(i) = "%" & Hex(CharCode)
            End Select
        Next i
        URLEncode2 = Join(result, "")
    End If
End Function
Public Function URL_decode(sEncodedURL As String) As String

    On Error Resume Next

    Dim iLoop As Long
    Dim sRtn As String
    Dim sTmp As String
    Dim sTmp2 As String

    If Len(sEncodedURL) > 0 Then
        ' Loop through each char
        For iLoop = 1 To Len(sEncodedURL)
            sTmp = Mid(sEncodedURL, iLoop, 1)
            sTmp = Replace(sTmp, "+", " ")
            ' If char is % then get next two chars
            ' and convert from HEX to decimal
            If sTmp = "%" And Len(sEncodedURL) + 1 > iLoop + 2 Then
                sTmp2 = vbNullString
                sTmp2 = Chr(CDbl("&H" & Mid(sEncodedURL, iLoop + 1, 2)))
                If sTmp2 <> vbNullString Then
                    sTmp = sTmp2
                    ' Increment loop by 2
                    iLoop = iLoop + 2
                End If
            End If

            sRtn = sRtn & sTmp
        Next
        URL_decode = sRtn
    End If

End Function
Public Function chrDecode(ByVal str As String) As String

    str = Replace(str, "chr124", Chr$(124))
    str = Replace(str, "chr60", Chr$(60))
    str = Replace(str, "chr62", Chr$(62))
    str = Replace(str, "chr61", Chr$(61))

    str = Replace(str, "chr38", Chr$(38))
    str = Replace(str, "chr47", Chr$(47))
    str = Replace(str, "chr34", Chr$(34))
    str = Replace(str, "chr39", Chr$(39))
    str = Replace(str, "chr42", Chr$(42))
    str = Replace(str, "chr63", Chr$(63))
    str = Replace(str, "chr35", Chr$(35))

    str = Replace(str, "chr64", Chr$(64))

    str = Replace(str, "chr92", Chr$(92))
    str = Replace(str, "chr58", Chr$(58))
    str = Replace(str, "chr46", Chr$(46))
    str = Replace(str, "chr37", Chr$(37))
    str = Replace(str, "chr45", Chr$(45))

    chrDecode = str

End Function

Public Function UTF8_Encode(ByVal sStr As String) As String
    On Error Resume Next
    Dim l As Long
    Dim lChar&
    Dim sUTF8$
    For l = 1 To Len(sStr)
        lChar& = AscW(Mid(sStr, l, 1))
        If lChar& < 128 Then
            sUTF8$ = sUTF8$ + Mid(sStr, l, 1)
        ElseIf ((lChar& > 127) And (lChar& < 2048)) Then
            sUTF8$ = sUTF8$ + Chr(((lChar& \ 64) Or 192))
            sUTF8$ = sUTF8$ + Chr(((lChar& And 63) Or 128))
        Else
            sUTF8$ = sUTF8$ + Chr(((lChar& \ 144) Or 234))
            sUTF8$ = sUTF8$ + Chr((((lChar& \ 64) And 63) Or 128))
            sUTF8$ = sUTF8$ + Chr(((lChar& And 63) Or 128))
        End If
    Next l
    UTF8_Encode = sUTF8$
End Function

Function UTF8_Decode(ByVal sStr As String) As String
    On Error Resume Next
    Dim l As Long, sUTF8 As String, iChar As Integer, iChar2 As Integer
    For l = 1 To Len(sStr)
        iChar = Asc(Mid(sStr, l, 1))
        If iChar > 127 Then
            If Not iChar And 32 Then    ' 2 chars
                iChar2 = Asc(Mid(sStr, l + 1, 1))
                sUTF8 = sUTF8 & ChrW$(((31 And iChar) * 64 + (63 And iChar2)))
                l = l + 1
            Else
                Dim iChar3 As Integer
                iChar2 = Asc(Mid(sStr, l + 1, 1))
                iChar3 = Asc(Mid(sStr, l + 2, 1))
                sUTF8 = sUTF8 & ChrW$(((iChar And 15) * 16 * 256) + ((iChar2 And 63) * 64) + (iChar3 And 63))
                l = l + 2
            End If
        Else
            sUTF8 = sUTF8 & Chr$(iChar)
        End If
    Next l
    UTF8_Decode = sUTF8
End Function


Public Function mPW(ByVal pw As String) As String

    Dim maskedPW As String
    Dim kirjain As Long
    Dim i As Long
    Dim num As Long

    Randomize

    For kirjain = 1 To Len(pw)
        maskedPW = maskedPW & Mid(pw, Len(pw) - kirjain + 1, 1)
    Next kirjain

    For i = 1 To 3
        Select Case i
        Case 1
            num = Int((90 - 65 + 1) * Rnd + 65)
        Case 2
            num = Int((57 - 48 + 1) * Rnd + 48)
        Case 3
            num = Int((122 - 97 + 1) * Rnd + 97)
        End Select
        maskedPW = Chr(num) & maskedPW
    Next i
    For i = 1 To 3
        Select Case i
        Case 1
            num = Int((90 - 65 + 1) * Rnd + 65)
        Case 2
            num = Int((57 - 48 + 1) * Rnd + 48)
        Case 3
            num = Int((122 - 97 + 1) * Rnd + 97)
        End Select
        maskedPW = maskedPW & Chr(num)
    Next i

    mPW = maskedPW

End Function

Public Function umPW(ByVal pw As String) As String

    Dim maskedPW As String
    Dim kirjain As Long

    pw = Left(pw, Len(pw) - 3)
    pw = Right(pw, Len(pw) - 3)

    For kirjain = 1 To Len(pw)
        maskedPW = maskedPW & Mid(pw, Len(pw) - kirjain + 1, 1)
    Next kirjain

    umPW = maskedPW

End Function


Public Function useEvaluateInFormula(formulaStr As String) As String

    formulaStr = Replace(formulaStr, " ", " & ")
    formulaStr = Replace(formulaStr, Chr(42), Chr(34) & Chr(42) & Chr(34))  '*
    formulaStr = Replace(formulaStr, "+", Chr(34) & "+" & Chr(34))
    formulaStr = Replace(formulaStr, "-", Chr(34) & "-" & Chr(34))
    formulaStr = Replace(formulaStr, "/", Chr(34) & "/" & Chr(34))
    formulaStr = Replace(formulaStr, "(", Chr(34) & "(" & Chr(34) & " & ")
    formulaStr = Replace(formulaStr, ")", " & " & Chr(34) & ")" & Chr(34))

    formulaStr = "evaluate(" & formulaStr & ")"

    useEvaluateInFormula = formulaStr

End Function

Sub testFind()


    Dim vrivi As Long
    Dim k As Long
    Dim i As Long
    Dim useMatch As Boolean
    Dim rivi1 As Long
    Dim rivi2 As Long
    Dim rivi3 As Long
    Dim rivi4 As Long
    Dim rivi5 As Long
    Dim rivi6 As Long
    Dim rivi7 As Long
    Dim rivi8 As Long
    vrivi = vikarivi(Cells(1, 10))

    For k = 1 To 1

        If k = 1 Then
            usingMacOSX = False
            useMatch = False
        ElseIf k = 2 Then
            usingMacOSX = True
            useMatch = False
        Else
            usingMacOSX = True
            useMatch = True
        End If

        aika = Timer

        For i = 1 To 10
            rivi1 = findRowWithValue(10, "~jhy", 5, ActiveSheet, 1, vrivi)
            rivi2 = findRowWithValue(10, "hjih6", 5, ActiveSheet, 1, vrivi)
            rivi3 = findRowWithValue(10, "oooo", 5, ActiveSheet, 1, vrivi)
            rivi4 = findRowWithValue(10, "hvordan opfylder man lavenergi 2015", 5, ActiveSheet, 1, vrivi)
            rivi5 = findRowWithValue(10, "*jkytoi", 5, ActiveSheet, 1, vrivi)
            rivi6 = findRowWithValue(10, "????", 5, ActiveSheet, 1, vrivi)
            '            rivi7 = findRowWithValue(10, "mikael thuneberg excel spredshee?", 5, ActiveSheet, 1, vrivi)
            '            rivi8 = findRowWithValue(10, "Ì´ser snÌülast", 10000, ActiveSheet, 1, vrivi)
        Next i

        Debug.Print rivi1 & "|" & rivi2 & "|" & rivi3 & "|" & rivi4 & "|" & rivi5 & "|" & rivi6 & "|" & rivi7 & "|" & rivi8
        Debug.Print "AIKA:" & Timer - aika
    Next k

End Sub

Sub copyValues(rngSource As Range, rngTarget As Range)
    rngTarget.Resize(rngSource.Rows.Count, rngSource.Columns.Count).value = rngSource.value
End Sub
Public Function findRowWithValue(ByVal col As Long, ByVal val As Variant, ByVal prevrow As Long, ws As Worksheet, ByVal erivi As Long, Optional ByVal vrivi As Long) As Long

    Dim dataRivi As Long
    On Error Resume Next

    Dim i As Integer
    Dim replacedChars As Boolean
    Dim origLen As Integer
    Dim foundValue As Variant
    Dim searchRng As Range


    prevrow = prevrow - 50
    If prevrow < 1 Then prevrow = 1

    If IsMissing(erivi) Or erivi = 0 Then erivi = 1
    If IsMissing(vrivi) Or vrivi = 0 Then vrivi = vikarivi(ws.Cells(1, col))

    origLen = Len(val)
    val = Replace(val, "~", "~~")
    val = Replace(val, Chr(63), "~" & Chr(63))
    val = Replace(val, Chr(42), "~" & Chr(42))
    If Len(val) <> origLen Then
        replacedChars = True
    Else
        replacedChars = False
    End If
    dataRivi = 0



    With ws
        Set searchRng = .Cells(erivi, col).Resize(vrivi - erivi + 1)
        dataRivi = erivi + Application.Match(CStr(val), searchRng, 0) - 1
        If IsNumeric(val) And (IsError(dataRivi) Or dataRivi = 0) Then dataRivi = erivi + Application.Match(CDbl(val), searchRng, 0) - 1
        If IsError(dataRivi) Or dataRivi = 0 Then
            findRowWithValue = 0
        Else
            foundValue = CStr(.Cells(dataRivi, col).value)
            If Not replacedChars Then
                If foundValue = CStr(val) Then
                    findRowWithValue = dataRivi
                ElseIf LCase(foundValue) = LCase(CStr(val)) Then
                    dataRivi = findRowWithValue(col, val, dataRivi + 1, ws, dataRivi + 1, vrivi)
                    findRowWithValue = dataRivi
                Else
                    dataRivi = 0
                End If
            Else
                If Replace(Replace(Replace(CStr(foundValue), "~", "~~"), Chr(63), "~" & Chr(63)), Chr(42), "~" & Chr(42)) = CStr(val) Then
                    findRowWithValue = dataRivi
                Else
                    dataRivi = findRowWithValue(col, val, dataRivi + 1, ws, dataRivi + 1, vrivi)
                End If
            End If
        End If
    End With


End Function


Public Function findRangeName(rng As Range, Optional wb As Workbook, Optional nth As Integer = 1) As String
    On Error Resume Next
    Dim n As Name
    Dim y As Variant
    Dim namesFound As Integer
    namesFound = 0
    If IsMissing(wb) = True Or wb Is Nothing Then Set wb = ThisWorkbook
    findRangeName = vbNullString
    For Each n In wb.Names
        If InStr(1, n.RefersTo, rng.Worksheet.Name, vbTextCompare) > 0 Then
            Set y = Nothing
            Set y = Intersect(rng, Range(n.RefersTo))
            If Not y Is Nothing Then
                namesFound = namesFound + 1
                findRangeName = n.Name
                If nth = namesFound Then Exit Function
            End If
        End If
    Next
    findRangeName = vbNullString
End Function




Public Function WeekNumberAbsolute(DT As Date) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WeekNumberAbsolute
' This returns the week number of the date in DT based on Week 1 starting
' on January 1 of the year of DT, regardless of what day of week that
' might be.
' Formula equivalent:
'       =TRUNC(((DT-DATE(YEAR(DT),1,0))+6)/7)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    WeekNumberAbsolute = Int(((DT - DateSerial(Year(DT), 1, 0)) + 6) / 7)
End Function


Public Function ISOWeekNum(AnyDate As Date, Optional includeYear As Boolean = True, Optional yearOnly As Boolean = False) As String
' WhichFormat: missing or <> 2 then returns week number,
'                                = 2 then YYWW
'
    Dim ThisYear As Long
    Dim PreviousYearStart As Date
    Dim ThisYearStart As Date
    Dim NextYearStart As Date
    Dim YearNum As Long

    ThisYear = Year(AnyDate)
    ThisYearStart = YearStart(ThisYear)
    PreviousYearStart = YearStart(ThisYear - 1)
    NextYearStart = YearStart(ThisYear + 1)
    Select Case AnyDate
    Case Is >= NextYearStart
        ISOWeekNum = (AnyDate - NextYearStart) \ 7 + 1
        '    YearNum = Year(AnyDate) + 1
    Case Is < ThisYearStart
        ISOWeekNum = (AnyDate - PreviousYearStart) \ 7 + 1
        '   YearNum = Year(AnyDate) - 1
    Case Else
        ISOWeekNum = (AnyDate - ThisYearStart) \ 7 + 1
        '    YearNum = Year(AnyDate)
    End Select

    If WeekDay(AnyDate) = 1 Then
        YearNum = Year(AnyDate - 3)
    Else
        YearNum = Year(AnyDate + 5 - WeekDay(AnyDate))
    End If

    If yearOnly = True Then
        ISOWeekNum = CStr(YearNum)
    ElseIf includeYear = True Then
        ISOWeekNum = CStr(CStr(YearNum) & "|" & Format(ISOWeekNum, "00"))
    Else
        ISOWeekNum = CStr(Format(ISOWeekNum, "00"))
    End If

End Function

Public Function YearStart(WhichYear As Long) As Date

    Dim WeekDay As Long
    Dim NewYear As Date

    NewYear = DateSerial(WhichYear, 1, 1)
    WeekDay = (NewYear - 2) Mod 7    'Generate weekday index where Monday = 0

    If WeekDay < 4 Then
        YearStart = NewYear - WeekDay
    Else
        YearStart = NewYear - WeekDay + 7
    End If

End Function

Public Function getQuarter(Optional dDate As Variant = 0, Optional dMonth As Variant = 0) As Variant
    If dMonth = 0 Then
        If dDate <> 0 Then
            dMonth = Month(dDate)
        Else
            getQuarter = "Error"
        End If
    End If
    If dMonth <= 3 Then
        getQuarter = "Q1"
    ElseIf dMonth <= 6 Then
        getQuarter = "Q2"
    ElseIf dMonth <= 9 Then
        getQuarter = "Q3"
    ElseIf dMonth <= 12 Then
        getQuarter = "Q4"
    End If
End Function


Function encrypt(str As String, Optional level As Integer = 1) As String
    Dim i As Long
    Dim str2 As String
    Dim orig As Long
    Dim s As Integer
    If level = 1 Then
        s = 51
    Else
        s = Int((99 - 10 + 1) * Rnd + 10)
    End If
    For i = 1 To Len(str)
        orig = Asc(Mid(str, i, 1)) - 33
        If orig >= 0 And orig <= 93 Then
            str2 = str2 & Chr(33 + ((orig + s) Mod 93))
        Else
            str2 = str2 & Chr(orig)
        End If
    Next i
    If level = 1 Then
        encrypt = "CH1_" & str2
    Else
        encrypt = "CH" & level & "_" & genRandomString(5) & s & str2 & genRandomString(4)
    End If
End Function
Function decrypt(str As String) As String
    Dim i As Long
    Dim str2 As String
    Dim orig As Long
    Dim level As Integer
    Dim s As Integer
    level = Mid(str, 3, 1)
    If level = 1 Then
        str = Replace(str, "CH1_", "")
        s = 51
    Else
        str = Replace(str, "CH" & level & "_", "")
        str = Right(str, Len(str) - 5)
        str = Left(str, Len(str) - 4)
        s = Left(str, 2)
        str = Right(str, Len(str) - 2)
    End If
    For i = 1 To Len(str)
        orig = Asc(Mid(str, i, 1)) - 33
        If orig >= 0 And orig <= 93 Then
            str2 = str2 & Chr(33 + truemod((orig - s), 93))
        Else
            str2 = str2 & Chr(orig)
        End If
    Next i
    decrypt = str2
End Function
Function truemod(num As Integer, modby As Integer) As Integer
    truemod = (modby + (num Mod modby)) Mod modby
End Function


Sub pArr(arr As Variant, Optional maxStrLength As Long = 100)
    Call printArr(arr, maxStrLength)
End Sub
Sub printArr(arr As Variant, Optional maxStrLength As Long = 100)
    On Error Resume Next
    Dim i As Long
    Dim k As Long
    Dim str As String
    Dim dimensions As Integer

    If maxStrLength = 0 Then maxStrLength = 100

    dimensions = NumberOfDimensions(arr)

    If dimensions = 0 Then Exit Sub
    If dimensions = 1 Then
        Debug.Print "Printing arr, " & dimensions & " dimensions, " & LBound(arr) & " to " & UBound(arr)
        For i = LBound(arr) To UBound(arr)
            If IsArray(arr(i)) Then
                Debug.Print i & ":: ARR:"
                pArr (arr(i))
            Else
                Debug.Print i & ":: " & Left(arr(i), maxStrLength)
            End If
        Next i
    Else
        Debug.Print "Printing arr, " & dimensions & " dimensions, " & LBound(arr, 1) & " to " & UBound(arr, 1) & " * " & LBound(arr, 2) & " to " & UBound(arr, 2)
        For i = LBound(arr, 1) To UBound(arr, 1)
            str = ""
            For k = LBound(arr, 2) To UBound(arr, 2)
                str = str & " " & k & ": " & Left(CStr(arr(i, k)), maxStrLength) & "  "
            Next k

            Debug.Print i & ":: " & str
        Next i
    End If
End Sub

Function genRandomString(Optional length As Integer = 50)
    Dim str As String
    Dim i As Integer

    For i = 1 To length
        If i Mod 2 = 0 Then
            str = Chr(Int((90 - 65 + 1) * Rnd + 65)) & str
        Else
            str = Int((9 * Rnd) + 1) & str
        End If
    Next i
    genRandomString = str
End Function

Function NumberOfDimensions(arr As Variant) As Integer
    Dim intDim As Integer
    Dim DimNum As Integer
    Dim ErrorCheck As Boolean

    On Error GoTo endEx
    For DimNum = 1 To 5
        ErrorCheck = LBound(arr, DimNum)
    Next DimNum
endEx:
    NumberOfDimensions = DimNum - 1
End Function

Public Function testNumberOfCharsThatCanBeReturnedToCell() As Long
    On Error Resume Next
    testNumberOfCharsThatCanBeReturnedToCell = 255
    Dim val As Variant
    val = Range("numberOfCharsThatCanBeReturnedToCell").value
    If val <> "" And val > 0 And IsNumeric(val) Then
        testNumberOfCharsThatCanBeReturnedToCell = val
    Else
        testNumberOfCharsThatCanBeReturnedToCell = 255
    End If
End Function

'    On Error Resume Next
'    Dim max As Long
'    Dim str As String
'    Dim i As Long
'    Dim iteration As Integer
'    Dim val As Variant
'    val = Range("numberOfCharsThatCanBeReturnedToCell").value
'
'    If val <> "" And val > 0 Then
'        numberOfCharsThatCanBeReturnedToCell = val
'        Exit Sub
'    End If
'
''
'    For iteration = 1 To 2
'        If iteration = 1 Then
'            max = 261
'        ElseIf iteration = 2 Then
'            max = 517
'            '        Else
'            '            max = 1029
'        End If
'        str = ""
'        For i = 1 To max
'            str = str & "i"
'        Next i
'        Range("testCell").value = str
'        If Range("testCell").value <> str Then
'            numberOfCharsThatCanBeReturnedToCell = max - 6
'            Range("numberOfCharsThatCanBeReturnedToCell").value = numberOfCharsThatCanBeReturnedToCell
'            Exit Sub
'        End If
'    Next iteration
'    numberOfCharsThatCanBeReturnedToCell = 517
'    Range("numberOfCharsThatCanBeReturnedToCell").value = numberOfCharsThatCanBeReturnedToCell


Public Function getToday() As Date
'non-volatile today function
    getToday = Date
End Function
Public Function getNow() As Date
'non-volatile today function
    getNow = Now
End Function
Public Function arrayReplaceRSCL(arr As Variant, Optional untilColumn As Integer = -1) As Variant
    Dim rivi As Long
    Dim sar As Long
    If untilColumn = -1 Then untilColumn = UBound(arr, 2)
    If untilColumn <= 0 Then
        arrayReplaceRSCL = arr
        Exit Function
    End If
    For rivi = LBound(arr) To UBound(arr)
        For sar = LBound(arr, 2) To untilColumn
            arr(rivi, sar) = replaceRSCL(arr(rivi, sar))
        Next sar
    Next rivi
    arrayReplaceRSCL = arr
End Function

Function replaceRSCL(str As Variant) As String
    str = Replace(str, "%rscL1%", rscL1)
    str = Replace(str, "%rscL2%", rscL2)
    str = Replace(str, "%rscL3%", rscL3)
    replaceRSCL = str
End Function


