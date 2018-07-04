Attribute VB_Name = "MatchText_FuzzyLookup"
Option Explicit


Public Function FuzzyVLookup(Lookup_Value As String, _
                             Table_Array As Variant, _
                             Optional Col_Index_Num As Integer = 1, _
                             Optional Compare As VbCompareMethod = vbTextCompare _
                            ) As Variant

' Find the best match for a given string in column 1 of an array of data obtained from an Excel range
' This is functionally similar to VLookup, but it returns the best match, not the first exact match
' This function is not case-sensitive, unless you specify 'Compare' as 0 or vbBinaryCompare

' If your data quality is poor, you are advised to display the retrieved index value from column 1
' and use the FuzzyMatchScore() function on this index value to reveal the fuzzy-matching 'score' and
' discard all results below a threshold value. Feel free to code up a 'threshold' parameter!

' If you are looking up names and addresses, use the NormaliseAddress() function on your search term and
' searched population to standardise abbreviations and word-order conventions used in British addresses.

' THIS CODE IS IN THE PUBLIC DOMAIN

Application.Volatile False

Dim dblBestMatch As Double

Dim iRowBest    As Integer
Dim dblMatch    As Double
Dim iRow        As Integer
Dim strTest     As String
Dim strInput    As String

Dim iStartCol   As Integer
Dim iEndCol     As Integer
Dim iOffset     As Integer

If InStr(TypeName(Table_Array), "(") + InStr(1, TypeName(Table_Array), "Range", vbTextCompare) < 1 Then
     'Table_Array is not an array
     FuzzyVLookup = "#VALUE"
    Exit Function
End If

If InStr(1, TypeName(Table_Array), "Range", vbTextCompare) > 0 Then
    Table_Array = Table_Array.Value
End If

' If you get a subscript-out-of-bounds error here, you're using a vector instead
' of the 2-dimensional array that is the default 'Value' property of an Excel range.

iStartCol = LBound(Table_Array, 2)
iEndCol = UBound(Table_Array, 2)
iOffset = 1 - iStartCol


Col_Index_Num = Col_Index_Num - iOffset

If Col_Index_Num > iEndCol Or Col_Index_Num < iStartCol Then
     'Out-of-bounds
     FuzzyVLookup = "#VALUE"
    Exit Function
End If



    strInput = UCase(Lookup_Value)

    iRowBest = -1
    dblBestMatch = 0

    For iRow = LBound(Table_Array, 1) To UBound(Table_Array, 1)

        strTest = ""
        strTest = Table_Array(iRow, iStartCol)

        dblMatch = 0
        dblMatch = FuzzyMatchScore(strInput, strTest, Compare)

        If dblMatch = 1 Then ' Bail out on finding an exact match
            iRowBest = iRow
            Exit For
        End If

        If dblMatch > dblBestMatch Then
            dblBestMatch = dblMatch
            iRowBest = iRow
        End If

    Next iRow


    If iRowBest = -1 Then
        FuzzyVLookup = "#NO MATCH"
        Exit Function
    End If

    FuzzyVLookup = Table_Array(iRowBest, Col_Index_Num)

End Function
   
Public Function FuzzyHLookup(Lookup_Value As String, _
                             Table_Array As Variant, _
                             Optional Row_Index_Num As Integer = 1, _
                             Optional Compare As VbCompareMethod = vbTextCompare)

' Find the best match for a given string in Row 1 of an array of data obtained from an Excel range
' This is functionally similar to HLookup, but it returns the best match, not the first exact match
' This function is not case-sensitive, unless you specify 'Compare' as vbTextBinary.

' If your data quality is poor, you are advised to display the retrieved index value from row 1
' and use the FuzzyMatchScore() function on this index value to reveal the fuzzy-matching 'score' and
' discard all results below a threshold value. Feel free to code up a 'threshold' parameter!

' If you are looking up names and addresses, use the NormaliseAddress() function on your search term and
' searched population to standardise abbreviations and word-order conventions used in British addresses.

' THIS CODE IS IN THE PUBLIC DOMAIN

Application.Volatile False

Dim dblBestMatch As Double

Dim iColBest    As Integer
Dim dblMatch    As Double
Dim iCol        As Integer
Dim strTest     As String
Dim strInput    As String

Dim iStartRow   As Integer
Dim iEndRow     As Integer
Dim iOffset     As Integer

If InStr(TypeName(Table_Array), "(") + InStr(1, TypeName(Table_Array), "Range", vbTextCompare) < 1 Then
     'Table_Array is not an array
     FuzzyHLookup = "#VALUE"
    Exit Function
End If

If InStr(1, TypeName(Table_Array), "Range", vbTextCompare) > 0 Then
    Table_Array = Table_Array.Value
End If

' If you get a subscript-out-of-bounds error here, you're using a vector instead
' of the 2-dimensional array that is the default 'Value' property of an Excel range.

iStartRow = LBound(Table_Array, 1)
iEndRow = UBound(Table_Array, 1)
iOffset = 1 - iStartRow


Row_Index_Num = Row_Index_Num - iOffset

If Row_Index_Num > iEndRow Or Row_Index_Num < iStartRow Then
     'Out-of-bounds
     FuzzyHLookup = "#VALUE"
    Exit Function
End If


    strInput = UCase(Lookup_Value)

    iColBest = -1
    dblBestMatch = 0

    For iCol = LBound(Table_Array, 2) To UBound(Table_Array, 2)

        strTest = ""
        strTest = Table_Array(iStartRow, iCol)

        dblMatch = 0
        dblMatch = FuzzyMatchScore(strInput, strTest, Compare)

        If dblMatch = 1 Then ' Bail out on finding an exact match
            iColBest = iCol
            Exit For
        End If

        If dblMatch > dblBestMatch Then
            dblBestMatch = dblMatch
            iColBest = iCol
        End If

    Next iCol


    If iColBest = -1 Then
        FuzzyHLookup = "#NO MATCH"
        Exit Function
    End If

    FuzzyHLookup = Table_Array(Row_Index_Num, iColBest)

End Function


Public Function FuzzyMatchScore(ByVal str1 As String, _
                                ByVal str2 As String, _
                                Optional Compare As VbCompareMethod = vbTextCompare _
                                ) As Double

' Returns an estimate of how closely word 1 matches word 2: this is best displayed as a percentage
' This is calculated as the fraction of the longer string that is made up of recognisable fragments of the shorter string
' There is no support for wildcards and regular expressions. Case-sensitivity is determined by the 'compare' parameter

' THIS CODE IS IN THE PUBLIC DOMAIN

Application.Volatile False

Dim maxLen As Integer
Dim minLen As Integer

    If str1 = str2 Then
        FuzzyMatchScore = 1#
        Exit Function
    End If

    If Len(str1) > Len(str2) Then
        maxLen = Len(str1)
        minLen = Len(str2)
    Else
        maxLen = Len(str2)
        minLen = Len(str1)
    End If

    If Len(str1) = 0 Or Len(str2) = 0 Then
        FuzzyMatchScore = 0#
    Else

        FuzzyMatchScore = 0#
        FuzzyMatchScore = SumOfCommonStrings(str1, str2, Compare) / maxLen

    End If
   
End Function

Public Function SumOfCommonStrings( _
                            ByVal s1 As String, _
                            ByVal s2 As String, _
                            Optional Compare As VBA.VbCompareMethod = vbTextCompare, _
                            Optional iScore As Integer = 0 _
                                ) As Integer

Application.Volatile False

' N.Heffernan 06 June 2006 (somewhere over Newfoundland)
' THIS CODE IS IN THE PUBLIC DOMAIN


' Function to measure how much of String 1 is made up of substrings found in String 2

' This function uses a modified Longest Common String algorithm.
' Simple LCS algorithms are unduly sensitive to single-letter
' deletions/changes near the midpoint of the test words, eg:
' Wednesday is obviously closer to WedXesday on an edit-distance
' basis than it is to WednesXXX. So it would be better to score
' the 'Wed' as well as the 'esday' and add up the total matched

' Watch out for strings of differing lengths:
'
'    SumOfCommonStrings("Wednesday", "WednesXXXday")
'
' This scores the same as:
'
'     SumOfCommonStrings("Wednesday", "Wednesday")
'
' So make sure the calling function uses the length of the longest
' string when calculating the degree of similarity from this score.


' This is coded for clarity, not for performance.

Dim arr() As Integer    ' Scoring matrix
Dim n As Integer        ' length of s1
Dim m As Integer        ' length of s2
Dim i As Integer        ' start position in s1
Dim j As Integer        ' start position in s2
Dim subs1 As String     ' a substring of s1
Dim len1 As Integer     ' length of subs1

Dim sBefore1            ' documented in the code
Dim sBefore2
Dim sAfter1
Dim sAfter2

Dim s3 As String


SumOfCommonStrings = iScore

n = Len(s1)
m = Len(s2)

If s1 = s2 Then
    SumOfCommonStrings = n
    Exit Function
End If

If n = 0 Or m = 0 Then
    Exit Function
End If

's1 should always be the shorter of the two strings:
If n > m Then
    s3 = s2
    s2 = s1
    s1 = s3
    n = Len(s1)
    m = Len(s2)
End If

n = Len(s1)
m = Len(s2)

' Special case: s1 is n exact substring of s2
If InStr(1, s2, s1, Compare) Then
    SumOfCommonStrings = n
    Exit Function
End If

For len1 = n To 1 Step -1

    For i = 1 To n - len1 + 1

        subs1 = Mid(s1, i, len1)
        j = 0
        j = InStr(1, s2, subs1, Compare)
       
        If j > 0 Then
       
            ' We've found a matching substring...
            iScore = iScore + len1

          ' Now clip out this substring from s1 and s2...
          ' And search the fragments before and after this excision:

       
            If i > 1 And j > 1 Then
                sBefore1 = Left(s1, i - 1)
                sBefore2 = Left(s2, j - 1)
                iScore = SumOfCommonStrings(sBefore1, _
                                            sBefore2, _
                                            Compare, _
                                            iScore)
            End If
   
   
            If i + len1 < n And j + len1 < m Then
                sAfter1 = Right(s1, n + 1 - i - len1)
                sAfter2 = Right(s2, m + 1 - j - len1)
                iScore = SumOfCommonStrings(sAfter1, _
                                            sAfter2, _
                                            Compare, _
                                            iScore)
            End If
   
   
            SumOfCommonStrings = iScore
            Exit Function

        End If

    Next


Next


End Function


Private Function Minimum(ByVal a As Integer, _
                         ByVal b As Integer, _
                         ByVal c As Integer) As Integer
Dim min As Integer

  min = a

  If b < min Then
        min = b
  End If

  If c < min Then
        min = c
  End If

  Minimum = min

End Function



Public Function NormaliseAddress(ByVal strAddress As String) As String
Application.Volatile False
' This function is intended to remove or standardise common phrases
' and abbreviations used in British postal addresses, allowing the use
' of string-comparison algorithms in lists of names and addresses.

' Developers in other countries should review the word list used here,
' as conventions probably differ in your local language or dialect.

strAddress = " " & UCase(strAddress) & " "

strAddress = Substitute(strAddress, ",", " ")
strAddress = Substitute(strAddress, ".", " ")
strAddress = Substitute(strAddress, "-", " ")
strAddress = Substitute(strAddress, vbCrLf, " ")
strAddress = Substitute(strAddress, " BLVD ", " BOULEVARD ")
strAddress = Substitute(strAddress, " BVD ", " BOULEVARD ")
strAddress = Substitute(strAddress, " AV ", " AVENUE ")
strAddress = Substitute(strAddress, " AVE ", " AVENUE ")
strAddress = Substitute(strAddress, " RD ", " ROAD ")
strAddress = Substitute(strAddress, " WY ", " WAY ")
strAddress = Substitute(strAddress, " EST ", " ESTATE ")
strAddress = Substitute(strAddress, " PL ", " PLACE ")
strAddress = Substitute(strAddress, " PK ", " PARK ")
strAddress = Substitute(strAddress, " HSE ", " HOUSE ")
strAddress = Substitute(strAddress, " H0 ", " HOUSE ")
strAddress = Substitute(strAddress, " GDNS ", " GARDENS ")

strAddress = Substitute(strAddress, "&", "AND")
strAddress = Substitute(strAddress, " LIMITED ", " LTD ")
strAddress = Substitute(strAddress, " COMPANY ", " CO ")
strAddress = Substitute(strAddress, " CORPORATION ", " CORP ")
strAddress = Substitute(strAddress, " T/A ", " TA ")
strAddress = Substitute(strAddress, " TRADING AS ", " TA ")

' Common personal titles: these are often applied inconsistently or
' omitted, and must therefore be removed. Specific applications may
' require additional titles and their abbreviations - military rank,
' academic titles and degrees, courtesy titles of the aristocracy,
' knighthoods and honours (particularly for lists of civil servants)

strAddress = Substitute(strAddress, " ESQ ", " ")
strAddress = Substitute(strAddress, " MR ", " ")
strAddress = Substitute(strAddress, " MRS ", " ")
strAddress = Substitute(strAddress, " MISS ", " ")
strAddress = Substitute(strAddress, " MS ", " ")
strAddress = Substitute(strAddress, " MESSRS ", " ")
strAddress = Substitute(strAddress, " SIR ", " ")
strAddress = Substitute(strAddress, " OF ", " ")
strAddress = Substitute(strAddress, " DR ", " ")
strAddress = Substitute(strAddress, " OR ", " ")
strAddress = Substitute(strAddress, " IN ", " ")
strAddress = Substitute(strAddress, " THE ", " ")
strAddress = Substitute(strAddress, " REVEREND ", " REV ")
strAddress = Substitute(strAddress, " REVERENT ", " REV ")
strAddress = Substitute(strAddress, " HONOURABLE ", " HON ")
strAddress = Substitute(strAddress, " BROS ", " BROTHERS ")
strAddress = Substitute(strAddress, " ASSOC ", " ASSOCIATION ")
strAddress = Substitute(strAddress, " ASSN ", " ASSOCIATION ")

' Standardising 'St.', 'St', and 'Street'. Note that there are over 40 English
' towns and place names that contain or consist entirely of the word 'Street'.
' In addition, 'St' is a common abbreviation for 'Saint' in addresses.

' I have never seen a list of addresses where 'Street' and 'St' were used in a
' consistent way, and the only workable solution is to delete them all:

strAddress = Substitute(strAddress, " STREET ", " ")
strAddress = Substitute(strAddress, " ST ", " ")
strAddress = Substitute(strAddress, " STR ", " ")

Do While InStr(strAddress, "  ") > 0
    strAddress = Substitute(strAddress, "  ", " ")
Loop

strAddress = Trim(strAddress)

NormaliseAddress = strAddress

End Function


Public Function StripChars(myString As String, ParamArray Exceptions()) As String

' Strip out all non-alphanumeric characters from a string in a single pass
' Exceptions parameters allow you to retain specific characters (eg: spaces)

' THIS CODE IS IN THE PUBLIC DOMAIN

Application.Volatile False

Dim i As Integer
Dim iLen As Integer
Dim chrA As String * 1
Dim intA As Integer
Dim j As Integer
Dim iStart As Integer
Dim iEnd As Integer

If Not IsEmpty(Exceptions()) Then
    iStart = LBound(Exceptions)
    iEnd = UBound(Exceptions)
End If

iLen = Len(myString)

For i = 1 To iLen
    chrA = Mid(myString, i, 1)
    intA = Asc(chrA)
    Select Case intA
    Case 48 To 57, 65 To 90, 97 To 122
        StripChars = StripChars & chrA
    Case Else
        If Not IsEmpty(Exceptions()) Then
            For j = iStart To iEnd
                If chrA = Exceptions(j) Then
                    StripChars = StripChars & chrA
                    Exit For ' j
                End If
            Next j
        End If
    End Select
Next i



End Function


Private Function Substitute(ByVal Text As String, _
                            ByVal Old_Text As String, _
                            ByVal New_Text As String, _
                            Optional Instance As Long = 0, _
                            Optional Compare As VbCompareMethod = vbTextCompare _
                            ) As String


' Replace all instances (or the nth instance ) of 'Old' text with 'New'
' Unlike VB.Mid$ this method is not sensitive to length and can replace ALL instances
' This is not exposed as a Public function because there is an Excel Worksheet function
' called Substitute(). However, Workheet Functions have length constraints.

' THIS CODE IS IN THE PUBLIC DOMAIN

Dim iStart As Long
Dim iEnd As Long
Dim iLen As Long
Dim iInstance As Long
Dim strOut As String

iLen = Len(Old_Text)

If iLen = 0 Then
    Substitute = Text
    Exit Function
End If

iEnd = 0
iStart = 1

iEnd = InStr(iStart, Text, Old_Text, Compare)

If iEnd = 0 Then
    Substitute = Text
    Exit Function
End If


strOut = ""

Do Until iEnd = 0

    strOut = strOut & Mid$(Text, iStart, iEnd - iStart)
    iInstance = iInstance + 1

    If Instance = 0 Or Instance = iInstance Then
        strOut = strOut & New_Text
    Else
        strOut = strOut & Mid$(Text, iEnd, Len(Old_Text))
    End If

    iStart = iEnd + iLen
    iEnd = InStr(iStart, Text, Old_Text, Compare)

Loop

iLen = Len(Text)
strOut = strOut & Mid$(Text, iStart, iLen - iEnd)

Substitute = strOut

End Function



