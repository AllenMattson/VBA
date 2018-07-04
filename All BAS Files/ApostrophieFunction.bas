Attribute VB_Name = "ApostrophieFunction"


'**************************************
'Windows API/Global Declarations for :Ap
'     rostrophe
'**************************************
'None
        
'**************************************
' Name: Aprostrophe
' Description:Have you ever try so send
'     a string variable to MS Access that have
'
'apostrophes using a SQL Statement? If YES you will get a run time ERROR
'Here is your solution....A function that formats the
'variable before sending it to the database.
' By: Gaetan Savoie (from psc cd)
'
'
' Inputs:sFieldString
'
' Returns:Aphostrophe
'
'Assumes:This code should be used in you
'     r Classes.


'For example :
    'let say myVar=" Gaetan's"
    'the follwing statement will give you errors:
    'SSQL = "INSERT INTO tablename (FirstName) VALUES (" & Chr(39) & myvar & Chr(39) & ")"
    'To fix it do the following:
    'myvar = Apostrophe(myvar)
    'SSQL = "INSERT INTO tablename (FirstName) VALUES (" & Chr(39) & myvar & Chr(39) & ")"
'
'Side Effects:None
'**************************************

'***************************************
'     ********************************
' Function: Apostrophe
' Argument: sFieldString
' Description: This subroutine will fill


'     format the field we
    ' want to store in the database if there
    '     is some apostrophes
    ' in the field.
    '***************************************
    '     ********************************


Public Function Apostrophe(sFieldString As String) As String


    If InStr(sFieldString, "'") Then
        Dim iLen As Integer
        Dim ii As Integer
        Dim apostr As Integer
        iLen = Len(sFieldString)
        ii = 1


        Do While ii <= iLen


            If Mid(sFieldString, ii, 1) = "'" Then
                apostr = ii
                sFieldString = Left(sFieldString, apostr) & "'" & _
                Right(sFieldString, iLen - apostr)
                iLen = Len(sFieldString)
                ii = ii + 1
            End If
            ii = ii + 1
        Loop
    End If
    Apostrophe = sFieldString
End Function
        


