Attribute VB_Name = "Module1"
Option Explicit

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetProfileStringA Lib "kernel32" _
        (ByVal lpAppName As String, _
         ByVal lpKeyName As String, _
         ByVal lpDefault As String, _
         ByVal lpReturnedString As _
         String, ByVal nSize As Long) As Long
#Else
    Private Declare Function GetProfileStringA Lib "kernel32" _
        (ByVal lpAppName As String, _
         ByVal lpKeyName As String, _
         ByVal lpDefault As String, _
         ByVal lpReturnedString As _
         String, ByVal nSize As Long) As Long
#End If

Sub DefaultPrinterInfo()
    Dim strLPT As String * 255
    Dim Result As String
    Dim ResultLength As Integer
    Dim Comma1 As Integer
    Dim Comma2 As Integer
    Dim Printer As String
    Dim Driver As String
    Dim Port As String
    Dim Msg As String
    
    Call GetProfileStringA _
       ("Windows", "Device", "", strLPT, 254)
    
    Result = Application.Trim(strLPT)
    ResultLength = Len(Result)

    Comma1 = InStr(1, Result, ",", 1)
    Comma2 = InStr(Comma1 + 1, Result, ",", 1)

'   Gets printer's name
    Printer = Left(Result, Comma1 - 1)

'   Gets driver
    Driver = Mid(Result, Comma1 + 1, Comma2 - Comma1 - 1)

'   Gets last part of device line
    Port = Right(Result, ResultLength - Comma2)

'   Build message
    Msg = "Printer:" & Chr(9) & Printer & Chr(13)
    Msg = Msg & "Driver:" & Chr(9) & Driver & Chr(13)
    Msg = Msg & "Port:" & Chr(9) & Port

'   Display message
    MsgBox Msg, vbInformation, "Default Printer Information"
End Sub


