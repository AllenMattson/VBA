Attribute VB_Name = "IsArray_Function"
Option Explicit

Sub IsThisArray()
  ' declare a dynamic array
  Dim sheetNames() As String
      Dim totalSheets As Integer
  Dim counter As Integer

  ' count the sheets in the current workbook
  totalSheets = ActiveWorkbook.Sheets.Count

  ' specify the size of the array
  ReDim sheetNames(1 To totalSheets)

  ' enter and show the names of sheets
  For counter = 1 To totalSheets
    sheetNames(counter) = ActiveWorkbook.Sheets(counter).Name
    MsgBox sheetNames(counter)
  Next counter

  ' check if this is indeed an array
  If IsArray(sheetNames) Then
    MsgBox "The sheetNames variable is an array."
  End If
End Sub

