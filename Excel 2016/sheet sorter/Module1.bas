Attribute VB_Name = "Module1"
Option Explicit

Sub SortSheets()
Attribute SortSheets.VB_ProcData.VB_Invoke_Func = "S\n14"
'   This routine sorts the sheets of the
'   active workbook in ascending order.
'   Use Ctrl+Shift+S to execute

    Dim SheetNames() As String
    Dim i As Long
    Dim SheetCount As Long
    Dim OldActive As Object
    
    If ActiveWorkbook Is Nothing Then Exit Sub ' No active workbook
    SheetCount = ActiveWorkbook.Sheets.Count
    
'   Check for protected workbook structure
    If ActiveWorkbook.ProtectStructure Then
        MsgBox ActiveWorkbook.Name & " is protected.", _
           vbCritical, "Cannot Sort Sheets."
        Exit Sub
    End If

'   Make user verify
    If MsgBox("Sort the sheets in the active workbook?", _
      vbQuestion + vbYesNo) <> vbYes Then Exit Sub

'   Disable Ctrl+Break
    Application.EnableCancelKey = xlDisabled
       
'   Get the number of sheets
    SheetCount = ActiveWorkbook.Sheets.Count
    
'   Redimension the array
    ReDim SheetNames(1 To SheetCount)

'   Store a reference to the active sheet
    Set OldActive = ActiveSheet
   
'   Fill array with sheet names
    For i = 1 To SheetCount
        SheetNames(i) = ActiveWorkbook.Sheets(i).Name
    Next i
   
'   Sort the array in ascending order
    Call BubbleSort(SheetNames)
   
'   Turn off screen updating
    Application.ScreenUpdating = False
    
'   Move the sheets
    For i = 1 To SheetCount
        ActiveWorkbook.Sheets(SheetNames(i)).Move _
            Before:=ActiveWorkbook.Sheets(i)
    Next i

'   Reactivate the original active sheet
    OldActive.Activate
    
End Sub


Sub BubbleSort(List() As String)
    Dim First As Long, Last As Long
    Dim i As Long, j As Long
    Dim Temp As String
    First = LBound(List)
    Last = UBound(List)
    For i = First To Last - 1
        For j = i + 1 To Last
            If UCase(List(i)) > UCase(List(j)) Then
                Temp = List(j)
                List(j) = List(i)
                List(i) = Temp
            End If
        Next j
    Next i
End Sub

