Attribute VB_Name = "Module5"
Option Explicit

Sub AddAStyle()
    Dim newStyleName As String
    Dim curStyle As Variant
    Dim i As Integer
    
    newStyleName = "SimpleFormat"
    i = 0
    
    For Each curStyle In ActiveWorkbook.Styles
        i = i + 1
        If curStyle.Name = newStyleName Then
            MsgBox "This style " & "(" & newStyleName & _
               ") already exists. " & Chr(13) & _
               "It's the " & i & " style in the Styles collection."
            Exit Sub
        End If
    Next
    
    With ActiveWorkbook.Styles.Add(newStyleName)
        .Font.Name = "Arial Narrow"
        .Font.Size = "12"
        .Borders.LineStyle = xlThin
        .NumberFormat = "$#,##0_);[Red]($#,##0)"
        .IncludeAlignment = False
    End With
End Sub

Sub GetStyleNames()
    Dim i As Integer

    For i = 1 To ActiveWorkbook.Styles.count
        Debug.Print "Style " & i & ":" & _
         ActiveWorkbook.Styles(i).Name
    Next i
End Sub

Sub AddSelectionStyle()
    Dim newStyleName As String
     
    newStyleName = "InvoiceAmount"
    ActiveWorkbook.Styles.Add Name:=newStyleName, _
                BasedOn:=ActiveCell
End Sub


