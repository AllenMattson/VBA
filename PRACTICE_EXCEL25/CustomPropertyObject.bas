Attribute VB_Name = "CustomPropertyObject"
Option Explicit

Sub StoreScores()
    Dim mySheet As Worksheet
    Dim custPrp As CustomProperty
    Dim i As Integer
    Dim rng As Range
    Dim totalCount As Integer

    Set mySheet = ThisWorkbook.Sheets(1)

    ' find out if custom properties exist
    If mySheet.CustomProperties.Count > 0 Then
        ' Display custom properties
        totalCount = mySheet.CustomProperties.Count

        For i = 1 To totalCount
            With mySheet.CustomProperties(1)
                Debug.Print .Name & vbTab; .Value
                Set rng = mySheet.Range("A:A").Find(what:=.Name)
                ' Delete the custom property
                 If Not rng Is Nothing Then .Delete
            End With
        Next
    End If

    mySheet.Activate
    Cells(2, 1).Select
    Do While ActiveCell <> ""
        If Not IsEmpty(ActiveCell) Then
            Set custPrp = mySheet.CustomProperties.Add( _
            Name:=ActiveCell.Text, _
            Value:=ActiveCell.Offset(0, 1).Text)
            Debug.Print custPrp.Name & vbTab & custPrp.Value
            ActiveCell.Offset(1, 0).Select
        End If
    Loop

    If mySheet.CustomProperties.Count > 0 Then
        ' Display custom properties
        For i = 1 To mySheet.CustomProperties.Count
            With mySheet.CustomProperties(i)
                Debug.Print .Name & vbTab; .Value
            End With
        Next
    End If
End Sub



