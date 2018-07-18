Attribute VB_Name = "Module1"
Sub ShowChartAsGrayscale()
Attribute ShowChartAsGrayscale.VB_ProcData.VB_Invoke_Func = " \n14"
'   Copies the active chart as a grayscale picture
'   Embedded charts only
    If ActiveChart Is Nothing Then
        MsgBox "Select a chart."
        Exit Sub
    End If
    ActiveChart.Parent.CopyPicture
    ActiveChart.Parent.TopLeftCell.Select
    ActiveSheet.Pictures.Paste
    ActiveSheet.Pictures(ActiveSheet.Pictures.Count). _
      ShapeRange.PictureFormat.ColorType = msoPictureGrayscale
End Sub
