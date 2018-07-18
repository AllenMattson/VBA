Attribute VB_Name = "Module1"
Option Explicit

Sub ShowWaterMark()
    With Worksheets("Sheet1").PageSetup.LeftHeaderPicture
        .Filename = "C:\Excel2013_HandsOn\cd.bmp"
        .Height = 75
        .Width = 75
        .Brightness = 0.25
        .ColorType = msoPictureWatermark
        .Contrast = 0.45
    End With
    ' Display the picture in the center header.
    ActiveSheet.PageSetup.LeftHeader = "&G"
End Sub

Sub ShowPageSettings()
    With ActiveSheet.PageSetup
        Debug.Print "Orientation="; .Orientation
        Debug.Print "Paper Size = "; .PaperSize
        Debug.Print "Print Gridlines = "; .PrintGridlines
        Debug.Print "Horizontal Print Quality = "; .PrintQuality(1)
        Debug.Print "Print Area = "; .PrintArea
    End With
End Sub

Sub ShowPageSettings_Modified()
    With ActiveSheet.PageSetup
        Debug.Print "Orientation="; .Orientation
        Debug.Print "Paper Size = "; .PaperSize
        Debug.Print "Print Gridlines = "; .PrintGridlines
        Debug.Print "Horizontal Print Quality = "; .PrintQuality(1)
        Cells(1, 1).Select
          .PrintArea = ActiveCell.CurrentRegion.Address
        Debug.Print "Print Area = "; .PrintArea;
        .CenterHeader = Chr(10) & "Bonus Information Sheet"
    End With
    Application.Dialogs(xlDialogPrintPreview).Show
End Sub

Sub FormatSheet()
    Dim curReg As Range
    Set curReg = ActiveCell.CurrentRegion

    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$1"
        Cells(1, 1).Select

        .PrintArea = curReg.Offset(1, 0). _
            Resize(curReg.Rows.Count - 1, _
            curReg.Columns.Count).Address

        Debug.Print "Print Area = "; .PrintArea;
        .CenterHeader = Chr(10) & "Bonus Information Sheet"
        .PrintGridlines = True
    End With
    Application.Dialogs(xlDialogPrintPreview).Show
End Sub



