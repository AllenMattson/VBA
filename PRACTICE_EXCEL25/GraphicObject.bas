Attribute VB_Name = "GraphicObject"
Option Explicit

Sub AddWatermarkImage()
    Dim strFilename As String

    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Custom image selection"
        .AllowMultiSelect = False
        .Filters.Add "Pictures", "*.gif; *.jpg; *.jpeg; *.bmp", 1
        .InitialView = msoFileDialogViewThumbnail
        If .Show = -1 Then
            strFilename = .SelectedItems(1)
            With ActiveSheet.PageSetup
                With .LeftHeaderPicture
                    .Filename = strFilename
                    .Brightness = 0.85
                    .ColorType = msoPictureWatermark
                    .Contrast = 0.15
                    .Height = 72
                    .Width = 72
                End With
                .TopMargin = Application.InchesToPoints(1.25)
                .LeftHeader = "&G"
            End With
        End If
    End With
End Sub


