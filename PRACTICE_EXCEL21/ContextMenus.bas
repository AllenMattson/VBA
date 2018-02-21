Attribute VB_Name = "ContextMenus"
Sub ContextMenus()
  Dim myBar As CommandBar
  Dim counter As Integer

  For Each myBar In CommandBars
    If myBar.Type = msoBarTypePopup Then
      counter = counter + 1
      Debug.Print counter & ": " & myBar.Name
    End If
  Next
End Sub


Sub AddToCellMenu()
  With Application.CommandBars("Cell")
    .Reset
    .Controls.Add(Type:=msoControlButton, _
        Before:=2).Caption = "Insert Picture..."
    .Controls("Insert Picture...").OnAction = "InsertPicture"
  End With
End Sub

Sub InsertPicture()
  CommandBars.ExecuteMso ("PictureInsertFromFile")
End Sub


Sub AddToCellMenu2()
  Dim ct As CommandBarButton

  With Application.CommandBars("Cell")
      .Reset
      Set ct = .Controls.Add(Type:=msoControlButton, _
          Before:=11, Temporary:=True)
  End With
  With ct
      .Caption = "Insert Picture..."
      .OnAction = "InsertPicture"
      .Picture = Application.CommandBars. _
          GetImageMso("PictureInsertFromFile", 16, 16)
      .Style = msoButtonIconAndCaption

  End With
End Sub

Sub DeleteInsertPicture()
      Dim c As CommandBarControl
      On Error Resume Next
      Set c = CommandBars("Cell").Controls("Insert Pict&ure...")
      c.Delete
End Sub


Sub Show_ShortMenu()
  Dim shortMenu As Object

  Set shortMenu = Application.CommandBars("MyComputer")
  With shortMenu
     .ShowPopup
  End With
End Sub


Sub Delete_ShortMenu()
   Application.CommandBars("MyComputer").Delete
End Sub


Sub Images()
    Dim i As Integer
    Dim j As Integer
    Dim total As Integer
    Dim buttonId As Integer
    Dim buttonName As String
    Dim myControl As CommandBarControl
    Dim bar As CommandBar

    On Error GoTo ErrorHandler

    Workbooks.Add
    Range("A1").Select
    With ActiveCell
      .Value = "Image"
      .Offset(0, 1) = "Index"
      .Offset(0, 2) = "Name"
      .Offset(0, 3) = "FaceID"
      .Offset(0, 4) = "CommandBar Name (Index)"
    End With

    For j = 1 To Application.CommandBars.Count

        Set bar = CommandBars(j)
        total = bar.Controls.Count

        With bar
          For i = 1 To total
              buttonName = .Controls(i).Caption
              buttonId = .Controls(i).ID

              Set myControl = CommandBars.FindControl(ID:=buttonId)
              myControl.CopyFace ' error could occur here
              ActiveCell.Offset(1, 0).Select
              Sheets(1).Paste

              With ActiveCell
                  .Offset(0, 1).Value = buttonId
                  .Offset(0, 2).Value = buttonName
                  .Offset(0, 3).Value = myControl.FaceId
                  .Offset(0, 4).Value = bar.Name & " (" & j & ")"
              End With
StartNext:
          Next i
        End With
    Next j

    Columns("A:E").EntireColumn.AutoFit
    Exit Sub
ErrorHandler:
      Resume StartNext
End Sub

Sub OpSystem()
    MsgBox Application.OperatingSystem, , "Operating System"
End Sub

Sub ActivePrinter()
    MsgBox Application.ActivePrinter
End Sub

Sub ActiveWorkbook()
    MsgBox Application.ActiveWorkbook.Name
End Sub

Sub ActiveSheet()
    MsgBox Application.ActiveSheet.Name
End Sub

Sub Create_ContextMenu()
  Dim sm As Object

  Set sm = Application.CommandBars.Add("MyComputer", msoBarPopup)
  With sm
    .Controls.Add(Type:=msoControlButton). _
        Caption = "Operating System"
    With .Controls("Operating System")
        .FaceId = 1954
        .OnAction = "OpSystem"
    End With
   .Controls.Add(Type:=msoControlButton).Caption = "Active Printer"
    With .Controls("Active Printer")
        .FaceId = 4
        .OnAction = "ActivePrinter"
    End With
   .Controls.Add(Type:=msoControlButton).Caption = "Active Workbook"
    With .Controls("Active Workbook")
        .FaceId = 247
        .OnAction = "ActiveWorkbook"
    End With
    .Controls.Add(Type:=msoControlButton).Caption = "Active Sheet"
    With .Controls("Active Sheet")
        .FaceId = 18
        .OnAction = "ActiveSheet"
    End With
  End With
End Sub


