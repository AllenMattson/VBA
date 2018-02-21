Attribute VB_Name = "Module18"
Option Explicit
Dim myClickEvent As clsCmdBarEvents

Sub ListVBECmdBars()
    Dim objCmdBar As CommandBar
    Dim strCmdType As String
    Dim c As Variant

    Workbooks.Add
    Range("A1").Select

    With ActiveCell
        .Offset(0, 0) = "CommandBar Name"
        .Offset(0, 1) = "Control Caption"
        .Offset(0, 2) = "Control ID"
    End With

    For Each objCmdBar In Application.VBE.CommandBars
        Select Case objCmdBar.Type
            Case 0
                strCmdType = "toolbar"
            Case 1
                strCmdType = "menu bar"
            Case 2
                strCmdType = "popup menu"
        End Select

        ActiveCell.Offset(1, 0) = objCmdBar.Name & _
            " (" & strCmdType & ")"

        For Each c In objCmdBar.Controls
            ActiveCell.Offset(1, 0).Select
            With ActiveCell
                .Offset(0, 1) = c.Caption
                .Offset(0, 2) = c.ID
            End With
        Next
    Next

    Columns("A:C").AutoFit

    Set objCmdBar = Nothing
End Sub

Sub AddCmdButton_ToVBE()
    Dim objCmdBar As CommandBar
    Dim objCmdBtn As CommandBarButton
    'Dim myClickEvent As clsCmdBarEvents
    
    ' get the reference to the Tools menu in the VBE
    Set objCmdBar = Application.VBE.CommandBars.FindControl _
        (ID:=30007).CommandBar

    ' add a button to the Tools menu
    Set objCmdBtn = objCmdBar.Controls.Add(msoControlButton)

    ' set the new button's properties
    With objCmdBtn
        .Caption = "List VBE menus and toolbars"
        .OnAction = "ListVBECmdBars"
     End With
     
     
    ' create an instance of the clsCmdEvents class
    Set myClickEvent = New clsCmdBarEvents

    ' hook up the class instance to the newly added button
    Set myClickEvent.cmdBtnEvents = objCmdBtn

    Set objCmdBtn = Nothing
    Set objCmdBar = Nothing
     
End Sub

Sub RemoveCmdButton_FromVBE()
    Dim objCmdBar As CommandBar
    Dim objCmdBarCtrl As CommandBarControl

    ' get the reference to the Tools menu in the VBE
    Set objCmdBar = Application.VBE.CommandBars("Tools")

    ' loop through the Tools menu controls
    ' and delete the control with the matching caption
    For Each objCmdBarCtrl In objCmdBar.Controls
        If objCmdBarCtrl.Caption = "List VBE menus and toolbars" Then
            objCmdBarCtrl.Delete
        End If
    Next

    Set objCmdBarCtrl = Nothing
    Set objCmdBar = Nothing
End Sub



