Attribute VB_Name = "Module1"
Public objRibbon As IRibbonUI
Private strUserTxt As String
Public blnEnabled As Boolean


'callback for the onLoad attribute of customUI
Public Sub RefreshMe(Ribbon As IRibbonUI)
    Set objRibbon = Ribbon
End Sub


Public Sub getEditBoxText(control As IRibbonControl, _
            ByRef text)
    text = UCase(strUserTxt)
End Sub

Public Sub onFullNameChangeToUcase(ByVal control As IRibbonControl, _
                text As String)
    If text <> "" Then
        strUserTxt = text
        objRibbon.InvalidateControl "txtFullName"
    End If
End Sub


Public Sub OpenNotepad(ctl As IRibbonControl)
    Shell "Notepad.exe", vbNormalFocus
End Sub

Public Sub OpenCharmap(ctl As IRibbonControl)
    Shell "Charmap.exe", vbNormalFocus
End Sub

Public Sub OnLoadImage(imgName As String, ByRef image)
    Dim strImgFileName As String
    strImgFileName = "C:\Excel2013_HandsOn\Extra Images\" & imgName
    Set image = LoadPicture(strImgFileName)
End Sub

Public Sub OpenCalculator(ctl As IRibbonControl)
    Shell "Calc.exe", vbNormalFocus
End Sub

Sub onGetPressed(control As IRibbonControl, _
            ByRef pressed)
    If control.id = "tglR1C1" Then
        pressed = False
    End If

    If control.id = "chkGridlines" And _
        ActiveWindow.DisplayGridlines = True Then
        pressed = True
    ElseIf control.id = "chkGridlines" And _
        ActiveWindow.DisplayGridlines = False Then
        pressed = False
    End If

    If control.id = "chkFormulaBar" And _
        Application.DisplayFormulaBar = True Then
        pressed = True
    ElseIf control.id = "chkFormulaBar" And _
        Application.DisplayFormulaBar = False Then
        pressed = False
    End If
End Sub


Sub SwitchRefStyle(control As IRibbonControl, _
            pressed As Boolean)
    If pressed Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If
End Sub

Sub GoToSpecial(control As IRibbonControl)
    On Error Resume Next
    Range("A1").Select

    If control.id = "btnFormulas" Then
        Selection.SpecialCells(xlCellTypeFormulas, 23).Select
    ElseIf control.id = "btnNumbers" Then
        Selection.SpecialCells(xlCellTypeConstants, 1).Select
    ElseIf control.id = "btnText" Then
        Selection.SpecialCells(xlCellTypeConstants, 2).Select
    ElseIf control.id = "btnBlanks" Then
        Selection.SpecialCells(xlCellTypeBlanks).Select
    ElseIf control.id = "btnLast" Then
        Selection.SpecialCells(xlCellTypeLastCell).Select
    End If
End Sub

Sub DoSomething(ctl As IRibbonControl, _
                pressed As Boolean)
    If ctl.id = "chkGridlines" And pressed Then
       ActiveWindow.DisplayGridlines = True
    ElseIf ctl.id = "chkGridlines" And Not pressed Then
        ActiveWindow.DisplayGridlines = False
    ElseIf ctl.id = "chkFormulaBar" And pressed Then
        Application.DisplayFormulaBar = True
    ElseIf ctl.id = "chkFormulaBar" And Not pressed Then
        Application.DisplayFormulaBar = False
    End If
End Sub

Public Sub onFullNameChange(ctl As IRibbonControl, _
                    text As String)
    If text <> "" Then
       MsgBox "You've entered '" & text & _
         "' in the edit box."
    End If
End Sub

Public Sub onChangeDept(ctl As IRibbonControl, _
          text As String)
     MsgBox "You selected " & text & " department."
End Sub

Public Sub onActionBoro(ctl As IRibbonControl, _
         ByRef selectedId As String, _
         ByRef selectedIndex As Integer)
    MsgBox "Index=" & selectedIndex & " ID=" & selectedId
End Sub

Public Sub onGetImage(ctl As IRibbonControl, ByRef image)
    Select Case ctl.id
      Case "glHolidays"
         Set image = LoadPicture( _
         "C:\Excel2013_HandsOn\Extra Images\Square0.gif")
    End Select
End Sub

Public Sub onGetItemCount(ctl As IRibbonControl, ByRef count)
  count = 12
End Sub

Public Sub onGetItemLabel(ctl As IRibbonControl, _
           index As Integer, ByRef label)
    label = MonthName(index + 1)
End Sub

Public Sub onGetItemImage(ctl As IRibbonControl, _
        index As Integer, ByRef image)
    Dim imgPath As String

    imgPath = "C:\Excel2013_HandsOn\Extra Images\square"
    Set image = LoadPicture(imgPath & index + 1 & ".gif")
End Sub

Public Sub onGetItemID(ctl As IRibbonControl, _
           index As Integer, ByRef id)
    id = MonthName(index + 1)
End Sub

Public Sub onSelectedItem(ctl As IRibbonControl, _
                selectedId As String, _
                selectedIndex As Integer)
    Select Case selectedIndex
        Case 6
            MsgBox "Holiday 1: Independence Day, July 4th", _
            vbInformation + vbOKOnly, _
            selectedId & " Holidays"
        Case 11
            MsgBox "Holiday 1: Christmas Day, December 25th", _
            vbInformation + vbOKOnly, _
            selectedId & " Holidays"
        Case Else
            MsgBox "Please program holidays for " & selectedId & ".", _
            vbInformation + vbOKOnly, _
            " Under Construction"
    End Select
End Sub

Public Sub onActionLaunch(ctl As IRibbonControl)
    Application.Dialogs(xlDialogAutoCorrect).Show
End Sub




Public Sub onGetEnabled(ctl As IRibbonControl, _
                ByRef returnedVal)

    returnedVal = blnEnabled

End Sub

Sub DisableNameManager(ctl As IRibbonControl, _
                       ByRef cancelDefault)
    MsgBox "You are not authorized to use this function."
    cancelDefault = True
End Sub

Public Sub CopyPicture(ctl As IRibbonControl, _
                ByRef cancelDefault)
    If ActiveSheet.Name = "Sheet1" Then
        ' display the CopyPicture dialog box instead
        Application.Dialogs(xlDialogCopyPicture).Show
    Else
        cancelDefault = False
    End If
End Sub


Sub onGetBitmap(ctl As IRibbonControl, ByRef image)
   Set image = Application.CommandBars. _
     GetImageMso("ResearchPane", 16, 16)
End Sub

Sub DoDefaultPlus(ctl As IRibbonControl)
    If Not IsNumeric(ActiveCell.Value) Then
        Application.CommandBars.ExecuteMso "Thesaurus"
    Else
        MsgBox "To use Thesaurus, select a cell " & _
        "containing text.", _
        vbOKOnly + vbInformation, "Action Required"
    End If
End Sub

Sub onActionExecHyperlink(ctl As IRibbonControl)
    Select Case ctl.id
        Case "YouTube"
            ThisWorkbook.FollowHyperlink Address:="http://www.YouTube.com", _
                NewWindow:=True
        Case "amazon"
            ThisWorkbook.FollowHyperlink Address:="http://www.amazon.com", _
                NewWindow:=True
        Case "merc"
            ThisWorkbook.FollowHyperlink Address:="http://www.merclearning.com", _
                NewWindow:=True
        Case "msft"
            ThisWorkbook.FollowHyperlink Address:="http://www.Microsoft.com", _
                NewWindow:=True
        Case Else
            MsgBox "You clicked control id " & ctl.id & _
                " that has not been programmed!"
    End Select
End Sub

Sub onActionCopyToArchive(ctl As IRibbonControl)
    Archive
End Sub

Sub Archive()
    Dim folderName As String
    Dim MyDrive As String
    Dim BackupName As String
    
    Application.DisplayAlerts = False
    
    On Error GoTo ErrorHandler
    
    folderName = ActiveWorkbook.Path
    
    If folderName = "" Then
        MsgBox "You can't copy this file. " & Chr(13) _
            & "This file has not been saved.", _
        vbInformation, "File Archive"
    Else
        With ActiveWorkbook
            If Not .Saved Then .Save
            MyDrive = InputBox("Enter the Pathname:" & _
                Chr(13) & "(for example: D:\, " & _
                    "E:\MyFolder\, etc.)", _
                    "Archive Location?", "D:\")
            If MyDrive <> "" Then
                If Right(MyDrive, 1) <> "\" Then
                    MyDrive = MyDrive & "\"
                End If
                BackupName = MyDrive & .Name
                .SaveCopyAs Filename:=BackupName
                MsgBox .Name & " was copied to: " _
                    & MyDrive, , "End of Archiving"
            End If
        End With
  End If
  GoTo ProcEnd
ErrorHandler:
    MsgBox "Visual Basic cannot find the " & _
        "specified path (" & MyDrive & ")" & Chr(13) & _
        "for the archive. Please try again.", _
        vbInformation + vbOKOnly, "Disk Drive or " & _
        "Folder does not exist"
ProcEnd:
    Application.DisplayAlerts = True
End Sub



