Attribute VB_Name = "Module1"
Option Explicit
Option Private Module
Public ChartData() As String
Public Const APPNAME As String = "Export Charts"

Public SelectedChartIndex As Long
Public UserRow As Long, UserCol As Long

'Callback for ec1 onAction
Sub ExportCharts(control As IRibbonControl)
    Call StartExportCharts
End Sub

'Callback for Button1 onAction
Sub ShowHelpFromRibbon(control As IRibbonControl)
    Call ShowHelp
End Sub



Sub StartExportCharts()
    Dim ChartCount As Integer, i As Integer
    Dim objChart As Object
    Dim SelectedChartObjectName As String
    Dim FileExt As String
    If ActiveWorkbook Is Nothing Then Exit Sub
    If GetSetting(APPNAME, "Settings", "RememberSettings", 1) = 1 Then
        Select Case GetSetting(APPNAME, "Settings", "ExportFormatCombo", 0)
            Case 0: FileExt = ".gif"
            Case 1: FileExt = ".jpg"
            Case 2: FileExt = ".tif"
            Case 3: FileExt = ".png"
        End Select
    Else
        FileExt = ".gif"
    End If
    
    On Error GoTo NoCanDo
    If TypeName(ActiveSheet) = "Chart" Then
        With UserForm1
            .Label1.Caption = "Select the chart sheet(s) to export:"
            .ScrollToChartButton.Caption = "Go to"
            .ScrollToChartButton.Accelerator = "G"
        End With
        ChartCount = -1
        SelectedChartObjectName = ActiveSheet.Name
        For Each objChart In ActiveWorkbook.Charts
            ChartCount = ChartCount + 1
            ReDim Preserve ChartData(1, ChartCount)
            ChartData(0, ChartCount) = objChart.Name
            ChartData(1, ChartCount) = LCase(Replace(objChart.Name, " ", "_") & FileExt)
        Next objChart
    End If
    
    If TypeName(ActiveSheet) = "Worksheet" Then
        SelectedChartObjectName = ""
        If TypeName(Selection) = "ChartObject" Then
            SelectedChartObjectName = Selection.Name
            Selection.Activate
        End If
        If Not ActiveChart Is Nothing Then
            SelectedChartObjectName = ActiveChart.Parent.Name
            ActiveWindow.Visible = False 'deselect the chart
        End If
        If ActiveSheet.ChartObjects.Count = 0 Then
            MsgBox "The active worksheet contains no embedded charts.", vbInformation, APPNAME
            Exit Sub
        Else
            ' remember scroll postition of window, in case "Scroll To" button is used
            On Error Resume Next
            UserRow = ActiveWindow.ScrollRow
            UserCol = ActiveWindow.ScrollColumn
            On Error GoTo 0
            UserForm1.Label1.Caption = "Select the chart object(s) to export:"
            ChartCount = -1
            For Each objChart In ActiveSheet.ChartObjects
                ChartCount = ChartCount + 1
                ReDim Preserve ChartData(1, ChartCount)
                ChartData(0, ChartCount) = objChart.Name
                ChartData(1, ChartCount) = LCase(Replace(objChart.Name, " ", "_") & FileExt)
            Next objChart
        End If
    End If
    
    UserForm1.ChartList.Column = ChartData
    If SelectedChartObjectName = "" Then 'select all
        For i = 0 To UserForm1.ChartList.ListCount - 1
            UserForm1.ChartList.Selected(i) = True
        Next i
    Else
        For i = 0 To UserForm1.ChartList.ListCount - 1
            If UserForm1.ChartList.List(i, 0) = SelectedChartObjectName Then UserForm1.ChartList.Selected(i) = True
        Next i
    End If
    With UserForm1
      .StartUpPosition = 0
      .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
      .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
      .Show
    End With
    Exit Sub
NoCanDo:
    MsgBox "Cannot export charts.", vbCritical, APPNAME
End Sub

Sub ShowHelp()
      Application.Help ThisWorkbook.path & "\export charts.chm", 0
End Sub
