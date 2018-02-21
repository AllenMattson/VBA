VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVFileClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'CSVFileClass
'''''''''''''
'PROPERTIES
'  ExportRange
'  ImportRange

'METHOD
'  Import
'  Export

Private RangeToExport As Range
Private ImportToCell As Range

Property Get ExportRange() As Range
    Set ExportRange = RangeToExport
End Property

Property Let ExportRange(rng As Range)
    Set RangeToExport = rng
End Property

Property Get ImportRange() As Range
    Set ImportRange = ImportToCell
End Property

Property Let ImportRange(rng As Range)
    Set ImportToCell = rng
End Property

Sub Export(CSVFileName)
'   Exports a range to CSV file
    Dim ExpBook As Workbook
    
    If RangeToExport Is Nothing Then
        MsgBox "ExportRange not specified"
        Exit Sub
    End If
    
    On Error GoTo ErrHandle
    Application.ScreenUpdating = False
    Set ExpBook = Workbooks.Add(xlWorksheet)
    RangeToExport.Copy
    Application.DisplayAlerts = False
    With ExpBook
        .Sheets(1).Paste
        .SaveAs FileName:=CSVFileName, FileFormat:=xlCSV
        .Close SaveChanges:=False
    End With
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
ErrHandle:
    ExpBook.Close SaveChanges:=False
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error " & Err & vbCrLf & vbCrLf & Error(Err), _
      vbCritical, "Export Method Error"
End Sub

Sub Import(CSVFileName)
'   Imports a CSV file to a range
    Dim CSVFile As Workbook
    
    If ImportToCell Is Nothing Then
        MsgBox "ImportRange not specified"
        Exit Sub
    End If
    
    If CSVFileName = "" Then
        MsgBox "Import FileName not specified"
        Exit Sub
    End If
    
    On Error GoTo ErrHandle
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Workbooks.Open CSVFileName
    Set CSVFile = ActiveWorkbook
    ActiveSheet.UsedRange.Copy Destination:=ImportToCell
    CSVFile.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
ErrHandle:
    CSVFile.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error " & Err & vbCrLf & vbCrLf & Error(Err), _
      vbCritical, "Import Method Error"
End Sub


