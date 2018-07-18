Attribute VB_Name = "Module1"
Option Explicit

Sub MakePivotTables()
Attribute MakePivotTables.VB_Description = "Macro recorded 12/28/1998 by John Walkenbach"
Attribute MakePivotTables.VB_ProcData.VB_Invoke_Func = " \n14"
'   This procedure creates 28 pivot tables
    Dim PTCache As PivotCache
    Dim pt As PivotTable
    Dim SummarySheet As Worksheet
    Dim ItemName As String
    Dim Row As Long, Col As Long, i As Long
    
    Application.ScreenUpdating = False
    
'   Delete Summary sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Summary").Delete
    On Error GoTo 0
    
'   Add Summary sheet
    Set SummarySheet = Worksheets.Add
    ActiveSheet.Name = "Summary"
    
'   Create Pivot Cache
    Set PTCache = ActiveWorkbook.PivotCaches.Create( _
      SourceType:=xlDatabase, _
      SourceData:=Sheets("SurveyData").Range("A1"). _
        CurrentRegion)
    
    Row = 1
    For i = 1 To 14
      For Col = 1 To 6 Step 5 '2 columns
        ItemName = Sheets("SurveyData").Cells(1, i + 2)
        With Cells(Row, Col)
            .Value = ItemName
            .Font.Size = 16
        End With

'       Create pivot table
        Set pt = ActiveSheet.PivotTables.Add( _
          PivotCache:=PTCache, _
          TableDestination:=SummarySheet.Cells(Row + 1, Col))
        
'       Add the fields
        If Col = 1 Then 'Frequency tables
            With pt.PivotFields(ItemName)
              .Orientation = xlDataField
              .Name = "Frequency"
              .Function = xlCount
            End With
        Else ' Percent tables
        With pt.PivotFields(ItemName)
            .Orientation = xlDataField
            .Name = "Percent"
            .Function = xlCount
            .Calculation = xlPercentOfColumn
            .NumberFormat = "0.0%"
        End With
        End If
   
        pt.PivotFields(ItemName).Orientation = xlRowField
        pt.PivotFields("Sex").Orientation = xlColumnField
        pt.TableStyle2 = "PivotStyleMedium2"
        pt.DisplayFieldCaptions = False
        If Col = 6 Then
'           add data bars to the last column
            pt.ColumnGrand = False
            pt.DataBodyRange.Columns(3).FormatConditions. _
            AddDatabar
            With pt.DataBodyRange.Columns(3).FormatConditions(1)
                .BarFillType = xlDataBarFillSolid
                .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
                .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1
            End With
        End If
     Next Col
        Row = Row + 10
   Next i
   
'   Replace numbers with descriptive text
    With Range("A:A,F:F")
        .Replace "1", "Strongly Disagree"
        .Replace "2", "Disagree"
        .Replace "3", "Undecided"
        .Replace "4", "Agree"
        .Replace "5", "Strongly Agree"
    End With
End Sub

