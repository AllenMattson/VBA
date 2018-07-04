Attribute VB_Name = "FullContactSocialMediaAPI"
Option Explicit

Sub GetData()
    Dim DataSheet As Worksheet
    Dim EndDate As Date
    Dim StartDate As Date
    Dim Symbol As String
    Dim qurl As String
    Dim nQuery As Name
    Dim LastRow As Integer
    Dim myAPI As String: myAPI = "efe47aa5fbfdb0dc"
    Dim myEMAIL As String: myEMAIL = "Allen.Mattson@outlook.com"
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    ActiveSheet.Cells.Clear
    
    Set DataSheet = ActiveSheet 'Sheets("CandleChart")
  
        
        qurl = "https://api.fullcontact.com/v2/person.json?email=" & myEMAIL & "&apiKey=" & myAPI
                   
QueryQuote:
             With ActiveSheet.QueryTables.Add(Connection:="URL;" & qurl, Destination:=ActiveSheet.Range("a1"))
                .BackgroundQuery = True
                .TablesOnlyFromHTML = False
                .Refresh BackgroundQuery:=False
                .SaveData = True
            End With
            
            ActiveSheet.Range("a1").CurrentRegion.TextToColumns Destination:=ActiveSheet.Range("a1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=True, Space:=False, other:=False
                
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
            
End Sub

