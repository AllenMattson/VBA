Attribute VB_Name = "Query_NCAAData"
'Conference Trends: http://stats.ncaa.org/rankings/conference_trends
'Team Trends: http://stats.ncaa.org/rankings/institution_trends

''''''http://stats.ncaa.org/team/369/12480'''''<Game Stats?
'http://stats.ncaa.org/team/746.0/12480
'http://stats.ncaa.org/team/schedule_list?academic_year=2017&conf_id=-1&division=1&sport_code=MBB
'http://stats.ncaa.org/rankings?academic_year=2017&division=1&sport_code=MBB
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''
''CONFERENCE TRENDS''
'http://stats.ncaa.org/rankings/conference_trends' NOT WORKING
''INSTITUTION TRENDS''
'http://stats.ncaa.org/rankings/institution_trends'' NOT WORKING
''RANKINGS DATA''
'http://stats.ncaa.org/rankings?academic_year=2017&division=1&sport_code=MBB WORKING!!
''''''''''''''''''''''''''''''''''''
Public Sub TeamData()
Dim vws As Worksheet
For Each vws In ThisWorkbook.Worksheets
    If vws.Visible = xlSheetHidden Then vws.Visible = xlSheetVisible
Next
If ActiveCell.Value = "" Then
    MsgBox "Active Cell Value Must be a Team", vbOKOnly, "No team selected..."
    Exit Sub
End If




Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    If Left(ws.Name, 4) = "Shee" Then ws.Delete
Next ws
Application.DisplayAlerts = True

Dim WB As Workbook: Set WB = ThisWorkbook



Dim NewWb As Workbook
Dim MyVal As String
Dim cell As Range, MyCell As String
On Error GoTo ErrGuy
MyCell = Selection.Value
For Each cell In Sheets("Attributes_Teams").Range("teams")
    If MyCell = cell.Value Then
        'GET TEAM STATS
        TeamStatic_Query (cell.Offset(0, 1).Value)
        Range("A1").Value = cell.Value
        ActiveSheet.Move
        Set NewWb = ActiveWorkbook
        
        'GET TEAM BIG STATS
        TeamStatDataSheet
        'GET PLAYERS
        PlayerStatic_Query (cell.Offset(0, 1).Value)
        Range("A1").Value = cell.Value
        ActiveSheet.Name = "Player Stats"
        WB.Activate
        Sheets("Bracket").Activate
    End If
Next cell


ErrGuy:
    If Err.Number = 13 Then MsgBox "Click cell that has a team in it before searching for team data", vbOKOnly, "Click a cell"
    
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Sheets("Attributes_Teams").Visible = xlSheetHidden

End Sub

Sub TeamStatDataSheet()
ActiveSheet.Name = "Team Stats"

Dim TeamStats As String
Dim i As Integer, LR As Integer
Dim home As Integer, away As Integer
Dim away_win As Integer, away_loss As Integer, home_win As Integer, home_loss As Integer
home = 0: away = 0: away_win = 0: away_loss = 0: home_win = 0: home_loss = 0
Dim Count As Integer, Spread As Integer, TotSpread As Integer
Count = 0: Spread = 0: TotSpread = 0
LR = Cells(Rows.Count, 1).End(xlUp).Row
For i = 5 To LR
'TotSpread = 0
    If Cells(i, 1).Value <> "" Then
    If Left(Cells(i, 1).Offset(0, 1).Value, 1) = "@" Then
        away = away + 1
        If Left(Cells(i, 1).Offset(0, 2).Value, 1) = "W" Then away_win = away_win + 1
        If Left(Cells(i, 1).Offset(0, 2).Value, 1) = "L" Then away_loss = away_loss + 1
    Else
        home = home + 1
        If Left(Cells(i, 1).Offset(0, 2).Value, 1) = "W" Then home_win = home_win + 1
        If Left(Cells(i, 1).Offset(0, 2).Value, 1) = "L" Then home_loss = home_loss + 1
    End If
    team_score = Mid(Cells(i, 1).Offset(0, 2).Value, 3, 2)
    opponent_score = Mid(Cells(i, 1).Offset(0, 2).Value, 8, 2)
    If Len(team_score) > 0 Then
        Spread = Spread + Abs(team_score - opponent_score)
        Count = Count + 1
        'Debug.Print team_score & " " & opponent_score & " Spread: " & team_score - opponent_score
    Else
        'CORRECTS ANY OFFSET BY THE LOOP
        'Debug.Print
    End If
    End If

Next
On Error Resume Next
Application.EnableEvents = True
Application.ScreenUpdating = True
myteam = Range("A1").Value

Sheets.Add
ActiveSheet.Name = "Gamblin Stuff"

TeamStats = vbNewLine & vbNewLine & "                                           " & myteam & vbNewLine
TeamStats = TeamStats & vbNewLine & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbNewLine
TeamStats = TeamStats & vbNewLine & "                                       **HOME**    " & vbNewLine
TeamStats = TeamStats & vbNewLine & "   Home Games Played: " & home & vbNewLine
TeamStats = TeamStats & vbNewLine & "   Home Won: " & home_win & vbNewLine
TeamStats = TeamStats & vbNewLine & "   Home Loss: " & home_loss & vbNewLine
TeamStats = TeamStats & vbNewLine & "                                       **AWAY**    " & vbNewLine
TeamStats = TeamStats & vbNewLine & "   Away Games Played: " & away & vbNewLine
TeamStats = TeamStats & vbNewLine & "   Away Won: " & away_win & vbNewLine
TeamStats = TeamStats & vbNewLine & "   Away Loss: " & away_loss & vbNewLine
TeamStats = TeamStats & vbNewLine & "                                       ***STATS*** " & vbNewLine
TeamStats = TeamStats & vbNewLine & "   Sum of Spread: " & Spread & " Number of Games: " & Count & vbNewLine
TeamStats = TeamStats & vbNewLine & "   Average Spread: " & Spread / Count & vbNewLine
TeamStats = TeamStats & vbNewLine & "   Win Percentage: " & (away_win + home_win) / Count & vbNewLine
TeamStats = TeamStats & vbNewLine & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbNewLine & vbNewLine
'MsgBox TeamStats

Range("A1:M33").Select: Selection.Merge: Selection.WrapText = True
ActiveCell.Value = TeamStats



'debug.print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"

'debug.print "       **HOME**"
'debug.print " Home Games Played: " & home
'debug.print "Home Won: " & home_win
'debug.print "Home Loss: " & home_loss
'debug.print "       **AWAY**"
'debug.print "Away Games Played: " & away
'debug.print "Away Won: " & away_win
'debug.print "Away Loss: " & away_loss
'debug.print "       ***STATS***"
'debug.print "Sum of Spread: " & Spread & " Number of Games: " & Count
'debug.print "Average Spread: " & Spread / Count
'debug.print "Win Percentage: " & (away_win + home_win) / Count
'debug.print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
'debug.print
End Sub

Private Function TeamStatic_Query(Att_arr As Integer)
Dim TeamData As Integer
TeamData = Att_arr
Sheets.Add
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
On Error Resume Next
'code = Sheet1.Range("tag").Value
   With ActiveSheet.QueryTables.Add(Connection:= _
      "URL;http://stats.ncaa.org/team/" & TeamData & "/12480", _
         Destination:=Range("a2"))
      .BackgroundQuery = True
      .TablesOnlyFromHTML = True
      .Refresh BackgroundQuery:=False
      .SaveData = True
   End With
ActiveSheet.Range("a1").CurrentRegion.TextToColumns Destination:=ActiveSheet.Range("a1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=True, Space:=False, other:=False
ActiveSheet.Columns.AutoFit
End Function
Private Function PlayerStatic_Query(Att_arr As Integer)
Dim TeamData As Integer
TeamData = Att_arr
Sheets.Add
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
On Error Resume Next

   With ActiveSheet.QueryTables.Add(Connection:= _
      "URL;http://stats.ncaa.org/team/" & TeamData & "/stats/12480", _
         Destination:=Range("a2"))
      .BackgroundQuery = True
      .TablesOnlyFromHTML = True
      .Refresh BackgroundQuery:=False
      .SaveData = True
   End With
ActiveSheet.Range("a1").CurrentRegion.TextToColumns Destination:=ActiveSheet.Range("a1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=True, Space:=False, other:=False
ActiveSheet.Columns.AutoFit
End Function
