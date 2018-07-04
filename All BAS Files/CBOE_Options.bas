Attribute VB_Name = "CBOE_Options"
Option Explicit
Sub GetOptionChain()

    Dim myBrowser As Object, myHtml As String
    Dim htmlTables As Object, htmlTable As Object
    Dim htmlRows As Object, htmlRow As Object
    Dim htmlCells As Object, htmlCell As Object
    Dim xlRow As Long, xlCol As Integer
    Dim ExitAt As Date, Symbol As String
    Dim i As Integer

    On Error GoTo ErrHdl
    Const myUrl = "http://www.cboe.com/DelayedQuote/QuoteTable.aspx"

    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    Symbol = UCase(Trim(Range("Ticker").Text))
    With Range("Ticker").Worksheet
        Range(Range("Ticker").Offset(1, 0), Cells(Rows.Count, Range("Ticker").Column + 15)).ClearContents
    End With

    Set myBrowser = CreateObject("InternetExplorer.Application")
    myBrowser.Navigate myUrl

    While myBrowser.Busy Or myBrowser.ReadyState <> 4
        DoEvents
    Wend

    With myBrowser.Document.all
        On Error Resume Next
        ExitAt = Now + TimeValue("00:00:05")
        Do
           .Item("ctl00_ctl00_AllContent_ContentMain_ucQuoteTableCtl_txtSymbol").value = Symbol
             If Err.Number = 0 Then Exit Do
             Err.Clear
             DoEvents
             If Now > ExitAt Then Exit Do
        Loop
        .Item("ctl00_ctl00_AllContent_ContentMain_ucQuoteTableCtl_optAll").Checked = True
        .Item("ctl00_ctl00_AllContent_ContentMain_ucQuoteTableCtl_btnSubmit").Click
    End With

    While myBrowser.Busy Or myBrowser.ReadyState <> 4
        DoEvents
    Wend

    ExitAt = Now + TimeValue("00:00:05")
    Do
        myHtml = myBrowser.Document.body.innerhtml
        If InStr(myHtml, "Last Sale") > 0 Then Exit Do
        DoEvents
        If Now > ExitAt Then Exit Do
    Loop

    myHtml = myBrowser.Document.body.innerhtml

    On Error GoTo ErrHdl
    If InStr(myHtml, "Last Sale") = 0 Then Err.Raise vbObjectError + 513, , "No data for '" & Symbol & "'..."

    Set htmlTables = myBrowser.Document.all.tags("TABLE")
    xlRow = Range("Ticker").Row + 1

    For i = 6 To 8
        With htmlTables(i)
            Set htmlRows = .Rows
            xlRow = xlRow + 1
            For Each htmlRow In htmlRows
                xlCol = Range("Ticker").Column
                Set htmlCells = htmlRow.Cells
                For Each htmlCell In htmlCells
                    Range("Ticker").Worksheet.Cells(xlRow, xlCol) = htmlCell.innerText
                    xlCol = xlCol + 1
                Next htmlCell
                If Not IsEmpty(Range("Ticker").Worksheet.Cells(xlRow, Range("Ticker").Column)) Then xlRow = xlRow + 1
            Next htmlRow
        End With
    Next i
    
    
ErrHdl:

    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    If Err.Number Then MsgBox Err.Description, vbCritical, "Get Data"
    On Error Resume Next
    myBrowser.Quit
    Set myBrowser = Nothing
    Set htmlTables = Nothing

End Sub
