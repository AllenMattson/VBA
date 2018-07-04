Attribute VB_Name = "WordDoc_Manipulation"
Private Word As Word.Application
Private Excel As Excel.Application
Sub Extract_WildCard_Matches()

Dim sWildCard As String
Dim sDir
Dim oWD As New Word.Document
Dim sPath As String
sWildCard = "\[[!\[\]]{1,}\]"

sPath = "C:\Documents\"
sDir = Dir$(sPath & "*.docx", vbNormal)

Do Until LenB(sDir) = 0

 Set oWD = Documents.Open(sPath & sDir)

    Open "C:\Match_Output.txt" For Append As #1
    
        Selection.HomeKey wdStory, wdMove
        
        Selection.Find.Execute FindText:=sWildCard, MatchWildcards:=True
        
        Do While Selection.Find.Found
            
            Print #1, ActiveDocument.Name & vbTab & Selection.Range.Text
            
            Selection.Range.Collapse wdCollapseEnd
            
            Selection.Find.Execute
        Loop
        
    Close #1

 oWD.Close False

 sDir = Dir$

Loop




End Sub
Sub sAddTableOfContents()
    
    Dim wdApp As Word.Application
    Set wdApp = New Word.Application
    
    'Alternatively, we can use Late Binding
    'Dim wdApp As Object
    'Set wdApp = CreateObject("word.Application")
    
    Dim wdDoc As Word.Document
    Set wdDoc = wdApp.Documents.Add
    
    ' Note we define a Word.range, as the default range wouled be an Excel range!
    Dim myWordRange As Word.Range
    Dim Counter As Integer
    
    wdApp.Visible = True
    wdApp.Activate
    
    'Insert Some Headers
    With wdApp
        For Counter = 1 To 5
            .Selection.TypeParagraph
            .Selection.Style = "Heading 1"
            .Selection.TypeText "A Heading Level 1"
            .Selection.TypeParagraph
            .Selection.TypeText "Some details"
        Next
    End With

    ' We want to put table of contents at the top of the page
    Set myWordRange = wdApp.ActiveDocument.Range(0, 0)
    
    wdApp.ActiveDocument.TablesOfContents.Add _
     Range:=myWordRange, _
     UseFields:=False, _
     UseHeadingStyles:=True, _
     LowerHeadingLevel:=3, _
     UpperHeadingLevel:=1

End Sub

Public Sub sWriteMicrosoftTabs()

    'In Tools > References, add reference to "Microsoft Word XX.X Object Library" before running.
    
    'Early Binding
    Dim wdApp As Word.Application
    Set wdApp = New Word.Application
    
    'Alternatively, we can use Late Binding
    'Dim wdApp As Object
    'Set wdApp = CreateObject("word.Application")
        
    With wdApp
        .Visible = True
        .Activate
        .Documents.Add
    
        For Counter = 1 To 3
            .Selection.TypeText Text:=Counter & " - Tab 1 "
            
            ' position to 2.5 inches
            .Selection.Paragraphs.TabStops.Add Position:=Application.InchesToPoints(2.5), _
                Leader:=Counter, Alignment:=wdAlignTabLeft
            
            .Selection.TypeText Text:=vbTab & " - Tab 2 "
            
            ' position to 5 inches
            .Selection.Paragraphs.TabStops.Add Position:=Application.InchesToPoints(5), _
                Leader:=Counter, Alignment:=wdAlignTabLeft
            
            .Selection.TypeText Text:=vbTab & " - Tab 3 "
                    
            .Selection.TypeParagraph
        Next Counter
        
    End With
End Sub
Sub sWriteMSWordTable()
 
    'In Tools > References, add reference to "Microsoft Word XX.X Object Library" before running.
    
    'Early Binding
    Dim wdApp As Word.Application
    Set wdApp = New Word.Application
    
    'Alternatively, we can use Late Binding
    'Dim wdApp As Object
    'Set wdApp = CreateObject("word.Application")
        
    With wdApp
        .Visible = True
        .Activate
        .Documents.Add
        
        With .Selection
        
            .Tables.Add _
                    Range:=wdApp.Selection.Range, _
                    NumRows:=1, NumColumns:=3, _
                    DefaultTableBehavior:=wdWord9TableBehavior, _
                    AutoFitBehavior:=wdAutoFitContent
            
            For Counter = 1 To 12
                .TypeText Text:="Cell " & Counter
                If Counter <> 12 Then
                    .MoveRight Unit:=wdCell
                End If
            Next
        
        End With
        
    End With

End Sub
Sub BulletedLists()

    'In Tools > References, add reference to "Microsoft Word XX.X Object Library" before running.
    
    'Early Binding
    Dim wdApp As Word.Application
    Set wdApp = New Word.Application
    
    'Alternatively, we can use Late Binding
    'Dim wdApp As Object
    'Set wdApp = CreateObject("word.Application")
    
    With wdApp
        .Visible = True
        .Activate
        .Documents.Add
        ' turn on bullets
        .ListGalleries(wdBulletGallery).ListTemplates(1).Name = ""
        .Selection.Range.ListFormat.ApplyListTemplate ListTemplate:=.ListGalleries(wdBulletGallery).ListTemplates(1), _
            continuepreviouslist:=False, applyto:=wdListApplyToWholeList, defaultlistbehavior:=wdWord9ListBehavior
        
        With .Selection
            .ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Font.Bold = False
            .Font.Name = "Century Gothic"
            .Font.Size = 12
            .TypeText ("some details")
            .TypeParagraph
            .TypeText ("some details")
            .TypeParagraph
        End With
        
        ' turn off bullets
        .Selection.Range.ListFormat.RemoveNumbers wdBulletGallery
        
        With .Selection
            .ParagraphFormat.Alignment = wdAlignParagraphLeft
            .TypeText ("some details")
            .TypeParagraph
            .TypeText ("some details")
            .TypeParagraph
            
        End With
        
        ' turn on outline numbers
        .ListGalleries(wdOutlineNumberGallery).ListTemplates(1).Name = ""
        .Selection.Range.ListFormat.ApplyListTemplate ListTemplate:=.ListGalleries(wdOutlineNumberGallery).ListTemplates(1), _
            continuepreviouslist:=False, applyto:=wdListApplyToWholeList, defaultlistbehavior:=wdWord9ListBehavior
        
        With .Selection
            .ParagraphFormat.Alignment = wdAlignParagraphLeft
            .TypeText ("some details")
            .TypeParagraph
            .TypeText ("some details")
            
        End With
        
    End With
End Sub
Sub WriteToTables2()
'In Tools > References, add reference to "Microsoft Word XX.X Object Library" before running.
    Dim wdApp As Word.Application
    Dim wdDoc As Word.Document

    Set wdApp = New Word.Application
    wdApp.Visible = True
    
    
    Dim x As Integer
    Dim y As Integer
    
    wdApp.Visible = True
    wdApp.Activate
    wdApp.Documents.Add
            
    wdApp.ActiveDocument.Tables.Add Range:=wdApp.Selection.Range, NumRows:=2, NumColumns:= _
        2, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
                
    With wdApp.Selection.Tables(1)
        If .Style <> "Table Grid" Then
            .Style = "Table Grid"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
    End With
            
    With wdApp.Selection
        
        For x = 1 To 2
            ' set style name
            .Style = "Heading 1"
            .TypeText "Subject" & x
            .TypeParagraph
            .Style = "No Spacing"
            For y = 1 To 20
                .TypeText "paragraph text "
            Next y
            .TypeParagraph
        Next x
    
        ' new paragraph
        .TypeParagraph
        
        ' toggle bold on
        .Font.Bold = wdToggle
        .TypeText Text:="show some text in bold"
        .TypeParagraph
        
        'toggle bold off
        .Font.Bold = wdToggle
        .TypeText "show some text in regular front weight"
        .TypeParagraph
        
        
    End With
End Sub
Sub GenerateWordTable()
'In Tools > References, add reference to "Microsoft Word XX.X Object Library" before running.

    Dim wdApp As Word.Application
    Dim wdDoc As Word.Document
    Dim r As Integer

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True

    Set wdDoc = wdApp.Documents.Add
    wdApp.Activate

    Dim wdTbl As Word.Table
    Set wdTbl = wdDoc.Tables.Add(Range:=wdDoc.Range, NumRows:=5, NumColumns:=1)

    With wdTbl

        .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
        .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
        .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
        .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
        .Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
        
        For r = 1 To 5
            .cell(r, 1).Range.Text = ActiveSheet.Cells(r, 1).Value
        Next r
    End With
End Sub
