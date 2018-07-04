Attribute VB_Name = "ExtractToWord"
Option Explicit
Dim wdApp As Word.Application
 
Sub ExtractToWord()

   'In Tools > References, add reference to "Microsoft Word 12 Object Library" before running.
   
    Dim lastCell
    Dim rng As Range
    Dim row As Range
    Dim cell As Range
    Dim arrayOfColumns
    arrayOfColumns = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    Dim thisRow As Range
    Dim thisCell As Range
    Dim myStyle As String
    
    ' get last cell in column B
    lastCell = getLastCell()
    
    Set rng = Range("B2:H" & lastCell)
    
    'iterate through rows
    For Each thisRow In rng.Rows
            
        'iterate through cells in row row
        For Each thisCell In thisRow.Cells

            If thisCell.Value = arrayOfColumns(thisCell.Column) Or thisCell.Value = "" Then
            ' do nothing
                ''frWriteLine thisCell.Value, "Normal"
                ''frWriteLine arrayOfColumns(thisCell.Column), "Normal"
                  If thisCell.Value = arrayOfColumns(thisCell.Column) Or thisCell.Value = "" Then
                  End If
                  
            Else
                myStyle = "Normal"
                Select Case thisCell.Column
                    Case 2
                        myStyle = "Heading 1"
                    Case 3
                        myStyle = "Heading 2"
                    Case 4
                        myStyle = "Heading 3"
                    Case Is > 5
                        myStyle = "Normal"
                    
                End Select
                    
                frWriteLine thisCell.Value, myStyle
            End If
        
        arrayOfColumns(thisCell.Column) = thisCell.Value
    
      Next thisCell
    Next thisRow
    
End Sub

Public Function getLastCell() As Integer

    Dim lastRowNumber As Long
    Dim lastRowString As String
    Dim lastRowAddress As String
         
    With ActiveSheet
        getLastCell = .Cells(.Rows.Count, 2).End(xlUp).row
    End With
    
End Function

Public Function frWriteLine(someData As Variant, myStyle As String)
    
    If wdApp Is Nothing Then
        
        Set wdApp = New Word.Application
        With wdApp
            .Visible = True
            .Activate
            .Documents.Add
        End With
            
    End If
    
    With wdApp
        
        With .Selection
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Style = myStyle
            .TypeText (someData)
            .TypeParagraph
        End With
    End With
    
End Function

