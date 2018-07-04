Attribute VB_Name = "PDFtoExcel_Scrape"

Sub PDFtoExcel()

    ' import pdf file text into excel, using msWord as a proxy

    ' set reference to microsoft word object library

    Dim wdApp As Word.Application
    Set wdApp = New Word.Application

    Dim file As String
    file = "C:\Users\Allen\Documents\ALLEN HOLLY COURT FOLDER\CASE PAYMENTS\2018.pdf"

    Dim wdDoc As Word.Document
    Set wdDoc = wdApp.Documents.Open( _
                    Filename:=file, ConfirmConversions:=False, _
                    ReadOnly:=True, AddToRecentFiles:=False, _
                    PasswordDocument:="", PasswordTemplate:="", Revert:=False, _
                    WritePasswordDocument:="", WritePasswordTemplate:="", _
                    Format:=wdOpenFormatAuto, XMLTransform:="")

'   wdApp.Visible = false                   ' can make msWord visible if you want ... would help in determining location of data

    Dim cel As Range
    Set cel = Range("d2")                   ' put paragraph text in column D

    Dim prgf As Paragraph
    For Each prgf In wdDoc.Paragraphs
        cel = prgf.Range.Text               ' put paragraph into worksheet cell
        Set cel = cel.Offset(1)             ' point to next cell down
    Next prgf

    Set cel = Range("b2")                   ' put word text in column D

    Dim wrd As Word.Range
    For Each wrd In wdDoc.Words
        cel = wrd.Text
        Set cel = cel.Offset(1)
    Next wrd

    wdDoc.Close False
    Set wdDoc = Nothing

    wdApp.Quit
    Set wdApp = Nothing

End Sub



