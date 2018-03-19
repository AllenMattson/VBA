Attribute VB_Name = "QuickChangePrinter"
'expression.PrintOut(Background, Append, Range, _
'OutputFileName, From, To, Item, Copies, Pages, _
'PageType, PrintToFile, Collate, FileName, _
'ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, _
'PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight)
Sub QuickChangePrinter()
Dim sNewPrinter As String
sNewPrinter = ActivePrinter
ActivePrinter = "Enter the Windows name of the printer here"
Application.PrintOut Filename:=""
ActivePrinter = sNewPrinter
End Sub
