Function SaveFileAs( _
  FileName as String
  FilePath as string
  Optional WorkbookName as Workbook = ActiveWorkbook.Name) as Boolean
  
  
  
  Workbooks(WorkBookName).SaveAs _
    FileName:=FileName
    
    
    ActiveWorkbook.SaveAs Filename:=thisWb.Path & "\new workbook.xls"
    ActiveWorkbook.Close savechanges:=False
End Sub
