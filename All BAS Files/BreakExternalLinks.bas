Sub BreakExternalLinks(Optional BreakLinks As Boolean = True)
     
    Dim ExternalLink As Variant
    Dim i As Long
        
    If Not BreakLinks = False Then
        ExternalLink = ActiveWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
          
        'Loops through all external links and break/remove them
        For i = 1 To UBound(ExternalLink)
             
            ActiveWorkbook.BreakLink _
                Name:=ExternalLink(i), _
                Type:=xlLinkTypeExcelLinks
        
        Next i
        
        Debug.Print "BreakExternalLinks set " & BreakLinks & ". " & i - 1 & " external link(s) broken"
    
    Else
        
        Debug.Print "BreakExternalLinks set " & BreakLinks & ". " & i - 1 & " external link(s) broken"
    
    End If

End Sub
