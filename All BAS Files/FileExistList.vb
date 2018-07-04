Sub FileExistList()
    
    'A variation of the FillFileNamesFromList routine. 
    'Takes a list of file names and tests whether they exist in a given 
    '  directory. A "Y" is returned if the file exists, and a "N" is 
    '  returned if it does not exist. These values can be changed
    '  in the definitions below.
    '
    'Requires the following functions:
    '  FindLastRow
    '  FileExist

    'Define all the variables/options.
    Dim i As Long           'Iteration counter
    Dim LastRow As Long     'Last row to evaluate
    Dim FirstRow As Long    'First row to evaluate
    Dim fPath As String     'Directory where files should be
    Dim fName As String     'File name (pulled from spreadsheet)
    Dim fType As String     'File type (required by FileExist; can be used as array)
    Dim nCol As Long        'Column where file names live
    Dim rCol As Long        'Column where results are printed
    Dim rTrue As String     'Text to return if file exists
    Dim rFalse As String    'Text to return if file does not exist


    FirstRow = 2
    LastRow = FindLastRow(3)
    
    nCol = 3
    rCol = 2

    fPath = "V:\Corporate\Tax\Public\Axip\Tx_Audit\Invoices\"
    
    rTrue = "Y"
    rFalse = "File not found"


    'Loop through row [fRow] through [lRow]
    '  take the text in column [nCol] and look for it within [fPath]
    '  If the file exists return [rTrue]; if it does not
    '  return [rFalse] in cell (i, rCol)
    
    For i = FirstRow To LastRow
        fName = Cells(i, nCol)
        
        If FileExist(fPath, fName) Then
            Cells(i, rCol) = rTrue
        Else
            Cells(i, rCol) = rFalse
        End If
    Next i


End Sub
