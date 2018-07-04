Attribute VB_Name = "SavePDFOtherFormat"
Option Explicit
Sub test()
Dim F As String, X As String
F = "C:\Users\Allen\Desktop\TIM MATTSON PDF FILES\"
X = "xlsx"
Debug.Print Dir(F)
Call SavePDFAsOtherFormat(X, F)
End Sub

Public Sub SavePDFAsOtherFormat(PDFPath As String, FileExtension As String)
   
    'Saves a PDF file as another format using Adobe Professional.
   
    'By Christos Samaras
    'http://www.myengineeringworld.net
   
    'In order to use the macro you must enable the Acrobat library from VBA editor:
    'Go to Tools -> References -> Adobe Acrobat xx.0 Type Library, where xx depends
    'on your Acrobat Professional version (i.e. 9.0 or 10.0) you have installed to your PC.
   
    'Alternatively you can find it Tools -> References -> Browse and check for the path
    'C:\Program Files\Adobe\Acrobat xx.0\Acrobat\acrobat.tlb
    'where xx is your Acrobat version (i.e. 9.0 or 10.0 etc.).
   
    Dim objAcroApp      As Acrobat.AcroApp
    Dim objAcroAVDoc    As Acrobat.AcroAVDoc
    Dim objAcroPDDoc    As Acrobat.AcroPDDoc
    Dim objJSO          As Object
    Dim boResult        As Boolean
    Dim ExportFormat    As String
    Dim NewFilePath     As String
   
    'Check if the file exists.
    If Dir(PDFPath) = "" Then
        MsgBox "Cannot find the PDF file!" & vbCrLf & "Check the PDF path and retry.", _
                vbCritical, "File Path Error"
        Exit Sub
    End If
   
    'Check if the input file is a PDF file.
    If LCase(Right(PDFPath, 3)) <> "pdf" Then
        MsgBox "The input file is not a PDF file!", vbCritical, "File Type Error"
        Exit Sub
    End If
   
    'Initialize Acrobat by creating App object.
    Set objAcroApp = CreateObject("AcroExch.App")
   
    'Set AVDoc object.
    Set objAcroAVDoc = CreateObject("AcroExch.AVDoc")
   
    'Open the PDF file.
    boResult = objAcroAVDoc.Open(PDFPath, "")
       
    'Set the PDDoc object.
    Set objAcroPDDoc = objAcroAVDoc.GetPDDoc
   
    'Set the JS Object - Java Script Object.
    Set objJSO = objAcroPDDoc.GetJSObject
   
    'Check the type of conversion.
    Select Case LCase(FileExtension)
        Case "eps": ExportFormat = "com.adobe.acrobat.eps"
        Case "html", "htm": ExportFormat = "com.adobe.acrobat.html"
        Case "jpeg", "jpg", "jpe": ExportFormat = "com.adobe.acrobat.jpeg"
        Case "jpf", "jpx", "jp2", "j2k", "j2c", "jpc": ExportFormat = "com.adobe.acrobat.jp2k"
        Case "docx": ExportFormat = "com.adobe.acrobat.docx"
        Case "doc": ExportFormat = "com.adobe.acrobat.doc"
        Case "png": ExportFormat = "com.adobe.acrobat.png"
        Case "ps": ExportFormat = "com.adobe.acrobat.ps"
        Case "rft": ExportFormat = "com.adobe.acrobat.rft"
        Case "xlsx": ExportFormat = "com.adobe.acrobat.xlsx"
        Case "xls": ExportFormat = "com.adobe.acrobat.spreadsheet"
        Case "txt": ExportFormat = "com.adobe.acrobat.accesstext"
        Case "tiff", "tif": ExportFormat = "com.adobe.acrobat.tiff"
        Case "xml": ExportFormat = "com.adobe.acrobat.xml-1-00"
        Case Else: ExportFormat = "Wrong Input"
    End Select
    
    'Check if the format is correct and there are no errors.
    If ExportFormat <> "Wrong Input" And Err.Number = 0 Then
        
        'Format is correct and no errors.
        
        'Set the path of the new file. Note that Adobe instead of xls uses xml files.
        'That's why here the xls extension changes to xml.
        If LCase(FileExtension) <> "xls" Then
            NewFilePath = WorksheetFunction.Substitute(PDFPath, ".pdf", "." & LCase(FileExtension))
        Else
            NewFilePath = WorksheetFunction.Substitute(PDFPath, ".pdf", ".xml")
        End If
        
        'Save PDF file to the new format.
        boResult = objJSO.SaveAs(NewFilePath, ExportFormat)
        
        'Close the PDF file without saving the changes.
        boResult = objAcroAVDoc.Close(True)
        
        'Close the Acrobat application.
        boResult = objAcroApp.Exit
        
        'Inform the user that conversion was successfully.
        MsgBox "The PDf file:" & vbNewLine & PDFPath & vbNewLine & vbNewLine & _
        "Was saved as: " & vbNewLine & NewFilePath, vbInformation, "Conversion finished successfully"
         
    Else
       
        'Something went wrong, so close the PDF file and the application.
       
        'Close the PDF file without saving the changes.
        boResult = objAcroAVDoc.Close(True)
       
        'Close the Acrobat application.
        boResult = objAcroApp.Exit
       
        'Inform the user that something went wrong.
        MsgBox "Something went wrong!" & vbNewLine & "The conversion of the following PDF file FAILED:" & _
        vbNewLine & PDFPath, vbInformation, "Conversion failed"

    End If
       
    'Release the objects.
    Set objAcroPDDoc = Nothing
    Set objAcroAVDoc = Nothing
    Set objAcroApp = Nothing
       
End Sub

Public Sub ExportAllPDFs()
    
    'Convert all the PDF files that their paths are on column B of
    'the worksheet "Paths" into a different file format.
    'By Christos Samaras
    'http://www.myengineeringworld.net
    
    Dim FileFormat As String
    Dim LastRow As Long
    Dim i As Integer
    
    'Change this according to your own needs.
    'Available formats: eps html, htm jpeg, jpg, jpe jpf, jpx, jp2,
    'j2k, j2c, jpc, docx, doc, png, ps, rft, xlsx, xls, txt, tiff, tif and xml.
    'In this example the PDF file will be saved as text file.
    FileFormat = "txt"
    
    If FileFormat = "" Then
        shPaths.Range("B2").Select
        MsgBox "There are no file paths to convert!", vbInformation, "File paths missing"
        Exit Sub
    End If

    shPaths.Activate
    
    'Find the last row.
    With shPaths
        LastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
    End With
    
    'Check that there are available file paths.
    If LastRow < 2 Then
        shPaths.Range("B2").Select
        MsgBox "There are no file paths to convert!", vbInformation, "File paths missing"
        Exit Sub
    End If
    
    'For each cell in the range "B2:B" & last row convert the pdf file
    'into different format (here to text - txt).
    For i = 2 To LastRow
        SavePDFAsOtherFormatNoMsg Cells(i, 2).Value, FileFormat
    Next i
    
    'Inform the user that conversion finished.
    MsgBox "All files were converted successfully!", vbInformation, "Finished"
    
End Sub



