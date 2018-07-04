Attribute VB_Name = "VBA_Adobe_ReadWrite_PDFs"
Sub ReadAdobeFields()
Dim row_number As Integer: row_number = 1
Dim AcrobatApplication As Acrobat.CAcroApp
Dim AcrobatDocument As Acrobat.CAcroAVDoc
Dim fcount As Long
Dim sFieldName As String

On Error Resume Next

Set AcrobatApplication = CreateObject("AcroExch.App")
Set AcrobatDocument = CreateObject("AcroExch.AvDoc")

If AcrobatDocument.Open("C:\Users\Allen\Desktop\CC2016_Offline_Activation Method With Screenshots.pdf", "") Then
    AcrobatApplication.Show
    Set acroform = CreateObject("AformAut.App")
    Set Fields = acroform.Fields
    fcount = Fields.Count 'number of fields
    
        For Each field In Fields
        row_number = row_number + 1
            sFieldName = field.Name
            'msgbox sfieldname
            
            Sheet1.Range("B" & row_number) = field.Name
            Sheet1.Range("C" & row_number) = field.Value
            Sheet1.Range("D" & row_number) = field.Style
            
        Next field
        
Else
    MsgBox "Failure"
End If

AcrobatApplication.exit
Set AcrobatApplication = Nothing
Set AcrobatDocument = Nothing
Set field = Nothing
Set Fields = Nothing
Set acroform = Nothing
End Sub
Sub WriteToAdobeFields()

Dim AcrobatApplication As Acrobat.CAcroApp
Dim AcrobatDocument As Acrobat.CAcroAVDoc
Dim fcount As Long
Dim sFieldName As String

Set AcrobatApplication = CreateObject("AcroExch.App")
Set AcrobatDocument = CreateObject("AcroExch.AvDoc")

If AcrobatDocument.Open("Path and Name .pdf", "") Then
    AcrobatApplication.Show
    Set acroform = CreateObject("AformAut.App")
    Set Fields = acroform.Fields
    fcount = Fields.Count 'number of fields
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Fields you are filling in
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Fields("topmostSubform[0].page1[0].line1[0].f1_09_0_[0]").Value = "First Name"
Fields("topmostSubform[0].page1[0].p1-cb1[0]").Value = "off"
        
Else
    MsgBox "Failure"
End If

AcrobatApplication.exit
Set AcrobatApplication = Nothing
Set AcrobatDocument = Nothing
Set field = Nothing

End Sub
