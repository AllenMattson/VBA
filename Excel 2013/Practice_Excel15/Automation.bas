Attribute VB_Name = "Automation"
Option Explicit

Sub PrintWordDoc()
    Dim objWord As Object
    Set objWord = CreateObject("Word.Application")

    With objWord
        .Visible = True
        .Documents.Open "C:\Excel2013_HandsOn\LinkOrEmbed.docx"
        .Options.PrintBackground = False
        .ActiveDocument.PrintOut
        .Documents.Close
        .Quit
    End With

    Set objWord = Nothing
End Sub

Sub WriteLetter()
    Dim wordAppl As Word.Application
    Dim strFolder As String
    Dim strFileName As String
    Dim flag As Boolean

    On Error GoTo ErrorHandler
    flag = True
    strFolder = "C:\Excel2013_ByExample\"
    strFileName = "Invite.docx"
    
    Set wordAppl = CreateObject("Word.Application")
    
    With wordAppl
        .Visible = True
        .StatusBar = "Creating a new document..."
        .Documents.Add
        .ActiveDocument.Paragraphs(1).Range.InsertBefore "Invitation"
        .StatusBar = "Saving document..."
        .ActiveDocument.SaveAs2 _
            Filename:=strFolder & strFileName
        .StatusBar = "Exiting Word..."
        .Quit
    End With
ExitHere:
    If flag Then MsgBox "The Document file " & _
        strFileName & Chr(13) & "was saved in " & _
        Left(strFolder, Len(strFolder) - 1) & ".", _
        vbInformation, "Document Created and Saved"
    Set wordAppl = Nothing
    Exit Sub
ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Number & ":" & Err.Description
        flag = False
    End If
    Resume ExitHere
End Sub


Sub CenterText()
    Dim wordDoc As Word.Document
    Dim wordAppl As Word.Application
    Dim strDoc As String
    Dim myAppl As String

    On Error GoTo ErrorHandler

    strDoc = "C:\Excel2013_ByExample\Invite.docx"
    myAppl = "Word.Application"

    ' first find out whether the specified document exists
    If Not DocExists(strDoc) Then
       MsgBox strDoc & " does not exist." & Chr(13) & Chr(13) _
        & "Please run the WriteLetter procedure to create " & _
        strDoc & "."
        Exit Sub
    End If

    ' now check if Word is running
    If Not IsRunning(myAppl) Then
       MsgBox "Word is not running -> will create " & _
        "a new instance of Word. "
       Set wordAppl = CreateObject("Word.Application")
       Set wordDoc = wordAppl.Documents.Open(strDoc)
    Else
       MsgBox "Word is running -> will get the specified document. "
       ' bind the wordDoc variable to a specific Word document
       Set wordDoc = GetObject(strDoc)
    End If
    ' center the 1st paragraph horizontally on page
    With wordDoc.Paragraphs(1).Range
       .ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With
    wordDoc.Application.Quit SaveChanges:=True
    Set wordDoc = Nothing
    Set wordAppl = Nothing
    MsgBox "The document " & strDoc & " was reformatted."
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbCritical, "Error: " & Err.Number
End Sub


Function DocExists(ByVal mydoc As String) As Boolean
    On Error Resume Next
    If Dir(mydoc) <> "" Then
        DocExists = True
    Else
        DocExists = False
    End If
End Function

Function IsRunning(ByVal myAppl As String) As Boolean
    Dim applRef As Object
    On Error Resume Next

    Set applRef = GetObject(, myAppl)
    If Err.Number = 429 Then
        IsRunning = False
    Else
        IsRunning = True
    End If
    ' clear the object variable
    Set applRef = Nothing
End Function

Sub GetContacts()
    Dim objOut As Outlook.Application
    Dim objNspc As Namespace
    Dim objItem As ContactItem
    Dim r As Integer ' row index
    Dim Headings As Variant
    Dim i As Integer ' array element
    Dim cell As Variant
    
    r = 2
    Set objOut = New Outlook.Application
    Set objNspc = objOut.GetNamespace("MAPI")
    
    Headings = Array("Full Name", "Street", "City", _
    "State", "Zip Code", "E-Mail")
    Workbooks.Add
    Sheets(1).Activate
    For Each cell In Range("A1:F1")
        cell.FormulaR1C1 = Headings(i)
        i = i + 1
    Next
    
    For Each objItem In objNspc.GetDefaultFolder _
    (olFolderContacts).Items
        With ActiveSheet
            .Cells(r, 1).Value = objItem.FullName
            .Cells(r, 2).Value = objItem.BusinessAddress
            .Cells(r, 3).Value = objItem.BusinessAddressCity
            .Cells(r, 4).Value = objItem.BusinessAddressState
            .Cells(r, 5).Value = objItem.BusinessAddressPostalCode
            .Cells(r, 6).Value = objItem.Email1Address
        End With
        r = r + 1
    Next objItem
    
    Set objItem = Nothing
    Set objNspc = Nothing
    Set objOut = Nothing
    MsgBox "Your contacts have been dumped to Excel."
End Sub


