Attribute VB_Name = "Module1"
Option Explicit

Sub MakeMemos()
Attribute MakeMemos.VB_ProcData.VB_Invoke_Func = " \n14"
'   Creates memos in word using Automation (late binding)
    Dim WordApp As Object
    Dim Data As Range, message As String
    Dim Records As Integer, i As Integer
    Dim Region As String, SalesAmt As String, SalesNum As String
    Dim SaveAsName As String
    
    'On Error GoTo ErrorCode
    
'   Start Word and create an object
    Set WordApp = CreateObject("Word.Application")

'   Information from worksheet
    Set Data = Sheets("Sheet1").Range("A1")
    message = Sheets("Sheet1").Range("Message")
    
'   Cycle through all records in Sheet1
    Records = Application.CountA(Sheets("Sheet1").Range("A:A"))
    For i = 1 To Records
'       Update status bar progress message
        Application.StatusBar = "Processing Record " & i

'       Assign current data to variables
        Region = Data.Cells(i, 1).Value
        SalesNum = Data.Cells(i, 2).Value
        SalesAmt = Format(Data.Cells(i, 3).Value, "#,000")

        
'       Determine the file name
        SaveAsName = Application.DefaultFilePath & "\" & Region & ".docx"

'       Send commands to Word
        With WordApp
            .Documents.Add
            With .Selection
                .Font.Size = 14
                .Font.Bold = True
                .ParagraphFormat.Alignment = 1
                .TypeText Text:="M E M O R A N D U M"
                .TypeParagraph
                .TypeParagraph
                .Font.Size = 12
                .ParagraphFormat.Alignment = 0
                .Font.Bold = False
                .TypeText Text:="Date:" & vbTab & _
                    Format(Date, "mmmm d, yyyy")
                .TypeParagraph
                .TypeText Text:="To:" & vbTab & Region & " Region Manager"
                .TypeParagraph
                .TypeText Text:="From:" & vbTab & _
                   Application.UserName
                .TypeParagraph
                .TypeParagraph
                .TypeText message
                .TypeParagraph
                .TypeText Text:="Units Sold:" & vbTab & SalesNum
                .TypeParagraph
                .TypeText Text:="Amount:" & vbTab & _
                  Format(SalesAmt, "$#,##0")
            End With
                .ActiveDocument.SaveAs FileName:=SaveAsName
        End With
    Next i

ErrorCode:
    If Err.Number = 0 Then
        MsgBox Records & " memos were created in Word " & WordApp.Version & " and saved in " & Application.DefaultFilePath
    Else
        MsgBox "An error occurred."
    End If

'   Kill the object
    WordApp.Quit False
    Set WordApp = Nothing

'   Reset status bar
    Application.StatusBar = ""

'   Show the folder
    Shell "explorer.exe " & Application.DefaultFilePath, vbNormalFocus

End Sub


