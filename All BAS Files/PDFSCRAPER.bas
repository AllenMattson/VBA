Attribute VB_Name = "PDFSCRAPER"

' (c) Dan Elgaard (www.EXCELGAARD.dk)

' Public constants for accessing the PDFTables website
  Public Const URL_Purchase_Pages As String = "https://pdftables.com/pricing"
  Public Const URL_Obtain_API_Key As String = "https://pdftables.com/pdf-to-excel-api"

' Module level constants
'uxhtdvclnkto
  Private Const URL_API_Remaining  As String = "https://pdftables.com/api/remaining?key=uxhtdvclnkto"
  Private Const URL_API_PDF2Excel  As String = "https://pdftables.com/api?key=uxhtdvclnkto&format=xlsx-single"

' Private Const HTTPRequest_Client As String = "Microsoft.XMLHTTP"            ' Faster, but gets corrupted more often, and will cache requests, thus cannot be used for fetching remaining pages.
  Private Const HTTPRequest_Client As String = "WinHttp.WinHttpRequest.5.1"   ' Slower, but more robust, and will not cache requests

' Public variable to test, if we were successful in converting/importing from PDF document to Excel worksheet
  Public PDFTables2ActiveSheet_Success As Boolean

' API Function
  #If VBA7 Then
      Private Declare PtrSafe Function DownloadURLToFile Lib "URLMon.DLL" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
  #Else
      Private Declare Function DownloadURLToFile Lib "URLMon.DLL" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
  #End If

Option Explicit
Sub PDF2Workbook()
'Make sure correct workbook is selected
If ActiveWorkbook.Name = "PDFSCRAPER.xlsm" Then
    Worksheets(1).Activate
    Dim wb As Workbook
    'Run module for file folder picker on pdf files
    Application.Run "PDFTables2Workbook", , True
    'Change new workbook page names to sheet1
    If applications.Workbooks.Count > 1 Then
        For Each wb In Application.Workbooks
          If wb.Name <> "PDFSCRAPER.xlsm" Then
              wb.Worksheets(1).Name = "Sheet1"
          End If
        Next wb
    'Display how many more times the macro is able to run before free api usage is full
        ShowPagesLeft
    End If
Else
    MsgBox "Make sure the activeworkbook name is PDFSCRAPER.xlsm"
End If
End Sub

Sub ShowPagesLeft()

      MsgBox PDFTablesPages, vbOKOnly + vbInformation, " PDFTables"

End Sub
Sub PDFTables2Workbook(Optional ByVal InitialFolderFile As String = vbNullString, Optional ByVal AllowMultiSelect As Boolean = False)

' This macro will allow you to select PDF documents to be converted into Excel, using online API from pdftables.com.
'
' For this to work, these things MUST be in order:
'
' - The workbook must have VB Reference to 'Microsoft Office Object Library' (MSOL)
' - There must an active working Internet connection available, when using this macro.
' - You must be a registered user at pdftables.com and have a key for their API (you can try it out for free, for the first 50 pages)
'
' The argument, 'InitialFolderFile', can be either a file to open or a folder to open from.
' The argument, 'AllowMultiSelect', can be used to allow the user to convert more than one PDF at the same time.
'
' The macro will create (and open) a new workbook for each converted PDF document.


' * ' Initialize
      On Error Resume Next

      If Len(PDFTablesKey) <> 12 Then Exit Sub        ' No key found/given
      If PDFTablesPages(5) < 1 Then Exit Sub          ' No pages on the account


' * ' Define variables
      Dim SelectedItem As Variant

      Dim FDO As Office.FileDialog                    ' File Dialog Object
      Set FDO = Application.FileDialog(msoFileDialogFilePicker)


' * ' Request PDF file(s) to convert into Excel
      With FDO
            .InitialFileName = InitialFolderFile
            .AllowMultiSelect = AllowMultiSelect
            .InitialView = msoFileDialogViewList

            .Filters.Clear
            .Filters.Add "PDF Documents", "*.PDF"     ' Make sure you can only pick PDF documents

            If .Show = -1 Then
                  For Each SelectedItem In .SelectedItems
                        Application.Run "PDFTables2Excel_CreateExcelFile", SelectedItem, True
                  Next
            End If
      End With


ES: ' End of Sub
      Set FDO = Nothing
      If IsArray(SelectedItem) Then Erase SelectedItem

End Sub
Sub PDFTables2ActiveSheet(Optional ByVal PDFFile As String = vbNullString)

' This macro will convert/import a PDF document into the active worksheet, using the pdftables.com API.
'
' If you don't provide a PDF document, as argument, or, if the given PDF document can not be found, the user is asked for at PDF document, using a standard file dialog.
'
' The macro accepts PDF documents located on-line, on a 'http' URL, like "http://www.EXCELGAARD.dk/Files/PDFs/Extern%20Data.PDF"


' * ' Initialize
      On Error Resume Next

      PDFTables2ActiveSheet_Success = False

      If TypeName(ActiveSheet) <> "Worksheet" Then GoTo ES:                         ' Active sheet is not a worksheet (probably a chart :-)
      If Len(PDFTablesKey) <> 12 Then GoTo ES:                                      ' No key found/given
      If PDFTablesPages(5) < 1 Then GoTo ES:                                        ' No pages on the account


' * ' Define variables
      Dim TempFile As String
      If LCase$(Left$(PDFFile, 4)) = "http" Then                                    ' An online PDF document is given - try to download it
            TempFile = Environ("TMP")
            If TempFile = vbNullString Then TempFile = Environ("TMP")
            If TempFile = vbNullString Then TempFile = ThisWorkbook.Path
            If Right$(TempFile, 1) <> Application.PathSeparator Then TempFile = TempFile & Application.PathSeparator
            TempFile = TempFile & "TempFile.PDF"

            SetAttr TempFile, vbNormal
            Kill TempFile

            If DownloadURLToFile(0, PDFFile, TempFile, 16, 0) <> 0 Then GoTo ES:    ' Download of online PDF document failed
            PDFFile = TempFile
      End If

      If Len(PDFFile) < 2 Or Len(Dir(PDFFile, vbHidden + vbSystem)) < 2 Then        ' If no PDF document is given, then ask for one
            Application.ScreenUpdating = True
            Application.Interactive = True
            Application.Cursor = xlDefault
            PDFFile = Application.GetOpenFilename("PDF (*.PDF), *.PDF")             ' Ask for PDF file to convert/import
            If UCase$(PDFFile) = "FALSE" Then GoTo ES:                              ' User clicked [Cancel]
      End If


' * ' Convert PDF document into Excel workbook
      PDFFile = Application.Run("PDFTables2Excel_CreateExcelFile", PDFFile, False)  ' Create temporary Excel file
      If Dir(PDFFile) = vbNullString Then GoTo ES:                                  ' Something went wrong...


' * ' Transfer converted PDF to active worksheet
      Call PDFTablesTransfer2Sheet(PDFFile)                                         ' Transfer to active worksheet


ES: ' End of Sub
      Kill PDFFile                                                                  ' Delete temporary file

End Sub
Private Sub PDFTablesTransfer2Sheet(ByVal ExcelFile As String)

' This macro will transfer the contents of the first worksheet in the given Excel workbook to the active worksheet.


' * ' Initialize
      On Error Resume Next


' * ' Define variables
      Dim HXL As Excel.Application                                                  ' HXL = Hidden Excel
      Set HXL = New Excel.Application

      With HXL                                                                      ' \
            .Visible = False                                                        '  \
            .EnableEvents = False                                                   '   > Set properties of hidden instance of Excel,
            .DisplayAlerts = False                                                  '   > so that it doesn't 'disturb' us while opened.
            .ScreenUpdating = False                                                 '  /
      End With                                                                      ' /

      Dim HWB As Excel.Workbook                                                     ' HWB = Hidden Workbook
      Set HWB = HXL.Workbooks.Open(ExcelFile)

      Dim HWS As Excel.Worksheet                                                    ' HWS = Hidden Worksheet
      Set HWS = HWB.Sheets(1)

      Dim DWS As Excel.Worksheet                                                    ' DWS = Destination Worksheet
      Set DWS = ActiveSheet


' * ' Prepare active worksheet
      DWS.Unprotect
      If DWS.ProtectContents = True Then GoTo ES:                                   ' Can not copy to a protected worksheet

      DWS.Select
      DWS.DisplayPageBreaks = False
      DWS.Cells.Delete

      Range("A1").Select


' * ' Copy contents of PDF to active worksheet
      Err.Clear
      HWS.Cells.Copy
      DWS.Paste
      If Err.Number = 0 Then
            PDFTables2ActiveSheet_Success = True
      Else
            PDFTables2ActiveSheet_Success = False
            Debug.Print "PDFTablesTransfer2Sheet : " & Err.Number & ", " & Err.Description
      End If


ES: ' End of Sub
      HWB.Saved = True                                                              ' In case of volatile functions in the hidden workbook
      HWB.Close SaveChanges:=False

      HXL.Quit                                                                      ' We can just quit the hidden instance of Excel, since we don't have any more workbooks opened in it

      Set HWS = Nothing                                                             ' \
      Set HWB = Nothing                                                             '  > Clean up object variables
      Set HXL = Nothing                                                             ' /
      Set DWS = Nothing                                                             '/

End Sub
Function PDFTablesKey(Optional ByVal DisplayWarning As VbTriState = vbUseDefault) As String

' This little User-Defined Function (UDF) allows you to set the key for your pdftables.com subscription.
'
' You can either hardcode the key into your project (not recommended), or
' you can set the key into an environment variable, or
' your can save the key to the Registry (recommended).
'
' If you don't have a key for using the PDFTables API, you can obtain one from https://pdftables.com/
'
' DisplayWarning  =  Default  =  Ask for key, if missing, using InputBox
'                 =  False    =  Don't display anything
'                 =  True     =  Display warning, but don't request a key


' * ' Initialize
      On Error Resume Next


' * ' Define variables
      Dim MsgBoxPrompt As String
      MsgBoxPrompt = "You must have an API key!" & vbNewLine & vbNewLine & "Please, get an API key at " & URL_Obtain_API_Key & vbNewLine & vbNewLine & "Visit website now?"

      Dim InputBoxPrompt As String
      InputBoxPrompt = "You must have an API key! " & vbNewLine & vbNewLine & "Enter key:"


' * ' Set the key to the function
      PDFTablesKey = "uxhtdvclnkto"                                                                   ' Insert your own key here, if you want to hardcode the key into the project (not recommended)
      If PDFTablesKey = "" Then PDFTablesKey = Environ("PDFTables Key")                   ' Key stored in an environment variable
      PDFTablesKey = GetSetting("Credentials", "PDFTables", "Token, Key", PDFTablesKey)   ' Key saved to Registry (recommended)


' * ' Request key
      If PDFTablesKey = "" And DisplayWarning = vbUseDefault Then
            PDFTablesKey = Application.InputBox(Prompt:=InputBoxPrompt, Title:=" PDFTables", Type:=2)
            If Len(PDFTablesKey) <> 12 Then GoTo EF:
            SaveSetting "Credentials", "PDFTables", "Token, Key", PDFTablesKey
      End If


' * ' Display warning
      If Len(PDFTablesKey) <> 12 And DisplayWarning <> False Then
            If MsgBox(MsgBoxPrompt, vbYesNo + vbCritical, " PDFTables") = vbYes Then     ' Visit website to obtain a key
                  ActiveWorkbook.FollowHyperlink Address:=URL_Obtain_API_Key, NewWindow:=True, AddHistory:=True
            End If
      End If


EF: ' End of Function
      If Len(PDFTablesKey) <> 12 Then PDFTablesKey = vbNullString

End Function
Function PDFTablesPages(Optional ByVal DisplayWarningBelow As Long = -1, Optional ByVal AccountKey As String = vbNullString) As Long

' This little function will return the number of pages left on the subscribtion at pdftables.com


' * ' Initialize
      On Error Resume Next


' * ' Define variables
      Dim BodyText As String

      Dim MsgBoxPrompt As String
      MsgBoxPrompt = "There are !PAGES! pages left." & vbNewLine & vbNewLine & "You can purchase more pages at " & URL_Purchase_Pages & vbNewLine & vbNewLine & "Visit website now?"

      If Len(AccountKey) <> 12 Then AccountKey = PDFTablesKey
      If Len(AccountKey) <> 12 Then GoTo EF:

' * ' Read number of pages left
      With CreateObject(HTTPRequest_Client)
            .Open "GET", Replace(URL_API_Remaining, "uxhtdvclnkto", AccountKey), False
            .Send
            If .Status <> 200 Then
                If MsgBox(.ResponseText, vbOKOnly + vbCritical, "Error") = vbOK Then
                    Exit Function
                End If
            Else
                BodyText = .ResponseText
            End If
      End With


' * ' Display warning
      If Len(BodyText) > 0 And Val(BodyText) <= DisplayWarningBelow Then
            If MsgBox(Replace(MsgBoxPrompt, "!PAGES!", Val(BodyText)), vbYesNo + vbExclamation, " PDFTables") = vbYes Then
                  ActiveWorkbook.FollowHyperlink Address:=URL_Purchase_Pages, NewWindow:=True, AddHistory:=True
            End If
      End If


EF: ' End of Function
      PDFTablesPages = Val(BodyText)

End Function
Private Function PDFTables2Excel_CreateExcelFile(ByVal PDFFile As String, Optional ByVal OpenAfterCreation As Boolean = False) As String

' This function will convert a given PDF file into an Excel workbook, using the pdftables.com API.
'
' For this to work, these things MUST be in order:
'
' - The workbook must have VB Reference to 'Microsoft Office Object Library' (MSOL)
' - There must an active working Internet connection available, when using this macro.
' - You must be a registered user at pdftables.com and have a key for their API (you can try it out for free, for the first 50 pages)


' * ' Initialize
      On Error Resume Next


' * ' Define variables
      Dim ExcelFileData As Variant
      ExcelFileData = PDFTables2Excel_ConvertFile(Replace(URL_API_PDF2Excel, "uxhtdvclnkto", PDFTablesKey), PDFFile)

      Dim DataArray() As Byte
      DataArray = ExcelFileData

      Dim FileNumber As Long
      FileNumber = InStrRev(PDFFile, ".")
      PDFFile = Left$(PDFFile, FileNumber) & "XLSx"
      FileNumber = -1
      FileNumber = InStrRev(PDFFile, Application.PathSeparator)
      If FileNumber > 0 Then PDFFile = Mid$(PDFFile, FileNumber + 1)
      PDFFile = UCase$(Left$(PDFFile, 1)) & Mid$(PDFFile, 2)

      Dim ExcelFile As String
      ExcelFile = CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2)
      If Right$(ExcelFile, 1) <> Application.PathSeparator Then ExcelFile = ExcelFile & Application.PathSeparator
      ExcelFile = ExcelFile & PDFFile


' * ' Create temporary temporary Excel file
      FileNumber = FreeFile
      Open ExcelFile For Binary Lock Read Write As #FileNumber
            Put #FileNumber, , DataArray
      Close #FileNumber


EF: ' End of Function
      PDFTables2Excel_CreateExcelFile = ExcelFile                             ' Return the path of the created Excel file as result of the function

      If OpenAfterCreation = True Then Workbooks.Open (ExcelFile)

      Erase DataArray

End Function
Private Function PDFTables2Excel_ConvertFile(ByVal APIURL As String, ByVal PDFFile As String) As Variant

' This function will upload (post) the given PDF document to the given online API at pdftables.com
' and retrieve the converted Excel file back, and place the Excel file in the Windows temporary directory.


' * ' Initialize
      Const Boundary As String = "3fbd04f5Rb1edX4060q99b9Nfca7ff59c113"

      On Error Resume Next


' * ' Define variables
      Dim BufferArray() As Byte
      Dim BodyData As String

      Dim FileNumber As Long
      FileNumber = FreeFile


' * ' Read file
      Open PDFFile For Binary Access Read As #FileNumber
            If LOF(FileNumber) > 0 Then
                  ReDim BufferArray(0 To LOF(FileNumber) - 1) As Byte
                  Get FileNumber, , BufferArray
                  BodyData = StrConv(BufferArray, vbUnicode)
            End If
      Close #FileNumber


' * ' Prepare body
      BodyData = "--" & Boundary & vbCrLf & _
                 "Content-Disposition: form-data; name=""uploadfile""; filename=""" & Mid$(PDFFile, InStrRev(PDFFile, Application.PathSeparator) + 1) & """" & vbCrLf & _
                 "Content-Type: application/octet-stream" & vbCrLf & vbCrLf & BodyData & vbCrLf & _
                 "--" & Boundary & "--"


' * ' Post (upload) the PDF document, and retrieve the converted Excel file, and the .ResponseText of the URL
      With CreateObject(HTTPRequest_Client)
            .Open "POST", APIURL, False
            .SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & Boundary
            .Send PDFTables2Excel_BodyElements2Bytes(BodyData)
            If .Status <> 200 Then
                If MsgBox(.ResponseText, vbOKOnly + vbCritical, "Error") = vbOK Then
                    Exit Function
                End If
            Else
                PDFTables2Excel_ConvertFile = .ResponseBody
            End If
      End With


EF: ' End of Function
      Erase BufferArray

End Function
Private Function PDFTables2Excel_BodyElements2Bytes(ByVal BodyData As String) As Byte()

' This sub-function converts the body string into an array (byte array)

      On Error Resume Next
      PDFTables2Excel_BodyElements2Bytes = StrConv(BodyData, vbFromUnicode)

End Function

