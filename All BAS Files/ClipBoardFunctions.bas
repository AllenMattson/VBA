Attribute VB_Name = "ClipBoardFunctions"
'''''''''''''''''''''''''''''''
'AWESOME WAY TO GET WHATEVER WAS COPIED SOMEWHERE ELSE! SO FAST
'
'FROM: http://www.cpearson.com/excel/Clipboard.aspx
'
'''''''''''''''''''''''''''''''


Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function CloseClipboard Lib "user32" () As Long

Sub ClearClipboard()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
End Sub
Public Function ArrayToClipboardString(Arr As Variant) As String
End Function
Public Function PutInClipboard(S As String, Optional FormatID As Variant) As Boolean
End Function
Public Function GetFromClipboard(Optional FormatID As Variant) As String
End Function
Public Function RangeToClipboardString(RR As Range) As String
End Function



Sub test()
Dim ReturnString As String
Sheets(2).Cells(1, 1).CurrentRegion.Select
Selection.Copy
ReturnString = RangeToClipboardString(Selection)
Debug.Print ReturnString
End Sub
