VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, _
            Cancel As Boolean)
    Dim strSearch As String

    strSearch = Sheets(2).Range("C3").Formula
    If Target = Range("C3") Then
        Cancel = True
        ActiveWorkbook.FollowHyperlink _
        Address:="http://www.bing.com/search", _
        ExtraInfo:="q=" & strSearch, _
        Method:=msoMethodGet
    End If
End Sub



