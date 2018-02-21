Attribute VB_Name = "Module1"
Option Explicit

Sub FollowMe()
    Dim myRange As Range
    Set myRange = Sheets(1).Range("A1")

    myRange.Hyperlinks.Add Anchor:=myRange, _
       Address:="http://search.yahoo.com/", _
       ScreenTip:="Search Yahoo", _
       TextToDisplay:="Click here"
End Sub



