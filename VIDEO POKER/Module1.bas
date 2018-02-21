Attribute VB_Name = "Module1"
Option Explicit

Public CardPics(1 To 58) As New Class1
Public Const REGKEYNAME As String = "Excel Games"
Public Const APPNAME As String = "Video Poker"
Type CardType
    Image As Image
    Name As String * 3
    Suit As String * 1
    Val As Long
    Keep As Boolean
End Type

Type CardSort
    CardName As String
    Random As Long
End Type

Public CardTop
Public CardLeft(5)
Public Const StackLeft = 0
Public Const StackTop = 205
Public Cards(1 To 5) As CardType
Public CardNames(1 To 53) As CardSort
Public FirstDeal As Boolean
Public UserScore
Public GameInProgress
Sub VideoPoker()
    If GameInProgress Then
        GameForm.Show
        Exit Sub
    End If
    Initialize
    Randomize
    GameInProgress = True
    GameForm.Show
End Sub

Sub Initialize()
    Dim i As Long
    
    For i = 1 To 13
        CardNames(i).CardName = "C" & Format(i, "00")
    Next i
    For i = 14 To 26
        CardNames(i).CardName = "D" & Format(i - 13, "00")
    Next i
    For i = 27 To 39
        CardNames(i).CardName = "H" & Format(i - 26, "00")
    Next i
    For i = 40 To 52
        CardNames(i).CardName = "S" & Format(i - 39, "00")
    Next i
    
    CardNames(53).CardName = "J00" 'Joker
    With ThisWorkbook.Worksheets("ScoreHistory")
        .Range("A:B").ClearContents
        .Range("A1").Value = 0
        .Range("B1").Value = 0
        .Range("A1").Name = "Hands"
        .Range("B1").Name = "Points"
    End With
End Sub
Sub ShuffleDeck()
'   Shuffles the deck before dealing
    Dim temp As CardSort
    Dim i As Long, j As Long

'   Assign random numbers
    For i = 1 To 53
        CardNames(i).Random = Int(Rnd * 32767)
    Next i
    
'   Sort cards by the random number
    For i = 1 To 53
        For j = 1 To 52
            If CardNames(j).Random > CardNames(j + 1).Random Then
                temp.Random = CardNames(j).Random
                temp.CardName = CardNames(j).CardName
                CardNames(j).Random = CardNames(j + 1).Random
                CardNames(j).CardName = CardNames(j + 1).CardName
                CardNames(j + 1).Random = temp.Random
                CardNames(j + 1).CardName = temp.CardName
            End If
        Next j
    Next i
    
'   Stash the joker if it is not being used
    If GameForm.cbGame.ListIndex = 0 Then
        For i = 1 To 10
            If CardNames(i).CardName = "J00" Then
                CardNames(i).CardName = CardNames(53).CardName
                CardNames(53).CardName = "J00"
            End If
        Next i
    End If
'   Initially, all cards will be kept
    For i = 1 To 5
        Cards(i).Keep = True
    Next i
End Sub

Sub ReturnToGame()
    Application.DisplayAlerts = False
    Windows("Video Poker Chart").Parent.Close
    GameForm.Show
End Sub

