Attribute VB_Name = "Module2"
Option Explicit
Option Private Module

Public HasJoker As Boolean
Public FirstDeal As Boolean
Public JokerGame As Boolean
Public PossibleRoyal As Boolean

Sub EvaluateHand(DealNum)
    Dim i As Long

    Dim FlushTrue As Boolean
    Dim StraightTrue As Boolean
    
    If DealNum = 1 Then FirstDeal = True Else FirstDeal = False
    If GameForm.cbGame.ListIndex = 1 Then JokerGame = True Else JokerGame = False

'   Determine if Joker present
    HasJoker = False
    For i = 1 To 5
        If Cards(i).Suit = "J" Then HasJoker = True
    Next i

    PossibleRoyal = False
    FlushTrue = IsFlush()
    StraightTrue = IsStraight()

    If FlushTrue And PossibleRoyal Then
        Call UpdateScore("Royal Flush")
        Exit Sub
    End If
    
    If IsFiveOfAKind() Then
        Call UpdateScore("Five of a Kind")
        Exit Sub
    End If
    
    If FlushTrue And StraightTrue Then
        Call UpdateScore("Straight Flush")
        Exit Sub
    End If
    
    If IsFourOfAKind() Then
        Call UpdateScore("Four of a Kind")
        Exit Sub
    End If
    
    Select Case HasJoker
        Case False
            If IsThreeofaKind() And IsPair() Then
                Call UpdateScore("Full House")
                Exit Sub
            End If
        Case True
            If IsJokerFullHouse() Then
                Call UpdateScore("Full House")
                Exit Sub
            End If
    End Select
    
    If FlushTrue Then
        Call UpdateScore("Flush")
        Exit Sub
    End If
    
    If StraightTrue Then
        Call UpdateScore("Straight")
        Exit Sub
    End If
        
    If IsThreeofaKind() Then
        Call UpdateScore("Three of a Kind")
        Exit Sub
    End If
    
    If IsTwoPair() Then
        Call UpdateScore("Two Pair")
        Exit Sub
    End If
    
    If Not JokerGame Then
        If IsJacksOrBetter() Then
            Call UpdateScore("Jacks or Better")
            Exit Sub
        End If
    End If
    
    If JokerGame Then
        If IsPairofAces() Then
            Call UpdateScore("Pair of Aces")
            Exit Sub
        End If
    End If
   
'   If it gets this far, no points
    Call UpdateScore("")
End Sub

Function IsPairofAces()
    Dim Count As Long
    Dim i As Long
    
    IsPairofAces = False
    Count = 0
    For i = 1 To 5
        If Cards(i).Val = 1 Then Count = Count + 1
    Next i
    If Count = 2 And Not HasJoker Then IsPairofAces = True
    If Count = 1 And HasJoker Then IsPairofAces = True
End Function

Function IsPair()
    Dim Count As Long
    Dim i As Long, j As Long
    
    IsPair = False
    Select Case HasJoker
        Case True
            IsPair = True
        Case False
            For i = 1 To 5
                Count = 0
                For j = 1 To 5
                    If Cards(j).Val = Cards(i).Val Then Count = Count + 1
                Next j
                If Count = 2 Then
                    IsPair = True
                    Exit Function
                End If
            Next i
    End Select
End Function


Function IsTwoPair()
    Dim Skip
    Dim Count As Long
    Dim Found1 As Boolean, Found2 As Boolean
    Dim i As Long, j As Long
    
    IsTwoPair = False
    Found1 = False
    Found2 = False
'   Look for one pair
    For i = 1 To 5
        Count = 0
        For j = 1 To 5
            If Cards(j).Val = Cards(i).Val Then Count = Count + 1
        Next j
        If Count = 2 Then
            Found1 = True
            Skip = Cards(i).Val
            GoTo NextOne
        End If
    Next i

NextOne:
    For i = 1 To 5
        Count = 0
        If Cards(i).Val <> Skip Then
            For j = 1 To 5
               If Cards(j).Val = Cards(i).Val Then Count = Count + 1
            Next j
        End If
        If Count = 2 Then
            Found2 = True
            GoTo Finish
        End If
    Next i

Finish:
    If Found1 And Found2 Then IsTwoPair = True
End Function

Function IsFlush()
    Dim cc(5)
    Dim i As Long, j As Long
    
    IsFlush = False
    Select Case HasJoker
        Case True
            j = 1
            For i = 1 To 5
                If Cards(i).Suit <> "J" Then
                    cc(j) = Cards(i).Suit
                    j = j + 1
                End If
            Next i
            For i = 2 To 4
                If cc(i) <> cc(1) Then Exit Function
            Next i
            IsFlush = True
        Case False
            For i = 2 To 5
            If Cards(i).Suit <> Cards(1).Suit Then Exit Function
        Next i
        IsFlush = True
    End Select
End Function

Function IsStraight()
    Dim Ace As Long
    Dim AceValue As Long
    Dim i As Long, j As Long
    
    PossibleRoyal = False
    IsStraight = False
    Dim JokerValue As Long
    Dim temp
    Dim cc(5)
    
    Select Case HasJoker
        Case True
'           See if there is an ace
            Ace = 1
            For i = 1 To 5
                If Cards(i).Val = 1 Then Ace = 2
            Next i
        
'           Will loop once if no ace, twice if ace present
            For AceValue = 1 To Ace * 7 Step 13
'               Loop once for each possible joker value
                For JokerValue = 1 To 14
'               Assign numbers to cards
                    For i = 1 To 5
                        Select Case Cards(i).Val
                            Case 11: cc(i) = 11
                            Case 12: cc(i) = 12
                            Case 13: cc(i) = 13
                            Case 1: cc(i) = AceValue '1 or 14
                            Case 0: cc(i) = JokerValue
                            Case Else: cc(i) = Val(Cards(i).Val)
                        End Select
                    Next i
'                   Sort them
                    For i = 1 To 5
                        For j = 1 To 4
                             If cc(j) > cc(j + 1) Then
                                temp = cc(j)
                                cc(j) = cc(j + 1)
                                cc(j + 1) = temp
                            End If
                        Next j
                    Next i
'                   See if they are consecutive
                    If (cc(1) = cc(2) - 1) And (cc(1) = cc(3) - 2) And (cc(1) = cc(4) - 3) And (cc(1) = cc(5) - 4) Then
                        IsStraight = True
                        If cc(5) = 14 Then PossibleRoyal = True
                        Exit Function
                    End If
                Next JokerValue
            Next AceValue
        
        Case False 'No joker
'           See if there is an ace
            Ace = 1
            For i = 1 To 5
                If Cards(i).Val = 1 Then Ace = 2
            Next i
        
'           Will loop once if no ace, twice if ace present
            For AceValue = 1 To Ace * 7 Step 13
'               Assign numbers to cards
                For i = 1 To 5
                    Select Case Cards(i).Val
                        Case 11: cc(i) = 11
                        Case 12: cc(i) = 12
                        Case 13: cc(i) = 13
                        Case 1: cc(i) = AceValue '1 or 14
                        Case Else: cc(i) = Val(Cards(i).Val)
                    End Select
                Next i
'               Sort them
                For i = 1 To 5
                    For j = 1 To 4
                         If cc(j) > cc(j + 1) Then
                            temp = cc(j)
                            cc(j) = cc(j + 1)
                            cc(j + 1) = temp
                        End If
                    Next j
                Next i
'               See if they are consecutive
                If (cc(1) = cc(2) - 1) And (cc(1) = cc(3) - 2) And (cc(1) = cc(4) - 3) And (cc(1) = cc(5) - 4) Then
                    IsStraight = True
                    If cc(5) = 14 Then PossibleRoyal = True
                    Exit Function
                End If
            Next AceValue
    End Select
End Function

Function IsJacksOrBetter()
    Dim cc(5)
    Dim i As Long, j As Long
    
    IsJacksOrBetter = False
    Dim Count As Long
    Select Case HasJoker
        Case True
            j = 1
            For i = 1 To 5
                If Cards(i).Suit <> 0 Then
                    cc(j) = Cards(i).Val
                    j = j + 1
                End If
            Next i
            For i = 1 To 4
                If cc(i) = 11 Or cc(i) = 12 Or cc(i) = 13 Or cc(i) = 1 Then
                    Count = 0
                    For j = 1 To 4
                        If cc(j) = cc(i) Then Count = Count + 1
                    Next j
                End If
               
                If Count = 1 Then
                    IsJacksOrBetter = True
                    Exit Function
                End If
            Next i
        
    Case False
        For i = 1 To 5
            If Cards(i).Val = 11 Or Cards(i).Val = 12 Or Cards(i).Val = 13 Or Cards(i).Val = 1 Then
                Count = 0
                For j = 1 To 5
                    If Cards(j).Val = Cards(i).Val Then Count = Count + 1
                Next j
            End If
            If Count = 2 Then
                IsJacksOrBetter = True
                Exit Function
            End If
        Next i
    End Select
End Function

Function IsFullHouse()
    Dim Found1 As Long, Found2 As Long
    Dim Skip As Long
    Dim Count As Long
    Dim i As Long, j As Long
    Dim cc(5)
    
    IsFullHouse = False
    Found1 = False
    Found2 = False
    
    Select Case HasJoker
        Case True
            j = 1
            For i = 1 To 5
                If Cards(i).Val <> 0 Then
                    cc(j) = Cards(i).Val
                    j = j + 1
                End If
            Next i
'           Look for one pair
            For i = 1 To 4
                Count = 0
                For j = 1 To 4
                    If cc(j) = cc(i) Then Count = Count + 1
                Next j
                If Count = 2 Then
                    Found1 = True
                    Skip = cc(i)
                    GoTo NextOne
                End If
            Next i
NextOne:
            For i = 1 To 4
                Count = 0
                If cc(i) <> Skip Then
                    For j = 1 To 4
                       If cc(j) = cc(i) Then Count = Count + 1
                    Next j
                End If
                If Count = 2 Then
                    Found2 = True
                    GoTo Finish
                End If
            Next i
Finish:
            If Found1 And Found2 Then IsFullHouse = True
        
        Case False
            If IsPair And IsThreeofaKind Then IsFullHouse = True
    End Select
End Function

Function IsFourOfAKind()
    Dim Count As Long
    Dim i As Long, j As Long
    
    Select Case HasJoker
        Case True
            IsFourOfAKind = False
            For i = 1 To 5
                Count = 0
                For j = 1 To 5
                    If Cards(j).Val = Cards(i).Val Then Count = Count + 1
                Next j
                If Count = 3 Then
                    IsFourOfAKind = True
                    Exit Function
                End If
            Next i
        Case False
            IsFourOfAKind = False
            For i = 1 To 5
                Count = 0
                For j = 1 To 5
                    If Cards(j).Val = Cards(i).Val Then Count = Count + 1
                Next j
                If Count = 4 Then
                    IsFourOfAKind = True
                    Exit Function
                End If
            Next i
    End Select
End Function

Function IsThreeofaKind()
    Dim Count As Long
    Dim i As Long, j As Long
    
    IsThreeofaKind = False
    Select Case HasJoker
        Case True
            For i = 1 To 5
                    Count = 0
                    For j = 1 To 5
                        If Cards(j).Val = Cards(i).Val Then Count = Count + 1
                    Next j
                If Count = 2 Then
                    IsThreeofaKind = True
                    Exit Function
                End If
            Next i
        Case False
            For i = 1 To 5
                Count = 0
                For j = 1 To 5
                    If Cards(j).Val = Cards(i).Val Then Count = Count + 1
                Next j
                If Count = 3 Then
                    IsThreeofaKind = True
                    Exit Function
                End If
            Next i
    End Select
End Function

Function IsFiveOfAKind()
    Dim Count As Long
    Dim i As Long, j As Long
    
    Select Case HasJoker
        Case True
            IsFiveOfAKind = False
            For i = 1 To 5
                Count = 0
                For j = 1 To 5
                    If Cards(j).Val = Cards(i).Val Then Count = Count + 1
                Next j
                If Count = 4 Then
                    IsFiveOfAKind = True
                    Exit Function
                End If
            Next i
        Case False
            IsFiveOfAKind = False
    End Select
End Function

Function IsJokerFullHouse()
    If IsTwoPair() Then IsJokerFullHouse = True Else IsJokerFullHouse = False
End Function

Sub UpdateScore(outcome)
    Dim PayoffRange As Range
    Dim TheBet As Long
    Dim Points As Long
    Dim NextBlankRow As Long
    
    Select Case FirstDeal
        Case True
            GameForm.ResultLabel.ForeColor = RGB(192, 192, 192)
            GameForm.ResultLabel.Caption = outcome
        Case False
            GameForm.ResultLabel.ForeColor = RGB(0, 255, 0)
            If outcome = "" Then
                GameForm.ResultLabel.Caption = "Game Over"
                GameForm.ResultLabel.ForeColor = RGB(255, 255, 255)
            Else
                If JokerGame Then Set PayoffRange = Sheet3.Range("JokerPayoffs") Else Set PayoffRange = Sheet3.Range("JacksPayoffs")
                TheBet = Right(GameForm.cbBet.Value, 1)
                Points = Application.WorksheetFunction.VLookup(outcome, PayoffRange, 2, False) * TheBet
                GameForm.ResultLabel.Caption = outcome & ": " & Points
                UserScore = UserScore + Points
                GameForm.ScoreLabel.Caption = UserScore
            End If
            NextBlankRow = Application.WorksheetFunction.CountA(Sheet2.Range("A:A")) + 1
            With ThisWorkbook.Worksheets("ScoreHistory")
                .Cells(NextBlankRow, 1).Value = Sheet2.Cells(NextBlankRow - 1, 1).Value + 1
                .Cells(NextBlankRow, 2).Value = UserScore
                .Range(.Cells(1, 1), .Cells(NextBlankRow, 1)).Name = "Hands"
                .Range(.Cells(1, 2), .Cells(NextBlankRow, 2)).Name = "Points"
            End With
    End Select
End Sub

