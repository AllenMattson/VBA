VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GameForm 
   Caption         =   "GameForm"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   OleObjectBlob   =   "GameForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Initialize()
    Dim i As Long
    Me.Width = 312
    Me.Height = 176
'   Fill the combo boxes
    With cbBet
        .AddItem "Bet 1"
        .AddItem "Bet 2"
        .AddItem "Bet 3"
        .AddItem "Bet 4"
        .AddItem "Bet 5"
        .ListIndex = 4
    End With
    With cbGame
        .AddItem "Jacks or Better"
        .AddItem "Joker's Wild"
        .ListIndex = 0
    End With
    
    For i = 1 To 13
        Set CardPics(i).CardPicture = Me.Controls("H" & Format(i, "00"))
        Set CardPics(i + 13).CardPicture = Me.Controls("C" & Format(i, "00"))
        Set CardPics(i + 26).CardPicture = Me.Controls("S" & Format(i, "00"))
        Set CardPics(i + 39).CardPicture = Me.Controls("D" & Format(i, "00"))
    Next i
    
    Set CardPics(53).CardPicture = Me.Controls("CardBack1")
    Set CardPics(54).CardPicture = Me.Controls("CardBack2")
    Set CardPics(55).CardPicture = Me.Controls("CardBack3")
    Set CardPics(56).CardPicture = Me.Controls("CardBack4")
    Set CardPics(57).CardPicture = Me.Controls("CardBack5")
    Set CardPics(58).CardPicture = Me.Controls("J00")

    
    DealButton.Caption = "Deal"
    ResultLabel.Caption = ""
    UserScore = 0
    ScoreLabel.Caption = UserScore

'   Coordinates for displayed cards
    CardTop = 6
    CardLeft(1) = 7.5
    CardLeft(2) = 66.15
    CardLeft(3) = 124.75
    CardLeft(4) = 183.35
    CardLeft(5) = 242
    
    If GetSetting(REGKEYNAME, "Settings", "RememberSettings", 1) = 1 Then
        cbBet.ListIndex = GetSetting(REGKEYNAME, APPNAME, "cbBet", 4)
        cbGame.ListIndex = GetSetting(REGKEYNAME, APPNAME, "cbGame", 0)
    Else
        cbBet.ListIndex = 4
        cbGame.ListIndex = 0
    End If
    Call cbGame_Change
End Sub

Private Sub cbGame_Change()
    If cbGame.ListIndex = 0 Then
        GameForm.Caption = "Video Poker - Jacks or Better"
    Else
        GameForm.Caption = "Video Poker - Joker's Wild"
    End If
End Sub

Private Sub ChartButton_Click()
    Worksheets("ScoreHistory").Activate
End Sub

Private Sub CommandButton1_Click()
    GameForm.Hide
End Sub

Private Sub DiscardAllButton_Click()
    Dim i As Long
    For i = 1 To 5
        Call Card_Click(Controls(Cards(i).Name).Name)
    Next i
End Sub

Private Sub KeepAllButton_Click()
    Dim i As Long
    For i = 1 To 5
        Call Card_Click(Controls("CardBack" & i).Name)
    Next i
End Sub

Private Sub DealButton_Click()
    Dim i As Long
    Dim NextCard As Long
    
    On Error Resume Next
    OpeningFrame.Visible = False
    Select Case DealButton.Caption
        Case "Deal"
            UserScore = UserScore - Right(cbBet.Value, 1)
            ScoreLabel.Caption = UserScore
            ShuffleDeck
'           Stash cards that are showing
            For i = 1 To 53
                With Controls(CardNames(i).CardName)
                    .Top = StackTop
                    .Left = StackLeft
                End With
             Next i
            For i = 1 To 5
                Controls("CardBack" & i).Top = StackTop
                Controls("CardBack" & i).Left = StackLeft
            Next i
            
            For i = 1 To 5
                Controls(CardNames(i).CardName).Top = CardTop
                Controls(CardNames(i).CardName).Left = CardLeft(i)
                Controls(CardNames(i).CardName).Tag = i 'stores the position
                Cards(i).Name = CardNames(i).CardName
                Cards(i).Suit = UCase(Left(CardNames(i).CardName, 1))
                Cards(i).Val = Right(CardNames(i).CardName, 2)
                Cards(i).Keep = True
            Next i
            DealButton.Caption = "Get New Cards"
            cbBet.Enabled = False
            cbGame.Enabled = False
            KeepAllButton.Enabled = True
            DiscardAllButton.Enabled = True
            Call EvaluateHand(1)
       
        Case "Get New Cards"
            NextCard = 6
            For i = 1 To 5
                If Not Cards(i).Keep Then
                    Controls("Cardback" & i).Top = StackTop
                    Controls("Cardback" & i).Left = StackLeft
                    Controls(CardNames(NextCard).CardName).Left = CardLeft(i)
                    Controls(CardNames(NextCard).CardName).Top = CardTop
                    Controls(CardNames(NextCard).CardName).Tag = i 'stores the position
                    Cards(i).Name = CardNames(NextCard).CardName
                    Cards(i).Suit = UCase(Left(CardNames(NextCard).CardName, 1))
                    Cards(i).Val = Right(CardNames(NextCard).CardName, 2)
                    NextCard = NextCard + 1
                End If
            Next i
            DealButton.Caption = "Deal"
            cbBet.Enabled = True
            cbGame.Enabled = True
            KeepAllButton.Enabled = False
            DiscardAllButton.Enabled = False
            Call EvaluateHand(2)
    End Select
    On Error GoTo 0
End Sub


Private Sub NewGameButton_Click()
    Dim i As Long
    Initialize
    For i = 1 To 53
        With Controls(CardNames(i).CardName)
            .Top = StackTop
            .Left = StackLeft
        End With
     Next i
    For i = 1 To 5
        Controls("CardBack" & i).Top = StackTop
        Controls("CardBack" & i).Left = StackLeft
    Next i
    DealButton.Caption = "Deal"
    ResultLabel.Caption = ""
    UserScore = 0
    ScoreLabel.Caption = UserScore
    OpeningFrame.Visible = True
    cbBet.Enabled = True
    cbGame.Enabled = True
End Sub

Private Sub OptionsButton_Click()
    Select Case OptionsButton.Caption
        Case "Options >>"
            GameForm.Width = 398
            OptionsButton.Caption = "Options <<"
        Case "Options <<"
            GameForm.Width = 312
            OptionsButton.Caption = "Options >>"
    End Select
End Sub

Private Sub PayoffsButton_Click()
    PayoffsForm.Show
End Sub

Private Sub QuitButton_Click()
    GameInProgress = False
    SaveSetting REGKEYNAME, APPNAME, "cbBet", cbBet.ListIndex
    SaveSetting REGKEYNAME, APPNAME, "cbGame", cbGame.ListIndex
    Unload Me
End Sub

Sub Card_Click(c As String)
    Dim Index As String
    
    On Error Resume Next
    If DealButton.Caption = "Deal" Then Exit Sub
    Controls(c).Top = StackTop
    Controls(c).Left = StackLeft
    Index = Controls(c).Tag
    Select Case Len(c)
        Case 3 'card showing
            Controls("CardBack" & Index).Top = CardTop
            Controls("CardBack" & Index).Left = CardLeft(Index)
            Controls("CardBack" & Index).Tag = Index
            Cards(Index).Keep = False
        Case Else 'card back showing
            Controls(Cards(Index).Name).Top = CardTop
            Controls(Cards(Index).Name).Left = CardLeft(Index)
            Cards(Index).Keep = True
    End Select
End Sub


Private Sub HelpButton_Click()
    Worksheets("Help").Activate
End Sub

Private Sub HelpButton2_Click()
    Worksheets("Help").Activate
End Sub

