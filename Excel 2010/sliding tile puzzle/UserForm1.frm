VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Sliding Tile Puzzle"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SizeChange As Boolean

Private Sub cbGameSize_Change()
    SizeChange = True
    Select Case cbGameSize.ListIndex
        Case 0: GameSize = 3
        Case 1: GameSize = 4
        Case 2: GameSize = 5
    End Select
    Call NewButton_Click
End Sub

Private Sub NewButton_Click()
    Dim i As Long
    Dim k As Long
    Dim r As Long, c As Long
    Dim x As Integer
    Dim ClickedRow As Long, ClickedCol As Long
    Dim Clicked As Control, Blank As Control
    
'   Reset click counter
    LabelMoves.Caption = 0
    
'   If the game size changed, make new tiles
    If SizeChange Then
        Call NewTiles
        SizeChange = False
    End If
    
'   Shuffle the tiles (simple randomization can result in unsolvable games)
    Randomize
    For i = 1 To GameSize * 400
'       simulate clicks
        ClickedRow = Application.RandBetween(1, GameSize)
        ClickedCol = Application.RandBetween(1, GameSize)
        Set Blank = UserForm1.Controls("cb" & BlankRow & BlankCol)
        Set Clicked = UserForm1.Controls("cb" & ClickedRow & ClickedCol)
        If Abs(BlankCol - ClickedCol) + Abs(BlankRow - ClickedRow) = 1 Then 'valid?
'       Swap captions
            Blank.Caption = Clicked.Caption
            Clicked.Caption = ""
'           Change visible proberty
            Blank.Visible = True
            Clicked.Visible = False
            Blank.SetFocus
'           Specify info for the new blank tile
            BlankRow = ClickedRow
            BlankCol = ClickedCol
        End If
    Next i
 End Sub

Private Sub QuitButton_Click()
    Unload Me
End Sub

Private Sub UserForm_Activate()
    NewButton_Click
End Sub


