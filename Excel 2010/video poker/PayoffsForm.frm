VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PayoffsForm 
   Caption         =   "Video Poker Payoffs"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   OleObjectBlob   =   "PayoffsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PayoffsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OKButton_Click()
    Unload PayoffsForm
End Sub


Private Sub SpinButton1_Change()
    BetLabel.Caption = SpinButton1.Value
    Call DisplayPayoffs(SpinButton1.Value)
End Sub

Private Sub UserForm_Initialize()
    CurrentBet = Right(GameForm.cbBet.Value, 1)
    Call DisplayPayoffs(CurrentBet)
    SpinButton1.Value = CurrentBet
    MultiPage1.Value = GameForm.cbGame.ListIndex
End Sub

Sub DisplayPayoffs(betamount)
    NameLabel1.Caption = ""
    For Each cell In Sheet3.Range("JacksPayoffs").Columns(1).Cells
        NameLabel1.Caption = NameLabel1.Caption & cell.Text & Chr(13)
    Next cell
    PointLabel1.Caption = ""
    For Each cell In Sheet3.Range("JacksPayoffs").Columns(2).Cells
        PointLabel1.Caption = PointLabel1.Caption & cell.Value * betamount & Chr(13)
    Next cell
    NameLabel2.Caption = ""
    For Each cell In Sheet3.Range("JokerPayoffs").Columns(1).Cells
        NameLabel2.Caption = NameLabel2.Caption & cell.Text & Chr(13)
    Next cell
    PointLabel2.Caption = ""
    For Each cell In Sheet3.Range("JokerPayoffs").Columns(2).Cells
        PointLabel2.Caption = PointLabel2.Caption & cell.Value * betamount & Chr(13)
    Next cell
End Sub
