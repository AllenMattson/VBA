VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents GameButton As CommandButton
Attribute GameButton.VB_VarHelpID = -1

Private Sub GameButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Clicked As Control, Blank As Control
    Dim ClickedRow As Long, ClickedCol As Long
    
    Set Clicked = UserForm1.Controls(GameButton.Name)
    Set Blank = UserForm1.Controls("cb" & BlankRow & BlankCol)
    ClickedRow = Mid(GameButton.Name, 3, 1)
    ClickedCol = Mid(GameButton.Name, 4, 1)
'   The statement below determines if the move is valid.
    If Abs(BlankCol - ClickedCol) + Abs(BlankRow - ClickedRow) = 1 Then
'       Swap captions
        Blank.Caption = Clicked.Caption
        Clicked.Caption = ""

'       Change visible proberty
        Blank.Visible = True
        Clicked.Visible = False
        Blank.SetFocus
        
'       Specify info for the new blank tile
        BlankRow = ClickedRow
        BlankCol = ClickedCol

'       Update Number of moves
        UserForm1.LabelMoves.Caption = UserForm1.LabelMoves.Caption + 1
    Else
        Beep
    End If

End Sub

