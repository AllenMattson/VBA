Attribute VB_Name = "Module1"
Option Explicit

Type ButtonData
    Button As New Class1
    Caption As String
    Visible As Boolean
    Row As Integer
    Col As Integer
End Type

Public Buttons() As ButtonData
Public BlankRow As Integer, BlankCol As Integer
Public GameSize As Integer ' number of rows and columns
Public OldGameSize As Integer 'previous number of rows and columns

Sub PlayGame()
'   This sub initializes and displays the UserForm
    Randomize

'   Create the game buttons
    GameSize = 4 'default
    Call NewTiles

'   Set up the combo box
    With UserForm1.cbGameSize
        .AddItem "3 x 3"
        .AddItem "4 x 4"
        .AddItem "5 x 5"
        .ListIndex = 1
    End With
    
'   Show the form
    UserForm1.Show
End Sub

Sub NewTiles()
'   Creates the tiles

    Dim Buttonsize As Integer
    Dim ButtonNumber As Integer
    Dim r As Integer, c As Integer
    Dim b As Control
    
'   Button size depends on number of rows and cols
    Buttonsize = 140 / GameSize
    ReDim Buttons(1 To GameSize * GameSize)
        
'   Delete old buttons, if any
    For Each b In UserForm1.Controls
        If TypeName(b) = "CommandButton" Then
            If b.Tag = "GameButton" Then UserForm1.Controls.Remove b.Name
        End If
    Next b
    
'   Create the buttons
    ButtonNumber = 1
    For r = 1 To GameSize
        For c = 1 To GameSize
            Set Buttons(ButtonNumber).Button.GameButton = _
                UserForm1.Controls.Add("forms.commandbutton.1")
            With Buttons(ButtonNumber).Button.GameButton
                .Width = Buttonsize
                .Height = Buttonsize
                .Top = 10 + (r - 1) * .Height
                .Left = 10 + (c - 1) * .Width
                .Caption = ButtonNumber
                .Name = "cb" & r & c
                .Font.Bold = True
                .Font.Name = "Arial"
                Select Case GameSize
                    Case 3: .Font.Size = 22
                    Case 4: .Font.Size = 16
                    Case 5: .Font.Size = 13
                    Case 6: .Font.Size = 11
                End Select
                .TakeFocusOnClick = False
                .Tag = "GameButton"
            End With
            ButtonNumber = ButtonNumber + 1
        Next c
    Next r
    
'   Set up the "blank" button
     BlankRow = GameSize
     BlankCol = GameSize
     Buttons(BlankRow * BlankCol).Visible = False

'   Reset the click counter
    UserForm1.LabelMoves.Caption = "0"
End Sub

