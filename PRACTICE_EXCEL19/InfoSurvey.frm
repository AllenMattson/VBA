VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InfoSurvey 
   Caption         =   "Info Survey"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   OleObjectBlob   =   "InfoSurvey.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InfoSurvey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim r As Integer
    Me.Hide
    
    r = Application.CountA(Range("A:A"))
    Range("A1").Offset(r + 1, 0) = Me.lboxSystems.Value
    
    If Me.optHard.Value = True Then
        Range("A1").Offset(r + 1, 1) = "*"
    End If
    If Me.optSoft.Value = True Then
        Range("A1").Offset(r + 1, 2) = "*"
    End If
    If Me.chkIBM.Value = True Then
        Range("A1").Offset(r + 1, 3) = "*"
    End If
    If Me.chkNote.Value = True Then
        Range("A1").Offset(r + 1, 4) = "*"
    End If
    If Me.chkMac.Value = True Then
        Range("A1").Offset(r + 1, 5) = "*"
    End If
    Range("A1").Offset(r + 1, 6) = Me.cboxWhereUsed.Value
    Range("A1").Offset(r + 1, 7) = Me.txtPercent.Value
    
    If Me.optMale.Value = True Then
        Range("A1").Offset(r + 1, 8) = "*"
    End If
    If Me.optFemale.Value = True Then
        Range("A1").Offset(r + 1, 9) = "*"
    End If
   Unload Me
End Sub

Private Sub optHard_Change()
    Me.lboxSystems.Clear
    Call ListHardware
    Me.lboxSystems.ListIndex = 0
    Me.picImage.Picture = LoadPicture("C:\Excel2013_HandsOn\cd.bmp")
End Sub

Private Sub optSoft_Change()
    Me.lboxSystems.Clear
    Call ListSoftware
    Me.lboxSystems.ListIndex = 0
    Me.picImage.Picture = LoadPicture("C:\Excel2013_HandsOn\books.bmp")
End Sub

Private Sub spPercent_Change()
    txtPercent.Value = spPercent.Value
End Sub

Private Sub txtPercent_Change()
    Dim entry As String
       
    On Error Resume Next
   
    entry = Me.txtPercent.Value
        If entry > 100 Then
            entry = 0
            Me.txtPercent.Value = entry
        End If
    spPercent.Value = txtPercent.Value
End Sub

Private Sub UserForm_Initialize()
'select the Hardware option
    optHard.Value = True
'turn off the Software option and all the check boxes
    optSoft.Value = False
    chkIBM.Value = False
    chkNote.Value = False
    chkMac.Value = False
    'display a zero in the text box
    txtPercent.Value = 0
' call ListHardware procedure
    Call ListHardware
'populate the combo box
    With Me.cboxWhereUsed
        .AddItem "Home"
        .AddItem "Work"
        .AddItem "School"
        .AddItem "Work/home"
        .AddItem "Home/school"
        .AddItem "Work/home/school"
    End With
 'select the first element in the list box
    Me.cboxWhereUsed.ListIndex = 0
'load a picture file for the Hardware option
    Me.picImage.Picture = LoadPicture("c:\Excel2013_HandsOn\cd.bmp")
End Sub





























