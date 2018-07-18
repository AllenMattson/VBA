VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Charts"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    obColumn.Value = True
End Sub

Private Sub obArea_Click()
    Call UpdateChart(xlAreaStacked)
End Sub

Private Sub obColumn_Click()
    Call UpdateChart(xlColumnClustered)
End Sub

Private Sub obLine_Click()
    Call UpdateChart(xlLineMarkers)
End Sub

Private Sub CloseButton_Click()
    Kill Application.DefaultFilePath & Application.PathSeparator & "temp.gif"
    Unload Me
End Sub

Private Sub UpdateChart(chtype)
    Dim currentChart As Chart
    Dim Fname As String
    
    Set currentChart = Sheets("Data").ChartObjects(1).Chart
    currentChart.ChartType = chtype

'   Save chart as GIF
    Fname = Application.DefaultFilePath & Application.PathSeparator & "temp.gif"
    currentChart.Export Filename:=Fname, FilterName:="GIF"

'   Show the chart
    Image1.Picture = LoadPicture(Fname)
End Sub


