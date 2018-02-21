VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmbChartClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents myChartClass As Chart
Attribute myChartClass.VB_VarHelpID = -1

Private Sub myChartClass_MouseDown(ByVal Button As Long, _
  ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    
    Dim IDnum As Long
    Dim a As Long, b As Long
    
'   The next statement returns valus for
'   IDNum, a, and b
    myChartClass.GetChartElement X, Y, IDnum, a, b

'   Was a series clicked?
    If IDnum = xlSeries Then
        Select Case b
            Case 1
                Sheets("North").Activate
            Case 2
                Sheets("South").Activate
            Case 3
                Sheets("West").Activate
        End Select
    End If
    Range("A1").Select
End Sub

