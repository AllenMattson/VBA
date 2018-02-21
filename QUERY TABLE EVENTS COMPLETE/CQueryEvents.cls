VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CQueryEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents qt As QueryTable
Attribute qt.VB_VarHelpID = -1

Public Property Get QTable() As QueryTable
    Set QTable = qt
End Property

Public Property Set QTable(rQTable As QueryTable)
    Set qt = rQTable
End Property

Private Sub qt_AfterRefresh(ByVal Success As Boolean)
    If Success Then
        Me.QTable.Parent.Range("A1").Value = _
            "Last updated: " & Format(Now, "mm-dd-yyyy hh:mm:ss")
    End If
End Sub

