VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQryTbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents qryTbl As QueryTable
Attribute qryTbl.VB_VarHelpID = -1

Private Sub qryTbl_BeforeRefresh(Cancel As Boolean)
   Response = MsgBox("Are you sure you " _
        & " want to refresh now?", vbYesNoCancel)
    If Response = vbNo Then Cancel = True
End Sub


Private Sub qryTbl_AfterRefresh(ByVal Success As _
         Boolean)
  If Success Then
      MsgBox "The data has been refreshed."
  Else
      MsgBox "The query failed."
  End If
End Sub


