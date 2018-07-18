VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents AppEvents As Application
Attribute AppEvents.VB_VarHelpID = -1

Private Sub AppEvents_WorkbookOpen _
  (ByVal Wb As Excel.Workbook)
    Call UpdateLogFile(Wb)
End Sub


