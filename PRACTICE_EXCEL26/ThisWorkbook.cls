VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    Dim objVBProj As VBProject
    Dim objRef As Reference
    Dim refBroken As Boolean

    Set objVBProj = ThisWorkbook.VBProject

    ' Loop through the selected references in
    ' the References dialog box
    For Each objRef In objVBProj.References
    ' If the reference is broken, get its name and its GUID
        If objRef.IsBroken Then
          Debug.Print objRef.Name
          Debug.Print objRef.GUID
          refBroken = True
       End If
    Next
    If refBroken = False Then
         Debug.Print "All references are valid."
    End If
End Sub


