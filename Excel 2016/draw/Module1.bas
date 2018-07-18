Attribute VB_Name = "Module1"
Option Explicit

Function DRAWONE(Rng As Variant, Optional Recalc As Boolean = False)
Attribute DRAWONE.VB_Description = "Displays the contents of a random cell from a range"
Attribute DRAWONE.VB_ProcData.VB_Invoke_Func = " \n5"
'    Chooses one cell at random from a range

'    Make function volatile if Recalc is True
     Application.Volatile Recalc

'    Determine a random cell
     DRAWONE = Rng(Int((Rng.Count) * Rnd + 1))
End Function



Sub DescribeFunction()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim FuncCat As Long
    Dim Arg1Desc As String, Arg2Desc As String

    FuncName = "DRAWONE"
    FuncDesc = "Displays the contents of a random cell from a range"
    FuncCat = 5 'Lookup & Reference
    Arg1Desc = "The range that contains the values"
    Arg2Desc = "(Optional) If False or missing, a new cell is not "
    Arg2Desc = Arg2Desc & "selected when recalculated. If True, a "
    Arg2Desc = Arg2Desc & "new cell is selected when recalculated."

    Application.MacroOptions _
        Macro:=FuncName, _
        Description:=FuncDesc, _
        Category:=FuncCat, _
        ArgumentDescriptions:=Array(Arg1Desc, Arg2Desc)
End Sub

