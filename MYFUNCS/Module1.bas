Attribute VB_Name = "Module1"
Option Explicit

Function AddTwo(num1, num2)
Attribute AddTwo.VB_Description = "Returns the sum of two numbers"
Attribute AddTwo.VB_HelpID = 1000
Attribute AddTwo.VB_ProcData.VB_Invoke_Func = " \n14"
    AddTwo = num1 + num2
End Function

Function Squared(num)
Attribute Squared.VB_Description = "Returns the square of an argument"
Attribute Squared.VB_HelpID = 2000
Attribute Squared.VB_ProcData.VB_Invoke_Func = " \n14"
    Squared = num * num
End Function

 Sub SetOptions()
'   Set options for the AddTwo function
    Application.MacroOptions Macro:="AddTwo", _
        Description:="Returns the sum of two numbers", _
        Category:=14, _
        HelpContextID:=1000, _
        HelpFile:=ThisWorkbook.Path & "\Myfuncs.chm", _
        ArgumentDescriptions:=Array("The first number to add", _
          "The second number to add")
        
'   Set options for the Squared function
    Application.MacroOptions Macro:="Squared", _
        Description:="Returns the square of an argument", _
        Category:=14, _
        HelpContextID:=2000, _
        HelpFile:=ThisWorkbook.Path & "\Myfuncs.chm", _
        ArgumentDescriptions:=Array("The number to be squared")
End Sub

