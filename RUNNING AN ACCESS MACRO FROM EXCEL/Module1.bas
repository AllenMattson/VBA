Attribute VB_Name = "Module1"
Sub Macro93()

'Step 1:  Declare your variables
    Dim AC As Access.Application
    

'Step 2:  Start Access and open the target database
    Set AC = New Access.Application
             AC.OpenCurrentDatabase _
            ("C:\Temp\YourAccessDatabse.accdb")
    

'Step 3:  Open the target report and send to Word
    With AC
        .DoCmd.RunMacro "MyMacro"
        .Quit
    End With
      
End Sub

