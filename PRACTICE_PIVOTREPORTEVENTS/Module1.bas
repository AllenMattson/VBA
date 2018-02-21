Attribute VB_Name = "Module1"
' declare a module-level object variable to point to the instance of the class
Private xlPiv As clsXLAppl

' write a procedure that creates a new instance of the class
' you must run this procedure to enable events
' this procedure will trigger the Initialize event procedure in the class module
' the object variable will be loaded into memory and ready for events

Public Sub InitEvents()
    Set xlPiv = New clsXLAppl
End Sub

' after you've run the above procedure, change some data in the pivot table source
' click anywhere in the pivot table report and choose the Refresh Data button
' on the Pivot Table toolbar (or right-click in the pivot table and choose Refresh
' from the shortcut menu)
