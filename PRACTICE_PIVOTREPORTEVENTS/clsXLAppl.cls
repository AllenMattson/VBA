VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsXLAppl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Using the WithEvents keyword declare an object variable
' of type Application to respond to events
' the WithEvents keyword will cause the events associated with
' the application object to be exposed
' VBA will add the variable name to the Object drop down list
Private WithEvents pivTbl As Application
Attribute pivTbl.VB_VarHelpID = -1

' assign a reference to the object variable in the Class_Initialize event procedure
Private Sub Class_Initialize()
    ' the initialize event occurs as soon as a new instance of the class is created
    ' we create a new instance of the class in a standard module (Module1)
    ' set the object variable to point to the current instance of the application
    Set pivTbl = Application
End Sub

' create an event procedure for the event you want to trap
' this procedure displays a message after the PivotTable report has been updated
Private Sub pivTbl_SheetPivotTableUpdate( _
        ByVal Sh As Object, _
        ByVal Target As PivotTable)
    MsgBox Target.Name & _
      " report has been updated." & vbCrLf _
    & "The PivotReport is located in cells " & _
        Target.DataBodyRange.Address
End Sub



