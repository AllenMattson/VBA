VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSalesReps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSalesReps As New Collection

Public Sub Add(clsSalesRep As CSalesRep)
    mSalesReps.Add clsSalesRep, CStr(clsSalesRep.SalesRepID)
End Sub

Public Property Get Count() As Long
    Count = mSalesReps.Count
End Property

Public Property Get Item(lId As Long) As CSalesRep
    Set Item = mSalesReps(lId)
End Property
