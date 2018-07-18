VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSalesRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSalesRepID As Long
Private mSalesRep As String
Private mCommissionRate As Double
Private mThreshold As Double
Private mInvoices As New CInvoices

Public Property Get Invoices() As CInvoices
    Set Invoices = mInvoices
End Property
Public Property Let SalesRepID(ByVal lSalesRepID As Long)
    mSalesRepID = lSalesRepID
End Property
Public Property Get SalesRepID() As Long
    SalesRepID = mSalesRepID
End Property
Public Property Let SalesRep(ByVal sSalesRep As String)
    mSalesRep = sSalesRep
End Property
Public Property Get SalesRep() As String
    SalesRep = mSalesRep
End Property
Public Property Let CommissionRate( _
    ByVal dCommissionRate As Double)
    mCommissionRate = dCommissionRate
End Property
Public Property Get CommissionRate() As Double
    CommissionRate = mCommissionRate
End Property
Public Property Let Threshold(ByVal dThreshold As Double)
    mThreshold = dThreshold
End Property
Public Property Get Threshold() As Double
    Threshold = mThreshold
End Property


Public Property Get Commission() As Double
    If Me.Invoices.Total < Me.Threshold Then
        Commission = 0
    Else
        Commission = (Me.Invoices.Total - Me.Threshold) _
            * Me.CommissionRate
    End If
End Property
