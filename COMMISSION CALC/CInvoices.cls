VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CInvoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mInvoices As New Collection

Public Sub Add(clsInvoice As CInvoice)
    mInvoices.Add clsInvoice, clsInvoice.Invoice
End Sub

Public Property Get Count() As Long
    Count = mInvoices.Count
End Property

Public Property Get Total() As Double
    Dim i As Long
    For i = 1 To mInvoices.Count
        Total = Total + mInvoices.Item(i).Amount
    Next i
End Property
