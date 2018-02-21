VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mInvoice As String
Private mInvoiceDate As Date
Private mAmount As Double

Public Property Let Invoice(ByVal sInvoice As String)
    mInvoice = sInvoice
End Property
Public Property Get Invoice() As String
    Invoice = mInvoice
End Property
Public Property Let InvoiceDate(ByVal dtInvoiceDate As Date)
    mInvoiceDate = dtInvoiceDate
End Property
Public Property Get InvoiceDate() As Date
    InvoiceDate = mInvoiceDate
End Property
Public Property Let Amount(ByVal dAmount As Double)
    mAmount = dAmount
End Property
Public Property Get Amount() As Double
    Amount = mAmount
End Property
