Attribute VB_Name = "Module1"
Option Explicit

Public Sub CalculateCommission()
    Dim clsSalesReps As CSalesReps
    Dim i As Long
    'Create a new parent object and fill it with child objects
    Set clsSalesReps = New CSalesReps
    FillSalesReps clsSalesReps
    
    'Loop through all the reps and print commissions
    For i = 1 To clsSalesReps.Count
        With clsSalesReps.Item(i)
            Debug.Print .SalesRep, _
                Format(.Commission, "$#,##0.00")
        End With
    Next i
End Sub

Public Sub FillSalesReps(ByRef clsSalesReps As CSalesReps)
    Dim i As Long
    Dim clsSalesRep As CSalesRep
    Dim loReps As ListObject
    
    Set loReps = Sheet1.ListObjects(1)
    'loop through all the sales reps
    For i = 1 To loReps.ListRows.Count
        'create a new sales rep object
        Set clsSalesRep = New CSalesRep
        'Set the properties
        With loReps.ListRows(i).Range
            clsSalesRep.SalesRepID = .Cells(1).Value
            clsSalesRep.SalesRep = .Cells(2).Value
            clsSalesRep.CommissionRate = .Cells(3).Value
            clsSalesRep.Threshold = .Cells(4).Value
        End With
        'Add the child to the parent class
        clsSalesReps.Add clsSalesRep
        'Fill invoices for this rep
        FillInvoices clsSalesRep
    Next i
End Sub

Public Sub FillInvoices(ByRef clsSalesRep As CSalesRep)
    Dim i As Long
    Dim clsInvoice As CInvoice
    Dim loInv As ListObject
    
    'create a variable for the table
    Set loInv = Sheet2.ListObjects(1)
    'loop through the invoices table
    For i = 1 To loInv.ListRows.Count
        With loInv.ListRows(i).Range
            'Only if it's for this rep, add it
            If .Cells(4).Value = clsSalesRep.SalesRepID Then
                Set clsInvoice = New CInvoice
                clsInvoice.Invoice = .Cells(1).Value
                clsInvoice.InvoiceDate = .Cells(2).Value
                clsInvoice.Amount = .Cells(3).Value
                
                clsSalesRep.Invoices.Add clsInvoice
            End If
        End With
    Next i
End Sub
