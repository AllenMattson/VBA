Attribute VB_Name = "HaveRecords"
Sub Test()
Dim rstTestData As ADODB.Recordset    'Recordset object
Dim lngRecords As Long                'Long variable to return
                                      'recordcount
Set rstTestData = New ADODB.Recordset
rstTestData.Open "SELECT * FROM Customers", ADOConnection

If HaveRecords(rstTestData) Then
   '...
End If

rstTestData.Close
Set rstTestData = Nothing
End Sub
Public Function HaveRecords(rstData As ADODB.Recordset, Optional ByRef lngRecordCount As Long) As Boolean
'*******************************************************************************
'* Name:  HaveRecords
'*
'* Description: This function will return true if the recordset has records and optionally return the amount of records
'* Date Created:  02/11/2001
'*
'* Created By: Adrian Henning
'*
'* Modified:
'*
'**************************************************************

    On Error GoTo HaveRecords_ERR

    HaveRecords = False

    If Not rstData Is Nothing Then
        If rstData.State = adStateOpen Then
            If (Not rstData.EOF) And (Not rstData.BOF) Then
                rstData.MoveFirst
                lngRecordCount = rstData.RecordCount
                HaveRecords = True
            End If
        Else
            HaveRecords = False
        End If
    Else
        HaveRecords = False
    End If

    Exit Function

HaveRecords_ERR:
    MsgBox "[" & Err.Number & "] " & Err.Description, vbInformation, "HaveRecords - Error"

End Function
