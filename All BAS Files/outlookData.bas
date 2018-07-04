Attribute VB_Name = "outlookData"
Option Explicit
Const sBook = "Global Address List"
Public Sub getOutLookData()
    Dim od As cOutlookAddressbook
    Dim oa As cOutlookApp
    Dim rData As Range, dSets As cDataSets, ds As cDataSet
    ' kick off an outlook session
    Set oa = New cOutlookApp
    If oa.Init Then
        ' get the required addresslist
        Set od = New cOutlookAddressbook
        If od.Init(oa, sBook) Then
            ' you could replace this with a dialogue to get the range to be processed
            Set rData = getLikelyColumnRange
            'create a cdataset of that range
            Set dSets = New cDataSets
            With dSets
                .Create
                .Init rData, , "data"
            End With
            Set ds = dSets.DataSet("data")
            ' now populate - will only actually commit to the sheet if all was well
            ' using the alias field/column as the key
            If od.Populate(ds, "alias") Then
                ds.Commit
            End If

            od.Destroy
        Else
            MsgBox ("Couldnt open address book " & sBook)
            
        End If
        ' clear up
        oa.Destroy
        Set od = Nothing
        Set oa = Nothing
    Else
        MsgBox ("Couldnt start outlook")
    End If
End Sub
