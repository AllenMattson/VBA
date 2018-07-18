Attribute VB_Name = "QueryTableListObj"
Dim sampleQry As New clsQryTbl


Public Sub Auto_Open()
   ' connect the class module and its objects with the Query object
   Set sampleQry.qryTbl = ActiveSheet.ListObjects(1).QueryTable
End Sub




