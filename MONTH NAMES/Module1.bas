Attribute VB_Name = "Module1"
Option Explicit

Function MonthNames(Optional MIndex)
    Dim AllNames As Variant
    Dim MonthVal As Long
    
    AllNames = Array("Jan", "Feb", "Mar", _
        "Apr", "May", "Jun", "Jul", "Aug", _
        "Sep", "Oct", "Nov", "Dec")
    If IsMissing(MIndex) Then
        MonthNames = AllNames
        Else
        Select Case MIndex
            Case Is >= 1
'            Determine month value (for example, 13=1)
             MonthVal = ((MIndex - 1) Mod 12)
             MonthNames = AllNames(MonthVal)
          Case Is <= 0 ' Vertical array
             MonthNames = Application.Transpose(AllNames)
         End Select
    End If
End Function

