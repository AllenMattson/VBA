Attribute VB_Name = "Module1"
Option Explicit

Function MONTHNAMES(Optional MIndex)
    Dim AllNames As Variant
    Dim MonthVal As Long
    
    AllNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    If IsMissing(MIndex) Then
        MONTHNAMES = AllNames
        Else
        Select Case MIndex
            Case Is >= 1
'            Determine month value (for example, 13=1)
             MonthVal = ((MIndex - 1) Mod 12)
             MONTHNAMES = AllNames(MonthVal)
          Case Is <= 0 ' Vertical array
             MONTHNAMES = Application.Transpose(AllNames)
         End Select
    End If
End Function

