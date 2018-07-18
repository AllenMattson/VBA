Attribute VB_Name = "Compatibility"
Option Explicit

Function IsCompatible(strVersion As String) As Boolean
    If CInt(strVersion) >= 12 Then
        If ActiveWorkbook.Excel8CompatibilityMode Then
            IsCompatible = False
        Else
            IsCompatible = True
        End If
    End If
End Function


Sub CheckCompatibility()
    Windows("SalesRegions.xls").Activate
    If Not IsCompatible(Application.Version) Then
        MsgBox "Excel 2007-2013 features will not work " & _
            "in this workbook.", vbCritical, _
            "Excel 97-2003 Compatibility Workbook"
    Else
        MsgBox "The are no compatibility issues."
    End If
End Sub


