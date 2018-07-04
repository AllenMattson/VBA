Attribute VB_Name = "GetAllFilesandFolders"
Option Explicit
'Requires a reference to Microsoft Scripting Runtime.
Public Function get_all_directory_files_with_wildcard _
 (ByVal tfolder As String, _
 ByVal getsubdirs As Boolean, _
 ByVal wildcard As String) _
 As String

'made by Alexander Triantafyllou
'alextriantf@yahoo.gr
'Athens-Greece

    
    Dim objfile As File
    Dim objfolder As Folder
    Dim fso As New FileSystemObject
   Dim kokovar As Variant
   Dim k As Long
   Dim wildext As String
   Dim wildexts As String
   Dim wildfirst As String
   Dim wildexte As String
   Dim wildfirsts As String
      Dim wildfirste As String
Dim examfirst As String
Dim examext As String
Dim afl_filetext As String

kokovar = Split(wildcard, ",")

    If tfolder <> "" Then


        For Each objfile In fso.GetFolder(tfolder).Files
            'do the stuff we want with the files
            For k = 0 To UBound(kokovar)
         wildext = LCase(cutgetExtension(kokovar(k)))
         wildfirst = LCase(Mid(kokovar(k), 1, Len(kokovar(k)) - Len(wildext) - 1))
            
       If InStr(1, wildext, "*") = 0 Then
       wildexts = "888NONE888"
       wildexte = "888NONE888"
       Else
       wildexts = Mid(wildext, 1, InStr(1, wildext, "*") - 1)
       wildexte = Mid(wildext, InStr(1, wildext, "*") + 1, Len(wildext) - InStr(1, wildext, "*"))
       End If
       
        If InStr(1, wildfirst, "*") = 0 Then
       wildfirsts = "888NONE888"
       wildfirste = "888NONE888"
       Else
       wildfirsts = Mid(wildfirst, 1, InStr(1, wildfirst, "*") - 1)
       wildfirste = Mid(wildfirst, InStr(1, wildfirst, "*") + 1, Len(wildfirst) - InStr(1, wildfirst, "*"))
       End If
            
        examfirst = LCase(cutgetName(cutfilename(CStr(objfile))))
        examext = LCase(cutgetExtension(CStr(objfile)))
            
        If wildexts = "888NONE888" Then
'we do not have a wildcard in the extension
        If wildfirsts = "888NONE888" Then
        'we do not have a wildcard neither on the beggining or the

'extension
        If examfirst = wildfirst And examext = wildext Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        End If
                
        Else
        
        'we do have a wildcard in the beggining but not in
'the extension
        If Mid(examfirst, 1, Len(wildfirsts)) = wildfirsts And _
        Mid(examfirst, Len(wildfirst) - Len(wildfirste) + 1, Len(wildfirste)) = wildfirste And wildext = examext Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        End If
        
        End If
        
        Else
        'we do not have a wildcard in the extension
        If wildfirsts = "888NONE888" Then
        'we do have a wildcard in the beggining but not in the
'extension
        If Mid(examext, 1, Len(wildexts)) = wildexts And _
        Mid(examext, Len(wildext) - Len(wildexte) + 1, Len(wildexte)) = wildexte Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        End If
            
        Else
        'we have a wildcard in both beggining and extension
        
            If Mid(examext, 1, Len(wildexts)) = wildexts And _
        Mid(examext, Len(wildext) - Len(wildexte) + 1, Len(wildexte)) = wildexte _
         And Mid(examfirst, 1, Len(wildfirsts)) = wildfirsts And _
        Mid(examfirst, Len(wildfirst) - Len(wildfirste) + 1, Len(wildfirste)) = wildfirste Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        End If
            
            End If
            
            End If
            
            'telos if
                    
            Next k
            
        Next

If getsubdirs Then

        For Each objfolder In fso.GetFolder(tfolder).SubFolders
           afl_filetext = afl_filetext & get_all_directory_files_with_wildcard(CStr(objfolder), getsubdirs, wildcard)
        Next
       
    End If
End If

Set fso = Nothing
get_all_directory_files_with_wildcard = afl_filetext
End Function

Public Function cutfilename(ByVal fname As String) As String
Dim spos As Integer
Dim ffn As String
spos = InStrRev(fname, "\")
ffn = Mid(fname, spos + 1, Len(fname) - spos)
cutfilename = ffn

End Function

Public Function cutgetExtension(ByVal fname As String)
Dim spos As Integer
Dim koko As String

spos = InStrRev(fname, ".")
If spos <> 0 Then
koko = Mid(fname, spos + 1, Len(fname) - spos)
End If

cutgetExtension = koko

End Function


Public Function cutgetName(ByVal fname As String)
Dim spos As Integer
Dim koko As String

spos = InStrRev(fname, ".")
If spos <> 0 Then
koko = Mid(fname, 1, spos - 1)
End If
cutgetName = koko

End Function



