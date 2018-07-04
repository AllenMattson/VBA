Attribute VB_Name = "Installer"
''
' Excel-Installer v0.0.0
' (c) Tim Hall
'
' Install/upgrade modules in Excel/Access
'
' Errors:
' 10021 - Project not found at path
' 10022 - VBA project object model must be trusted
' 10023 - Failed to remove backups after install
' 10024 - Failed to restore backups after error
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Enum ApplicationType
    ExcelApplication
    AccessApplication
End Enum

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public ProgressCallback As String
Public ShowProgress As Boolean

' ============================================= '
' Public Methods
' ============================================= '

''
' Install module in project
' - Safe: Backup existing module (if found) before install and restore on error
' - Smart: Keep project open if it's already open
'
' @param {String} ProjectPath
' @param {ModuleWrapper} Module
''
Public Sub InstallModule(ProjectPath As String, Module As InstallerModule)
    Dim Modules As New Collection
    Modules.Add Module
    
    InstallModules ProjectPath, Modules
End Sub

''
' Install modules in project
' - Safe: Backup existing module (if found) before install and restore on error
' - Smart: Keep project open if it's already open
'
' @param {String} ProjectPath
' @param {Collection of ModuleWrapper} Modules
''
Public Sub InstallModules(ProjectPath As String, Modules As Collection)
    Dim Project As New InstallerProject
    Project.Path = ProjectPath
    Project.HideProgress = Not ShowProgress
    Project.ProgressCallback = ProgressCallback
    
    Project.InstallModules Modules
End Sub

''
' Export modules from project
'
' @param {String} ProjectPath
' @param {Collection of ModuleWrapper} Modules
''
Public Sub ExportModules(ProjectPath As String, Modules As Collection)
    Dim Project As New InstallerProject
    Project.Path = ProjectPath
    Project.HideProgress = Not ShowProgress
    Project.ProgressCallback = ProgressCallback
    
    Project.ExportModules Modules
End Sub

' ============================================= '
' Private Functions
' ============================================= '

Public Function FullPath(RelativePath As String) As String
    FullPath = ThisWorkbook.Path & Application.PathSeparator & VBA.Replace$(RelativePath, "/", Application.PathSeparator)
End Function

Public Function GetFilename(Filepath As String) As String
    Dim FilepathParts() As String
    
    FilepathParts = VBA.Split(Filepath, Application.PathSeparator)
    GetFilename = FilepathParts(UBound(FilepathParts))
End Function

Public Function RemoveExtension(Filename As String) As String
    Dim FilenameParts() As String
    
    FilenameParts = VBA.Split(Filename, ".")
    If UBound(FilenameParts) > LBound(FilenameParts) Then
        ReDim Preserve FilenameParts(UBound(FilenameParts) - 1)
    End If
    
    RemoveExtension = VBA.Join(FilenameParts, ".")
End Function

Public Function FileExists(Filepath As String) As Boolean
#If Mac Then
    Dim Script As String
    Script = "tell application ""Finder""" & Chr(13) & _
        "exists file """ & Filepath & """" & Chr(13) & _
        "end tell" & Chr(13)
    FileExists = MacScript(Script)
#Else
    FileExists = VBA.Len(VBA.Dir(Filepath)) <> 0
#End If
End Function

Public Function DeleteFile(Filepath As String)
    If FileExists(Filepath) Then
#If Mac Then
        ' Use AppleScript to avoid long filename issues with Kill on Mac
        ' see: http://www.rondebruin.nl/mac/mac011.htm
        On Error Resume Next
        MacScript "tell application ""Finder""" & Chr(13) & _
            "do shell script ""rm "" & quoted form of posix path of """ & Filepath & """" & Chr(13) & _
            "end tell"
        On Error GoTo 0
#Else
        SetAttr Filepath, vbNormal
        Kill Filepath
#End If
    End If
End Function

Public Function GetExtension(Filepath As String)
    Dim FilenameParts() As String
    FilenameParts = VBA.Split(Filepath, ".")
    GetExtension = FilenameParts(UBound(FilenameParts))
End Function

