Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Public vA() As String
Public N As Long


Sub MakeList()
    'loads an array with details of the files in the selected folder.
    
    Dim sFolder As String, bRecurse As Boolean
    
    'NOTE
    'The Windows folders My Music, My Videos, and My Pictures
    'generate (handled) error numbers 70,90,91 respectively, so are avoided.
    'These folders are denied access, but not the files, so you can copy
    'any files from these into user-produced folders for listing.
    
    'set folder and whether or not recursive search applies
    sFolder = "C:\Users\Allen\" '"C:\Users\My Folder\Documents\Computer Data\"
    bRecurse = True

    'erase any existing contents of the array
    Erase vA()  'public string array
        
    'this variable will accumulate the result of all recursions
    N = 0 'initialize an off-site counting variable
            
    'status bar message for long runs
    Application.StatusBar = "Loading array...please wait."
    
    'run the folder proc
    LoadArray sFolder, bRecurse
        
    If N = 0 Then
       Application.StatusBar = "No Files were found!"
       MsgBox "NO FILES FOUND"
       Application.StatusBar = ""
       Exit Sub
    Else
       'status bar message for long runs
       Application.StatusBar = "Done!"
       MsgBox "Done!" & vbCrLf & N & " Files listed."
       Application.StatusBar = ""
       Exit Sub
    End If

End Sub

Sub LoadArray(sFolder As String, bRecurse As Boolean)
    'loads dynamic public array vA() with recursive or flat file listing
       
    'The Windows folders My Music, My Videos, and My Pictures
    'generate error numbers 70,90,91 respectively, and are best avoided.
    
    Dim FSO As Object, SourceFolder As Object, sSuff As String, vS As Variant
    Dim SubFolder As Object, FileItem As Object, sPath As String
    Dim r As Long, Count As Long, m As Long, sTemp As String
    
    'm counts items in each folder run
    'N (public) accumulates m for recursive runs
    m = m + N
        
    On Error GoTo Errorhandler
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set SourceFolder = FSO.GetFolder(sFolder)
    
    For Each FileItem In SourceFolder.Files
        DoEvents
        sTemp = CStr(FileItem.Name)
        sPath = CStr(FileItem.Path)
        
        'get suffix from fileitem
        vS = Split(CStr(FileItem.Name), "."): sSuff = vS(UBound(vS))
        
        If Not FileItem Is Nothing Then 'add other file filter conditions to this existing one here
            m = m + 1 'increment this sourcefolder's file count
            'reset the array bounds
            ReDim Preserve vA(1 To 6, 0 To m)
            r = UBound(vA, 2)
                'store details for one file on the array row
                vA(1, r) = CStr(FileItem.Name)
                vA(2, r) = CStr(FileItem.Path)
                vA(3, r) = CLng(FileItem.Size)
                vA(4, r) = CDate(FileItem.DateCreated)
                vA(5, r) = CDate(FileItem.DateLastModified)
                vA(6, r) = CStr(sSuff)
        End If
    Next FileItem
    
    'increment public counter with this sourcefolder count
    N = m  'N is public
    
    'this bit is responsible for the recursion
    If bRecurse Then
        For Each SubFolder In SourceFolder.SubFolders
            LoadArray SubFolder.Path, True
        Next SubFolder
    End If
       
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    Set FSO = Nothing
    Exit Sub

Errorhandler:
    If Err.Number <> 0 Then
        Select Case Err.Number
        Case 70 'access denied
            'MsgBox "error 70"
            Err.Clear
            Resume Next
        Case 91 'object not set
            'MsgBox "error 91"
            Err.Clear
            Resume Next
        Case Else
            'MsgBox "When m = " & m & " in LoadArray" & vbCrLf & _
            "Error Number :  " & Err.Number & vbCrLf & _
            "Error Description :  " & Err.Description
            Err.Clear
            Exit Sub 'goes to next subfolder - recursive
        End Select
    End If

End Sub

