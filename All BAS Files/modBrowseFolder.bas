Attribute VB_Name = "modBrowseFolder"
Option Explicit
Option Private Module
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modBrowseFolder
' This contains the BrowseFolder function, which displays the standard Windows Browse For Folder
' dialog. It return the complete path of the selected folder or vbNullString if the user cancelled.
' It also contains the function BrowseFolderExplorer which presents the user with a Windows
' Explorer-like interface to pick the folder.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_DONTGOBELOWDOMAIN As Long = &H2
Private Const BIF_RETURNFSANCESTORS As Long = &H8
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Private Const BIF_BROWSEFORPRINTER As Long = &H2000
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000


Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszINSTRUCTIONS As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type


Private Declare Function SHGetPathFromIDListA Lib "shell32.dll" (ByVal pidl As Long, _
    ByVal pszBuffer As String) As Long

Private Declare Function SHBrowseForFolderA Lib "shell32.dll" (lpBrowseInfo As _
    BROWSEINFO) As Long


Private Const MAX_PATH = 260 ' Windows mandated


Function BrowseFolder(Optional ByVal DialogTitle As String) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BrowseFolder
' This displays the standard Windows Browse Folder dialog. It returns
' the complete path name of the selected folder or vbNullString if the
' user cancelled.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If DialogTitle = vbNullString Then
    DialogTitle = "Select A Folder"
End If

Dim uBrowseInfo As BROWSEINFO
Dim szBuffer As String
Dim lID As Long
Dim lRet As Long


With uBrowseInfo
    .hOwner = 0
    .pidlRoot = 0
    .pszDisplayName = String$(MAX_PATH, vbNullChar)
    .lpszINSTRUCTIONS = DialogTitle
    .ulFlags = BIF_RETURNONLYFSDIRS ' + BIF_USENEWUI
    .lpfn = 0
End With
szBuffer = String$(MAX_PATH, vbNullChar)
lID = SHBrowseForFolderA(uBrowseInfo)

If lID Then
    ''' Retrieve the path string.
    lRet = SHGetPathFromIDListA(lID, szBuffer)
    If lRet Then
        BrowseFolder = Left$(szBuffer, InStr(szBuffer, vbNullChar) - 1)
    End If
End If

End Function

Function BrowseFolderExplorer(Optional DialogTitle As String, _
    Optional ViewType As MsoFileDialogView = _
        MsoFileDialogView.msoFileDialogViewSmallIcons, _
    Optional InitialDirectory As String) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BrowseFolderExplorer
' This provides an Explorer-like Folder Open dialog.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim fDialog  As Office.FileDialog
    Dim varFile As Variant
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    fDialog.InitialView = ViewType
    With fDialog
        If Dir(InitialDirectory, vbDirectory) <> vbNullString Then
            .InitialFileName = InitialDirectory
        Else
            .InitialFileName = CurDir
        End If
        .Title = DialogTitle
        
        If .Show = True Then
            ' user picked a folder
            BrowseFolderExplorer = .SelectedItems(1)
        Else
            ' user cancelled
            BrowseFolderExplorer = vbNullString
        End If
    End With
End Function

