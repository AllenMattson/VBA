Attribute VB_Name = "Dev"
Private Property Get BlankWBPath() As String
    BlankWBPath = GetFullPath("VBA-Web - Blank.xlsm")
End Property
Private Property Get ExampleWBPath() As String
    ExampleWBPath = GetFullPath("examples/VBA-Web - Example.xlsm")
End Property
Private Property Get SpecsWBPath() As String
    SpecsWBPath = GetFullPath("specs/VBA-Web - Specs.xlsm")
End Property
Private Property Get AsyncSpecsWBPath() As String
    AsyncSpecsWBPath = GetFullPath("specs/VBA-Web - Specs - Async.xlsm")
End Property

Public Sub Release(Version As String)
    If VBA.InStr(1, Version, "v") <> 1 Then
        Version = "v" & Version
    End If
    
    Debug.Print vbNewLine & "Releasing " & Version & "..."
    
    Debug.Print "1. Releasing Blank..."
    ReleaseBlank
    Debug.Print "2. Releasing Example..."
    ' For some strange reason the backups for src are not removed on the first try for example
    ' but they are removed when run twice...
    ReleaseExample
    ReleaseExample
    Debug.Print "3. Releasing Specs..."
    ReleaseSpecs
    Debug.Print "4. Releasing Async..."
    ReleaseAsyncSpecs
    Debug.Print "5. Releasing Installer..."
    ReleaseInstaller Version
    
    Debug.Print "DONE!"
End Sub

Public Sub Specs()
    Dim Selections As VBAWebSelections
    Selections.Src = True
    Selections.Auth = True
    Selections.Specs = True
    Selections.AuthSpecs = True
    
    VBAWebInstaller.ExportSelections SpecsWBPath, Selections, False
End Sub

Public Sub Async()
    Dim Selections As VBAWebSelections
    Selections.Src = True
    Selections.AsyncWrapper = True
    Selections.AsyncSpecs = True
    
    VBAWebInstaller.ExportSelections AsyncSpecsWBPath, Selections, False
End Sub

Public Sub Example()
    Dim Selections As VBAWebSelections
    Selections.Src = True
    Selections.Auth = True
        
    VBAWebInstaller.ExportSelections ExampleWBPath, Selections, False
End Sub

Public Sub Import(Selection As String, ToWB As String)
    Dim WorkbookPath As String
    Dim Selections As VBAWebSelections
    
    Select Case ToWB
    Case "Specs"
        WorkbookPath = SpecsWBPath
    Case "Async"
        WorkbookPath = AsyncSpecsWBPath
    Case "Example"
        WorkbookPath = ExampleWBPath
    End Select
    
    Select Case Selection
    Case "Src"
        Selections.Src = True
    Case "Specs"
        Selections.Specs = True
    Case "Auth"
        Selections.Auth = True
    Case "All"
        Selections.Src = True
        Selections.Auth = True
        Selections.Specs = True
        Selections.AuthSpecs = True
    End Select

    VBAWebInstaller.InstallSelections WorkbookPath, Selections, False
End Sub

Public Sub Export(Src As String, FromWB As String)
    ' TODO
End Sub

Private Sub ReleaseBlank()
    Dim Selections As VBAWebSelections
    Selections.Src = True
    Selections.VBADictionary = True
    
    VBAWebInstaller.InstallSelections BlankWBPath, Selections, False
End Sub

Private Sub ReleaseSpecs()
    Dim Selections As VBAWebSelections
    Selections.Src = True
    Selections.Auth = True
    Selections.Specs = True
    Selections.AuthSpecs = True
    
    VBAWebInstaller.InstallSelections SpecsWBPath, Selections, False
End Sub

Private Sub ReleaseAsyncSpecs()
    Dim Selections As VBAWebSelections
    Selections.Src = True
    Selections.AsyncWrapper = True
    Selections.AsyncSpecs = True
    
    VBAWebInstaller.InstallSelections AsyncSpecsWBPath, Selections, False
End Sub

Public Sub ReleaseExample()
    Dim Selections As VBAWebSelections
    Selections.Src = True
    Selections.Auth = True
    
    VBAWebInstaller.InstallSelections ExampleWBPath, Selections, False
End Sub

Private Sub ReleaseInstaller(Version)
    ThisWorkbook.Sheets("Install VBA-Web").Range("Version") = Version
    ThisWorkbook.Sheets("Install VBA-Web").Reset
End Sub

Private Function SrcToSelections(Src As String) As VBAWebSelections
    Dim Selections As VBAWebSelections

    Select Case VBA.UCase$(Src)
    Case "SRC"
        Selections.Src = True
    Case "AUTH"
        Selections.Auth = True
    Case "ASYNC"
        Selections.AsyncWrapper = True
    Case "SPECS"
        Selections.Specs = True
    Case "AUTH-SPECS"
        Selections.AuthSpecs = True
    Case "ASYNC-SPECS"
        Selections.AsyncSpecs = True
    End Select
    
    SrcToSelections = Selections
End Function

Private Function WBToPath(Wb As String) As String
    Select Case VBA.UCase$(Wb)
    Case "BLANK"
        WBToPath = BlankWBPath
    Case "EXAMPLE"
        WBToPath = ExampleWBPath
    Case "SPECS"
        WBToPath = SpecsWBPath
    Case "ASYNC-SPECS"
        WBToPath = AsyncSpecsWBPath
    Case Else
        WBToPath = Wb
    End Select
End Function

Private Function GetFullPath(RelativePath As String) As String
    GetFullPath = ThisWorkbook.Path & Application.PathSeparator & VBA.Replace$(RelativePath, "/", Application.PathSeparator)
End Function
