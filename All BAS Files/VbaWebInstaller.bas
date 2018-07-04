Attribute VB_Name = "VBAWebInstaller"
Public Type VBAWebSelections
    Src As Boolean
    VBADictionary As Boolean
    AsyncWrapper As Boolean
    HttpBasicAuthenticator As Boolean
    OAuth1Authenticator As Boolean
    OAuth2Authenticator As Boolean
    DigestAuthenticator As Boolean
    WindowsAuthenticator As Boolean
    GoogleAuthenticator As Boolean
    FacebookAuthenticator As Boolean
    TwitterAuthenticator As Boolean
    TodoistAuthenticator As Boolean
    EmptyAuthenticator As Boolean
    
    ' Dev-only
    Auth As Boolean
    Specs As Boolean
    AuthSpecs As Boolean
    AsyncSpecs As Boolean
End Type

Public Sub InstallSelections(WorkbookPath As String, Selections As VBAWebSelections, Optional ShowProgress As Boolean = True)
    Installer.ShowProgress = True
    Installer.ProgressCallback = "VBAWebInstaller.ShowProgress"
    
    Installer.InstallModules WorkbookPath, GetModulesForSelections(Selections)
End Sub

Public Sub ShowProgress(TotalCount As Long, CompletedCount As Long)
    ThisWorkbook.Sheets("Install VBA-Web").ShowProgress TotalCount, CompletedCount
End Sub

Public Sub ExportSelections(WorkbookPath As String, Selections As VBAWebSelections, Optional ShowProgress As Boolean = False)
    Installer.ShowProgress = ShowProgress
    Installer.ProgressCallback = "VBAWebInstaller.ShowProgress"
    
    Installer.ExportModules WorkbookPath, GetModulesForSelections(Selections)
End Sub

Private Function GetModulesForSelections(Selections As VBAWebSelections) As Collection
    Dim Module As InstallerModule
    Dim Modules As New Collection
    
    With Selections
        If .Src Then
            AddModule Modules, "WebHelpers", "src/WebHelpers.bas"
            AddModule Modules, "WebClient", "src/WebClient.cls"
            AddModule Modules, "WebRequest", "src/WebRequest.cls"
            AddModule Modules, "WebResponse", "src/WebResponse.cls"
            AddModule Modules, "IWebAuthenticator", "src/IWebAuthenticator.cls"
        End If
        If .AsyncWrapper Then
            AddModule Modules, "WebAsyncWrapper", "src/WebAsyncWrapper.cls"
        End If
        If .VBADictionary Then
            AddModule Modules, "Dictionary", "Dictionary.cls", FromLocal:=True
        End If
        If .HttpBasicAuthenticator Or .Auth Then
            AddModule Modules, "HttpBasicAuthenticator", "authenticators/HttpBasicAuthenticator.cls"
        End If
        If .OAuth1Authenticator Or .Auth Then
            AddModule Modules, "OAuth1Authenticator", "authenticators/OAuth1Authenticator.cls"
        End If
        If .OAuth2Authenticator Or .Auth Then
            AddModule Modules, "OAuth2Authenticator", "authenticators/OAuth2Authenticator.cls"
        End If
        If .DigestAuthenticator Or .Auth Then
            AddModule Modules, "DigestAuthenticator", "authenticators/DigestAuthenticator.cls"
        End If
        If .WindowsAuthenticator Or .Auth Then
            AddModule Modules, "WindowsAuthenticator", "authenticators/WindowsAuthenticator.cls"
        End If
        If .GoogleAuthenticator Or .Auth Then
            AddModule Modules, "GoogleAuthenticator", "authenticators/GoogleAuthenticator.cls"
        End If
        If .FacebookAuthenticator Or .Auth Then
            AddModule Modules, "FacebookAuthenticator", "authenticators/FacebookAuthenticator.cls"
        End If
        If .TwitterAuthenticator Or .Auth Then
            AddModule Modules, "TwitterAuthenticator", "authenticators/TwitterAuthenticator.cls"
        End If
        If .TodoistAuthenticator Or .Auth Then
            AddModule Modules, "TodoistAuthenticator", "authenticators/TodoistAuthenticator.cls"
        End If
        If .EmptyAuthenticator Or .Auth Then
            AddModule Modules, "EmptyAuthenticator", "authenticators/EmptyAuthenticator.cls"
        End If
        
        If .Specs Then
            AddModule Modules, "Specs_WebClient", "specs/Specs_WebClient.bas"
            AddModule Modules, "Specs_WebRequest", "specs/Specs_WebRequest.bas"
            AddModule Modules, "Specs_WebResponse", "specs/Specs_WebResponse.bas"
            AddModule Modules, "Specs_WebHelpers", "specs/Specs_WebHelpers.bas"
        End If
        If .AuthSpecs Then
            AddModule Modules, "Specs_IWebAuthenticator", "specs/Specs_IWebAuthenticator.bas"
            AddModule Modules, "Specs_HttpBasicAuthenticator", "specs/Specs_HttpBasicAuthenticator.bas"
            AddModule Modules, "Specs_OAuth1Authenticator", "specs/Specs_OAuth1Authenticator.bas"
            AddModule Modules, "Specs_OAuth2Authenticator", "specs/Specs_OAuth2Authenticator.bas"
            AddModule Modules, "Specs_DigestAuthenticator", "specs/Specs_DigestAuthenticator.bas"
            AddModule Modules, "Specs_GoogleAuthenticator", "specs/Specs_GoogleAuthenticator.bas"
            AddModule Modules, "SpecAuthenticator", "specs/SpecAuthenticator.cls"
        End If
        If .AsyncSpecs Then
            AddModule Modules, "Specs_WebAsyncWrapper", "specs/Specs_WebAsyncWrapper.bas"
        End If
    End With
    
    Set GetModulesForSelections = Modules
End Function

Private Sub AddModule(ByRef Modules As Collection, Name As String, Path As String, Optional FromLocal As Boolean = False)
    Dim Module As New InstallerModule
    Module.Name = Name
    Module.Path = Path
    Module.FromLocal = FromLocal
    
    Modules.Add Module
End Sub
