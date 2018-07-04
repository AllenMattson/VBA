Attribute VB_Name = "ListModules_Procedures"
Sub PrintModuleProcedureSheet()
Application.DisplayAlerts = False
On Error Resume Next
Sheets("ModuleList").Delete
Sheets.Add: ActiveSheet.Name = "ModuleList"
Application.DisplayAlerts = True
ListModules




'headers
Cells(2, 1).Value = "Procedure Name"
Cells(2, 2).Value = "Procedure Type"
Cells(2, 3).Value = "Comments"
Cells(2, 4).Value = "Module Name"
Cells(2, 5).Value = "Module Type"

Range("A2:E2").Font.Bold = True
Range("A1:E1").Merge
Cells(1, 1).Value = "Complete list of Modules and Procedures from " & ThisWorkbook.Name

With Cells(1, 1).Font
    .Bold = True
    .Size = 14
    .ColorIndex = 3
End With

Columns.AutoFit
End Sub
Sub ListModules()
Dim boolIsSub As Boolean
Dim boolHasComment As Boolean
Dim lngCtLines As Long
Dim lngLineNbr As Long
Dim lngAmtLines As Long
Dim strCompType As String
Dim strProcName As String
Dim strFindSub As String
Dim strFindComment As String
Dim VBProj As Object
Dim VBEPart As Object
Dim WS As Worksheet
Dim rngOutputRow As Range
Dim aCodeMod As CodeModule
Dim ProcKind As VBIDE.vbext_ProcKind

Set VBEPart = Application.VBE.ActiveVBProject.VBComponents
Set VBProj = ActiveWorkbook.VBProject
Set WS = ActiveWorkbook.Worksheets("ModuleList")
Set rngOutputRow = WS.Range("A3")

strFindSub = "Sub"
strFindComment = "'* Module: *"

For Each VBEPart In VBProj.VBComponents
strCompType = ComponentTypeToString(VBEPart.Type)
If strCompType = "UserForm" Or strCompType = "Code Module" Then
Set aCodeMod = VBEPart.CodeModule
rngOutputRow.Cells(1, 4).Value = VBEPart.Name
rngOutputRow.Cells(1, 5).Value = ComponentTypeToString(VBEPart.Type)
With aCodeMod
lngLineNbr = .CountOfDeclarationLines + 1
lngCtLines = .CountOfLines
Do Until lngLineNbr >= .CountOfLines
strProcName = .ProcOfLine(lngLineNbr, ProcKind)
rngOutputRow.Cells(1, 1).Value = strProcName
GoSub GetDetails
lngLineNbr = .ProcStartLine(strProcName, ProcKind) + _
.ProcCountLines(strProcName, ProcKind) + 1
Set rngOutputRow = rngOutputRow.Cells(2, 1)
Loop
End With
Set rngOutputRow = rngOutputRow.Cells(2, 1)
Else
GoTo ContinueSearch
End If
ContinueSearch:
Next VBEPart
Exit Sub
GetDetails:
boolIsSub = aCodeMod.Find(Target:=strFindSub, StartLine:=aCodeMod.ProcStartLine(strProcName, ProcKind), _
StartColumn:=1, EndLine:=lngLineNbr, EndColumn:=20, _
Wholeword:=True, MatchCase:=False, Patternsearch:=False)
If boolIsSub = True Then
'Debug.Print strProcName + " is Sub"
rngOutputRow.Cells(1, 2).Value = "Sub"
Else
'Debug.Print strProcName + " is Function"
rngOutputRow.Cells(1, 2).Value = "Function"
End If
boolHasComment = aCodeMod.Find(Target:=strFindComment, StartLine:=aCodeMod.ProcStartLine(strProcName, ProcKind), _
StartColumn:=1, EndLine:=(aCodeMod.ProcStartLine(strProcName, ProcKind) + 5), _
EndColumn:=20, Wholeword:=True, MatchCase:=False, Patternsearch:=False)
If boolHasComment = True Then
'Debug.Print strProcName + " has Comment"
rngOutputRow.Cells(1, 3).Value = "has Comment"
Else
'Debug.Print strProcName + " Comment is missing"
rngOutputRow.Cells(1, 3).Value = "Comment is missing"
End If
boolIsSub = False
boolHasComment = False
Return

End Sub
Function ComponentTypeToString(ComponentType As VBIDE.vbext_ComponentType) As String

Select Case ComponentType
Case vbext_ct_ActiveXDesigner
ComponentTypeToString = "ActiveX Designer"
Case vbext_ct_ClassModule
ComponentTypeToString = "Class Module"
Case vbext_ct_Document
ComponentTypeToString = "Document Module"
Case vbext_ct_MSForm
ComponentTypeToString = "UserForm"
Case vbext_ct_StdModule
ComponentTypeToString = "Code Module"
Case Else
ComponentTypeToString = "Unknown Type: " & CStr(ComponentType)
End Select
End Function
Function ProcKindString(ProcKind As VBIDE.vbext_ProcKind) As String
Select Case ProcKind
Case vbext_pk_Get
ProcKindString = "Property Get"
Case vbext_pk_Let
ProcKindString = "Property Let"
Case vbext_pk_Set
ProcKindString = "Property Set"
Case vbext_pk_Proc
ProcKindString = "Sub Or Function"
Case Else
ProcKindString = "Unknown Type: " & CStr(ProcKind)
End Select
End Function

