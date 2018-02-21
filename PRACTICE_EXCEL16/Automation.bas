Attribute VB_Name = "Automation"
Option Explicit

Sub AccessViaAutomation()
  Dim objAccess As Access.Application
  Dim strPath As String

  On Error Resume Next

  Set objAccess = GetObject(, "Access.Application.15")
  If objAccess Is Nothing Then
      ' Get a reference to the Access Application object
      Set objAccess = New Access.Application
  End If

  strPath = "C:\Excel2013_HandsOn\Northwind 2007.accdb"

  ' Open the Employees table in the Northwind database
  With objAccess
      .OpenCurrentDatabase strPath
      .DoCmd.OpenTable "Employees", acViewNormal, acReadOnly
      If MsgBox("Do you want to make the Access " & vbCrLf _
          & "Application visible?", vbYesNo, _
          "Display Access") = vbYes Then
          .Visible = True
          MsgBox "Notice the Access Application icon " _
          & "now appears on the Windows taskbar."
      End If
      ' Close the database and quit Access
      .CloseCurrentDatabase
      .Quit
  End With

  Set objAccess = Nothing
End Sub



Sub OpenSecuredDB()
  Static objAccess As Access.Application
  Dim db As DAO.Database
  Dim strDb As String
    
  strDb = "C:\Excel2013_HandsOn\Med.mdb"

  Set objAccess = New Access.Application
  Set db = objAccess.DBEngine.OpenDatabase(Name:=strDb, _
      Options:=False, _
      ReadOnly:=False, _
      Connect:=";PWD=test")
  With objAccess
      .Visible = True
      .OpenCurrentDatabase strDb
  End With
  db.Close
  Set db = Nothing
End Sub


