Attribute VB_Name = "ImportCode"
Option Explicit

Sub ImportCodeModule()
      Dim Filt$, Title$, FileName$, Message As VbMsgBoxResult
      Do Until Message = vbNo
            'type of file to browse for
            Filt = "VB Files (*.bas; *.frm; *.cls)(*.bas; *.frm; *.cls)," & _
                   "*.bas;*.frm;*.cls"
            'caption for browser
            Title = "SELECT A FOLDER - CLICK OPEN TO IMPORT - " & _
                    "CANCEL TO QUIT"
            'browser
            FileName = Application.GetOpenFilename(FileFilter:=Filt, _
                                                   FilterIndex:=5, Title:=Title)
            On Error GoTo Finish    '< cancelled
            Application.VBE.ActiveVBProject.VBComponents.Import _
                        (FileName)
            'finished?
            Message = MsgBox(FileName & vbCrLf & " has been imported " & _
                             "- more imports?", vbYesNo, "More Imports?")
      Loop
Finish:
End Sub
