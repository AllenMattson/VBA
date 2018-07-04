Attribute VB_Name = "ListFilesandFolders_Recursion"
Public useitOnlytoSetFolder As FileSystemObject
Public mm As Folder
Public ttt As Folder
Public zzz As File
Function recursiveList(a As String, delimRcvd As String)

Dim delimToSend As String             'CANNOT BE A PUBLIC VARIABLE, MUST BE DECLARED HERE
    delimToSend = delimRcvd & "*   "

Set mm = useitOnlytoSetFolder.GetFolder(a)

If mm.Files.Count > 0 Then              '''---------------------------------------
                                        '''  This section is not essenial to the recursive algorithm
  List2.AddItem delimRcvd & mm.Path     '''  NOT NEEDED FOR RECURSION, ONLY USED TO LIST EACH FILE IN EACH FOLDER
                                        '''
  For Each zzz In mm.Files              '''
                                        '''
    List2.AddItem delimRcvd & zzz.Name    '  LISTS EACH FILE IN A LISTBOX
                                        '''
  Next                                  '''
                                        '''
End If                                  '''---------------------------------------


 If mm.SubFolders.Count > 0 Then

         
    For Each ttt In mm.SubFolders
    
      List1.AddItem delimRcvd & ttt.Path   ' LISTS EACH FOLDER IN A LISTBOX
        
        Call recursiveList(ttt.Path, delimToSend)  'Recursive Call
        
    Next

 End If

End Function


Private Sub Command1_Click()
Do While List1.ListCount > 0 '       Will clear the list box
List1.RemoveItem (0)         '       before we change a directory
Loop                         '       There could have been other values in the box from a prior directory listing

Call recursiveList(Dir1.Path, "") 'Begins the recursion process

End Sub

Private Sub Form_Load()
Dir1.Path = App.Path
Set useitOnlytoSetFolder = New FileSystemObject
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set useitOnlytoSetFolder = Nothing
End Sub
