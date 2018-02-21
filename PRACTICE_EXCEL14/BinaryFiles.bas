Attribute VB_Name = "BinaryFiles"
Sub EnterAndDisplay()
    Open "C:\Excel2013_ByExample\MyData.txt" _
        For Binary As #1                'Open the file MyData.txt for binary access as file number 1.
    MsgBox "Total bytes: " & LOF(1)     'Show the number of bytes on opening the file. (The file is currently empty.)
    fname = "Julitta"                   'Assign a value to the variable fname.
    Ln = Len(fname)                     'Assign to the variable ln the length of string stored in the variable fname.
    Put #1, , Ln                        'Enter the value of the variable ln in the binary file in the position of the next byte.
    MsgBox "The last byte: " & Loc(1)   'Display the position of the last byte.
    Put #1, , fname                     'Enter the contents of the variable fname in the next position.
    lname = "Korol"                     'Assign a value to the variable lname.
    Ln = Len(lname)                     'Assign to the variable ln the length of string stored in the variable lname.
    Put #1, , Ln                        'Enter the value of the variable ln in the binary file in the position of the next byte.
    Put #1, , lname                     'Enter the contents of the variable lname in the next byte position.
    MsgBox "The last byte: " & Loc(1)   'Display the position of the last byte.
    Get #1, 1, entry1                   'Read the value stored in the position of the first byte and assign it to the variable entry1.
    MsgBox entry1                       'Display the contents of the variable entry1.
    Get #1, , entry2                    'Read the next value and assign it to the variable entry2.
    MsgBox entry2                       'Display the contents of the variable entry2.
    Get #1, , entry3                    'Read the next value and store it in the variable entry3.
    MsgBox entry3                       'Display the contents of the variable entry3.
    Get #1, , entry4                    'Read the next value and store it in the variable entry4.
    MsgBox entry4                       'Display the contents of the variable entry4.
    Debug.Print entry1; entry2; entry3; entry4  'Print all the data in the Immediate window.
    Close #1                            'Close the file.
End Sub
