Attribute VB_Name = "StartChrome"
Sub StartChrome()
Dim ChromePath As String: ChromePath = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe""": Shell (ChromePath & " -url http:northwallapplications.com")
End Sub

