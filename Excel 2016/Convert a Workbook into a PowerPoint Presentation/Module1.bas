Attribute VB_Name = "Module1"
Sub Macro101()
    
'Step 1:  Declare your variables
    Dim pp As PowerPoint.Application
    Dim PPPres As PowerPoint.Presentation
    Dim PPSlide As PowerPoint.Slide
    Dim xlwksht As Excel.Worksheet
    Dim MyRange As String
    Dim MyTitle As String
    Dim Slidecount As Long
    
    
'Step 2:  Open PowerPoint, add a new presentation and make visible
    Set pp = New PowerPoint.Application
    Set PPPres = pp.Presentations.Add
    pp.Visible = True
    
        
'Step 3:  Set the ranges for your data and title
    MyRange = "A1:J29"
    
    
'Step 4:  Start the loop through each worksheet
    For Each xlwksht In ActiveWorkbook.Worksheets
    xlwksht.Select
    Application.Wait (Now + TimeValue("0:00:1"))
    MyTitle = xlwksht.Range("C20").Value
    

'Step 5:  Copy the range as picture
    xlwksht.Range(MyRange).CopyPicture _
    Appearance:=xlScreen, Format:=xlPicture
    
    
'Step 6:  Count slides and add new slide as next available slide number
    Slidecount = PPPres.Slides.Count
    Set PPSlide = PPPres.Slides.Add(Slidecount + 1, ppLayoutTitleOnly)
    PPSlide.Select
    
         
'Step 7:  Paste the picture and adjust its position
    PPSlide.Shapes.Paste.Select
    pp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
    pp.ActiveWindow.Selection.ShapeRange.Top = 100
    
        
'Step 8:  Add the title to the slide then move to next worksheet
    PPSlide.Shapes.Title.TextFrame.TextRange.Text = MyTitle
    Next xlwksht
    
            
'Step 9:  Memory Cleanup
    pp.Activate
    Set PPSlide = Nothing
    Set PPPres = Nothing
    Set pp = Nothing
            
End Sub
