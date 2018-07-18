Attribute VB_Name = "Module1"
Sub Macro99()

'Step 1:  Declare your variables
    Dim PP As PowerPoint.Application
    Dim PPPres As PowerPoint.Presentation
    Dim PPSlide As PowerPoint.Slide
    Dim SlideTitle As String
    
'Step 2:  Open PowerPoint and create new presentation
    Set PP = New PowerPoint.Application
    Set PPPres = PP.Presentations.Add
    PP.Visible = True
    
'Step 3:  Add new slide as slide 1 and set focus to it
    Set PPSlide = PPPres.Slides.Add(1, ppLayoutTitleOnly)
    PPSlide.Select

'Step 4:  Copy the range as a picture
    Sheets("Slide Data").Range("A1:J28").CopyPicture _
    Appearance:=xlScreen, Format:=xlPicture

'Step 5:  Paste the picture and adjust its position
    PPSlide.Shapes.Paste.Select
    PP.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
    PP.ActiveWindow.Selection.ShapeRange.Align msoAlignMiddles, True
 
'Step 6:  Add the title to the slide
    SlideTitle = "My First PowerPoint Slide"
    PPSlide.Shapes.Title.TextFrame.TextRange.Text = SlideTitle

'Step 7:  Memory Cleanup
    PP.Activate
    Set PPSlide = Nothing
    Set PPPres = Nothing
    Set PP = Nothing

End Sub


