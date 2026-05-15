Attribute VB_Name = "Module1"
Sub CreateMasterGridSlide()
    Dim oPres As Presentation
    Dim oNewSlide As Slide
    Dim oBg As Shape
    Dim i As Long, col As Long, row As Long
    Dim cols As Long, rows As Long
    Dim slideW As Single, slideH As Single
    Dim cellW As Single, cellH As Single
    Dim availW As Single, availH As Single
    Dim aspect As Single
    Dim margin As Single, gap As Single
    Dim gridW As Single, gridH As Single
    Dim gridLeft As Single, gridTop As Single
    Dim tempFolder As String, imgPath As String
    Dim totalSlides As Long
    
    '--- CONFIGURE HERE ---
    cols = 6                ' columns
    rows = 7                ' rows  (cols x rows must be >= number of slides)
    margin = 24             ' padding around entire grid (points)
    gap = 10                ' spacing between slides (points)
    Dim bgR As Integer: bgR = 202    ' light grey background, use 64 for darker grey
    Dim bgG As Integer: bgG = 202
    Dim bgB As Integer: bgB = 202
    '----------------------
    
    Set oPres = ActivePresentation
    totalSlides = oPres.Slides.Count
    
    tempFolder = Environ("HOME") & "/PPTGrid/"
    If Dir(tempFolder, vbDirectory) = "" Then MkDir tempFolder
    
    ' Export each original slide as PNG
    For i = 1 To totalSlides
        imgPath = tempFolder & "slide_" & Format(i, "000") & ".png"
        oPres.Slides(i).Export imgPath, "PNG"
    Next i
    
    slideW = oPres.PageSetup.SlideWidth
    slideH = oPres.PageSetup.SlideHeight
    aspect = slideW / slideH
    
    ' Add new blank slide for the grid
    Set oNewSlide = oPres.Slides.Add(totalSlides + 1, ppLayoutBlank)
    
    ' Grey background rectangle covering the slide
    Set oBg = oNewSlide.Shapes.AddShape(msoShapeRectangle, 0, 0, slideW, slideH)
    oBg.Fill.ForeColor.RGB = RGB(bgR, bgG, bgB)
    oBg.Line.Visible = msoFalse
    
    ' Available space per cell after margins + gaps
    availW = (slideW - 2 * margin - (cols - 1) * gap) / cols
    availH = (slideH - 2 * margin - (rows - 1) * gap) / rows
    
    ' Fit each cell to original slide aspect ratio (no stretching)
    If availW / availH > aspect Then
        cellH = availH
        cellW = cellH * aspect
    Else
        cellW = availW
        cellH = cellW / aspect
    End If
    
    ' Center the grid on the canvas
    gridW = cols * cellW + (cols - 1) * gap
    gridH = rows * cellH + (rows - 1) * gap
    gridLeft = (slideW - gridW) / 2
    gridTop = (slideH - gridH) / 2
    
    ' Place each slide image
    For i = 1 To totalSlides
        col = (i - 1) Mod cols
        row = (i - 1) \ cols
        imgPath = tempFolder & "slide_" & Format(i, "000") & ".png"
        oNewSlide.Shapes.AddPicture _
            FileName:=imgPath, _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=gridLeft + col * (cellW + gap), _
            Top:=gridTop + row * (cellH + gap), _
            Width:=cellW, _
            Height:=cellH
    Next i
    
    MsgBox "Done! Grid is on slide " & oNewSlide.SlideIndex
End Sub

