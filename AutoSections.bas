Attribute VB_Name = "Module1"
Sub AutoSections()

    Dim intSlide As Integer
    Dim strNotes As String
    Dim tabSectionNames() As String
    Dim tabSectionSlides() As Integer
    Dim sec As Integer
    Dim secNumber As Integer
    Dim thisSection As Integer
    Dim texte As String
    Dim debut As Integer
    Dim longueur As Integer

    Dim width As Integer
    Dim height As Integer
    Dim barWidth As Integer
    Dim Sep As String
    Dim normalColor
    Dim emphColor
    Dim BackgroundColor
    Dim font As String
    Dim size As Integer
    Dim j As Integer
    Dim visibleSlides As Integer

    visibleSlides = 0
    sec = 1

    ' PARAMETERS

    Sep = " "
    normalColor = RGB(173, 185, 202)
    emphColor = RGB(68, 84, 106)
    BackgroundColor = RGB(222, 235, 247)

    size = 12
    font = "Calibri"

    width = ActivePresentation.PageSetup.SlideWidth
    height = ActivePresentation.PageSetup.SlideHeight

    With ActivePresentation

        ' Identify sections from slide notes
        For intSlide = 1 To .Slides.Count

            If .Slides(intSlide).NotesPage.Shapes.Placeholders(2).TextFrame.HasText Then
                strNotes = .Slides(intSlide).NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Lines(1).Text
            Else
                strNotes = ""
            End If
            
            strNotes = Replace(strNotes, vbLf, "")
            strNotes = Replace(strNotes, vbCr, "")

            If InStr(strNotes, "Section:") = 1 Then
                ReDim Preserve tabSectionNames(sec + 1)
                ReDim Preserve tabSectionSlides(sec + 1)

                tabSectionNames(sec) = Mid(strNotes, 9)
                tabSectionSlides(sec) = intSlide

                sec = sec + 1
            End If

            If Not .Slides(intSlide).SlideShowTransition.Hidden = True Then
                visibleSlides = visibleSlides + 1
            End If

        Next intSlide

        secNumber = sec - 1

        ' Add progress bars and section labels
        For intSlide = 1 To .Slides.Count

            texte = ""

            For sec = 1 To secNumber

                thisSection = 0

                If intSlide >= tabSectionSlides(sec) Then

                    If sec = secNumber Then
                        thisSection = 1
                    ElseIf intSlide < tabSectionSlides(sec + 1) Then
                        thisSection = 1
                    End If

                End If

                If thisSection = 1 Then
                    debut = Len(texte) + 1
                    longueur = Len(tabSectionNames(sec)) + 1
                End If

                If Not texte = "" Then
                    texte = texte & Sep
                End If

                texte = texte & tabSectionNames(sec)

            Next sec

            ' Remove existing shapes if rerunning macro
            j = 1
            While j <= .Slides(intSlide).Shapes.Count

                If .Slides(intSlide).Shapes(j).Name = "MyText" Then
                    .Slides(intSlide).Shapes(j).Delete

                ElseIf .Slides(intSlide).Shapes(j).Name = "MyBar1" Then
                    .Slides(intSlide).Shapes(j).Delete

                ElseIf .Slides(intSlide).Shapes(j).Name = "MyBar2" Then
                    .Slides(intSlide).Shapes(j).Delete

                Else
                    j = j + 1
                End If

            Wend

            If intSlide >= tabSectionSlides(1) Then

                ' Section titles at top
                With .Slides(intSlide).Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=30, Top:=16, width:=width - 60, height:=40)

                    With .TextFrame.TextRange
                        .Text = texte
                        .font.size = size
                        .font.Name = font
                        .font.Color.RGB = normalColor
                        .Characters(debut, longueur).font.Bold = True
                        .Characters(debut, longueur).font.Color.RGB = emphColor
                    End With

                    .Name = "MyText"

                End With

                ' Background progress bar (bottom)
                With .Slides(intSlide).Shapes.AddShape( _
                    Type:=msoShapeRectangle, _
                    Left:=30, _
                    Top:=height - 10, _
                    width:=width - 60, _
                    height:=4)

                    .Fill.BackColor.RGB = BackgroundColor
                    .Fill.ForeColor.RGB = BackgroundColor
                    .Line.Visible = False
                    .Name = "MyBar1"

                End With

                ' Filled progress bar
                barWidth = CInt((width - 62#) * (intSlide - 1#) / (visibleSlides - 1#))

                With .Slides(intSlide).Shapes.AddShape( _
                    Type:=msoShapeRectangle, _
                    Left:=31, _
                    Top:=height - 10, _
                    width:=barWidth, _
                    height:=4)

                    .Fill.BackColor.RGB = emphColor
                    .Fill.ForeColor.RGB = emphColor
                    .Line.Visible = False
                    .Name = "MyBar2"

                End With

            End If

        Next intSlide

    End With

End Sub

