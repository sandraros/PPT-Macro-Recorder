Attribute VB_Name = "Factory"
Function New_iApplication(iApplication As Application) As iApplication
    Set New_iApplication = Utility.GetObject(iApplication)
    If New_iApplication Is Nothing Then
        Set New_iApplication = Utility.AddObject(iApplication, New iApplication)
    End If
End Function

Function New_iColorFormat(iColorFormat As PowerPoint.ColorFormat) As iColorFormat
    Set New_iColorFormat = Utility.GetObject(iColorFormat)
    If New_iColorFormat Is Nothing Then
        Set New_iColorFormat = Utility.AddObject(iColorFormat, New iColorFormat)
    End If
End Function

Function New_iDocumentWindow(iDocumentWindow As DocumentWindow) As iDocumentWindow
    Set New_iDocumentWindow = Utility.GetObject(iDocumentWindow)
    If New_iDocumentWindow Is Nothing Then
        Set New_iDocumentWindow = Utility.AddObject(iDocumentWindow, New iDocumentWindow)
    End If
End Function

Function New_iDocumentWindows(iDocumentWindows As DocumentWindows) As iDocumentWindows
    Set New_iDocumentWindows = Utility.GetObject(iDocumentWindows)
    If New_iDocumentWindows Is Nothing Then
        Set New_iDocumentWindows = Utility.AddObject(iDocumentWindows, New iDocumentWindows)
    End If
End Function

Function New_iFillFormat(iFillFormat As PowerPoint.FillFormat) As iFillFormat
    Set New_iFillFormat = Utility.GetObject(iFillFormat)
    If New_iFillFormat Is Nothing Then
        Set New_iFillFormat = Utility.AddObject(iFillFormat, New iFillFormat)
    End If
End Function

Function New_iFont(iFont As Font) As iFont
    Set New_iFont = Utility.GetObject(iFont)
    If New_iFont Is Nothing Then
        Set New_iFont = Utility.AddObject(iFont, New iFont)
    End If
End Function

Function New_iFont2(iFont2 As Font2) As iFont2
    Set New_iFont2 = Utility.GetObject(iFont2)
    If New_iFont2 Is Nothing Then
        Set New_iFont2 = Utility.AddObject(iFont2, New iFont2)
    End If
End Function

Function New_iGlowFormat(iGlowFormat As GlowFormat) As iGlowFormat
    Set New_iGlowFormat = Utility.GetObject(iGlowFormat)
    If New_iGlowFormat Is Nothing Then
        Set New_iGlowFormat = Utility.AddObject(iGlowFormat, New iGlowFormat)
    End If
End Function

Function New_iGradientStop(iGradientStop As GradientStop) As iGradientStop
    Set New_iGradientStop = Utility.GetObject(iGradientStop)
    If New_iGradientStop Is Nothing Then
        Set New_iGradientStop = Utility.AddObject(iGradientStop, New iGradientStop)
    End If
End Function

Function New_iGradientStops(iGradientStops As GradientStops) As iGradientStops
    Set New_iGradientStops = Utility.GetObject(iGradientStops)
    If New_iGradientStops Is Nothing Then
        Set New_iGradientStops = Utility.AddObject(iGradientStops, New iGradientStops)
    End If
End Function

Function New_iIAssistance(iIAssistance As IAssistance) As iIAssistance
    Set New_iIAssistance = Utility.GetObject(iIAssistance)
    If New_iIAssistance Is Nothing Then
        Set New_iIAssistance = Utility.AddObject(iIAssistance, New iIAssistance)
    End If
End Function

Function New_iLineFormat(iLineFormat As LineFormat) As iLineFormat
    Set New_iLineFormat = Utility.GetObject(iLineFormat)
    If New_iLineFormat Is Nothing Then
        Set New_iLineFormat = Utility.AddObject(iLineFormat, New iLineFormat)
    End If
End Function

Function New_iPresentation(iPresentation As Presentation) As iPresentation
    Set New_iPresentation = Utility.GetObject(iPresentation)
    If New_iPresentation Is Nothing Then
        Set New_iPresentation = Utility.AddObject(iPresentation, New iPresentation)
    End If
End Function

Function New_iPresentations(iPresentations As Presentations) As iPresentations
    Set New_iPresentations = Utility.GetObject(iPresentations)
    If New_iPresentations Is Nothing Then
        Set New_iPresentations = Utility.AddObject(iPresentations, New iPresentations)
    End If
End Function

Function New_iReflectionFormat(iReflectionFormat As ReflectionFormat) As iReflectionFormat
    Set New_iReflectionFormat = Utility.GetObject(iReflectionFormat)
    If New_iReflectionFormat Is Nothing Then
        Set New_iReflectionFormat = Utility.AddObject(iReflectionFormat, New iReflectionFormat)
    End If
End Function

Function New_iSelection(iSelection As Selection) As iSelection
    Set New_iSelection = Utility.GetObject(iSelection)
    If New_iSelection Is Nothing Then
        Set New_iSelection = Utility.AddObject(iSelection, New iSelection)
    End If
End Function

Function New_iShadowFormat(iShadowFormat As PowerPoint.ShadowFormat) As iShadowFormat
    Set New_iShadowFormat = Utility.GetObject(iShadowFormat)
    If New_iShadowFormat Is Nothing Then
        Set New_iShadowFormat = Utility.AddObject(iShadowFormat, New iShadowFormat)
    End If
End Function

Function New_iShape(iShape As Shape) As iShape
    Set New_iShape = Utility.GetObject(iShape)
    If New_iShape Is Nothing Then
        Set New_iShape = Utility.AddObject(iShape, New iShape)
    End If
End Function

Function New_iShapeRange(iShapeRange As ShapeRange) As iShapeRange
    Set New_iShapeRange = Utility.GetObject(iShapeRange)
    If New_iShapeRange Is Nothing Then
        Set New_iShapeRange = Utility.AddObject(iShapeRange, New iShapeRange)
    End If
End Function

Function New_iShapes(iShapes As Shapes) As iShapes
    Set New_iShapes = Utility.GetObject(iShapes)
    If New_iShapes Is Nothing Then
        Set New_iShapes = Utility.AddObject(iShapes, New iShapes)
    End If
End Function

Function New_iSlide(iSlide As Slide) As iSlide
    Set New_iSlide = Utility.GetObject(iSlide)
    If New_iSlide Is Nothing Then
        Set New_iSlide = Utility.AddObject(iSlide, New iSlide)
    End If
End Function

Function New_iSlideRange(iSlideRange As SlideRange) As iSlideRange
    Set New_iSlideRange = Utility.GetObject(iSlideRange)
    If New_iSlideRange Is Nothing Then
        Set New_iSlideRange = Utility.AddObject(iSlideRange, New iSlideRange)
    End If
End Function

Function New_iSlides(iSlides As Slides) As iSlides
    Set New_iSlides = Utility.GetObject(iSlides)
    If New_iSlides Is Nothing Then
        Set New_iSlides = Utility.AddObject(iSlides, New iSlides)
    End If
End Function

Function New_iTextFrame(iTextFrame As TextFrame) As iTextFrame
    Set New_iTextFrame = Utility.GetObject(iTextFrame)
    If New_iTextFrame Is Nothing Then
        Set New_iTextFrame = Utility.AddObject(iTextFrame, New iTextFrame)
    End If
End Function

Function New_iTextFrame2(iTextFrame2 As TextFrame2) As iTextFrame2
    Set New_iTextFrame2 = Utility.GetObject(iTextFrame2)
    If New_iTextFrame2 Is Nothing Then
        Set New_iTextFrame2 = Utility.AddObject(iTextFrame2, New iTextFrame2)
    End If
End Function

Function New_iTextRange(iTextRange As TextRange) As iTextRange
    Set New_iTextRange = Utility.GetObject(iTextRange)
    If New_iTextRange Is Nothing Then
        Set New_iTextRange = Utility.AddObject(iTextRange, New iTextRange)
    End If
End Function

Function New_iTextRange2(iTextRange2 As TextRange2) As iTextRange2
    Set New_iTextRange2 = Utility.GetObject(iTextRange2)
    If New_iTextRange2 Is Nothing Then
        Set New_iTextRange2 = Utility.AddObject(iTextRange2, New iTextRange2)
    End If
End Function

Function New_oColorFormat(oColorFormat As Office.ColorFormat) As oColorFormat
    Set New_oColorFormat = Utility.GetObject(oColorFormat)
    If New_oColorFormat Is Nothing Then
        Set New_oColorFormat = Utility.AddObject(oColorFormat, New oColorFormat)
    End If
End Function

Function New_oFillFormat(oFillFormat As Office.FillFormat) As oFillFormat
    Set New_oFillFormat = Utility.GetObject(oFillFormat)
    If New_oFillFormat Is Nothing Then
        Set New_oFillFormat = Utility.AddObject(oFillFormat, New oFillFormat)
    End If
End Function

Function New_oShadowFormat(oShadowFormat As Office.ShadowFormat) As oShadowFormat
    Set New_oShadowFormat = Utility.GetObject(oShadowFormat)
    If New_oShadowFormat Is Nothing Then
        Set New_oShadowFormat = Utility.AddObject(oShadowFormat, New oShadowFormat)
    End If
End Function

