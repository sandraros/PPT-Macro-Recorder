Attribute VB_Name = "FactoryFunctions"
Function New_UDiff(ObjectName As String, StopObject As Object, StartObject As Object) As UDiff
    Set New_UDiff = New UDiff
    New_UDiff.ObjectName = ObjectName
    Set New_UDiff.StartObject = StartObject
    Set New_UDiff.StopObject = StopObject
End Function

Function New_iApplication(iApplication As Application) As iApplication
    Set New_iApplication = GetMrsObject(iApplication)
    If New_iApplication Is Nothing Then
        Set New_iApplication = AddObject(iApplication, New iApplication)
    End If
End Function

Function New_iColorFormat(iColorFormat As PowerPoint.ColorFormat) As iColorFormat
    Set New_iColorFormat = GetMrsObject(iColorFormat)
    If New_iColorFormat Is Nothing Then
        Set New_iColorFormat = AddObject(iColorFormat, New iColorFormat)
    End If
End Function

Function New_iDocumentWindow(iDocumentWindow As DocumentWindow) As iDocumentWindow
    Set New_iDocumentWindow = GetMrsObject(iDocumentWindow)
    If New_iDocumentWindow Is Nothing Then
        Set New_iDocumentWindow = AddObject(iDocumentWindow, New iDocumentWindow)
    End If
End Function

Function New_iDocumentWindows(iDocumentWindows As DocumentWindows) As iDocumentWindows
    Set New_iDocumentWindows = GetMrsObject(iDocumentWindows)
    If New_iDocumentWindows Is Nothing Then
        Set New_iDocumentWindows = AddObject(iDocumentWindows, New iDocumentWindows)
    End If
End Function

Function New_iFillFormat(iFillFormat As PowerPoint.FillFormat) As iFillFormat
    Set New_iFillFormat = GetMrsObject(iFillFormat)
    If New_iFillFormat Is Nothing Then
        Set New_iFillFormat = AddObject(iFillFormat, New iFillFormat)
    End If
End Function

Function New_iFont(iFont As Font) As iFont
    Set New_iFont = GetMrsObject(iFont)
    If New_iFont Is Nothing Then
        Set New_iFont = AddObject(iFont, New iFont)
    End If
End Function

Function New_iFont2(Optional iFont2 As Font2 = Nothing) As iFont2

    On Error GoTo err_

    If iFont2 Is Nothing Then
        Set New_iFont2 = New iFont2
        Call New_iFont2.DefaultValues
    Else
        Set New_iFont2 = GetMrsObject(iFont2)
        If New_iFont2 Is Nothing Then
            Set New_iFont2 = AddObject(iFont2, New iFont2)
        End If
    End If

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function New_iGlowFormat(iGlowFormat As GlowFormat) As iGlowFormat
    Set New_iGlowFormat = GetMrsObject(iGlowFormat)
    If New_iGlowFormat Is Nothing Then
        Set New_iGlowFormat = AddObject(iGlowFormat, New iGlowFormat)
    End If
End Function

Function New_iGradientStop(iGradientStop As GradientStop) As iGradientStop
    Set New_iGradientStop = GetMrsObject(iGradientStop)
    If New_iGradientStop Is Nothing Then
        Set New_iGradientStop = AddObject(iGradientStop, New iGradientStop)
    End If
End Function

Function New_iGradientStops(iGradientStops As GradientStops) As iGradientStops
    Set New_iGradientStops = GetMrsObject(iGradientStops)
    If New_iGradientStops Is Nothing Then
        Set New_iGradientStops = AddObject(iGradientStops, New iGradientStops)
    End If
End Function

Function New_iIAssistance(iIAssistance As IAssistance) As iIAssistance
    Set New_iIAssistance = GetMrsObject(iIAssistance)
    If New_iIAssistance Is Nothing Then
        Set New_iIAssistance = AddObject(iIAssistance, New iIAssistance)
    End If
End Function

Function New_iLineFormat(Optional iLineFormat As LineFormat = Nothing) As iLineFormat

    On Error GoTo err_

    If iLineFormat Is Nothing Then
        Set New_iLineFormat = New iLineFormat
        Call New_iLineFormat.DefaultValues
    Else
        Set New_iLineFormat = GetMrsObject(iLineFormat)
        If New_iLineFormat Is Nothing Then
            Set New_iLineFormat = AddObject(iLineFormat, New iLineFormat)
        End If
    End If

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function New_iPresentation(iPresentation As presentation) As iPresentation
    Set New_iPresentation = GetMrsObject(iPresentation)
    If New_iPresentation Is Nothing Then
        Set New_iPresentation = AddObject(iPresentation, New iPresentation)
    End If
End Function

Function New_iPresentations(iPresentations As Presentations) As iPresentations
    Set New_iPresentations = GetMrsObject(iPresentations)
    If New_iPresentations Is Nothing Then
        Set New_iPresentations = AddObject(iPresentations, New iPresentations)
    End If
End Function

Function New_iReflectionFormat(iReflectionFormat As ReflectionFormat) As iReflectionFormat
    Set New_iReflectionFormat = GetMrsObject(iReflectionFormat)
    If New_iReflectionFormat Is Nothing Then
        Set New_iReflectionFormat = AddObject(iReflectionFormat, New iReflectionFormat)
    End If
End Function

Function New_iSelection(iSelection As Selection) As iSelection
    Set New_iSelection = GetMrsObject(iSelection)
    If New_iSelection Is Nothing Then
        Set New_iSelection = AddObject(iSelection, New iSelection)
    End If
End Function

Function New_iShadowFormat(iShadowFormat As PowerPoint.ShadowFormat) As iShadowFormat
    Set New_iShadowFormat = GetMrsObject(iShadowFormat)
    If New_iShadowFormat Is Nothing Then
        Set New_iShadowFormat = AddObject(iShadowFormat, New iShadowFormat)
    End If
End Function

Function New_iShape(Optional iShape As Shape = Nothing) As iShape

    On Error GoTo err_

    If iShape Is Nothing Then
        Set New_iShape = New iShape
        Call New_iShape.DefaultValues
    Else
        Set New_iShape = GetMrsObject(iShape)
        If New_iShape Is Nothing Then
            Set New_iShape = AddObject(iShape, New iShape)
        End If
    End If

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function New_iShapeRange(iShapeRange As shapeRange) As iShapeRange
    Set New_iShapeRange = GetMrsObject(iShapeRange)
    If New_iShapeRange Is Nothing Then
        Set New_iShapeRange = AddObject(iShapeRange, New iShapeRange)
    End If
End Function

Function New_iShapes(iShapes As Shapes) As iShapes
    Set New_iShapes = GetMrsObject(iShapes)
    If New_iShapes Is Nothing Then
        Set New_iShapes = AddObject(iShapes, New iShapes)
    End If
End Function

Function New_iSlide(iSlide As Slide) As iSlide
    Set New_iSlide = GetMrsObject(iSlide)
    If New_iSlide Is Nothing Then
        Set New_iSlide = AddObject(iSlide, New iSlide)
    End If
End Function

Function New_iSlideRange(iSlideRange As SlideRange) As iSlideRange
    Set New_iSlideRange = GetMrsObject(iSlideRange)
    If New_iSlideRange Is Nothing Then
        Set New_iSlideRange = AddObject(iSlideRange, New iSlideRange)
    End If
End Function

Function New_iSlides(iSlides As Slides) As iSlides
    Set New_iSlides = GetMrsObject(iSlides)
    If New_iSlides Is Nothing Then
        Set New_iSlides = AddObject(iSlides, New iSlides)
    End If
End Function

Function New_iTextFrame(iTextFrame As TextFrame) As iTextFrame
    Set New_iTextFrame = GetMrsObject(iTextFrame)
    If New_iTextFrame Is Nothing Then
        Set New_iTextFrame = AddObject(iTextFrame, New iTextFrame)
    End If
End Function

Function New_iTextFrame2(Optional iTextFrame2 As TextFrame2 = Nothing) As iTextFrame2

    On Error GoTo err_

    If iTextFrame2 Is Nothing Then
        Set New_iTextFrame2 = New iTextFrame2
        Call New_iTextFrame2.DefaultValues
    Else
        Set New_iTextFrame2 = GetMrsObject(iTextFrame2)
        If New_iTextFrame2 Is Nothing Then
            Set New_iTextFrame2 = AddObject(iTextFrame2, New iTextFrame2)
        End If
    End If

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function New_iTextRange(iTextRange As TextRange) As iTextRange
    Set New_iTextRange = GetMrsObject(iTextRange)
    If New_iTextRange Is Nothing Then
        Set New_iTextRange = AddObject(iTextRange, New iTextRange)
    End If
End Function

Function New_iTextRange2(Optional iTextRange2 As TextRange2 = Nothing) As iTextRange2

    On Error GoTo err_

    If iTextRange2 Is Nothing Then
        Set New_iTextRange2 = New iTextRange2
        Call New_iTextRange2.DefaultValues
    Else
        Set New_iTextRange2 = GetMrsObject(iTextRange2)
        If New_iTextRange2 Is Nothing Then
            Set New_iTextRange2 = AddObject(iTextRange2, New iTextRange2)
        End If
    End If

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function New_oColorFormat(oColorFormat As Office.ColorFormat) As oColorFormat
    Set New_oColorFormat = GetMrsObject(oColorFormat)
    If New_oColorFormat Is Nothing Then
        Set New_oColorFormat = AddObject(oColorFormat, New oColorFormat)
    End If
End Function

Function New_oFillFormat(oFillFormat As Office.FillFormat) As oFillFormat
    Set New_oFillFormat = GetMrsObject(oFillFormat)
    If New_oFillFormat Is Nothing Then
        Set New_oFillFormat = AddObject(oFillFormat, New oFillFormat)
    End If
End Function

Function New_oShadowFormat(oShadowFormat As Office.ShadowFormat) As oShadowFormat
    Set New_oShadowFormat = GetMrsObject(oShadowFormat)
    If New_oShadowFormat Is Nothing Then
        Set New_oShadowFormat = AddObject(oShadowFormat, New oShadowFormat)
    End If
End Function

