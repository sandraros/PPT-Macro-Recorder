Attribute VB_Name = "Factory"
' Factory could be simplified by defining Attribute VB_PredeclaredId = True in each class module (Export/Import)
Function New_MR_Diff(StopObject As Object, StartObject As Object) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    If ExistsInCollection(goDiffPtrs, CStr(ObjPtr(StopObject))) Then

        Set oDiff = goDiffPtrs(CStr(ObjPtr(StopObject)))

    Else

        Set oDiff = New MR_Diff
        Set oDiff.StartObject = StartObject
        Set oDiff.StopObject = StopObject
        Set oDiff.AddedObjects = New Collection
        Set oDiff.RemovedObjects = New Collection
        Set oDiff.ScalarProperties = New Collection
        Set oDiff.ObjectProperties = New Collection
        Set oDiff.MethodCalls = New Collection

        Call goDiffPtrs.Add(oDiff, CStr(ObjPtr(StartObject)))
        Call goDiffPtrs.Add(oDiff, CStr(ObjPtr(StopObject)))

    End If

    Set New_MR_Diff = oDiff

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function New_iApplication(Optional iApplication As Application = Nothing) As iApplication

    On Error GoTo err_

    If iApplication Is Nothing Then
        Set New_iApplication = New iApplication
        Call New_iApplication.DefaultValues
    Else
        Set New_iApplication = GetMRObject(iApplication)
        If New_iApplication Is Nothing Then
            Set New_iApplication = AddObject(iApplication, New iApplication)
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

Function New_iColorFormat(Optional iColorFormat As PowerPoint.ColorFormat = Nothing) As iColorFormat

    On Error GoTo err_

    If iColorFormat Is Nothing Then
        Set New_iColorFormat = New iColorFormat
        Call New_iColorFormat.DefaultValues
    Else
        Set New_iColorFormat = GetMRObject(iColorFormat)
        If New_iColorFormat Is Nothing Then
            Set New_iColorFormat = AddObject(iColorFormat, New iColorFormat)
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

Function New_iCustomLayout(Optional iCustomLayout As CustomLayout = Nothing) As iCustomLayout

    On Error GoTo err_

    If iCustomLayout Is Nothing Then
        Set New_iCustomLayout = New iCustomLayout
        Call New_iCustomLayout.DefaultValues
    Else
        Set New_iCustomLayout = GetMRObject(iCustomLayout)
        If New_iCustomLayout Is Nothing Then
            Set New_iCustomLayout = AddObject(iCustomLayout, New iCustomLayout)
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

Function New_iDocumentWindow(Optional iDocumentWindow As DocumentWindow = Nothing) As iDocumentWindow

    On Error GoTo err_

    If iDocumentWindow Is Nothing Then
        Set New_iDocumentWindow = New iDocumentWindow
        Call New_iDocumentWindow.DefaultValues
    Else
    Set New_iDocumentWindow = GetMRObject(iDocumentWindow)
        If New_iDocumentWindow Is Nothing Then
            Set New_iDocumentWindow = AddObject(iDocumentWindow, New iDocumentWindow)
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

Function New_iDocumentWindows(Optional iDocumentWindows As DocumentWindows = Nothing) As iDocumentWindows

    On Error GoTo err_

    If iDocumentWindows Is Nothing Then
        Set New_iDocumentWindows = New iDocumentWindows
        Call New_iDocumentWindows.DefaultValues
    Else
        Set New_iDocumentWindows = GetMRObject(iDocumentWindows)
        If New_iDocumentWindows Is Nothing Then
            Set New_iDocumentWindows = AddObject(iDocumentWindows, New iDocumentWindows)
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

Function New_iFillFormat(Optional iFillFormat As PowerPoint.FillFormat = Nothing) As iFillFormat

    On Error GoTo err_

    If iFillFormat Is Nothing Then
        Set New_iFillFormat = New iFillFormat
        Call New_iFillFormat.DefaultValues
    Else
        Set New_iFillFormat = GetMRObject(iFillFormat)
        If New_iFillFormat Is Nothing Then
            Set New_iFillFormat = AddObject(iFillFormat, New iFillFormat)
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

Function New_iFont(Optional iFont As Font = Nothing) As iFont

    On Error GoTo err_

    If iFont Is Nothing Then
        Set New_iFont = New iFont
        Call New_iFont.DefaultValues
    Else
        Set New_iFont = GetMRObject(iFont)
        If New_iFont Is Nothing Then
            Set New_iFont = AddObject(iFont, New iFont)
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

Function New_iFont2(Optional iFont2 As Font2 = Nothing) As iFont2

    On Error GoTo err_

    If iFont2 Is Nothing Then
        Set New_iFont2 = New iFont2
        Call New_iFont2.DefaultValues
    Else
        Set New_iFont2 = GetMRObject(iFont2)
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

Function New_iGlowFormat(Optional iGlowFormat As GlowFormat = Nothing) As iGlowFormat

    On Error GoTo err_

    If iGlowFormat Is Nothing Then
        Set New_iGlowFormat = New iGlowFormat
        Call New_iGlowFormat.DefaultValues
    Else
        Set New_iGlowFormat = GetMRObject(iGlowFormat)
        If New_iGlowFormat Is Nothing Then
            Set New_iGlowFormat = AddObject(iGlowFormat, New iGlowFormat)
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

Function New_iGradientStop(Optional iGradientStop As GradientStop = Nothing) As iGradientStop

    On Error GoTo err_

    If iGradientStop Is Nothing Then
        Set New_iGradientStop = New iGradientStop
        Call New_iGradientStop.DefaultValues
    Else
        Set New_iGradientStop = GetMRObject(iGradientStop)
        If New_iGradientStop Is Nothing Then
            Set New_iGradientStop = AddObject(iGradientStop, New iGradientStop)
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

Function New_iGradientStops(Optional iGradientStops As GradientStops = Nothing) As iGradientStops

    On Error GoTo err_

    If iGradientStops Is Nothing Then
        Set New_iGradientStops = New iGradientStops
        Call New_iGradientStops.DefaultValues
    Else
        Set New_iGradientStops = GetMRObject(iGradientStops)
        If New_iGradientStops Is Nothing Then
            Set New_iGradientStops = AddObject(iGradientStops, New iGradientStops)
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

Function New_iLineFormat(Optional iLineFormat As LineFormat = Nothing) As iLineFormat

    On Error GoTo err_

    If iLineFormat Is Nothing Then
        Set New_iLineFormat = New iLineFormat
        Call New_iLineFormat.DefaultValues
    Else
        Set New_iLineFormat = GetMRObject(iLineFormat)
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

Function New_iPresentation(Optional iPresentation As Presentation = Nothing) As iPresentation

    On Error GoTo err_

    If iPresentation Is Nothing Then
        Set New_iPresentation = New iPresentation
        Call New_iPresentation.DefaultValues
    Else
        Set New_iPresentation = GetMRObject(iPresentation)
        If New_iPresentation Is Nothing Then
            Set New_iPresentation = AddObject(iPresentation, New iPresentation)
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

Function New_iPresentations(Optional iPresentations As Presentations = Nothing) As iPresentations

    On Error GoTo err_

    If iPresentations Is Nothing Then
        Set New_iPresentations = New iPresentations
        Call New_iPresentations.DefaultValues
    Else
        Set New_iPresentations = GetMRObject(iPresentations)
        If New_iPresentations Is Nothing Then
            Set New_iPresentations = AddObject(iPresentations, New iPresentations)
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

Function New_iReflectionFormat(Optional iReflectionFormat As ReflectionFormat = Nothing) As iReflectionFormat

    On Error GoTo err_

    If iReflectionFormat Is Nothing Then
        Set New_iReflectionFormat = New iReflectionFormat
        Call New_iReflectionFormat.DefaultValues
    Else
        Set New_iReflectionFormat = GetMRObject(iReflectionFormat)
        If New_iReflectionFormat Is Nothing Then
            Set New_iReflectionFormat = AddObject(iReflectionFormat, New iReflectionFormat)
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

Function New_iSelection(Optional iSelection As Selection = Nothing) As iSelection

    On Error GoTo err_

    If iSelection Is Nothing Then
        Set New_iSelection = New iSelection
        Call New_iSelection.DefaultValues
    Else
        Set New_iSelection = GetMRObject(iSelection)
        If New_iSelection Is Nothing Then
            Set New_iSelection = AddObject(iSelection, New iSelection)
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

Function New_iShadowFormat(Optional iShadowFormat As PowerPoint.ShadowFormat = Nothing) As iShadowFormat

    On Error GoTo err_

    If iShadowFormat Is Nothing Then
        Set New_iShadowFormat = New iShadowFormat
        Call New_iShadowFormat.DefaultValues
    Else
        Set New_iShadowFormat = GetMRObject(iShadowFormat)
        If New_iShadowFormat Is Nothing Then
            Set New_iShadowFormat = AddObject(iShadowFormat, New iShadowFormat)
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

Function New_iShape(Optional iShape As Shape = Nothing) As iShape

    On Error GoTo err_

    If iShape Is Nothing Then
        Set New_iShape = New iShape
        Call New_iShape.DefaultValues
    Else
        Set New_iShape = GetMRObject(iShape)
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

Function New_iShapeRange(Optional iShapeRange As shapeRange = Nothing) As iShapeRange

    On Error GoTo err_

    If iShapeRange Is Nothing Then
        Set New_iShapeRange = New iShapeRange
        Call New_iShapeRange.DefaultValues
    Else
        Set New_iShapeRange = GetMRObject(iShapeRange)
        If New_iShapeRange Is Nothing Then
            Set New_iShapeRange = AddObject(iShapeRange, New iShapeRange)
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

Function New_iShapes(Optional iShapes As Shapes = Nothing) As iShapes

    On Error GoTo err_

    If iShapes Is Nothing Then
        Set New_iShapes = New iShapes
        Call New_iShapes.DefaultValues
    Else
        Set New_iShapes = GetMRObject(iShapes)
        If New_iShapes Is Nothing Then
            Set New_iShapes = AddObject(iShapes, New iShapes)
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

Function New_iSlide(Optional iSlide As Slide = Nothing) As iSlide

    On Error GoTo err_

    If iSlide Is Nothing Then
        Set New_iSlide = New iSlide
        Call New_iSlide.DefaultValues
    Else
        Set New_iSlide = GetMRObject(iSlide)
        If New_iSlide Is Nothing Then
            Set New_iSlide = AddObject(iSlide, New iSlide)
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

Function New_iSlideRange(Optional iSlideRange As SlideRange = Nothing) As iSlideRange

    On Error GoTo err_

    If iSlideRange Is Nothing Then
        Set New_iSlideRange = New iSlideRange
        Call New_iSlideRange.DefaultValues
    Else
        Set New_iSlideRange = GetMRObject(iSlideRange)
        If New_iSlideRange Is Nothing Then
            Set New_iSlideRange = AddObject(iSlideRange, New iSlideRange)
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

Function New_iSlides(Optional iSlides As Slides = Nothing) As iSlides

    On Error GoTo err_

    If iSlides Is Nothing Then
        Set New_iSlides = New iSlides
        Call New_iSlides.DefaultValues
    Else
        Set New_iSlides = GetMRObject(iSlides)
        If New_iSlides Is Nothing Then
            Set New_iSlides = AddObject(iSlides, New iSlides)
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

Function New_iTextFrame(Optional iTextFrame As TextFrame = Nothing) As iTextFrame

    On Error GoTo err_

    If iTextFrame Is Nothing Then
        Set New_iTextFrame = New iTextFrame
        Call New_iTextFrame.DefaultValues
    Else
        Set New_iTextFrame = GetMRObject(iTextFrame)
        If New_iTextFrame Is Nothing Then
            Set New_iTextFrame = AddObject(iTextFrame, New iTextFrame)
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

Function New_iTextFrame2(Optional iTextFrame2 As TextFrame2 = Nothing) As iTextFrame2

    On Error GoTo err_

    If iTextFrame2 Is Nothing Then
        Set New_iTextFrame2 = New iTextFrame2
        Call New_iTextFrame2.DefaultValues
    Else
        Set New_iTextFrame2 = GetMRObject(iTextFrame2)
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

Function New_iTextRange(Optional iTextRange As TextRange = Nothing) As iTextRange

    On Error GoTo err_

    If iTextRange Is Nothing Then
        Set New_iTextRange = New iTextRange
        Call New_iTextRange.DefaultValues
    Else
        Set New_iTextRange = GetMRObject(iTextRange)
        If New_iTextRange Is Nothing Then
            Set New_iTextRange = AddObject(iTextRange, New iTextRange)
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

Function New_iTextRange2(Optional iTextRange2 As TextRange2 = Nothing) As iTextRange2

    On Error GoTo err_

    If iTextRange2 Is Nothing Then
        Set New_iTextRange2 = New iTextRange2
        Call New_iTextRange2.DefaultValues
    Else
        Set New_iTextRange2 = GetMRObject(iTextRange2)
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

Function New_oColorFormat(Optional oColorFormat As Office.ColorFormat = Nothing) As oColorFormat

    On Error GoTo err_

    If oColorFormat Is Nothing Then
        Set New_oColorFormat = New oColorFormat
        Call New_oColorFormat.DefaultValues
    Else
        Set New_oColorFormat = GetMRObject(oColorFormat)
        If New_oColorFormat Is Nothing Then
            Set New_oColorFormat = AddObject(oColorFormat, New oColorFormat)
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

Function New_oFillFormat(Optional oFillFormat As Office.FillFormat = Nothing) As oFillFormat

    On Error GoTo err_

    If oFillFormat Is Nothing Then
        Set New_oFillFormat = New oFillFormat
        Call New_oFillFormat.DefaultValues
    Else
        Set New_oFillFormat = GetMRObject(oFillFormat)
        If New_oFillFormat Is Nothing Then
            Set New_oFillFormat = AddObject(oFillFormat, New oFillFormat)
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

Function New_oShadowFormat(Optional oShadowFormat As Office.ShadowFormat = Nothing) As oShadowFormat

    On Error GoTo err_

    If oShadowFormat Is Nothing Then
        Set New_oShadowFormat = New oShadowFormat
        Call New_oShadowFormat.DefaultValues
    Else
        Set New_oShadowFormat = GetMRObject(oShadowFormat)
        If New_oShadowFormat Is Nothing Then
            Set New_oShadowFormat = AddObject(oShadowFormat, New oShadowFormat)
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
