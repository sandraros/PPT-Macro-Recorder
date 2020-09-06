Attribute VB_Name = "Main"
Public snapshots As New Collection
Global snapshot As cSnapShot
Global AllObjectsCompared As Collection
Global MacroPresentation As String

Sub test()
    Dim Fill As Office.FillFormat
    Set Fill = ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Font.Fill
End Sub

Sub start_stop_recording()

    Dim code As String
    Dim Presentation As Presentation
    Dim VBProject As VBProject
    Dim VBComponent As VBComponent

    If snapshots.Count = 0 Then
        UserForm1.Caption = "Start recording"
        For Each VBProject In Application.vbe.VBProjects
            For Each VBComponent In VBProject.VBComponents
                If Len(VBComponent.Name) >= 6 Then
                    If Left(VBComponent.Name, 5) = "Macro" And CStr(Val(Mid(VBComponent.Name, 6))) = Mid(VBComponent.Name, 6) Then
                        macroNumber = CInt(Mid(VBComponent, 6))
                        If macroNumber > maxMacroNumber Then
                            maxMacroNumber = macroNumber
                        End If
                    End If
                End If
            Next
        Next
        UserForm1.MacroName = "Macro" & CStr(maxMacroNumber + 1)
        For Each Presentation In Application.Presentations
            Call UserForm1.Presentation.AddItem(Presentation.Name)
        Next
        UserForm1.Presentation.Value = Application.ActivePresentation.Name
        UserForm1.Presentation.Style = fmStyleDropDownList
        Call UserForm1.Show
        action = UserForm1.action
        MacroPresentation = UserForm1.Presentation.Value
        Call Unload(UserForm1)
        If action = enumAction.cancel Then
            Exit Sub
        End If
    End If

    take_snapshot

    If snapshots.Count = 2 Then

        ' Build collections MyObjPtrs and PptObjPtrs of all snapshots.
        Call BuildObjectIndexes

        Set AllObjectsCompared = New Collection

        code = compare_snapshots( _
            first:=snapshots.Item(snapshots.Count - 1), _
            last:=snapshots.Item(snapshots.Count))

        code = "Sub Macro1()" & Chr(13) _
                & code _
                & "End Sub"

        Call ExportCode(code)

        ' Clear the collection (can we trust the garbage collection?)
        snapshots.Remove 1
        snapshots.Remove 1
        Set snapshots = New Collection

    End If

End Sub

Sub ExportCode(code As String)

    Dim oVBComps As VBComponents
    Dim oVBComp As VBComponent

    Set oVBComps = Application.Presentations(MacroPresentation).VBProject.VBComponents
    On Error Resume Next
    Set oVBComp = oVBComps("NewMacros")
    If err.number <> 0 Then
        On Error GoTo 0
        Set oVBComp = oVBComps.Add(vbext_ct_StdModule)
        oVBComp.Name = "NewMacros"
    End If
    Call oVBComp.CodeModule.AddFromString(code)

End Sub

Sub take_snapshot()

    Set snapshot = New cSnapShot
    
    Set snapshot.iApplication = New_iApplication(Application)

    Call snapshots.Add(snapshot)

End Sub

Public Function compare_snapshots(first As cSnapShot, last As cSnapShot) As String

    ' parameter indent must be 4 minimum -> "With" will be output at position 0.
    compare_snapshots = last.iApplication.compare("Application", first.iApplication, 8)

End Function

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

Function New_iIAssistance(iIAssistance As IAssistance) As iIAssistance
    Set New_iIAssistance = Utility.GetObject(iIAssistance)
    If New_iIAssistance Is Nothing Then
        Set New_iIAssistance = Utility.AddObject(iIAssistance, New iIAssistance)
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
