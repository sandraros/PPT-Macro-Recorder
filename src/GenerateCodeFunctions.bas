Attribute VB_Name = "GenerateCodeFunctions"
Function WrapCodeIntoMacro(isCode As String) As String

    Dim strCode As String

    On Error GoTo err_

    strCode = ""
    ' Macro start (Sub)
    strCode = strCode _
        & "Sub " & macroName & "()" & Chr(13) _
        & "'" & Chr(13) _
        & "' " & macroName & " Macro" & Chr(13)

    ' Macro description
    astrMacroDescription = Split(macroDescription, Chr(13))
    For i = LBound(astrMacroDescription) To UBound(astrMacroDescription)
        strCode = strCode & "' " & astrMacroDescription(i) & Chr(13)
    Next
    strCode = strCode & "'" & Chr(13)

    ' Macro code
    strCode = strCode & isCode

    ' Macro end (End Sub)
    strCode = strCode & "End Sub"

    WrapCodeIntoMacro = strCode

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function GenerateCode(ioDiff As MR_Diff, ioStartSelection As iSelection, ioStopSelection As iSelection, ioDiffSelection As MR_Diff) As String

    'Dim oCode As MR_Code
    Dim strCode As String

    On Error GoTo err_

    ' If one object is currently selected, prefer to generate:
    '     Application.ActiveWindow.Selection.ShapeRange(1).Fill.BackColor = RGB(17, 15, 150)
    ' rather than
    '     Application.ActivePresentation.Shapes(27).Fill.BackColor = RGB(17, 15, 150)
    'Set oCode = New MR_Code
    'Call oDiff.AddDiff(CallUnselectAndSelectIfSelectionChanged())
    'Call oDiff.AddDiff(goStopSnapshot.iSelection.MR_Compare("ActiveWindow.Selection", goStartSnapshot.iSelection))
    'Call oDiff.AddDiff(goStopSnapshot.iPresentation.MR_Compare("ActivePresentation", goStartSnapshot.iPresentation))
    '' In the future, propose to record actions on several presentations
    'Call oDiff.AddDiff(goStopSnapshot.iPresentation.MR_Compare("Application", goStartSnapshot.iApplication))

    '============================
    ' SOURCE CODE AROUND
    '============================
    strCode = ""
    ' Macro start (Sub)
    strCode = strCode _
        & "Sub " & macroName & "()" & Chr(13) _
        & "'" & Chr(13) _
        & "' " & macroName & " Macro" & Chr(13)
    ' Macro description
    astrMacroDescription = Split(macroDescription, Chr(13))
    For i = LBound(astrMacroDescription) To UBound(astrMacroDescription)
        strCode = strCode & "' " & astrMacroDescription(i) & Chr(13)
    Next
    strCode = strCode & "'" & Chr(13)

    ' Macro code
    strCode = strCode & GetCodeForAddedObjects(ioDiff, "ActivePresentation").ConvertToString()
    strCode = strCode & GetCodeToSelectObjects(ioStartSelection, ioStopSelection).ConvertToString()
    If ioStopSelection.Type_ <> ppSelectionNone Then
        strCode = strCode & GetCodeForAllChangedElements(ioDiffSelection, "ActiveWindow.Selection").ConvertToString()
    End If
    strCode = strCode & GetCodeForUnselectedObjects(ioDiff, "ActivePresentation", ioStopSelection).ConvertToString()

    ' Macro end (End Sub)
    strCode = strCode & "End Sub"

    GenerateCode = strCode

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function GetCodeForAddedObjects(ioDiff As MR_Diff, ObjectName As String) As MR_Code

    Dim oCode As MR_Code
    Dim oAddedObject As UDiffAddedObject
    Dim oRemovedObject As UDiffRemovedObject
    Dim oScalarProperty As UDiffScalarProperty
    Dim oObjectProperty As UDiffObjectProperty
    Dim oMethodCall As UDiffMethodCall

    On Error GoTo err_

    Set oCode = New MR_Code

    For Each oItem In ioDiff.AddedObjects
        Set oAddedObject = oItem
        Call oCode.AddCode(oAddedObject.MRObject.create())
    Next

    For Each oItem In ioDiff.ObjectProperties
        Set oObjectProperty = oItem
        Call oCode.AddCode(GetCodeForAddedObjects(oObjectProperty.Diff, "." & oObjectProperty.ObjectName))
    Next

    Call oCode.Wrap(ObjectName)

    Set GetCodeForAddedObjects = oCode

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function GetCodeToSelectObjects(ioStartSelection As iSelection, ioStopSelection As iSelection) As MR_Code

    Dim oCode As MR_Code
    Dim bObjectIsUnselected As Boolean
    Dim oShape As iShape
    Dim oPresentation As iPresentation
    Dim oSlide As iSlide
    Dim lSlideId As Long
    Dim oSelectedSlide As Slide

    On Error GoTo err_

    Set oCode = New MR_Code

    ' Do "call activewindow.selection.unselect" if at least one selected object became unselected
    bObjectIsUnselected = False
    If ioStopSelection.Type_ = ppSelectionNone _
                And ioStartSelection.Type_ <> ppSelectionNone Then
        bObjectIsUnselected = True
    Else
        Select Case ioStartSelection.Type_
            Case ppSelectionShapes
                For Each oShape In ioStartSelection.shapeRange.Items
                    If Not IsObjectPartOfSelection(GetPptObject(goStartSnapshot, oShape), goStopSnapshot) Then
                        bObjectIsUnselected = True
                        Exit For
                    End If
                Next
            Case ppSelectionSlides
                For Each oSlide In ioStartSelection.SlideRange.Items
                    If Not IsObjectPartOfSelection(GetPptObject(goStartSnapshot, oSlide), goStopSnapshot) Then
                        bObjectIsUnselected = True
                        Exit For
                    End If
                Next
            Case ppSelectionText
                ' TODO
        End Select
    End If
    If bObjectIsUnselected Then
        Call oCode.Add("Call ActiveWindow.Selection.Unselect")
    End If

    ' Process newly selected objects
    Select Case ioStopSelection.Type_
        Case ppSelectionShapes
            Set oSelectedSlide = ioStopSelection.shapeRange.Items(1).Parent
            For i = 1 To ActivePresentation.Slides.Count
                If oSelectedSlide Is ActivePresentation.Slides(i) Then
                    lSlideId = i
                    Exit For
                End If
            Next
            For i = 1 To ioStopSelection.shapeRange.Items.Count 'oSelectedSlide.Shapes.Count
                Set oShape = ioStopSelection.shapeRange.Items(i) 'GetMRObject(oSelectedSlide.Shapes(i))
                If IsObjectNewlySelected(GetPptObject(goStopSnapshot, oShape)) Then
                    Call oCode.Add("Call ActivePresentation.Slides(" & CStr(lSlideId) & ").Shapes(" & CStr(i) & ").Select")
                    Exit For
                End If
            Next
        Case ppSelectionSlides
            ' TODO
        Case ppSelectionText
            ' TODO
    End Select

    Set GetCodeToSelectObjects = oCode

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function GetCodeForAllChangedElements(ioDiff As MR_Diff, ObjectName As String) As MR_Code

    Dim oCode As MR_Code
    Dim oAddedObject As UDiffAddedObject
    Dim oScalarProperty As UDiffScalarProperty
    Dim oObjectProperty As UDiffObjectProperty
    Dim oMethodCall As UDiffMethodCall

    On Error GoTo err_

    Set oCode = New MR_Code

    For Each oItem In ioDiff.AddedObjects
        Set oAddedObject = oItem
        Call oCode.AddCode(oAddedObject.MRObject.create())
    Next

    For Each oItem In ioDiff.ScalarProperties
        Set oScalarProperty = oItem
        Call oCode.Add("." & oScalarProperty.Name & " = " & oScalarProperty.Value)
    Next

    For Each oItem In ioDiff.ObjectProperties
        Set oObjectProperty = oItem
        Call oCode.AddCode(GetCodeForAllChangedElements(oObjectProperty.Diff, "." & oObjectProperty.Diff.ObjectName))
    Next

    For Each oItem In ioDiff.MethodCalls
        Set oMethodCall = oItem
        Call oCode.Add("Call ." & oMethodCall.ProcName & "(" & oMethodCall.Arguments & ")")
    Next

    Call oCode.Wrap(ObjectName)

    Set GetCodeForAllChangedElements = oCode

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function GetCodeForUnselectedObjects(ioDiff As MR_Diff, ObjectName As String, ioStopSelection As iSelection) As MR_Code

    Dim oCode As MR_Code
    Dim oScalarProperty As UDiffScalarProperty
    Dim oObjectProperty As UDiffObjectProperty
    Dim oMethodCall As UDiffMethodCall

    On Error GoTo err_

    Set oCode = New MR_Code

    For Each oItem In ioDiff.ScalarProperties
        Set oScalarProperty = oItem
        Call oCode.Add("." & oScalarProperty.Name & " = " & oScalarProperty.Value)
    Next

    For Each oItem In ioDiff.ObjectProperties
        Set oObjectProperty = oItem
        bIsObjectSelected = False
        Select Case TypeName(oObjectProperty.Diff.StopObject)
        Case "iShape"
        bIsObjectSelected = (IsObjectPartOfSelection(GetPptObject(goStopSnapshot, oObjectProperty.Diff.StopObject), goStopSnapshot))
        End Select
        If bIsObjectSelected = False Then
            Call oCode.AddCode(GetCodeForUnselectedObjects(oObjectProperty.Diff, "." & oObjectProperty.Diff.ObjectName, ioStopSelection))
        End If
    Next

    For Each oItem In ioDiff.MethodCalls
        Set oMethodCall = oItem
        Call oCode.Add("Call ." & oMethodCall.ProcName & "(" & oMethodCall.Arguments & ")")
    Next

    Call oCode.Wrap(ObjectName)

    Set GetCodeForUnselectedObjects = oCode

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function
