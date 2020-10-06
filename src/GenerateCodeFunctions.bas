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

Function GetCodeForInitiallySelectedObjects(ioDiff As MR_Diff) As MR_Code

    Dim oCode As MR_Code
    Dim oObjectProperty As UDiffObjectProperty

    On Error GoTo err_

    Set oCode = New MR_Code

    For Each oItem In ioDiff.ObjectProperties
        Set oObjectProperty = oItem
        If oObjectProperty.Diff.ObjectName <> "Selection" Then
            Call oCode.AddCode(GetCodeForInitiallySelectedObjects(oObjectProperty.Diff))
        Else
            Call oCode.AddCode(GetCodeForInitiallySelectedObjects2(oObjectProperty.Diff))
            Call oCode.Wrap("ActiveWindow.Selection")
        End If
    Next

    Set GetCodeForInitiallySelectedObjects = oCode

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function GetCodeForInitiallySelectedObjects2(ioDiff As MR_Diff) As MR_Code ', Optional ObjectName As String = "") As MR_Code

    Dim oCode As MR_Code
    Dim sObjectPrefix As String
    Dim oScalarProperty As UDiffScalarProperty
    Dim oObjectProperty As UDiffObjectProperty
    Dim oMethodCall As UDiffMethodCall

    On Error GoTo err_

    Set oCode = New MR_Code

    sObjectPrefix = GetObjectPrefix(ioDiff)

    For Each oItem In ioDiff.ScalarProperties
        Set oScalarProperty = oItem
        Call oCode.Add(sObjectPrefix & oScalarProperty.Name & " = " & oScalarProperty.Value)
    Next

    For Each oItem In ioDiff.ObjectProperties
        Set oObjectProperty = oItem
        bIsObjectSelected = (IsObjectPartOfSelection(GetPptObject(goStartSnapshot, oObjectProperty.Diff.StartObject), goStartSnapshot))
        If bIsObjectSelected = False Then
            Call oCode.AddCode(GetCodeForInitiallySelectedObjects2(oObjectProperty.Diff)) ', sObjectPrefix & oObjectProperty.Diff.ObjectName))
        End If
    Next

    For Each oItem In ioDiff.MethodCalls
        Set oMethodCall = oItem
        Call oCode.Add("Call " & sObjectPrefix & oMethodCall.ProcName & "(" & oMethodCall.Arguments & ")")
    Next

    Set GetCodeForInitiallySelectedObjects2 = oCode

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function GetCodeForAddedObjects(ioDiff As MR_Diff, Optional ObjectName As String = "") As MR_Code

    Dim oCode As MR_Code
    Dim oAddedObject As UDiffAddedObject
    Dim oRemovedObject As UDiffRemovedObject
    Dim oScalarProperty As UDiffScalarProperty
    Dim oObjectProperty As UDiffObjectProperty
    Dim oMethodCall As UDiffMethodCall

    On Error GoTo err_

    Set oCode = New MR_Code

    sObjectPrefix = GetObjectPrefix(ioDiff)

    For Each oItem In ioDiff.AddedObjects
        Set oAddedObject = oItem
        Call oCode.AddCode(oAddedObject.MRObject.create())
    Next

    For Each oItem In ioDiff.ObjectProperties
        Set oObjectProperty = oItem
        Call oCode.AddCode(GetCodeForAddedObjects(oObjectProperty.Diff, sObjectPrefix & oObjectProperty.ObjectName))
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

Function GetAddedObjects(ioDiff As MR_Diff) As Collection

    Dim oColl As Collection
    Dim oColl2 As Collection
    Dim oAddedObject As UDiffAddedObject
    Dim oObjectProperty As UDiffObjectProperty

    On Error GoTo err_

    Set oColl = New Collection

    For Each oItem In ioDiff.AddedObjects
        Set oAddedObject = oItem
        Call oColl.Add(oAddedObject.MRObject)
    Next

    For Each oItem In ioDiff.ObjectProperties
        Set oObjectProperty = oItem
        Set oColl2 = GetAddedObjects(oObjectProperty.Diff)
        For i = 1 To oColl2.Count
            Call oColl.Add(oColl2.Item(i))
        Next
    Next

    Set GetAddedObjects = oColl

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function GetCodeForChangedSelection(ioDiff As MR_Diff, ioStartSelection As iSelection, ioStopSelection As iSelection) As MR_Code

    Dim oCode As MR_Code
    Dim bObjectIsUnselected As Boolean
    Dim oAddedObjects As Collection
    Dim lSlideIndex As Long
    Dim oShape As iShape
    Dim oSlide As iSlide
    Dim oPresentation As iPresentation
    Dim oPptSelectedSlide As Slide
    Dim oPptShape As Shape
    Dim lShapeIndex As Long

    On Error GoTo err_

    Set oCode = New MR_Code

    ' Generate "call activewindow.selection.unselect" if all selected objects were unselected
    bObjectIsUnselected = False
    If ioStartSelection.Type_ <> ppSelectionNone _
                And ioStopSelection.Type_ = ppSelectionNone Then
        bObjectIsUnselected = True
    End If
    If 0 = 1 Then
        ' TODO DELETE THIS DEAD CODE
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
    Else
        ' Generate Select on newly selected objects except if they were just added
        '   (because in that case "Select" was added by GetCodeForAddedObjects)
        Set oAddedObjects = GetAddedObjects(ioDiff)
        Select Case ioStopSelection.Type_
            Case ppSelectionShapes
                ' Generate "Select" if shape was not added (see reason above)
                For i = 1 To ioStopSelection.shapeRange.Items.Count
                    Set oShape = ioStopSelection.shapeRange.Items(i)
                    If IsObjectNewlySelected(GetPptObject(goStopSnapshot, oShape)) Then
                        ' Was the shape selected because it was added?
                        bObjectIsAdded = False
                        For j = 1 To oAddedObjects.Count
                            If oShape Is oAddedObjects(j) Then
                                bObjectIsAdded = True
                                Exit For
                            End If
                        Next
                        If Not bObjectIsAdded Then
                            ' Determine the Slides item index of the slide containing the selected shape
                            Set oPptSelectedSlide = ioStopSelection.shapeRange.Items(1).Parent
                            For j = 1 To ActivePresentation.Slides.Count
                                If oPptSelectedSlide Is ActivePresentation.Slides(j) Then
                                    lSlideIndex = j
                                    Exit For
                                End If
                            Next
                            ' Get Shapes item index
                            Set oPptShape = GetPptObject(goStopSnapshot, oShape)
                            For j = 1 To oPptSelectedSlide.Shapes.Count
                                If oPptShape Is oPptSelectedSlide.Shapes(j) Then
                                    lShapeIndex = j
                                End If
                            Next
                            Call oCode.Add("Call ActivePresentation.Slides(" & CStr(lSlideIndex) & ").Shapes(" & CStr(lShapeIndex) & ").Select")
                        End If
                        Exit For
                    End If
                Next
            Case ppSelectionSlides
                ' TODO
            Case ppSelectionText
                ' TODO
        End Select
    End If

    Set GetCodeForChangedSelection = oCode

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function GetCodeForUnselectedObjects(ioDiff As MR_Diff, Optional ObjectName As String = "") As MR_Code

    Dim oCode As MR_Code
    Dim sObjectPrefix As String
    Dim oScalarProperty As UDiffScalarProperty
    Dim oObjectProperty As UDiffObjectProperty
    Dim oMethodCall As UDiffMethodCall

    On Error GoTo err_

    Set oCode = New MR_Code

    sObjectPrefix = GetObjectPrefix(ioDiff)

    For Each oItem In ioDiff.ScalarProperties
        Set oScalarProperty = oItem
        Call oCode.Add(sObjectPrefix & oScalarProperty.Name & " = " & oScalarProperty.Value)
    Next

    For Each oItem In ioDiff.ObjectProperties
        Set oObjectProperty = oItem
        ' TODO process other possible selected objects like iSlide, iTextRange2...
        bIsObjectSelected = False
        Select Case oObjectProperty.Diff.ObjectName
            Case "iShape"
                bIsObjectSelected = (IsObjectPartOfSelection(GetPptObject(goStopSnapshot, oObjectProperty.Diff.StopObject), goStopSnapshot))
        End Select
        If oObjectProperty.Diff.ObjectName <> "Selection" And bIsObjectSelected = False Then
            Call oCode.AddCode(GetCodeForUnselectedObjects(oObjectProperty.Diff, sObjectPrefix & oObjectProperty.Diff.ObjectName))
        Else
        End If
    Next

    For Each oItem In ioDiff.MethodCalls
        Set oMethodCall = oItem
        Call oCode.Add("Call " & sObjectPrefix & oMethodCall.ProcName & "(" & oMethodCall.Arguments & ")")
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

Function GetObjectPrefix(ioDiff As MR_Diff) As String

    If ioDiff.ObjectName = "Application" Then
        GetObjectPrefix = ""
    Else
        GetObjectPrefix = "."
    End If

End Function

Function OBSOLETE_GetCodeForAllChangedElements(ioDiff As MR_Diff, ObjectName As String) As MR_Code

    Dim oCode As MR_Code
    Dim oAddedObject As UDiffAddedObject
    Dim oScalarProperty As UDiffScalarProperty
    Dim oObjectProperty As UDiffObjectProperty
    Dim oMethodCall As UDiffMethodCall

    On Error GoTo err_

    Set oCode = New MR_Code

    sObjectPrefix = GetObjectPrefix(ioDiff)

    For Each oItem In ioDiff.AddedObjects
        Set oAddedObject = oItem
        Call oCode.AddCode(oAddedObject.MRObject.create())
    Next

    For Each oItem In ioDiff.ScalarProperties
        Set oScalarProperty = oItem
        Call oCode.Add(sObjectPrefix & oScalarProperty.Name & " = " & oScalarProperty.Value)
    Next

    For Each oItem In ioDiff.ObjectProperties
        Set oObjectProperty = oItem
        Call oCode.AddCode(OBSOLETE_GetCodeForAllChangedElements(oObjectProperty.Diff, sObjectPrefix & oObjectProperty.Diff.ObjectName))
    Next

    For Each oItem In ioDiff.MethodCalls
        Set oMethodCall = oItem
        Call oCode.Add("Call " & sObjectPrefix & oMethodCall.ProcName & "(" & oMethodCall.Arguments & ")")
    Next

    Call oCode.Wrap(ObjectName)

    Set OBSOLETE_GetCodeForAllChangedElements = oCode

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function OBSOLETE_GenerateCode(ioDiff As MR_Diff, ioStartSelection As iSelection, ioStopSelection As iSelection, ioDiffSelection As MR_Diff) As String

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
    strCode = strCode & GetCodeForChangedSelection(ioDiff, ioStartSelection, ioStopSelection).ConvertToString()
    If ioStopSelection.Type_ <> ppSelectionNone Then
        strCode = strCode & OBSOLETE_GetCodeForAllChangedElements(ioDiffSelection, "ActiveWindow.Selection").ConvertToString()
    End If
    strCode = strCode & GetCodeForUnselectedObjects(ioDiff).ConvertToString()

    ' Macro end (End Sub)
    strCode = strCode & "End Sub"

    OBSOLETE_GenerateCode = strCode

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

