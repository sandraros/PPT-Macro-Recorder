Attribute VB_Name = "GenerateCode"
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
    Dim oObjectProperty As MR_DiffObjectProperty

    On Error GoTo err_

    Set oCode = New MR_Code

    For Each oItem In ioDiff.ObjectProperties
        Set oObjectProperty = oItem
        If oObjectProperty.ObjectName <> "Selection" Then
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

Function GetCodeForInitiallySelectedObjects2(ioDiff As MR_Diff) As MR_Code

    Dim oSelection As iSelection
    Dim oCode As MR_Code
    Dim oShape As iShape
    Dim oDiff As MR_Diff
    Dim oCollDiffs As Collection
    Dim sObjectPrefix As String
    Dim oScalarProperty As MR_DiffScalarProperty
    Dim oObjectProperty As MR_DiffObjectProperty
    Dim oMethodCall As MR_DiffMethodCall

    On Error GoTo err_

    Set oSelection = ioDiff.StartObject
    Set oCollDiffs = New Collection
    Set oCode = New MR_Code

    Select Case oSelection.Type_
        Case ppSelectionNone

        Case ppSelectionShapes
            For Each oShape In oSelection.shapeRange.Items
                ' Shouldn't it exist always?
                If ExistsInCollection(goDiffPtrs, CStr(ObjPtr(oShape))) Then
                    Set oDiff = goDiffPtrs(CStr(ObjPtr(oShape)))
                    For Each oObjectProperty In oDiff.ObjectProperties
                        ' TODO
                        If Not ExistsInCollection(oCollDiffs, oObjectProperty.ObjectName) Then
                            Call oCollDiffs.Add(oDiff, oObjectProperty.ObjectName)
                        End If
                    Next
                End If
            Next

            For Each oDiff In oCollDiffs
                Call oCode.AddCode(GetCodeForAnyObject(oDiff))
            Next

            Call oCode.Wrap(".ShapeRange")

        Case ppSelectionSlides
            ' TODO

        Case ppSelectionText
            ' TODO
    End Select

    Set GetCodeForInitiallySelectedObjects2 = oCode

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function GetCodeForAnyObject(ioDiff As MR_Diff, Optional ObjectName As String = "") As MR_Code

    Dim oCode As MR_Code
    Dim oAddedObject As MR_DiffAddedObject
    Dim oRemovedObject As MR_DiffRemovedObject
    Dim oScalarProperty As MR_DiffScalarProperty
    Dim oObjectProperty As MR_DiffObjectProperty
    Dim oMethodCall As MR_DiffMethodCall

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
        Call oCode.AddCode(GetCodeForAnyObject(oObjectProperty.Diff, sObjectPrefix & oObjectProperty.ObjectName))
    Next

    For Each oItem In ioDiff.MethodCalls
        Set oMethodCall = oItem
        Call oCode.Add("Call " & sObjectPrefix & oMethodCall.ProcName & "(" & oMethodCall.Arguments & ")")
    Next

    Call oCode.Wrap(ObjectName)

    Set GetCodeForAnyObject = oCode

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
    Dim oAddedObject As MR_DiffAddedObject
    Dim oRemovedObject As MR_DiffRemovedObject
    Dim oScalarProperty As MR_DiffScalarProperty
    Dim oObjectProperty As MR_DiffObjectProperty
    Dim oMethodCall As MR_DiffMethodCall

    On Error GoTo err_

    Set oCode = New MR_Code

    sObjectPrefix = GetObjectPrefix(ioDiff)

    For Each oItem In ioDiff.AddedObjects
        Set oAddedObject = oItem
        If Not ExistsInCollection(goCollObjectsWithCodeGenerated, CStr(ObjPtr(oAddedObject.MRObject))) Then
            Call goCollObjectsWithCodeGenerated.Add(oAddedObject, CStr(ObjPtr(oAddedObject.MRObject)))
            Call oCode.AddCode(oAddedObject.MRObject.create())
        End If
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
    Dim oAddedObject As MR_DiffAddedObject
    Dim oObjectProperty As MR_DiffObjectProperty

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

Function GetCodeForChangedSelection(ioDiffSelection As MR_Diff, ioStartSelection As iSelection, ioStopSelection As iSelection) As MR_Code

    Dim oCode As MR_Code
    Dim bNoneObjectIsSelected As Boolean
    Dim oAddedObjects As Collection
    Dim oCollSelectedShapes As Collection
    Dim oCollUnselectedShapes As Collection
    Dim lSlideIndex As Long
    Dim oShape As iShape
    Dim oSlide As iSlide
    Dim oPresentation As iPresentation
    Dim oPptSelectedSlide As Slide
    Dim oPptShape As Shape
    Dim lShapeIndex As Long
    Dim bReplace As Boolean

    On Error GoTo err_

    Set oCode = New MR_Code

    If ioStartSelection.Type_ <> ppSelectionNone _
        Or ioStopSelection.Type_ <> ppSelectionNone Then

        ' Generate "call activewindow.selection.unselect" if :
        '   - all selected objects were unselected
        '   - or if at least one object was unselected (followed by Select of currently-selected objects)
        ' The only case where Unselect is not generated is that :
        '   - either there was no change in selection
        '   - or only one or more elements were added to the selection
        If ioStopSelection.Type_ = ppSelectionNone Then
            Call oCode.Add("Call ActiveWindow.Selection.Unselect")
        Else
            ' Generate Select on newly selected objects except if they were just added
            '   (because in that case "Select" was added by GetCodeForAddedObjects)
            Set oAddedObjects = GetAddedObjects(ioDiffSelection)
            Select Case ioStopSelection.Type_
                Case ppSelectionShapes

                    Set oCollSelectedShapes = New Collection
                    Set oCollUnselectedShapes = New Collection
                    bObjectIsAdded = False
                    For i = 1 To ioStopSelection.shapeRange.Items.Count
                        Set oShape = ioStopSelection.shapeRange.Items(i)
                        If IsObjectNewlySelected(GetPptObject(goStopSnapshot, oShape)) Then
                            Call oCollSelectedShapes.Add(oShape)
                            ' Was the shape selected because it was added?
                            For j = 1 To oAddedObjects.Count
                                If oShape Is oAddedObjects(j) Then
                                    bObjectIsAdded = True
                                    Exit For
                                End If
                            Next
                        End If
                    Next
                    If ioStartSelection.Type_ = ppSelectionShapes Then
                        For i = 1 To ioStartSelection.shapeRange.Items.Count
                            Set oShape = ioStartSelection.shapeRange.Items(i)
                            If IsObjectNewlyUnselected(GetPptObject(goStartSnapshot, oShape)) Then
                                Call oCollUnselectedShapes.Add(oShape)
                            End If
                        Next
                    End If

                    ' Determine the Slides item index of the slide containing any of the selected shapes.
                    lSlideIndex = 0
                    Set oPptSelectedSlide = ioStopSelection.shapeRange.Items(1).Parent
                    For j = 1 To ActivePresentation.Slides.Count
                        If oPptSelectedSlide Is ActivePresentation.Slides(j) Then
                            lSlideIndex = j
                            Exit For
                        End If
                    Next

                    If bObjectIsAdded Then
                        ' Do nothing (Select was previously done during Call ...Slides.Item(x).AddShape(...).Select)
                    Else
                        If ioStartSelection.Type_ = ppSelectionShapes Then
                            If oCollUnselectedShapes.Count > 0 Then
                                bReplace = True
                                For Each oShape In oCollSelectedShapes
                                    Call AddCodeToSelectShape(oCode, lSlideIndex, oPptSelectedSlide, oShape, ibReplace:=bReplace)
                                    bReplace = False
                                Next
                            Else
                                For Each oShape In oCollSelectedShapes
                                    Call AddCodeToSelectShape(oCode, lSlideIndex, oPptSelectedSlide, oShape, ibReplace:=False)
                                Next
                            End If
                        Else
                            bReplace = True
                            For Each oShape In oCollSelectedShapes
                                Call AddCodeToSelectShape(oCode, lSlideIndex, oPptSelectedSlide, oShape, ibReplace:=bReplace)
                                bReplace = False
                            Next
                        End If
                    End If

                Case ppSelectionSlides
                    ' TODO
                Case ppSelectionText
                    ' TODO
            End Select
        End If
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

Sub AddCodeToSelectShape(ioCode As MR_Code, ilSlideIndex As Long, ioPptSelectedSlide As Slide, ioShape As iShape, ibReplace As Boolean)

    Dim lShapeIndex As Long
    Dim oPptShape As Shape
    Dim j As Long

    ' Get Shapes item index
    lShapeIndex = 0
    Set oPptShape = GetPptObject(goStopSnapshot, ioShape)
    For j = 1 To ioPptSelectedSlide.Shapes.Count
        If oPptShape Is ioPptSelectedSlide.Shapes(j) Then
            lShapeIndex = j
        End If
    Next

    Call ioCode.Add("Call ActivePresentation.Slides(" & CStr(ilSlideIndex) & ").Shapes(" & CStr(lShapeIndex) & ").Select(Replace:=" & BooleanToVBA(ibReplace) & ")")

End Sub

Function GetCodeForUnselectedObjects(isObjectName As String, ioDiff As MR_Diff) As MR_Code

    Dim oCode As MR_Code
    Dim sObjectPrefix As String
    Dim oScalarProperty As MR_DiffScalarProperty
    Dim oObjectProperty As MR_DiffObjectProperty
    Dim oMethodCall As MR_DiffMethodCall

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
        Select Case oObjectProperty.ObjectName
            Case "iShape"
                bIsObjectSelected = (IsObjectPartOfSelection(GetPptObject(goStopSnapshot, oObjectProperty.Diff.StopObject), goStopSnapshot))
        End Select
        If oObjectProperty.ObjectName <> "Selection" And bIsObjectSelected = False Then
            Call oCode.AddCode(GetCodeForUnselectedObjects(sObjectPrefix & oObjectProperty.ObjectName, oObjectProperty.Diff))
        Else
        End If
    Next

    For Each oItem In ioDiff.MethodCalls
        Set oMethodCall = oItem
        Call oCode.Add("Call " & sObjectPrefix & oMethodCall.ProcName & "(" & oMethodCall.Arguments & ")")
    Next

    Call oCode.Wrap(isObjectName)

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

    If TypeName(ioDiff.StopObject) = "iApplication" Then
        GetObjectPrefix = ""
    Else
        GetObjectPrefix = "."
    End If

End Function

Function OBSOLETE_GetCodeForAllChangedElements(ioDiff As MR_Diff, ObjectName As String) As MR_Code

    Dim oCode As MR_Code
    Dim oAddedObject As MR_DiffAddedObject
    Dim oScalarProperty As MR_DiffScalarProperty
    Dim oObjectProperty As MR_DiffObjectProperty
    Dim oMethodCall As MR_DiffMethodCall

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
        Call oCode.AddCode(OBSOLETE_GetCodeForAllChangedElements(oObjectProperty.Diff, sObjectPrefix & oObjectProperty.ObjectName))
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
    'Call oDiff.AddDiff("ActiveWindow.Selection", goStopSnapshot.iSelection.MR_Compare(goStartSnapshot.iSelection))
    'Call oDiff.AddDiff("ActivePresentation", goStopSnapshot.iPresentation.MR_Compare(goStartSnapshot.iPresentation))
    '' In the future, propose to record actions on several presentations
    'Call oDiff.AddDiff("Application", goStopSnapshot.iPresentation.MR_Compare(goStartSnapshot.iApplication))

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
    strCode = strCode & GetCodeForUnselectedObjects("Application", ioDiff).ConvertToString()

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

