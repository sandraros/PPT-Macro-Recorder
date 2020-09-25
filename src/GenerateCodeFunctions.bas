Attribute VB_Name = "GenerateCodeFunctions"
Function GenerateCode(ioDiff As UDiff, ioStartSelection As iSelection, ioStopSelection As iSelection, ioDiffSelection As UDiff) As String

    'Dim oCode As cCode
    Dim strCode As String

    On Error GoTo err_

    ' If one object is currently selected, prefer to generate:
    '     Application.ActiveWindow.Selection.ShapeRange(1).Fill.BackColor = RGB(17, 15, 150)
    ' rather than
    '     Application.ActivePresentation.Shapes(27).Fill.BackColor = RGB(17, 15, 150)
    'Set oCode = New cCode
    'Call oCode.AddCode(CallUnselectAndSelectIfSelectionChanged())
    'Call oCode.AddCode(stopSnapShot.iSelection.Compare("ActiveWindow.Selection", startSnapShot.iSelection))
    'Call oCode.AddCode(stopSnapShot.iPresentation.Compare("ActivePresentation", startSnapShot.iPresentation))
    '' In the future, propose to record actions on several presentations
    'Call oCode.AddCode(stopSnapShot.iPresentation.compare("Application", startSnapShot.iApplication))

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
    If ioStopSelection.iType <> ppSelectionNone Then
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

Function GetCodeForAddedObjects(ioDiff As UDiff, ObjectName As String) As cCode

    Dim oCode As cCode
    Dim oAddedObject As UDiffAddedObject
    Dim oRemovedObject As UDiffRemovedObject
    Dim oScalarProperty As UDiffScalarProperty
    Dim oObjectProperty As UDiffObjectProperty
    Dim oMethodCall As UDiffMethodCall

    On Error GoTo err_

    Set oCode = New cCode

    For Each oItem In ioDiff.AddedObjects
        Set oAddedObject = oItem
        Call oCode.AddCode(oAddedObject.Object.create())
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

Function GetCodeToSelectObjects(ioStartSelection As iSelection, ioStopSelection As iSelection) As cCode

    Dim bObjectIsUnselected As Boolean
    Dim oShape As iShape
    Dim oPresentation As iPresentation
    Dim oSlide As iSlide
    Dim lSlideId As Long
    Dim oSelectedSlide As Slide

    On Error GoTo err_

    Set oCode = New cCode

    ' Do "call activewindow.selection.unselect" if at least one selected object became unselected
    bObjectIsUnselected = False
    If ioStopSelection.iType = ppSelectionNone _
                And ioStartSelection.iType <> ppSelectionNone Then
        bObjectIsUnselected = True
    Else
        Select Case ioStartSelection.iType
            Case ppSelectionShapes
                For Each oShape In ioStartSelection.shapeRange.Items
                    If Not IsObjectPartOfSelection(GetPptObject(startSnapShot, oShape), stopSnapShot) Then
                        bObjectIsUnselected = True
                        Exit For
                    End If
                Next
            Case ppSelectionSlides
                For Each oSlide In ioStartSelection.SlideRange.Items
                    If Not IsObjectPartOfSelection(GetPptObject(startSnapShot, oSlide), stopSnapShot) Then
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
    Select Case ioStopSelection.iType
        Case ppSelectionShapes
            Set oSelectedSlide = ioStopSelection.shapeRange.Items(1).Parent
            For i = 1 To ActivePresentation.Slides.Count
                If oSelectedSlide Is ActivePresentation.Slides(i) Then
                    lSlideId = i
                    Exit For
                End If
            Next
            For i = 1 To ioStopSelection.shapeRange.Items.Count 'oSelectedSlide.Shapes.Count
                Set oShape = ioStopSelection.shapeRange.Items(i) 'GetMrsObject(oSelectedSlide.Shapes(i))
                If IsObjectNewlySelected(GetPptObject(stopSnapShot, oShape)) Then
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

Function GetCodeForAllChangedElements(ioDiff As UDiff, ObjectName As String) As cCode

    Dim oCode As cCode
    Dim oScalarProperty As UDiffScalarProperty
    Dim oObjectProperty As UDiffObjectProperty
    Dim oMethodCall As UDiffMethodCall

    On Error GoTo err_

    Set oCode = New cCode

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

Function GetCodeForUnselectedObjects(ioDiff As UDiff, ObjectName As String, ioStopSelection As iSelection) As cCode

    Dim oCode As cCode
    Dim oScalarProperty As UDiffScalarProperty
    Dim oObjectProperty As UDiffObjectProperty
    Dim oMethodCall As UDiffMethodCall

    On Error GoTo err_

    Set oCode = New cCode

    For Each oItem In ioDiff.ScalarProperties
        Set oScalarProperty = oItem
        Call oCode.Add("." & oScalarProperty.Name & " = " & oScalarProperty.Value)
    Next

    For Each oItem In ioDiff.ObjectProperties
        Set oObjectProperty = oItem
        bIsObjectSelected = False
        Select Case TypeName(oObjectProperty.Diff.StopObject)
        Case "iShape"
        bIsObjectSelected = (IsObjectPartOfSelection(GetPptObject(stopSnapShot, oObjectProperty.Diff.StopObject), stopSnapShot))
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
