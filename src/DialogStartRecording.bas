Attribute VB_Name = "DialogStartRecording"
Function DialogStartRecorder() As enumAction

    Dim oPresentation As Presentation

    On Error GoTo err_

    UserForm1.macroName = DetermineMacroName(GetPresentationName(ActivePresentation))
    For Each oPresentation In Application.Presentations
        Call UserForm1.macroPresentation.AddItem(GetPresentationName(oPresentation))
    Next
    UserForm1.macroPresentation.Value = GetPresentationName(Application.ActivePresentation)
    UserForm1.macroPresentation.Style = fmStyleDropDownList

    Call UserForm1.Show

    ' Save form fields so that they can be used after Unload
    action = UserForm1.action
    ' Save form fields to global variables
    macroName = UserForm1.macroName.Value
    macroPresentation = UserForm1.macroPresentation.Value
    macroDescription = f
    arr = Split(UserForm1.macroDescription.Value, vbCrLf)
    For i = LBound(arr) To UBound(arr)
        macroDescription = macroDescription & "' " & arr(i) & Chr(10)
    Next

    Call Unload(UserForm1)

    DialogStartRecorder = action


    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function GetPresentationName(Presentation As Presentation) As String
        
    On Error GoTo err_

    If Presentation.Path <> "" Then
        GetPresentationName = Presentation.Name & " (in " & Presentation.Path & ")"
    Else
        GetPresentationName = Presentation.Name
    End If


    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function GetPresentation(PresentationName As String) As Presentation

    Dim Presentation As Presentation

    On Error GoTo err_

    For Each Presentation In Application.Presentations
        If GetPresentationName(Presentation) = PresentationName Then
            Exit For
        End If
    Next

    Set GetPresentation = Presentation


    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function DetermineMacroName(PresentationName As String) As String

    Dim objVBProject As VBProject
    Dim objVBComponent As VBComponent
    Dim objPresentation As Presentation
    Dim intMacroNumber As Integer
    Dim strProcName As String
    Dim enumProcKind As vbext_ProcKind

    On Error GoTo err_

    Set objPresentation = GetPresentation(PresentationName)

    On Error Resume Next
    Set objVBComponent = objPresentation.VBProject.VBComponents("NewMacros")
    errnum = err.number
    On Error GoTo 0

    If errnum <> 0 Then
        DetermineMacroName = "Macro1"
        Exit Function
    End If

    intLine = objVBComponent.CodeModule.CountOfDeclarationLines + 1
    While intLine <= objVBComponent.CodeModule.CountOfLines
        strProcName = objVBComponent.CodeModule.ProcOfLine(intLine, enumProcKind)
        If Left(strProcName, 5) = "Macro" And Mid(strProcName, 6) = CStr(Val(Mid(strProcName, 6))) Then
            If CInt(Mid(strProcName, 6)) > intMacroNumber Then
                intMacroNumber = CInt(Mid(strProcName, 6))
            End If
        End If
        intLine = intLine + objVBComponent.CodeModule.ProcCountLines(strProcName, enumProcKind)
    Wend

    DetermineMacroName = "Macro" & CStr(intMacroNumber + 1)


    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function


