Attribute VB_Name = "Main"
Public snapshots As New Collection
Global snapshot As cSnapShot
Global AllObjectsCompared As Collection
Global macroPresentation As String
Global macroName As String
Global macroDescription As String

Sub test()
    Dim Fill As Office.FillFormat
    Set Fill = ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Font.Fill
End Sub

Sub start_stop_recording()

    Dim code As String
    Dim code1 As String
    Dim code2 As String
    Dim first As cSnapShot
    Dim last As cSnapShot

    If snapshots.Count = 0 Then
        
        If DialogStartRecorder() = enumAction.cancel Then
            Exit Sub
        End If
    End If

    Call take_snapshot

    If snapshots.Count = 2 Then

        ' Build collections MyObjPtrs and PptObjPtrs of all snapshots.
        Call BuildObjectIndexes

        Set AllObjectsCompared = New Collection

        Set first = snapshots.Item(snapshots.Count - 1)
        Set last = snapshots.Item(snapshots.Count)

        code = ""
        Call addCode(code, "Sub " & macroName & "()")
        Call addCode(code, "'")
        Call addCode(code, "' " & macroName & " Macro")
        Call addCode(code, macroDescription)
        Call addCode(code, "'")

        code2 = ""
        Call addCode(code2, last.iSelection.compare("Application.ActiveWindow.Selection", first.iSelection))
        Call addCode(code2, last.iPresentation.compare("Application.ActivePresentation", first.iPresentation))
        Call addCode(code, Utility.indentCode(code2, 4))

        Call addCode(code, "End Sub")

        Call ExportCode(code)

        ' Clear the collection (can we trust the garbage collection?)
        Call snapshots.Remove(1)
        Call snapshots.Remove(1)
        Set snapshots = New Collection

    End If

End Sub

Function DialogStartRecorder() As enumAction

    Dim VBProject As VBProject
    Dim VBComponent As VBComponent
    Dim Presentation As Presentation

    MacroNumber = 0
    i = 1
    Do While True
        MacroAlreadyExists = False
        For Each VBProject In Application.vbe.VBProjects
            For Each VBComponent In VBProject.VBComponents
                If VBComponent.Name = "NewMacros" Then
                    Errnum = 0
                End If
                On Error Resume Next
                countLines = VBComponent.CodeModule.ProcCountLines("Macro" & CStr(i), vbext_pk_Proc)
                Errnum = err.number
                On Error GoTo 0
                If Errnum = 0 Then
                    MacroAlreadyExists = True
                    Exit For
                End If
            Next
            If Not MacroAlreadyExists Then
                Exit For
            End If
        Next
        If Not MacroAlreadyExists Then
            MacroNumber = i
            Exit Do
        End If
        i = i + 1
    Loop

    UserForm1.macroName = "Macro" & CStr(MacroNumber)
    For Each Presentation In Application.Presentations
        Call UserForm1.macroPresentation.AddItem(GetPresentationName(Presentation))
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

End Function

Function GetPresentationName(Presentation As Presentation) As String
        
    If Presentation.Path <> "" Then
        GetPresentationName = Presentation.Name & " (in " & Presentation.Path & ")"
    Else
        GetPresentationName = Presentation.Name
    End If

End Function

Function GetPresentation(PresentationName As String) As Presentation

    Dim Presentation As Presentation

    For Each Presentation In Application.Presentations
        If GetPresentationName(Presentation) = PresentationName Then
            Exit For
        End If
    Next

    Set GetPresentation = Presentation

End Function

Function DetermineMacroName(PresentationName As String) As String

    Dim VBProject As VBProject
    Dim VBComponent As VBComponent
    Dim Presentation As Presentation

    Set Presentation = GetPresentation(PresentationName)

    MacroNumber = 0
    i = 1
    Do While True
        MacroAlreadyExists = False
        For Each VBComponent In Presentation.VBProject.VBComponents
            If VBComponent.Name = "NewMacros" Then
                Errnum = 0
            End If
            On Error Resume Next
            countLines = VBComponent.CodeModule.ProcCountLines("Macro" & CStr(i), vbext_pk_Proc)
            Errnum = err.number
            On Error GoTo 0
            If Errnum = 0 Then
                MacroAlreadyExists = True
                Exit For
            End If
        Next
        If Not MacroAlreadyExists Then
            MacroNumber = i
            Exit Do
        End If
        i = i + 1
    Loop

    DetermineMacroName = "Macro" & CStr(MacroNumber)

End Function

Sub ExportCode(code As String)

    Dim oVBComps As VBComponents
    Dim oVBComp As VBComponent
    Dim Presentation As Presentation

    Set oVBComps = GetPresentation(macroPresentation).VBProject.VBComponents

    On Error Resume Next
    Set oVBComp = oVBComps("NewMacros")
    Errnum = err.number
    On Error GoTo 0
    If Errnum <> 0 Then
        Set oVBComp = oVBComps.Add(vbext_ct_StdModule)
        oVBComp.Name = "NewMacros"
    End If
    'oVBComp.CodeModule.Lines = oVBComp.CodeModule.Lines) & code
    Call oVBComp.CodeModule.InsertLines(oVBComp.CodeModule.CountOfLines + 1, code)

End Sub

Sub take_snapshot()

    Set snapshot = New cSnapShot

    'TODO for now, try to simplify = only changes in active presentation
    Set snapshot.iSelection = New_iSelection(Application.ActiveWindow.Selection)
    Set snapshot.iPresentation = New_iPresentation(Application.ActivePresentation)
    'Set snapshot.iApplication = New_iApplication(Application)

    Call snapshots.Add(snapshot)

End Sub

Public Function compare_snapshots(first As cSnapShot, last As cSnapShot) As String


End Function

