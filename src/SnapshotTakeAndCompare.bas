Attribute VB_Name = "SnapshotTakeAndCompare"
Function TakeSnapshot() As MR_Snapshot
    ' Called both by StartRecording and by TakeSnapshotCompareAndGenerateCode

    On Error GoTo err_

    Set Snapshot = New MR_Snapshot

    'TODO for now, try to simplify = only changes in active presentation
    'Set Snapshot.iSelection = New_iSelection(Application.ActiveWindow.Selection)
    'Set Snapshot.iPresentation = New_iPresentation(Application.ActivePresentation)
    Set Snapshot.iApplication = New_iApplication(Application)

    Call Snapshot.BuildObjectIndexes

    Set TakeSnapshot = Snapshot

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Sub TakeSnapshotCompareAndGenerateCode()
    ' Called both by ProcessEvent and StopRecording

    Dim oDiff As MR_Diff
    Dim oStartSelection As iSelection
    Dim oStopSelection As iSelection

    Set goStopSnapshot = TakeSnapshot()

    Set oDiff = CompareSnapshots()

    ' EXAMPLES:
    '
    '   After adding the first shape => Code for added elements (notice .Select after AddShape(...)!!!):
    '     Call ActivePresentation.Slides.Item(1).Shapes.AddShape(Type:=msoShapeRectangle, Left:=83.33331, Top:=50.66669, Width:=95.33339, Height:=75.33331).Select
    '
    '   EITHER After changing the color of the shape + Stop Macro
    '     ActiveWindow.Selection.shapeRange.Fill.ForeColor.ObjectThemeColor = Office.msoThemeColorAccent4
    '
    '   OR after adding the second shape => Code for changed elements previously selected + Code for other changed elements + Code for added elements
    '     With ActiveWindow.Selection.shapeRange
    '       .Fill.ForeColor.ObjectThemeColor = Office.msoThemeColorAccent4
    '       .Line.ForeColor.ObjectThemeColor = Office.msoThemeColorAccent4
    '     End With
    '     Call ActivePresentation.Slides.Item(1).Shapes.AddShape(Type:=msoShapeOval, Left:=206.6667, Top:=50.66669, Width:=80, Height:=75.33331).Select

    'Code for changed elements previously selected + Code for other changed elements + Code for added elements + Code for object Selection changed
    Set oStartSelection = goStartSnapshot.iApplication.ActiveWindow.Selection
    Set oStopSelection = goStopSnapshot.iApplication.ActiveWindow.Selection

    Call goCode.AddCode(GetCodeForInitiallySelectedObjects(oDiff))
    Call goCode.AddCode(GetCodeForUnselectedObjects(oDiff))
    Call goCode.AddCode(GetCodeForAddedObjects(oDiff))
    Call goCode.AddCode(GetCodeForChangedSelection(oDiff, oStartSelection, oStopSelection))

    Set goStartSnapshot = goStopSnapshot
    Set goStopSnapshot = Nothing

End Sub

Function CompareSnapshots() As MR_Diff

    On Error GoTo err_

    ' Build collections MRObjPtrs and PptObjPtrs of all snapshots.

    Set AllObjectsCompared = New Collection ' To not compare one object twice -- still needed?

    Set CompareSnapshots = goStopSnapshot.iApplication.MR_Compare("Application", goStartSnapshot.iApplication)

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Sub OBSOLETE_GetCodeForAddedObjects()

    Dim oStartSlide As iSlide
    Dim oStopShape As iShape

    On Error GoTo err_

    'With goStopSnapshot.iPresentation.Slides
    '    For i = 1 To .Count
    '        Set oStopSlide = .Items(i)
    '        Set oStartSlide = GetObjectInStartSnapshot(oStopSlide)
    '        If oStartSlide Is Nothing Then
    '            Set oStartSlide = goStartSnapshot.iPresentation.Slides.AddSlide(oStopSlide.SlideIndex, oStopSlide.CustomLayout)
    '        End If
    '        With .Shapes
    '            For j = 1 To .Count
    '                Set oStopShape = .Items(j)
    '                If GetObjectInStartSnapshot(oStopShape) Is Nothing Then
    '                    Select Case oStopShape.Type_
    '                        Case msoAutoShape
    '                            Call oStartSlide.Shapes.AddShape(Type_:=oStopSlide.AutoShapeType, _
    '                                        Left:=oStopSlide.Left, _
    '                                        Top:=oStopSlide.Top, _
    '                                        Width:=oStopSlide.Width, _
    '                                        Height:=oStopSlide.Height).Select
    '                        Case Else
    '                            err.Raise 9999, , "TODO iShape.Create"
    '                    End Select
    '                End If
    '            Next
    '        End With
    '    Next
    'End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

