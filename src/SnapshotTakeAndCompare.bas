Attribute VB_Name = "SnapshotTakeAndCompare"
Function TakeSnapshot() As MR_Snapshot

    On Error GoTo err_

    Set Snapshot = New MR_Snapshot

    'TODO for now, try to simplify = only changes in active presentation
    Set Snapshot.iSelection = New_iSelection(Application.ActiveWindow.Selection)
    Set Snapshot.iPresentation = New_iPresentation(Application.ActivePresentation)
    'Set snapshot.iApplication = New_iApplication(Application)

    Set TakeSnapshot = Snapshot

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Function CompareSnapshots() As MR_Diff '(oDiff As MR_Diff, oDiffSelection As MR_Diff)

    On Error GoTo err_

    ' Build collections MRObjPtrs and PptObjPtrs of all snapshots.
    Call BuildObjectIndexes

    'Set goStack = New cStack ' (still needed?)
    Set AllObjectsCompared = New Collection ' To not MR_Compare one object twice (still needed?)

    'firstSelectedObjectIsProcessed = False ' (still needed?)

    Set CompareSnapshots = goStopSnapshot.iPresentation.MR_Compare("ActivePresentation", goStartSnapshot.iPresentation)
    'Set oDiffSelection = goStopSnapshot.iSelection.MR_Compare("ActiveWindow.Selection", goStartSnapshot.iSelection)

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Sub GetCodeForAddedObjects()

    Dim oStartSlide As iSlide
    Dim oStopShape As iShape

    On Error GoTo err_

    With goStopSnapshot.iPresentation.Slides
        For i = 1 To .Count
            Set oStopSlide = .Items(i)
            Set oStartSlide = GetObjectIngoStartSnapshot(oStopSlide)
            If oStartSlide Is Nothing Then
                Set oStartSlide = goStartSnapshot.iPresentation.Slides.AddSlide(oStopSlide.SlideIndex, oStopSlide.CustomLayout)
            End If
            With .Shapes
                For j = 1 To .Count
                    Set oStopShape = .Items(j)
                    If GetObjectIngoStartSnapshot(oStopShape) Is Nothing Then
                        Select Case oStopShape.Type_
                            Case msoAutoShape
                                Call oStartSlide.Shapes.AddShape(Type_:=oStopSlide.AutoShapeType, _
                                            Left:=oStopSlide.Left, _
                                            Top:=oStopSlide.Top, _
                                            Width:=oStopSlide.Width, _
                                            Height:=oStopSlide.Height).Select
                            Case Else
                                err.Raise 9999, , "TODO iShape.Create"
                        End Select
                    End If
                Next
            End With
        Next
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

