Attribute VB_Name = "SnapshotTakeAndCompare"
Function TakeSnapshot() As cSnapShot

    On Error GoTo err_

    Set Snapshot = New cSnapShot

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

Sub CompareSnapshots(oDiff As UDiff, oDiffSelection As UDiff)

    On Error GoTo err_

    ' Build collections MrsObjPtrs and PptObjPtrs of all snapshots.
    Call BuildObjectIndexes

    Set goStack = New cStack ' (still needed?)
    Set AllObjectsCompared = New Collection ' To not compare one object twice (still needed?)
    firstSelectedObjectIsProcessed = False ' (still needed?)

    Set oDiff = stopSnapShot.iPresentation.Compare("ActivePresentation", startSnapShot.iPresentation)
    Set oDiffSelection = stopSnapShot.iSelection.Compare("ActiveWindow.Selection", startSnapShot.iSelection)

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub
