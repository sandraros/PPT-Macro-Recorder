VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MR_Snapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public iSelection As New iSelection
'Public iPresentation As New iPresentation
Public iApplication As New iApplication
Public allObjects As New Collection ' Built by function AddObject called by Factory methods of all MR objects
Public allObjectClasses As New Collection ' Just for debug purpose - Still needed?
Public MRObjPtrs As Collection ' Macro Recorder Snapshot objects (pseudo PowerPoint objects)
Public PptObjPtrs As Collection ' Native PowerPoint objects

Sub BuildObjectIndexes()
    ' Called by TakeSnapshot
    ' Can't it be done by AddObject, at same time allObjects is built?

    Dim oObjectPair As cObjectPair

    On Error GoTo err_

    Set MRObjPtrs = New Collection
    Set PptObjPtrs = New Collection
    For i = 1 To allObjects.Count
        Set oObjectPair = allObjects.Item(i)
        Call MRObjPtrs.Add(oObjectPair, CStr(ObjPtr(oObjectPair.MRObject)))
        Call PptObjPtrs.Add(oObjectPair, CStr(ObjPtr(oObjectPair.PptObject)))
    Next

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub
