VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MR_Snapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public iSelection As New iSelection
Public iPresentation As New iPresentation
'Public iApplication As New iApplication
Public allObjects As New Collection
Public allObjectClasses As New Collection
Public MRObjPtrs As Collection ' Macro Recorder Snapshot objects (pseudo PowerPoint objects)
Public PptObjPtrs As Collection ' Native PowerPoint objects
